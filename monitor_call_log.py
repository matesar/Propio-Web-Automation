#!/usr/bin/env python3
"""
Monitorea una tabla de historial de llamadas en una página ya abierta y logueada,
y guarda nuevas llamadas en Excel con cálculo de importe.

Backend principal: Playwright conectado a una sesión existente vía CDP.
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from playwright.sync_api import Browser, Error as PlaywrightError, Page, sync_playwright

# ---------------------------- Configuración global ---------------------------- #
RATE_PER_MINUTE = Decimal("0.11")
TABLE_ROW_SELECTOR = "mat-table.call-list-table mat-row"
EXCEL_HEADERS = [
    "Customer ID",
    "Call Date",
    "Call Start",
    "Duration (Minutes)",
    "Amount USD",
    "Detected At",
    "Unique Key",
]
SUMMARY_SHEET_NAME = "Daily Summary"
CALLS_SHEET_NAME = "Calls"
SUMMARY_HEADERS = ["Call Date", "Total Calls", "Total Minutes", "Total Amount USD"]


@dataclass(frozen=True)
class CallRecord:
    """Representa una llamada normalizada lista para persistir."""

    customer_id: str
    call_date: str
    call_start: str
    duration_minutes: int

    @property
    def amount_usd(self) -> Decimal:
        return (Decimal(self.duration_minutes) * RATE_PER_MINUTE).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )

    @property
    def unique_key(self) -> str:
        return build_unique_key(
            self.customer_id, self.call_date, self.call_start, self.duration_minutes
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Monitorea nuevas llamadas en una tabla HTML y guarda en Excel."
    )
    parser.add_argument(
        "--cdp-url",
        default="http://127.0.0.1:9222",
        help="URL CDP del navegador abierto con remote debugging (default: http://127.0.0.1:9222)",
    )
    parser.add_argument(
        "--url-contains",
        default="",
        help="Texto para encontrar la pestaña correcta por URL (ej: /call-history). Si se omite, usa la pestaña activa encontrada.",
    )
    parser.add_argument(
        "--excel",
        default="call_log.xlsx",
        help="Ruta del archivo Excel de salida (default: call_log.xlsx)",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=180,
        help="Segundos entre cada polling (default: 180)",
    )
    return parser.parse_args()


# ------------------------- Conexión y selección de página ------------------------- #

def connect_browser_via_cdp(cdp_url: str) -> Tuple[object, Browser]:
    """Conecta Playwright a un navegador Chromium/Chrome/Edge ya abierto vía CDP."""
    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(cdp_url)
    except Exception:
        pw.stop()
        raise
    return pw, browser


def find_target_page(browser: Browser, url_contains: str) -> Page:
    """Busca una página abierta que coincida con url_contains o toma la primera disponible."""
    candidates: List[Page] = []
    for context in browser.contexts:
        candidates.extend(context.pages)

    if not candidates:
        raise RuntimeError(
            "No se encontraron pestañas abiertas en la sesión conectada. "
            "Abrí la página de historial primero."
        )

    if url_contains:
        for p in candidates:
            if url_contains.lower() in p.url.lower():
                return p
        raise RuntimeError(
            f"No se encontró ninguna pestaña cuya URL contenga: {url_contains!r}."
        )

    # Si no se especifica filtro, usar la primera pestaña no vacía
    for p in candidates:
        if p.url and p.url != "about:blank":
            return p

    return candidates[0]


# ---------------------------- Lectura de tabla HTML ---------------------------- #

def wait_table_ready(page: Page, timeout_ms: int = 15_000) -> None:
    """Espera al menos una fila visible para confirmar que la tabla cargó."""
    page.wait_for_selector(TABLE_ROW_SELECTOR, timeout=timeout_ms)


def safe_text(locator) -> str:
    txt = locator.inner_text(timeout=2000)
    return " ".join(txt.split()).strip()


def parse_duration_to_int(raw: str) -> Optional[int]:
    """Convierte duración textual a entero de minutos (esperado: '5')."""
    cleaned = raw.strip()
    if not cleaned:
        return None
    # mantener solo dígitos por robustez ante espacios o sufijos accidentales
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    if not digits:
        return None
    try:
        return int(digits)
    except ValueError:
        return None


def build_unique_key(customer_id: str, call_date: str, call_start: str, duration_minutes: int) -> str:
    return f"{customer_id}|{call_date}|{call_start}|{duration_minutes}"


def extract_record_from_row(row) -> Optional[CallRecord]:
    """
    Extrae datos de una fila mat-row.

    Estrategia de selectores:
    1) Preferir columnas por clase (más estable en Angular Material):
       - .cdk-column-customerId
       - .cdk-column-callDate
       - .cdk-column-callStart
       - .cdk-column-callDuration
    2) Fallback por índice de celdas si cambian clases pero no orden de columnas.
    """
    customer_id = ""
    call_date = ""
    call_start = ""
    duration_raw = ""

    # 1) Selectores de columna por clase
    col_map = {
        "customer_id": ".cdk-column-customerId, .mat-column-customerId",
        "call_date": ".cdk-column-callDate, .mat-column-callDate",
        "call_start": ".cdk-column-callStart, .mat-column-callStart",
        "duration": ".cdk-column-callDuration, .mat-column-callDuration",
    }

    if row.locator(col_map["customer_id"]).count() > 0:
        customer_id = safe_text(row.locator(col_map["customer_id"]).first)
    if row.locator(col_map["call_date"]).count() > 0:
        call_date = safe_text(row.locator(col_map["call_date"]).first)
    if row.locator(col_map["call_start"]).count() > 0:
        call_start = safe_text(row.locator(col_map["call_start"]).first)
    if row.locator(col_map["duration"]).count() > 0:
        duration_raw = safe_text(row.locator(col_map["duration"]).first)

    # 2) Fallback por índice (Customer ID, Call Date, Call Start, Duration=índice 3)
    if not (customer_id and call_date and call_start and duration_raw):
        cells = row.locator("mat-cell")
        if cells.count() >= 4:
            customer_id = customer_id or safe_text(cells.nth(0))
            call_date = call_date or safe_text(cells.nth(1))
            call_start = call_start or safe_text(cells.nth(2))
            duration_raw = duration_raw or safe_text(cells.nth(3))

    duration_minutes = parse_duration_to_int(duration_raw)
    if not customer_id or not call_date or not call_start or duration_minutes is None:
        return None

    return CallRecord(
        customer_id=customer_id,
        call_date=call_date,
        call_start=call_start,
        duration_minutes=duration_minutes,
    )


def read_visible_records(page: Page) -> List[CallRecord]:
    """Relee todas las filas visibles en cada ciclo (estrategia robusta)."""
    records: List[CallRecord] = []
    rows = page.locator(TABLE_ROW_SELECTOR)
    row_count = rows.count()

    for i in range(row_count):
        row = rows.nth(i)
        rec = extract_record_from_row(row)
        if rec:
            records.append(rec)
    return records


def read_visible_records_with_recovery(page: Page, retries: int = 1) -> List[CallRecord]:
    """
    Lee filas visibles con recuperación simple ante rerender/cambios transitorios del DOM.
    Reintenta relocalizando la tabla antes de propagar error.
    """
    last_error: Optional[Exception] = None
    for attempt in range(retries + 1):
        try:
            wait_table_ready(page, timeout_ms=10_000)
            return read_visible_records(page)
        except Exception as e:
            last_error = e
            if attempt < retries:
                print(f"[WARN] Error leyendo tabla. Reintentando relocalizar ({attempt + 1}/{retries})...")
                time.sleep(1)
    if last_error:
        raise last_error
    return []


# ------------------------------- Manejo de Excel ------------------------------- #

def apply_basic_sheet_formatting(
    ws,
    headers: List[str],
    amount_header: Optional[str] = None,
    widths: Optional[Dict[str, int]] = None,
) -> None:
    """Aplica formato básico: header bold, freeze, autofilter y anchos."""
    bold = Font(bold=True)
    for i, header in enumerate(headers, start=1):
        ws.cell(1, i).font = bold
        if widths and header in widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[header]

    ws.freeze_panes = "A2"
    max_col = len(headers)
    max_row = max(ws.max_row, 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    if amount_header:
        amount_col = get_column_index(ws, amount_header, -1)
        if amount_col > 0:
            for r in range(2, ws.max_row + 1):
                ws.cell(r, amount_col).number_format = "#,##0.00"


def ensure_sheet_headers(ws, headers: List[str]) -> None:
    if ws.max_row < 1:
        ws.append(headers)
        return
    current = [ws.cell(1, i + 1).value for i in range(len(headers))]
    if current != headers:
        if ws.max_row == 1 and all(v is None for v in current):
            for i, h in enumerate(headers, start=1):
                ws.cell(1, i, h)
        else:
            print(f"[WARN] Encabezados en hoja '{ws.title}' no coinciden exactamente con lo esperado.")


def ensure_workbook(path: Path):
    """Crea/abre Excel y asegura hojas Calls + Daily Summary."""
    if not path.exists():
        wb = Workbook()
        ws_calls = wb.active
        ws_calls.title = CALLS_SHEET_NAME
        ws_calls.append(EXCEL_HEADERS)
        ws_summary = wb.create_sheet(SUMMARY_SHEET_NAME)
        ws_summary.append(SUMMARY_HEADERS)
        apply_basic_sheet_formatting(
            ws_calls,
            EXCEL_HEADERS,
            amount_header="Amount USD",
            widths={
                "Customer ID": 18,
                "Call Date": 14,
                "Call Start": 12,
                "Duration (Minutes)": 18,
                "Amount USD": 14,
                "Detected At": 20,
                "Unique Key": 50,
            },
        )
        apply_basic_sheet_formatting(
            ws_summary,
            SUMMARY_HEADERS,
            amount_header="Total Amount USD",
            widths={
                "Call Date": 14,
                "Total Calls": 12,
                "Total Minutes": 14,
                "Total Amount USD": 18,
            },
        )
        wb.save(path)
        return wb, ws_calls, ws_summary

    wb = load_workbook(path)
    ws_calls = wb[CALLS_SHEET_NAME] if CALLS_SHEET_NAME in wb.sheetnames else wb.active
    if ws_calls.title != CALLS_SHEET_NAME:
        ws_calls.title = CALLS_SHEET_NAME
    ws_summary = (
        wb[SUMMARY_SHEET_NAME]
        if SUMMARY_SHEET_NAME in wb.sheetnames
        else wb.create_sheet(SUMMARY_SHEET_NAME)
    )

    ensure_sheet_headers(ws_calls, EXCEL_HEADERS)
    ensure_sheet_headers(ws_summary, SUMMARY_HEADERS)
    apply_basic_sheet_formatting(
        ws_calls,
        EXCEL_HEADERS,
        amount_header="Amount USD",
        widths={
            "Customer ID": 18,
            "Call Date": 14,
            "Call Start": 12,
            "Duration (Minutes)": 18,
            "Amount USD": 14,
            "Detected At": 20,
            "Unique Key": 50,
        },
    )
    apply_basic_sheet_formatting(
        ws_summary,
        SUMMARY_HEADERS,
        amount_header="Total Amount USD",
        widths={
            "Call Date": 14,
            "Total Calls": 12,
            "Total Minutes": 14,
            "Total Amount USD": 18,
        },
    )
    wb.save(path)
    return wb, ws_calls, ws_summary


def get_column_index(ws, header_name: str, fallback_index: int) -> int:
    """Obtiene índice 1-based de una columna por header; usa fallback si no la encuentra."""
    headers = [ws.cell(1, i + 1).value for i in range(ws.max_column)]
    for idx, header in enumerate(headers, start=1):
        if str(header).strip() == header_name:
            return idx
    return fallback_index


def load_existing_unique_keys(ws) -> Set[str]:
    unique_key_col = get_column_index(ws, "Unique Key", len(EXCEL_HEADERS))
    keys: Set[str] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        key = row[unique_key_col - 1] if len(row) >= unique_key_col else None
        if key:
            keys.add(str(key))
    return keys


def calculate_total_amount(ws) -> Decimal:
    amount_col = get_column_index(
        ws, "Amount USD", EXCEL_HEADERS.index("Amount USD") + 1
    )
    total = Decimal("0")
    for row in ws.iter_rows(min_row=2, values_only=True):
        amount = row[amount_col - 1] if len(row) >= amount_col else None
        if amount is None:
            continue
        try:
            total += Decimal(str(amount))
        except (InvalidOperation, ValueError):
            continue
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def append_new_records(
    path: Path, ws_calls, wb, new_records_with_detection: List[Tuple[CallRecord, str]]
) -> Decimal:
    """Agrega registros nuevos al Excel y devuelve nuevo total acumulado."""
    for rec, detected_at in new_records_with_detection:
        ws_calls.append(
            [
                rec.customer_id,
                rec.call_date,
                rec.call_start,
                rec.duration_minutes,
                rec.amount_usd,
                detected_at,
                rec.unique_key,
            ]
        )
    apply_basic_sheet_formatting(ws_calls, EXCEL_HEADERS, amount_header="Amount USD")

    wb.save(path)
    return calculate_total_amount(ws_calls)


def rebuild_daily_summary(ws_calls, ws_summary) -> Dict[str, Dict[str, Decimal]]:
    """Reconstruye la hoja Daily Summary desde Calls (fuente de verdad)."""
    call_date_col = get_column_index(ws_calls, "Call Date", 2)
    minutes_col = get_column_index(ws_calls, "Duration (Minutes)", 4)
    amount_col = get_column_index(ws_calls, "Amount USD", 5)

    daily: Dict[str, Dict[str, Decimal]] = {}
    for row in ws_calls.iter_rows(min_row=2, values_only=True):
        if len(row) < max(call_date_col, minutes_col, amount_col):
            continue
        call_date = str(row[call_date_col - 1]).strip()
        if not call_date:
            continue
        minutes_raw = row[minutes_col - 1]
        amount_raw = row[amount_col - 1]
        try:
            minutes = int(str(minutes_raw)) if minutes_raw is not None else 0
        except ValueError:
            minutes = 0
        try:
            amount = Decimal(str(amount_raw)) if amount_raw is not None else Decimal("0")
        except (InvalidOperation, ValueError):
            amount = Decimal("0")

        if call_date not in daily:
            daily[call_date] = {
                "total_calls": Decimal("0"),
                "total_minutes": Decimal("0"),
                "total_amount": Decimal("0"),
            }
        daily[call_date]["total_calls"] += Decimal("1")
        daily[call_date]["total_minutes"] += Decimal(minutes)
        daily[call_date]["total_amount"] += amount

    ws_summary.delete_rows(2, ws_summary.max_row)
    for call_date in sorted(daily.keys()):
        ws_summary.append(
            [
                call_date,
                int(daily[call_date]["total_calls"]),
                int(daily[call_date]["total_minutes"]),
                daily[call_date]["total_amount"].quantize(Decimal("0.01")),
            ]
        )
    apply_basic_sheet_formatting(
        ws_summary, SUMMARY_HEADERS, amount_header="Total Amount USD"
    )
    return daily


def get_today_summary(daily_summary: Dict[str, Dict[str, Decimal]]) -> Tuple[int, int, Decimal]:
    """Obtiene resumen del día local actual basado en claves de Call Date parseables."""
    today = date.today()
    total_calls = 0
    total_minutes = 0
    total_amount = Decimal("0")

    for call_date, values in daily_summary.items():
        parsed_date: Optional[date] = None
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                parsed_date = datetime.strptime(call_date[:10], fmt).date()
                break
            except ValueError:
                continue
        if parsed_date == today:
            total_calls += int(values["total_calls"])
            total_minutes += int(values["total_minutes"])
            total_amount += values["total_amount"]

    return total_calls, total_minutes, total_amount.quantize(Decimal("0.01"))


# ------------------------------- Lógica principal ------------------------------- #

def print_new_call(rec: CallRecord, detected_at: str, total_amount: Decimal) -> None:
    print("\n[NEW CALL] Nueva llamada detectada")
    print(f"  Customer ID        : {rec.customer_id}")
    print(f"  Call Date          : {rec.call_date}")
    print(f"  Call Start         : {rec.call_start}")
    print(f"  Duration (Minutes) : {rec.duration_minutes}")
    print(f"  Detected At        : {detected_at}")
    print(f"  Amount USD         : {rec.amount_usd}")
    print(f"  Total acumulado    : {total_amount}")


def print_initial_sync_call(rec: CallRecord, detected_at: str, total_amount: Decimal) -> None:
    print("\n[INITIAL SYNC] Llamada visible importada")
    print(f"  Customer ID        : {rec.customer_id}")
    print(f"  Call Date          : {rec.call_date}")
    print(f"  Call Start         : {rec.call_start}")
    print(f"  Duration (Minutes) : {rec.duration_minutes}")
    print(f"  Detected At        : {detected_at}")
    print(f"  Amount USD         : {rec.amount_usd}")
    print(f"  Total acumulado    : {total_amount}")


def monitor_loop(page: Page, excel_path: Path, interval_seconds: int) -> None:
    wb, ws_calls, ws_summary = ensure_workbook(excel_path)
    seen_keys = load_existing_unique_keys(ws_calls)
    persisted_count_before_sync = len(seen_keys)

    # Sincronización inicial: importar llamadas visibles no presentes en Excel
    visible_records = read_visible_records_with_recovery(page, retries=1)
    initial_sync_records: List[CallRecord] = [
        rec for rec in visible_records if rec.unique_key not in seen_keys
    ]
    initial_sync_with_detection = [
        (rec, time.strftime("%Y-%m-%d %H:%M:%S")) for rec in initial_sync_records
    ]
    initial_imported_count = len(initial_sync_with_detection)

    if initial_sync_with_detection:
        total_amount = append_new_records(
            excel_path, ws_calls, wb, initial_sync_with_detection
        )
        for rec, _ in initial_sync_with_detection:
            seen_keys.add(rec.unique_key)
    else:
        total_amount = calculate_total_amount(ws_calls)

    daily_summary = rebuild_daily_summary(ws_calls, ws_summary)
    wb.save(excel_path)
    today_calls, today_minutes, today_amount = get_today_summary(daily_summary)

    print("=" * 72)
    print("Monitor iniciado")
    print(f"Archivo Excel         : {excel_path.resolve()}")
    print(f"Registros ya guardados en Excel (antes de sync): {persisted_count_before_sync}")
    print(f"Sincronización inicial: {initial_imported_count} llamadas visibles importadas")
    print(f"Total acumulado general USD (post-sync): {total_amount}")
    print(
        f"Resumen hoy (post-sync) -> llamadas: {today_calls} | minutos: {today_minutes} | importe USD: {today_amount}"
    )
    print(f"Polling cada          : {interval_seconds} segundos")
    print("Presioná Ctrl+C para detener.")
    print("=" * 72)
    for rec, detected_at in initial_sync_with_detection:
        print_initial_sync_call(rec, detected_at, total_amount)
    if initial_imported_count:
        print(f"[INITIAL SYNC] Completado: {initial_imported_count} llamadas visibles importadas.")
    else:
        print("[INITIAL SYNC] Completado: no hubo llamadas visibles nuevas para importar.")

    while True:
        cycle_start = time.strftime("%Y-%m-%d %H:%M:%S")
        try:
            all_records = read_visible_records_with_recovery(page, retries=1)
            new_records: List[CallRecord] = []

            # Detectar nuevos por unique key contra lo ya persistido
            for rec in all_records:
                if rec.unique_key not in seen_keys:
                    new_records.append(rec)

            if new_records:
                new_records_with_detection = [
                    (rec, time.strftime("%Y-%m-%d %H:%M:%S")) for rec in new_records
                ]
                total_amount = append_new_records(
                    excel_path, ws_calls, wb, new_records_with_detection
                )
                daily_summary = rebuild_daily_summary(ws_calls, ws_summary)
                wb.save(excel_path)
                today_calls, today_minutes, today_amount = get_today_summary(daily_summary)
                for rec, detected_at in new_records_with_detection:
                    seen_keys.add(rec.unique_key)
                    print_new_call(rec, detected_at, total_amount)
                print(
                    f"  Resumen hoy -> llamadas: {today_calls} | minutos: {today_minutes} | "
                    f"importe USD: {today_amount}"
                )
            else:
                print(
                    f"[{cycle_start}] Sin novedades. Filas visibles leídas: {len(all_records)} | "
                    f"Total acumulado USD: {total_amount}"
                )

        except PlaywrightError as e:
            print(f"[ERROR] Fallo Playwright durante lectura: {e}")
        except Exception as e:
            print(f"[ERROR] Excepción inesperada: {e}")

        time.sleep(interval_seconds)


def main() -> int:
    args = parse_args()
    excel_path = Path(args.excel)

    pw = None
    browser = None
    try:
        pw, browser = connect_browser_via_cdp(args.cdp_url)
        page = find_target_page(browser, args.url_contains)

        print(f"Conectado a pestaña: {page.url}")
        wait_table_ready(page)

        monitor_loop(page, excel_path, args.interval)
        return 0

    except KeyboardInterrupt:
        print("\nDetenido por usuario.")
        return 0
    except Exception as e:
        print(f"[FATAL] {e}")
        return 1
    finally:
        if pw:
            pw.stop()


if __name__ == "__main__":
    sys.exit(main())
