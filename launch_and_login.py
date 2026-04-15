#!/usr/bin/env python3
"""
Launcher de login para preparar sesión antes de ejecutar monitor_call_log.py.

Flujo:
1) Verifica CDP (localhost:9222)
2) Si no está, abre Chrome con remote debugging + user-data-dir
3) Se conecta por Playwright CDP
4) Navega a login_url, completa credenciales desde entorno (.env)
5) Hace login y espera página destino
6) Deja Chrome abierto
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional

from playwright.sync_api import Browser, Page, TimeoutError, sync_playwright

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

# ---------------------------------------------------------------------------
# CONFIGURACIÓN EDITABLE (reemplazá estos valores para tu sitio)
# ---------------------------------------------------------------------------
LOGIN_URL = "https://interpreters.propio-ls.com/Login?ReturnUrl=%2Fportal"
POST_LOGIN_URL_CONTAINS = "/portal"

SELECTOR_EMAIL = "input[name='username']"
SELECTOR_LOGIN_BUTTON = "input#wp-submit"

CDP_URL = "http://127.0.0.1:9222"
CDP_PORT = 9222
CHROME_USER_DATA_DIR = r"C:\temp\chrome-cdp"
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
AUTO_REDIRECT_TIMEOUT_MS = 8000
# ---------------------------------------------------------------------------


def log(msg: str) -> None:
    print(f"[launch_and_login] {msg}")


def load_env_file() -> None:
    if load_dotenv is None:
        log("python-dotenv no instalado; se usarán solo variables de entorno del sistema.")
        return
    env_path = Path(".env")
    if env_path.exists():
        load_dotenv(env_path)
        log("Archivo .env cargado.")


def get_required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta variable de entorno requerida: {name}")
    return value


def get_optional_env(name: str) -> str:
    return os.getenv(name, "").strip()


def cdp_available(cdp_url: str = CDP_URL, timeout_sec: float = 1.5) -> bool:
    version_url = f"{cdp_url}/json/version"
    try:
        with urllib.request.urlopen(version_url, timeout=timeout_sec) as resp:
            return resp.status == 200
    except (urllib.error.URLError, TimeoutError, ConnectionError, OSError):
        return False


def find_chrome_executable() -> Optional[str]:
    for candidate in CHROME_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return None


def launch_chrome_with_cdp() -> None:
    chrome_path = find_chrome_executable()
    if not chrome_path:
        raise RuntimeError("No se encontró chrome.exe en rutas típicas. Editá CHROME_CANDIDATES.")

    Path(CHROME_USER_DATA_DIR).mkdir(parents=True, exist_ok=True)
    args = [
        chrome_path,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}",
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    log("Chrome lanzado con remote debugging.")


def ensure_cdp_ready(max_wait_sec: int = 20) -> None:
    if cdp_available():
        log("CDP ya disponible (Chrome ya abierto).")
        return

    log("CDP no disponible, abriendo Chrome...")
    launch_chrome_with_cdp()

    deadline = time.time() + max_wait_sec
    while time.time() < deadline:
        if cdp_available():
            log("CDP disponible.")
            return
        time.sleep(0.8)

    raise RuntimeError("No se pudo habilitar CDP en tiempo esperado.")


def connect_cdp() -> tuple[object, Browser]:
    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(CDP_URL)
    except Exception:
        pw.stop()
        raise
    log("Conexión CDP exitosa.")
    return pw, browser


def pick_page(browser: Browser) -> Page:
    for ctx in browser.contexts:
        for page in ctx.pages:
            if page.url and page.url != "about:blank":
                return page

    if browser.contexts:
        ctx = browser.contexts[0]
    else:
        ctx = browser.new_context()
    return ctx.new_page()


def save_login_debug_artifacts(page: Page) -> None:
    try:
        page.screenshot(path="login_debug.png", full_page=True)
        log("Debug screenshot guardada: login_debug.png")
    except Exception as exc:
        log(f"No se pudo guardar screenshot de debug: {exc}")
    try:
        Path("login_debug.html").write_text(page.content(), encoding="utf-8")
        log("Debug HTML guardado: login_debug.html")
    except Exception as exc:
        log(f"No se pudo guardar HTML de debug: {exc}")


def do_login(page: Page, email: str) -> None:
    log(f"Navegando a login... {LOGIN_URL}")
    page.goto(LOGIN_URL, wait_until="domcontentloaded")

    log("Esperando redirección automática por sesión existente...")
    try:
        page.wait_for_url(
            lambda url: POST_LOGIN_URL_CONTAINS.lower() in url.lower(),
            timeout=AUTO_REDIRECT_TIMEOUT_MS,
        )
        log("Sesión ya iniciada, se omite login")
        log(f"Portal listo: {page.url}")
        return
    except TimeoutError:
        log("No hubo redirección automática, se procede con login manual")

    log("Completando credenciales...")
    if not email:
        raise RuntimeError("PROPIO_EMAIL es obligatoria cuando se requiere iniciar sesión.")
    try:
        page.wait_for_selector(SELECTOR_EMAIL, timeout=15000)
    except TimeoutError:
        save_login_debug_artifacts(page)
        raise RuntimeError(
            f"No se encontró el selector de email: {SELECTOR_EMAIL}. "
            "Se guardaron login_debug.png y login_debug.html"
        )
    page.fill(SELECTOR_EMAIL, email)

    log("Login mode: email-only")

    log("Enviando login...")
    page.click(SELECTOR_LOGIN_BUTTON)

    log(f"Esperando URL final que contenga: {POST_LOGIN_URL_CONTAINS}")
    page.wait_for_url(
        lambda url: POST_LOGIN_URL_CONTAINS.lower() in url.lower(),
        timeout=30000,
    )
    log(f"Página final detectada: {page.url}")


def main() -> int:
    load_env_file()
    email = get_optional_env("PROPIO_EMAIL")

    pw = None
    browser = None
    try:
        ensure_cdp_ready()
        pw, browser = connect_cdp()
        page = pick_page(browser)
        do_login(page, email)

        log("Login completado. Chrome queda abierto y listo para monitor_call_log.py")
        return 0

    except Exception as exc:
        log(f"ERROR: {exc}")
        return 1

    finally:
        # No cerrar el browser remoto (sesión del usuario)
        if pw:
            pw.stop()


if __name__ == "__main__":
    sys.exit(main())
