# Monitor de historial de llamadas (Windows)

Script: `monitor_call_log.py`

## Instalar dependencias

```powershell
py -m pip install --upgrade pip
py -m pip install playwright openpyxl
py -m playwright install chromium
```

## Abrir Chrome o Edge con Remote Debugging (misma sesión del usuario)

> Cerrá todas las ventanas del navegador primero, para evitar conflicto de perfil.

### Opción Chrome

```powershell
Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList '--remote-debugging-port=9222','--user-data-dir=C:\temp\chrome-debug-profile'
```

### Opción Edge

```powershell
Start-Process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList '--remote-debugging-port=9222','--user-data-dir=C:\temp\edge-debug-profile'
```

Luego:
1. Iniciá sesión manualmente en la web.
2. Abrí la página de historial de llamadas.

## Ejecutar el monitor

```powershell
py monitor_call_log.py --cdp-url http://127.0.0.1:9222 --url-contains "/call-history" --excel call_log.xlsx --interval 180
```

Parámetros útiles:
- `--url-contains`: filtro para elegir la pestaña correcta.
- `--interval`: segundos de polling (default 180 = 3 minutos).
- `--excel`: nombre/ruta del archivo Excel.

## Salida Excel

Encabezados:
- Customer ID
- Call Date
- Call Start
- Duration (Minutes)
- Amount USD
- Detected At
- Unique Key

El archivo se crea automáticamente si no existe y se agregan solo llamadas nuevas.

Además, se mantiene una hoja `Daily Summary` reconstruida desde `Calls` con:
- Call Date
- Total Calls
- Total Minutes
- Total Amount USD
- Total Available Minutes (sumado desde `Status Durations` donde `From Status = AV`)

Y se agregan hojas para estados del intérprete:
- `Status Log` (cambios reales de estado)
- `Status Durations` (duraciones entre transiciones de estado, incluyendo AV -> LO)

Nuevas hojas de reporting:
- `Monthly Call Summary` (Year-Month, Total Calls, Total Minutes, Total Amount USD)
- `Weekly Available Summary` (Week Start/End domingo-sábado, Total Available Hours, Target/Remaining/Over)
- `Weekly Available Detail` (detalle diario Sunday-Saturday por semana para gráfico diario de available)

Formato visual:
- Encabezados en azul oscuro, texto blanco, bordes finos, zebra stripes, filtros y freeze pane.
- KPIs y gráficos en hojas de resumen (`Daily Summary`, `Monthly Call Summary`, `Weekly Available Summary`).

## Launcher de login (opcional)

Archivos:
- `launch_and_login.py`
- `launch_and_login.ps1`
- `launch_and_login.bat` (wrapper)

### Variables de entorno
Copiá `.env.example` a `.env` y completá:
- `PROPIO_EMAIL` (obligatoria)

El launcher actual usa login **solo por correo electrónico** (sin campo password).

### Ejecutar desde PowerShell
```powershell
.\launch_and_login.ps1
```

O wrapper:
```powershell
.\launch_and_login.bat
```

Para lanzar también el monitor al terminar login:
```powershell
.\launch_and_login.ps1 -RunMonitor
```
