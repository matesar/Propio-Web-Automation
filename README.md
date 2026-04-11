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
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp\chrome-debug-profile"
```

### Opción Edge

```powershell
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp\edge-debug-profile"
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
