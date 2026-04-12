param(
  [switch]$RunMonitor
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

$cdpUrl = 'http://127.0.0.1:9222/json/version'
$chromePath = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
$userDataDir = 'C:\temp\chrome-cdp'

function Test-Cdp {
  try {
    $resp = Invoke-WebRequest -Uri $cdpUrl -UseBasicParsing -TimeoutSec 2
    return ($resp.StatusCode -eq 200)
  } catch {
    return $false
  }
}

Write-Host '[launcher] Verificando CDP...'
if (-not (Test-Cdp)) {
  Write-Host '[launcher] CDP no disponible. Abriendo Chrome con remote debugging...'
  Start-Process $chromePath -ArgumentList '--remote-debugging-port=9222',("--user-data-dir=$userDataDir")
  Start-Sleep -Seconds 2
}

Write-Host '[launcher] Ejecutando launch_and_login.py...'
py .\launch_and_login.py
if ($LASTEXITCODE -ne 0) {
  throw 'launch_and_login.py falló.'
}

$runMonitorFromEnv = $env:RUN_MONITOR_AFTER_LOGIN -eq '1'
if ($RunMonitor -or $runMonitorFromEnv) {
  Write-Host '[launcher] Ejecutando monitor_call_log.py...'
  py .\monitor_call_log.py --cdp-url http://127.0.0.1:9222 --url-contains '/call-history' --excel call_log.xlsx --interval 180
}

Write-Host '[launcher] Listo.'
