@echo off
setlocal

REM Wrapper para PowerShell (launcher principal)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launch_and_login.ps1"

endlocal
