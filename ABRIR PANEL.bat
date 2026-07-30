@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo No se ha encontrado la instalacion todavia.
    echo Ejecuta primero INSTALAR.bat con doble clic.
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File ".\run_desktop.ps1"
