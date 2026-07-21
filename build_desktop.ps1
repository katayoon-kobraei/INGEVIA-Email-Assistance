$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

& ".\.venv\Scripts\Activate.ps1"
python -m pip install pyinstaller

pyinstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name "INGEVIA Email Assistant" `
    --paths "." `
    --collect-all google.genai `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    desktop_app\main.py

Write-Host "Aplicación creada en: dist\INGEVIA Email Assistant" -ForegroundColor Green
Write-Host "Mantén el archivo .env junto al ejecutable o ejecuta la aplicación desde la carpeta del proyecto." -ForegroundColor Yellow
