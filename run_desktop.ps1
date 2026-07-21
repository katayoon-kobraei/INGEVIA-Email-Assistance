$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    Write-Host "No se encontró .venv. Créalo con: py -m venv .venv" -ForegroundColor Yellow
    exit 1
}

& ".\.venv\Scripts\Activate.ps1"
python -m desktop_app.main
