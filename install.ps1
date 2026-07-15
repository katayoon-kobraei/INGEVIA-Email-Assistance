# install.ps1 — one-time setup: creates the virtual environment and installs dependencies

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$VenvPath = Join-Path $ProjectRoot ".venv"
$PythonExe = Join-Path $VenvPath "Scripts\python.exe"

Write-Host "Setting up Email AI Assistant..."

if (-not (Test-Path $VenvPath)) {
    Write-Host "Creating virtual environment..."
    python -m venv $VenvPath
} else {
    Write-Host "Virtual environment already exists, skipping."
}

Write-Host "Installing dependencies..."
& $PythonExe -m pip install --upgrade pip --quiet
& $PythonExe -m pip install -e $ProjectRoot --quiet

Write-Host "Done. Environment is ready."
Write-Host "Next: run setup_scheduler.ps1 to register the automatic email check."