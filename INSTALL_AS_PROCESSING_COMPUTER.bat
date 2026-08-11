@echo off
setlocal
cd /d "%~dp0"
title Install INGEVIA Email Assistant - Processing Computer
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_installer\install.ps1" -PackageRoot "%~dp0." -InstallMode PROCESSOR
if errorlevel 1 (
    echo.
    echo Installation failed. Read the error message above.
    pause
    exit /b 1
)
echo.
echo Processing computer installation completed.
pause
