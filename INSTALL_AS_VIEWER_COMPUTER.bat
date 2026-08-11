@echo off
setlocal
cd /d "%~dp0"
title Install INGEVIA Email Assistant - Viewer Computer
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_installer\install.ps1" -PackageRoot "%~dp0." -InstallMode VIEWER
if errorlevel 1 (
    echo.
    echo Installation failed. Read the error message above.
    pause
    exit /b 1
)
echo.
echo Viewer computer installation completed.
pause
