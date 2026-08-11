@echo off
setlocal
cd /d "%~dp0"
title Prepare Offline INGEVIA Email Assistant Installer

echo ======================================================
echo   Prepare Fully Offline Installer Copy
echo ======================================================
echo.
echo Run this once on a Windows PC with internet access.
echo It downloads the official portable Python runtime and all
echo required dependency files into this installer folder.
echo You can then copy the complete folder to other PCs and
echo install without internet.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_installer\prepare_offline.ps1" -PackageRoot "%~dp0."
if errorlevel 1 (
    echo.
    echo Offline preparation failed. Read the error above.
    pause
    exit /b 1
)
echo.
pause
