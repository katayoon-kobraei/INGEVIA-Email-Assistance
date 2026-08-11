@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\uninstall.ps1"
echo.
pause
