@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\setup_scheduler.ps1"
echo.
pause
