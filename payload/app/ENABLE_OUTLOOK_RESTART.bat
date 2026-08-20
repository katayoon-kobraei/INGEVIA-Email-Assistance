@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\setup_outlook_restart_scheduler.ps1"
echo.
pause