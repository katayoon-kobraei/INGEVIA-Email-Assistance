@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\disable_outlook_restart_scheduler.ps1"
echo.
pause