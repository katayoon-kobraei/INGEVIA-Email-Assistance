@echo off
setlocal
cd /d "%~dp0"

echo ======================================================
echo   INGEVIA Email Assistant - Reprocess UNSORTED
echo ======================================================
echo.
echo This re-checks every email in this year's UNSORTED folder
echo WITHOUT opening Outlook. Emails that can now be classified
echo will be MOVED to their correct folder. Emails judged as
echo junk/spam/promotional will be DELETED permanently.
echo.
echo Recommended: preview first with a dry run (nothing is
echo changed), then run for real.
echo.
echo   1. Dry run (preview only, changes nothing)
echo   2. Run for real (moves and deletes)
echo.
set /p choice=Enter 1 or 2:
if "%choice%"=="1" goto dryrun
if "%choice%"=="2" goto realrun
echo.
echo Invalid selection.
pause
exit /b 1

:dryrun
"%~dp0..\runtime\python.exe" "%~dp0reprocess_unsorted.py" --dry-run
goto done

:realrun
"%~dp0..\runtime\python.exe" "%~dp0reprocess_unsorted.py"
goto done

:done
echo.
pause
