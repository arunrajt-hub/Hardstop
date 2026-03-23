@echo off
REM Schedule Valmo Hardstop at 4 PM and 10 PM IST
REM Run this once to create the scheduled tasks.

cd /d "%~dp0"
set "BAT_PATH=%~dp0run_valmo_hardstop.bat"

REM Remove existing tasks (if re-running)
schtasks /Delete /TN "ValmoHardstop_Daily" /F 2>nul

REM Create daily task - runs at 4:00 PM and 10:00 PM
REM schtasks supports only one trigger per task, so create two tasks
schtasks /Create /TN "ValmoHardstop_4PM" /TR "\"%BAT_PATH%\"" /SC DAILY /ST 16:00 /F
schtasks /Create /TN "ValmoHardstop_10PM" /TR "\"%BAT_PATH%\"" /SC DAILY /ST 22:00 /F

echo.
echo Scheduled:
echo   ValmoHardstop_4PM  - 4:00 PM daily
echo   ValmoHardstop_10PM - 10:00 PM daily
echo.
echo Make sure Windows time zone is set to India (IST).
echo.
pause
