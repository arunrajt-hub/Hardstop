@echo off
REM Run Valmo Hardstop automation (Gmail -> Google Sheet)
REM Scheduled: 4 PM & 10 PM IST via schedule_valmo_hardstop.bat

cd /d "%~dp0"
python valmo_hardstop_gmail_to_sheet.py
if errorlevel 1 (
    echo Valmo Hardstop run failed
    exit /b 1
)
echo Valmo Hardstop completed successfully
exit /b 0
