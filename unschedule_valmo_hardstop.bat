@echo off
REM Remove Valmo Hardstop scheduled tasks

schtasks /Delete /TN "ValmoHardstop_Daily" /F 2>nul
schtasks /Delete /TN "ValmoHardstop_4PM" /F 2>nul
schtasks /Delete /TN "ValmoHardstop_10PM" /F 2>nul

echo Removed Valmo Hardstop scheduled tasks.
pause
