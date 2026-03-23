# Schedule Valmo Hardstop - 4 PM & 10 PM IST
# Run this script once to create the scheduled tasks (requires Admin for system tasks, or runs as current user)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath = Join-Path $scriptDir "run_valmo_hardstop.bat"

if (-not (Test-Path $batPath)) {
    Write-Error "Not found: $batPath"
    exit 1
}

# Ensure full path
$batPath = (Resolve-Path $batPath).Path

# Remove existing tasks if present (so we can re-run this script to update)
$task1 = "ValmoHardstop_4PM_IST"
$task2 = "ValmoHardstop_10PM_IST"

Get-ScheduledTask -TaskName $task1 -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false
Get-ScheduledTask -TaskName $task2 -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false

# Create daily triggers at 4:00 PM and 10:00 PM (local time - set system to IST)
$action = New-ScheduledTaskAction -Execute $batPath -WorkingDirectory $scriptDir
$trigger1 = New-ScheduledTaskTrigger -Daily -At "4:00PM"
$trigger2 = New-ScheduledTaskTrigger -Daily -At "10:00PM"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Single task with multiple triggers
$triggers = @($trigger1, $trigger2)
Register-ScheduledTask -TaskName "ValmoHardstop_Daily" -Action $action -Trigger $triggers -Settings $settings -Description "Valmo Hardstop: Gmail -> Google Sheet at 4 PM & 10 PM IST"

Write-Host "Scheduled: ValmoHardstop_Daily" -ForegroundColor Green
Write-Host "  - 4:00 PM (IST)" -ForegroundColor Cyan
Write-Host "  - 10:00 PM (IST)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Task runs: $batPath" -ForegroundColor Gray
Write-Host "Make sure Windows is set to IST (India Standard Time) or adjust times in this script." -ForegroundColor Yellow
Write-Host ""
Write-Host "To remove: Unregister-ScheduledTask -TaskName 'ValmoHardstop_Daily'" -ForegroundColor Gray
