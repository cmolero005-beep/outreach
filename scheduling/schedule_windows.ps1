# Registers a Windows Task Scheduler job: Mon-Fri at 11:00 AM.
# If the laptop is off/asleep at 11:00, it runs as soon as it's back on (same day).
# Run once from PowerShell in the project folder:   powershell -ExecutionPolicy Bypass -File scheduling\schedule_windows.ps1
$project = Split-Path -Parent $PSScriptRoot
$python  = (Get-Command python).Source
$action  = New-ScheduledTaskAction -Execute $python -Argument "`"$project\outreach.py`" run" -WorkingDirectory $project
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 11:00am
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 2)
Register-ScheduledTask -TaskName "Networking Outreach" -Action $action -Trigger $trigger -Settings $settings -Force
Write-Host "Scheduled: 'Networking Outreach' runs Mon-Fri at 11:00 AM."
Write-Host "To remove:  Unregister-ScheduledTask -TaskName 'Networking Outreach'"
