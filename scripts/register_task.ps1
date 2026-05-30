$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\eklementeva\Cowork\stm-vercel-dashboard\scripts\daily_refresh.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At '08:45'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 4)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
try { Unregister-ScheduledTask -TaskName 'STM Dashboard Daily Refresh' -Confirm:$false -ErrorAction Stop } catch {}
Register-ScheduledTask -TaskName 'STM Dashboard Daily Refresh' -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'STM dashboard daily refresh at 08:45 MSK'
Get-ScheduledTask -TaskName 'STM Dashboard Daily Refresh' | Select-Object TaskName, State, @{n='NextRun';e={(Get-ScheduledTaskInfo $_).NextRunTime}} | Format-List
