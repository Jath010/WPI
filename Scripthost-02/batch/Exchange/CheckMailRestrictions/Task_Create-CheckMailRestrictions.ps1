Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "Exchange Online - Check for Mail Restrictions"
$TaskDescription = "Get list of adusers that have been modified in the last 24 hours and verify that there are no issues with mail restrictions, bad forwards, or hidden from GAL" 
$TaskUsername = 'ADMIN\exch_automation'
$TaskPath = '\WPI\Exchange\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\Exchange\CheckMailRestrictions\CheckMailRestrictions.ps1"' `
    -WorkingDirectory "D:\wpi\batch\Exchange\CheckMailRestrictions\"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 6am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
Sleep 5
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT1H"
$task | Set-ScheduledTask 
$task | Disable-ScheduledTask