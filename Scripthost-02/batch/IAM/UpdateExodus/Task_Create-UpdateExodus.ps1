Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "IAM - Update EXODUS Database"
$TaskDescription = "Runs every 15 minutes to validate that all active accounts have an entry in the SelfService database EXODUS" 
$TaskUsername = 'ADMIN\s_tcollins'
$TaskPath = '\WPI\IAM\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "d:\wpi\batch\IAM\UpdateExodus\UpdateExodus.ps1"' `
    -WorkingDirectory "d:\wpi\batch\IAM\UpdateExodus\"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 12am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT15M"
$task | Set-ScheduledTask 
$task | Disable-ScheduledTask