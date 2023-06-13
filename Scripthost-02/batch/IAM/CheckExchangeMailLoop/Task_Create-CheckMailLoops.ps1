Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "IAM - Check for Mail Loops"
$TaskDescription = "Get list of adusers that have been modified in the last 24 hours and verify that there are no mail loops" 
$TaskUsername = 'ADMIN\s_tcollins'
$TaskPath = '\WPI\IAM\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\IAM\CheckExchangeMailLoop\CheckMailLoops.ps1"' `
    -WorkingDirectory "d:\wpi\batch\IAM\CheckExchangeMailLoop\"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 6am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT6H"
$task | Set-ScheduledTask 
$task | Disable-ScheduledTask