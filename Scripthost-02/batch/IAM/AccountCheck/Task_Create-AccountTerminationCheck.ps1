Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "IAM - Check for Missed Terminations"
$TaskDescription = "Runs every day to validate that all active accounts have an corresponding UNIX account" 
$TaskUsername = 'ADMIN\s_tcollins'
$TaskPath = '\WPI\IAM\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\IAM\AccountCheck\AccountTerminationCheck.ps1"' `
    -WorkingDirectory "d:\wpi\batch\IAM\AccountCheck\"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 6am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task | Disable-ScheduledTask