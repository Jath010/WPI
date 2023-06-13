Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "IAM - Audit Account and Password Expirations"
$TaskDescription = "Runs every day to reset expired passwords and disable expired accounts" 
$TaskUsername = 'ADMIN\s_tcollins'
$TaskPath = '\WPI\IAM\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\IAM\AuditExpirations\Audit Password and Account expiration for Azure.ps1"' `
    -WorkingDirectory "D:\wpi\batch\IAM\AuditExpirations\"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 12:15am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task | Disable-ScheduledTask