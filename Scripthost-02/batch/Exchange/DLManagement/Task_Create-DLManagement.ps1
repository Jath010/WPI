Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$ScriptPath      = 'D:\wpi\batch\Exchange\DLManagement'
$ScriptFilename  = 'DLManagement.ps1'
$TaskName        = "Exchange Online - Update Standing List Distribution Groups"
$TaskDescription = "Update Distribution groups from Standing List JSON data" 
$TaskUsername    = 'ADMIN\exch_automation'
$TaskPath        = '\WPI\Exchange\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument "-NoProfile -WindowStyle Hidden -File `"$ScriptPath\$ScriptFilename`"" `
    -WorkingDirectory "$ScriptPath"

## Set Start Time    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 6pm

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval 
#$task.Triggers.Repetition.Duration = "P1D"
#$task.Triggers.Repetition.Interval = "PT1H"

## Set and disable the task to allow it to be manually enabled when it is ready
$task | Set-ScheduledTask 
$task | Disable-ScheduledTask