Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "Exchange Online - Manage Exchange Online Protection Blocked Senders"
$TaskDescription = "This will run on an interval to check if there are any blocked senders due to phishing/bulk mail violations.  
    If there are, it will see if the password has been changed following the block and once it has been a set number of hours since 
    the password change, the access to send will be restored."
$TaskUsername = 'ADMIN\exch_automation'
$TaskPath = '\WPI\Exchange\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\Exchange\ManagedBlockedUsers\ManageBlockedUsers.ps1"' `
    -WorkingDirectory "D:\wpi\batch\Exchange\ManagedBlockedUsers"

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT10M"

$task | Set-ScheduledTask 
$task | Disable-ScheduledTask