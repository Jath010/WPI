Clear-Host
## Create New Scheduled Task

## Set Task Configuration
$TaskName = "Exchange Online - Get Alumni Pilot Statistics"
$TaskDescription = "Monthly task to pull stats on usage for Alumni Pilot mailboxes"
$TaskUsername = 'ADMIN\exch_automation'
$TaskPath = '\WPI\Exchange\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "D:\wpi\batch\Exchange\AlumniPilotStats\Get-AlumniPilotStats.ps1"' `
    -WorkingDirectory "d:\wpi\batch\ExchangeMigration\"

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
#$task.Triggers.Repetition.Duration = "P1D"
#$task.Triggers.Repetition.Interval = "PT10M"
#$task | Set-ScheduledTask 

#$triggers = $TaskDefinition.Triggers
#$trigger = $triggers.Create(1) # Creates a "One time" trigger
#$trigger = $triggers.Create(4)
#$trigger.DaysOfMonth = 1

$task | Disable-ScheduledTask



