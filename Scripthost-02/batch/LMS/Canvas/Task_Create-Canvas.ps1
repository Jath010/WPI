Clear-Host
## Create New Scheduled Task

## Create task for PROD instance.
## ------------------------------

## Set Task Configuration
$TaskName = "Canvas - Process SIS Import - PROD"
$TaskDescription = "At :05 every hour the script will copy the updated Banner extract for Canvas and process it to update the Course and Enrollment data."
$TaskUsername = 'ADMIN\canvas_service'
$TaskPath = '\WPI\LMS\'

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "d:\wpi\batch\LMS\Canvas\SIS_Import_PROD\Canvas_SIS_Import.ps1"' `
    -WorkingDirectory "d:\wpi\batch\LMS\Canvas\SIS_Import_PROD\"

## Set Start Time   
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 6:05am

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT60M"
$task | Set-ScheduledTask
$task | Disable-ScheduledTask


## Create task for BETA instance.
## ------------------------------

## Set Task Configuration
$TaskName = "Canvas - Process SIS Import - BETA"

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "d:\wpi\batch\LMS\Canvas\SIS_Import_BETA\Canvas_SIS_Import.ps1"' `
    -WorkingDirectory "d:\wpi\batch\LMS\Canvas\SIS_Import_BETA\"

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT60M"
$task | Set-ScheduledTask
$task | Disable-ScheduledTask

## Create task for TEST instance.
## ------------------------------

## Set Task Configuration
$TaskName = "Canvas - Process SIS Import - TEST"

## Build Task
$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "d:\wpi\batch\LMS\Canvas\SIS_Import_TEST\Canvas_SIS_Import.ps1"' `
    -WorkingDirectory "d:\wpi\batch\LMS\Canvas\SIS_Import_TEST\"

## Register/Create Task in Task Scheduler
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername

## Get Task to make additional post-creation changes
$task = Get-ScheduledTask $TaskName

## Set recurring task duration/interval and then disable the task to allow it to be manually enabled when it is ready
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT60M"
$task | Set-ScheduledTask
$task | Disable-ScheduledTask
