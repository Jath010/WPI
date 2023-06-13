Clear-Host
## Create New Scheduled Task

## Set the task action
$TaskName = "Migrate Mailboxes to Exchange Online"
$TaskDescription = "Checks every 30 minutes and moves any local mailboxes to Office 365"
#$TaskUsername = New-ScheduledTaskPrincipal -UserID "exch_automation@wpi.edu" -LogonType Password 
$TaskUsername = 'ADMIN\exch_automation'
$TaskPath = '\WPI\Exchange\'

$TaskAction = New-ScheduledTaskAction `
    -Execute 'Powershell.exe' `
    -Argument '-NoProfile -WindowStyle Hidden -File "e:\wpi\batch\ExchangeMigration\Migrate_Local_Mailboxes_Online.ps1"'
    
$TaskTrigger =  New-ScheduledTaskTrigger -Daily -At 12am

<#
$Task = New-ScheduledTask -Action $TaskAction -Principal $TaskUsername -Trigger $TaskTrigger -Description $TaskDescription
    #>
#<#
Clear-Host
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger `
    -TaskPath $TaskPath -Description $TaskDescription -User $TaskUsername
    #>

#Register-ScheduledTask -TaskName $TaskName  -InputObject $Task -TaskPath $TaskPath
#<#
# Unregister-ScheduledTask $TaskName -Confirm:$false

$task = Get-ScheduledTask $TaskName
$task
$task.Triggers.Repetition.Duration = "P1D"
$task.Triggers.Repetition.Interval = "PT30M"
$task | Set-ScheduledTask 
$task | Disable-ScheduledTask

cls
Get-ScheduledTask $TaskName | Set-ScheduledTask -User $TaskUsername

#mdo95jIENLT95c3G3OeyMUcy6uRj8T9Z0zlU7hiRZ10vcXiNFigRaGGmYnQ74hB

#>