$saturdayTrigger = New-JobTrigger -At "04/22/2022 1:00:00"

Register-ScheduledJob -Name "Restart Computer" -Trigger $saturdayTrigger -ScriptBlock {
    Restart-Computer
}