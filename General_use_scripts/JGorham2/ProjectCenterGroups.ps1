$CSVpath = "C:\tmp\Project Center"
$GroupsData = Import-Csv "$($CSVpath)\Groups21-22.csv"

foreach ($group in $GroupsData) {
    $groupAddress = $Group.Group
    $DisplayName = ($Group.Group.split('@')[0])
    $Alias = $Group.Group.split('@')[0].substring(3)

    $groupid = New-UnifiedGroup -AutoSubscribeNewMembers:$true -Owner dfarmer@wpi.edu -AlwaysSubscribeMembersToCalendarEvents:$true -AccessType Private -Members dfusaro@wpi.edu, jrichard@wpi.edu, rmckeogh@wpi.edu, ebell@wpi.edu, gcollins@wpi.edu, nafay@wpi.edu, ccruta@wpi.edu -DisplayName $DisplayName -PrimarySmtpAddress $groupAddress

    Set-UnifiedGroup -Identity $groupid.id -emailaddresses @{add="smtp:$($alias)@wpi.edu"}
}