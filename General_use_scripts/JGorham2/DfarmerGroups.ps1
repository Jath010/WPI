

function Set-DfarmerGroups {
    [CmdletBinding()]
    param (
        
    )
    Write-Verbose "Getting lists"
    $path = "D:\tmp\dfarmerGroups"
    $lists = Get-ChildItem $path
    $counter = 0
    foreach ($list in $lists) {
        $counter++
        Write-Progress -Activity "Processing Lists" -CurrentOperation $List.Name -PercentComplete (($counter / $lists.count) * 100) -Id 0
        $GroupName = $list.name.split(".")[0]
        $members = Import-Csv -Path $path\$list

        try {
            $Object = New-UnifiedGroup -DisplayName $GroupName -PrimarySmtpAddress "$GroupName@wpi.edu" -AccessType Private
            if ($null -ne $Object) {
                Set-UnifiedGroup -Identity $Object.id -UnifiedGroupWelcomeMessageEnabled:$false
                Add-UnifiedGroupLinks -identity $Object.id -LinkType member -Links raalicea@wpi.edu
                Add-UnifiedGroupLinks -identity $Object.id -LinkType owner -Links raalicea@wpi.edu
                Remove-UnifiedGroupLinks -LinkType Owner -Confirm:$False -identity $Object.id -Links jmgorham2_prv@wpi.edu
                Remove-UnifiedGroupLinks -LinkType member -Confirm:$False -identity $Object.id -Links jmgorham2_prv@wpi.edu
            }
            
        }
        catch {
            Write-Host "$groupName creation errored"
        }

        try {
            $CurrentMembers = Get-UnifiedGroupLinks -LinkType member -Identity $groupName
            foreach ($member in $CurrentMembers) {
                if ($member.PrimarySmtpAddress -eq "dfarmer@wpi.edu") {
                    Remove-UnifiedGroupLinks -LinkType owner -Confirm:$False -identity $groupname -Links $member.PrimarySmtpAddress
                    Remove-UnifiedGroupLinks -LinkType member -Confirm:$False -identity $groupname -Links $member.PrimarySmtpAddress
                }
                if ($member.PrimarySmtpAddress -ne "raalicea@wpi.edu") {
                    Remove-UnifiedGroupLinks -LinkType member -Confirm:$False -identity $groupname -Links $member.PrimarySmtpAddress
                }
            }
        }
        catch {
            
        }

        if ($null -ne $CurrentMembers) {
            
        }
        $counter2 = 0
        foreach ($Member in $members) {
            $counter2++
            Write-Progress -Activity "Processing Members" -CurrentOperation $Member.'student email' -PercentComplete (($counter2 / $members.count) * 100) -Id 1 -ParentId 0
            Add-UnifiedGroupLinks -Identity $groupname -LinkType Members -Links $Member.'student email'
        }
    }
}