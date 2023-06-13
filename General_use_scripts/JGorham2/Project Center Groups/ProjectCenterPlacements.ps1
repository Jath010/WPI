

function Set-ProjectCenterGroups {
    [CmdletBinding()]
    param (
        
    )
    Write-Verbose "Getting lists"
    $path = "D:\tmp\dfarmerGroups"

    #Get the lists of users, groups, and owners
    $GroupFile = import-csv $path\Groups.csv
    $AllGroupMembersFile = import-csv $path\AllGroupMembers.csv
    $OwnersFile = import-csv $path\Owners.csv
    #Get the list of group from the main data file
    $groups = $groupFile.group | Sort-Object -Unique

    $counter = 0
    foreach ($groupName in $groups) {
        $counter++
        Write-Progress -Activity "Processing Groups" -CurrentOperation $GroupName -PercentComplete (($counter / $Groups.count) * 100) -Id 0
        
        try {
            #Create the Group
            $Object = New-UnifiedGroup -DisplayName $GroupName -PrimarySmtpAddress "$GroupName@wpi.edu" -AccessType Private -Notes "Project Center Group"
            if ($null -ne $Object) {
                #Set the requested features of the group
                Set-UnifiedGroup -Identity $Object.id -UnifiedGroupWelcomeMessageEnabled:$false -RequireSenderAuthenticationEnabled:$false -AutoSubscribeNewMembers:$true
                #Add the Owners before removing the account that ran the script
                foreach ($Owner in $OwnersFile) {
                    Add-UnifiedGroupLinks -identity $Object.id -LinkType member -Links $owner.email
                    Add-UnifiedGroupLinks -identity $Object.id -LinkType owner -Links $owner.email
                }
                Remove-UnifiedGroupLinks -LinkType Owner -Confirm:$False -identity $Object.id -Links jmgorham2_prv@wpi.edu
                Remove-UnifiedGroupLinks -LinkType member -Confirm:$False -identity $Object.id -Links jmgorham2_prv@wpi.edu
            }
            
        }
        catch {
            Write-Host "$groupName creation errored"
        }
        #make the owners only owners
        try {
            $CurrentMembers = Get-UnifiedGroupLinks -LinkType member -Identity $groupName
            foreach ($member in $CurrentMembers) {
                if ($ownerFile.Owners -contains $member.PrimarySmtpAddress) {
                    Remove-UnifiedGroupLinks -LinkType member -Confirm:$False -identity $groupname -Links $member.PrimarySmtpAddress
                }
            }
        }
        catch {
            
        }

        if ($null -ne $CurrentMembers) {
            
        }
        #get only the members of this group from the main data
        $Members = $groupfile | Where-Object { $_.group -eq $groupname }
        #add them to the group
        $counter2 = 0
        foreach ($Member in $members) {
            $counter2++
            if ($null -ne $members.count) {
                Write-Progress -Activity "Processing Members" -CurrentOperation $Member.Email -PercentComplete (($counter2 / $members.count) * 100) -Id 1 -ParentId 0
            }            
            Add-UnifiedGroupLinks -Identity $groupname -LinkType Members -Links $Member.Email
        }
        #add the people who are supposed to be members of every group
        $counter3 = 0
        foreach ($Member in $AllGroupMembersFile) {
            $counter3++
            if ($null -ne $AllGroupMembersFile.count) {
                Write-Progress -Activity "Processing Members Present in All Groups" -CurrentOperation $Member.Email -PercentComplete (($counter3 / $AllGroupMembersFile.count) * 100) -Id 2 -ParentId 0
            }
            Add-UnifiedGroupLinks -Identity $groupname -LinkType Members -Links $Member.Email
        }
    }
}

function Add-GroupToAdministrativeUnit {
    [CmdletBinding()]
    param (
        $GroupID,
        $AUID
    )
    
    begin {
        #Connect-MgGraph
    }
    
    process {
        $URI = "https://graph.microsoft.com/v1.0/directory/administrativeUnits/$AUID/members/"
        $Body = @{"@odata.id" = "https://graph.microsoft.com/v1.0/groups/$GroupID"} | ConvertTo-Json
        $Body
        Invoke-MgGraphRequest -Method POST -Uri $URI -Body $Body
    }
    
    end {
        
    }
}