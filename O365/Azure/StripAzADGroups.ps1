function Remove-AllAzureADGroups {
    [CmdletBinding()]
    param (
        $UPN,
        $user
    )
    
    begin {
        connect-azuread
    }
    
    process {
        if($null -ne $user){
            $UPN = "${User}@wpi.edu"
        }
        $UID = Get-AzureADUser -SearchString $UPN
        $Groups = $UID | Get-AzureADUserMembership
        foreach ($group in $groups) {
            $groupname = $group.displayName
            Write-Verbose "Removing $groupname"
            try { Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $UID.ObjectId }
            catch {Write-Verbose "User isn't a member of $groupname"}
            try { Remove-AzureADGroupOwner -ObjectId $group.ObjectId -OwnerId $UID.ObjectId }
            catch {Write-Verbose "User isn't an owner of $groupname"}
            try{Remove-DistributionGroupMember -Identity $group.ObjectId -Member $UID.ObjectId -Confirm:$false -ErrorAction Ignore}
            catch{}
        }
    }
    
    end {
        
    }
}