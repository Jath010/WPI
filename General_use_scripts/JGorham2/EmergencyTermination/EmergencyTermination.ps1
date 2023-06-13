#Removes all a user's groups, disabled them in AD, places them in the disabled ou, adds the "deny logon interactively" group

function Start-EmergencyTermination {
    [CmdletBinding()]
    param (
        $User
    )
    
    begin {
        #############################################
        # Logging
        # Set path for log files:
        $logPath = "D:\wpi\Logs\EmergencyTermination"

        # Get date for logging and file naming:
        $date = Get-Date
        $datestamp = $date.ToString("yyyyMMdd-HHmm")

        Start-Transcript -Append -Path "$($logPath)\$($User)$($datestamp)_EmergencyTermination.log" -Force
        #############################################

        Connect-AzureAD
        if ($user.endswith("@wpi.edu")) {
            $user = $user.split("@")[0]
        }
        $Account = get-aduser $user
        $UID = Get-AzureADUser -ObjectId $account.userprincipalname
    }
    
    process {
        #Disable the Account and move it to the Disabled OU
        Disable-ADAccount -Identity $account.samaccountname
        Move-ADObject -TargetPath "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Identity $Account.DistinguishedName

        #Disable Account in Azure and revoke refresh tokens
        Set-AzureADUser -ObjectId $uid.ObjectId -AccountEnabled $false
        Revoke-AzureADUserAllRefreshToken -ObjectId $uid.ObjectId
    
        # Section Removed AD groups
        $ADGroups = Get-ADPrincipalGroupMembership $Account.samaccountname | Where-Object {$_.Name -ne "Deny Logon Interactively" -and $_.Name -ne "Domain Users"}
        foreach ($group in $ADGroups) {
            Write-Host "Removing group: $($group.name)"
            try {
                Remove-ADGroupMember -Identity $group -Members $Account.samaccountname -Confirm:$false
            }
            catch {
                Write-Verbose "Could not remove $($group.name)"
            }
        }

        #Section Removes Azure Groups
        $Groups = $UID | Get-AzureADUserMembership | Where-Object {$_.DisplayName -ne "Deny Logon Interactively" -and $_.DisplayName -ne "Domain Users" -and $_.DisplayName -ne "License_Disabled"}
        foreach ($group in $groups) {
            if ($group.ObjectType -eq "Role") {
                Write-Host "Removing role: $($group.displayname)"
                Remove-AzureADDirectoryRoleMember -objectid $group.ObjectId -MemberId $uid.ObjectId    
            }
            else {
                $groupname = $group.displayName
                Write-Host "Removing group: $groupname"
                try { Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $UID.ObjectId -ErrorAction SilentlyContinue}
                catch { Write-Verbose "User isn't a member of $groupname" }
                try { Remove-AzureADGroupOwner -ObjectId $group.ObjectId -OwnerId $UID.ObjectId -ErrorAction SilentlyContinue}
                catch { Write-Verbose "User isn't an owner of $groupname" }
                try { Remove-DistributionGroupMember -Identity $group.ObjectId -Member $UID.ObjectId -Confirm:$false -ErrorAction Ignore }
                catch { Write-Verbose "$groupname isn't a Distribution Group" }
            }
        }

        #Add deny logon
        Add-ADGroupMember -Identity "Deny Logon Interactively" -Members $Account.samaccountname

        #Sync Azure
        $AADC = New-PSSession aadc-utl-p-w03
        Invoke-Command -Session $AADC -ScriptBlock { Import-module adsync }
        Invoke-Command -Session $AADC -ScriptBlock { start-adsyncsynccycle -policytype Delta }
        Remove-PSSession $AADC
    }
    
    end {
        Stop-Transcript
    }
}