<#

Create a script to clean out the groups from a disabled user

#>

function Clear-DisabledUserGroupMembership {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        #############################################
        # Logging
        # Set path for log files:
        $logPath = "D:\wpi\Logs\DisabledUserGroupClear"

        # Get date for logging and file naming:
        $date = Get-Date
        $datestamp = $date.ToString("yyyyMMdd-HHmm")

        Start-Transcript -Append -Path "$($logPath)\$($User)$($datestamp)_DisabledUserGroupClear.log" -Force
        #############################################
        
        import-module AzureADPreview -force

        #Creds for automagic login on ScriptHost-02
        $Credentials = $null
        if ($env:COMPUTERNAME -eq "SCRIPTHOST-02") {
            $Credentials = Import-Clixml -Path 'D:\wpi\XML\exch_automation\exch_automation@wpi.edu.xml'
            Connect-AzureAD -Credential $Credentials
            Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false
        }

    }
    
    process {
        $DisabledUsers = get-aduser -Filter "Enabled -eq '$false'" -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu"
        foreach ($user in $DisabledUsers) {
            $UID = Get-AzureADUser -ObjectId $user.userprincipalname
            $Groups = $uid.UserPrincipalName | Get-AzureADUserMembership | Where-Object { $_.DisplayName -ne "Deny Logon Interactively" -and $_.DisplayName -ne "Domain Users" -and $_.DisplayName -ne "License_Disabled" }
            Write-Host "Operating on: $($user.UserPrincipalName)"
            foreach ($group in $groups) {
                $groupname = $group.displayName
                Write-Host "Removing group: $groupname"
                if ($group.ObjectType -ne "Role") {
                    $graphData = get-azureadmsgroup -id $group.ObjectId
                }
                if ($group.ObjectType -eq "Role") {
                    Write-Host "Removing role: $($group.displayname)"
                    Remove-AzureADDirectoryRoleMember -objectid $group.ObjectId -MemberId $uid.ObjectId    
                }
                elseif ($graphData.grouptypes -eq "Unified") {
                    if ($null -eq (get-azureadgroupowner -ObjectId $group.ObjectId | where-object { $_.UserPrincipalName -eq $uid.userprincipalname })) {
                        try { Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $UID.ObjectId }
                        catch { Write-Verbose "User isn't a member of $groupname" }
                    }
                    else {
                        try { Remove-AzureADGroupOwner -ObjectId $group.ObjectId -OwnerId $UID.ObjectId }
                        catch { Write-Verbose "User isn't an owner of $groupname" }
                    }
                }
                elseif ($graphData.SecurityEnabled -eq $true) {
                    try { Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $UID.ObjectId }
                    catch { Write-Verbose "User isn't a member of $groupname" }
                }
                else {
                    try { Remove-DistributionGroupMember -Identity $group.ObjectId -Member $UID.ObjectId -Confirm:$false }
                    catch { Write-Verbose "$groupname isn't a Distribution Group" }
                }
            }
        }
        
    }
    
    end {
        Stop-Transcript
    }
}

Clear-DisabledUserGroupMembership