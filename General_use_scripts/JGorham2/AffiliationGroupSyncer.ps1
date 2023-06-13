function Sync-AffiliationGroup {
    [cmdletbinding()]
    param (
        $PrimaryAffiliation,
        $SubAffiliation
    )


    Write-Verbose "Gathering Email List"
    if($null -ne $SubAffiliation){
        $userlist = get-aduser -filter {ExtensionAttribute7 -eq $PrimaryAffiliation -and ExtensionAttribute3 -eq $SubAffiliation -and Enabled -eq $True} -Property ExtensionAttribute7,ExtensionAttribute3 -ResultSetSize 20000
        $groupid = "$($PrimaryAffiliation)_$($SubAffiliation)_Synced"

        try{Get-ADGroup $groupid}
        catch{
            New-ADGroup -Name $groupid -SamAccountName $groupid -GroupCategory Security -GroupScope Global -DisplayName $groupid -Path "OU=Resources,OU=Groups,DC=admin,DC=wpi,DC=edu"
            Add-ADGroupMember -Identity $groupid -Members jmgorham2_np
        }

        $GroupList = (Get-ADGroup "$($PrimaryAffiliation)_$($SubAffiliation)_Synced" -Properties members).Members |Get-ADUser|Select-Object samaccountname
    }
    else{
        $userlist = get-aduser -filter {ExtensionAttribute7 -eq $PrimaryAffiliation -and Enabled -eq $True} -Property ExtensionAttribute7 -ResultSetSize 20000
        $groupid = "$($PrimaryAffiliation)_Synced"
        
        try{Get-ADGroup $groupid}
        catch{
            New-ADGroup -Name $groupid -SamAccountName $groupid -GroupCategory Security -GroupScope Global -DisplayName $groupid -Path "OU=Resources,OU=Groups,DC=admin,DC=wpi,DC=edu"
            Add-ADGroupMember -Identity $groupid -Members jmgorham2_np
        }
        
        $GroupList = (Get-ADGroup "$($PrimaryAffiliation)_Synced" -Properties members).Members |Get-ADUser|Select-Object samaccountname
    }



    #$GroupList = (Get-ADGroup license_disabled -Properties members).Members | Get-ADUser | select-object samaccountname
    #$LOAList = get-aduser -filter * -searchbase "OU=Leave Of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu" |Select-Object UserPrincipalName
    Write-Verbose "Email Addresses Gathered"

    #$GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $userlist -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from $($GroupID)"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to $($GroupID)"

                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}