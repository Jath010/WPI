function Sync-WPIVPNStaff {
    [cmdletbinding()]
    param (
    )
    $groupid = "VPN-Staff"
      
    $UsersData = Get-ADUser -Filter {ExtensionAttribute7 -eq "Staff" -and Enabled -eq $True} -Properties ExtensionAttribute7 -ResultSetSize 50000

    $GroupList = (Get-ADGroup $groupID -Properties members).Members | Get-ADUser | select-object samaccountname
 
    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $UsersData -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"
 
    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from $GroupID"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to $GroupID"
 
                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
    
}

function Sync-WPIVPNAff {
    [cmdletbinding()]
    param (
    )
    $groupid = "VPN-Aff"
      
    $UsersData = Get-ADUser -Filter {ExtensionAttribute7 -eq "Affiliate" -and Enabled -eq $True} -Properties ExtensionAttribute7 -ResultSetSize 50000

    $GroupList = (Get-ADGroup $groupID -Properties members).Members | Get-ADUser | select-object samaccountname
 
    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $UsersData -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"
 
    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from $GroupID"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to $GroupID"
 
                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
    
}


function Sync-WPIVPNFaculty {
    [cmdletbinding()]
    param (
    )
    $groupid = "VPN-Faculty"
      
    $UsersData = Get-ADUser -Filter {ExtensionAttribute7 -eq "Faculty" -and Enabled -eq $True} -Properties ExtensionAttribute7 -ResultSetSize 50000

    $GroupList = (Get-ADGroup $groupID -Properties members).Members | Get-ADUser | select-object samaccountname
 
    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $UsersData -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"
 
    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from $GroupID"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to $GroupID"
 
                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
    
}

function Sync-WPIVPNGraduate {
    [cmdletbinding()]
    param (
    )
    $groupid = "VPN-Graduate"
      
    $UsersData = Get-ADUser -Filter {ExtensionAttribute7 -eq "Student" -and ExtensionAttribute3 -ne "Freshman" -and ExtensionAttribute3 -ne "Sophomore" -and ExtensionAttribute3 -ne "Junior" -and ExtensionAttribute3 -ne "Senior" -and Enabled -eq $True} -Properties ExtensionAttribute7,ExtensionAttribute3 -ResultSetSize 50000

    $GroupList = (Get-ADGroup $groupID -Properties members).Members | Get-ADUser | select-object samaccountname
 
    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $UsersData -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"
 
    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from $GroupID"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to $GroupID"
 
                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
    
}

Sync-WPIVPNStaff
Sync-WPIVPNFaculty
Sync-WPIVPNGraduate
Sync-WPIVPNAff