import-module D:\wpi\powershell\LicenseAssignment\BannerSearch.ps1 -Force
Import-Module ActiveDirectory

#Creds for automagic login on ScriptHost-02
$Credentials = $null
if($env:COMPUTERNAME -eq "SCRIPTHOST-02"){
    $Credentials = Import-Clixml -Path 'D:\wpi\powershell\LicenseAssignment\exch_autoCred.xml'
    Connect-AzureAD -Credential $Credentials
}
#Add-AzureADGroupMember -ObjectId (Get-AzureADGroup -SearchString license_staff).objectid -RefObjectId (Get-AzureADUser -ObjectId jmgorham2@wpi.edu).objectid

###############################################################################################
#   Current Saviynt View
$SaviyntTable = "GWVSVNT_EXPANDED"
#   I Tossed this in because GWVSVNT_EXPANDED was a temporary table to be replaced with GWVSVNT
###############################################################################################

# TODO: This has no handling for Leave of Absence and the Alumni piece is a bit fiddly because of the number of unsynced alumni users.
function Sync-WPIStaffLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_staff).objectid

    Write-Verbose "Gathering Email List"
    $BannerList = Get-WPIEmployeeUsernames
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Staff"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectID
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.UserPrincipalName) to License_Staff"
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{
                    Write-Verbose "User is not in Azure"
                    continue
                }

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPIStudentLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_student).objectid

    Write-Verbose "Gathering Email List"
    #$BannerList = Get-WPIStudentUsernames
    $BannerList = (get-aduser -filter * -searchbase "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu")
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Student"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectID
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.UserPrincipalName) to License_Student"
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{
                    Write-Verbose "User is not in Azure"
                    continue
                }

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPIAlumniLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_Alumni).objectid

    Write-Verbose "Gathering Email List"
    $BannerList = Get-WPIAlumniUsernames
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Alumni"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId ($Difference.UserPrincipalName)).objectid
            }
            #In Banner but not Azure
            "<=" {
                #Alumni don't have unix addresses so we need to compose an email address
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{continue}
                Write-Verbose "Adding user $($Difference.UserPrincipalName)@wpi.edu to License_Alumni"

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPIStudentLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_student).objectid

    Write-Verbose "Gathering Email List"
    $BannerList = Get-WPIStudentUsernames
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Student"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectID
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.UserPrincipalName) to License_Student"
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{
                    Write-Verbose "User is not in Azure"
                    continue
                }

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPILoALicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_LOA).objectid

    Write-Verbose "Gathering Email List"
    $LOAList = get-aduser -filter * -searchbase "OU=Leave Of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu" |Select-Object UserPrincipalName
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $LOAList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_LOA"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId ($Difference.UserPrincipalName)).objectid
            }
            #In Banner but not Azure
            "<=" {
                #Alumni don't have unix addresses so we need to compose an email address
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{continue}
                Write-Verbose "Adding user $($Difference.UserPrincipalName)@wpi.edu to License_LOA"

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPIFacultyLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_Faculty).objectid

    Write-Verbose "Gathering Email List"
    $BannerList = Get-WPIFacultyUsernames
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Faculty"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectID
            }
            #In Banner but not Azure
            "<=" {
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{continue}
                Write-Verbose "Adding user $($Difference.UserPrincipalName) to License_Faculty"

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }

    <# foreach ($Faculty in $FacultyList) {
                Write-Verbose "Adding user $($Faculty.UserPrincipalName) to License_Faculty"
                $FacultyID = (Get-AzureADUser -ObjectId $Faculty.UserPrincipalName).objectid

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $FacultyID }
                catch {
                    Write-Verbose "User is already in group."
                }
            } #>
}

function Sync-WPIDisabledLicense {
    [cmdletbinding()]
    param (
    )
    $groupid = "license_disabled"
    $Disabledlist = get-aduser -Filter * -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" |Select-Object samaccountname
    $GroupList = (Get-ADGroup license_disabled -Properties members).Members | Get-ADUser | select-object samaccountname

    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $Disabledlist -DifferenceObject $GroupList -Property samaccountname
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.samaccountname) from License_Disabled"
                Remove-ADGroupMember -Identity $groupid -Members $difference.samaccountname -Confirm:$false
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.samaccountname) to License_Disabled"

                try { Add-ADGroupMember -Identity $GroupID -Members $difference.samaccountname -Confirm:$false}
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
    
}

function Sync-WPIExchangeLicences {
    [cmdletbinding()]
    param (
        
    )
    $GroupID = (Get-AzureADGroup -SearchString license_mailbox).objectid

    Write-Verbose "Gathering Email List"
    $BannerList = Invoke-WPIBannerQuery -Query "SELECT distinct CONCAT(USERNAME,'@WPI.EDU') as UserPrincipalName FROM $SaviyntTable WHERE IS_ACTIVE = '1'"
    Write-Verbose "Email Addresses Gathered"

    $GroupList = Get-AzureADGroupMember -ObjectId $GroupID -All $true


    #Comparison can use -property to specify what to check, but they need to use the same one
    $comparison = Compare-Object -ReferenceObject $BannerList -DifferenceObject $GroupList -Property UserPrincipalName
    Write-Verbose "Comparison Complete"

    foreach ($difference in $comparison) {
        switch ($difference.SideIndicator) {
            #In GroupList(Azure) but not in Bannerlist(Banner)
            "=>" {
                Write-Verbose "Removing user $($Difference.UserPrincipalName) from License_Mailbox"
                Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectID
            }
            #In Banner but not Azure
            "<=" {
                Write-Verbose "Adding user $($Difference.UserPrincipalName) to License_Mailbox"
                try{$UserID = (Get-AzureADUser -ObjectId $difference.UserPrincipalName).objectid}
                catch{
                    Write-Verbose "User doesn't exist in Azure"
                    continue  
                }

                try { Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UserID }
                catch {
                    Write-Verbose "User is already in group."
                }
            }
        }
    }
}

function Sync-WPILicenseGroups {
    [cmdletbinding()]
    param(
        [switch]
        $SkipAlum
    )
    Write-Verbose "Syncing Exchange Access"
    Sync-WPIExchangeLicences

    # Write-Verbose "Syncing Faculty"
    # Sync-WPIFacultyLicences

    if (!$SkipAlum) {
        Write-Verbose "Syncing Alumni"
        Sync-WPIAlumniLicences
    }

    # Write-Verbose "Syncing Students"
    # Sync-WPIStudentLicences

    # Write-Verbose "Syncing Staff"
    # Sync-WPIStaffLicences

    Write-Verbose "Syncing Leave of Absence"
    Sync-WPILoALicences

    Write-Verbose "Syncing Disabled"
    Sync-WPIDisabledLicense
}


Sync-WPILicenseGroups

#Comparitor work

<# Function comparison-testTrash {

            #Banner Section

            $StaffList = Get-WPIEmployeeUsernames
            $stafflist = $stafflist | where-object { $_.UserPrincipalName -ne "roger@wpi.edu" }

            #Azure Section

            $GroupID = (Get-AzureADGroup -SearchString license_staff).objectid
            $StaffGroup = Get-AzureADGroupMember -ObjectId $GroupID -All $true
            $staffgroup = $staffgroup | where-object { $_.UserPrincipalName -ne "mtaylor@wpi.edu" }


            #Comparison can use -property to specify what to check, but they need to use the same one
            $comparison = Compare-Object -ReferenceObject $StaffList -DifferenceObject $StaffGroup -Property UserPrincipalName
            #roger@wpi.edu is in Staffgroup but not stafflist : =>
            #mtaylor@WPI.EDU is in Stafflist but not $staffgroup : <=
            foreach ($difference in $comparison) {
                switch ($difference.SideIndicator) {
                    #In StaffGroup(Azure) but not in Stafflist(Banner)
                    "=>" {
                        write-host "Remove-AzureADGroupMember -ObjectId $GroupID -MemberId (Get-AzureADUser -ObjectId $($difference.UserPrincipalName)).objectID"
                    }
                    #In Banner but not Azure
                    "<=" {
                        Write-Host "Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId (Get-AzureADUser -ObjectId $($difference.UserPrincipalName)).objectID"
                    }
                }
            }
        } 
    }
}#>
#$BannerList | ForEach-Object {Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId ((Get-AzureADUser -ObjectId $_).objectid)}
<#
foreach($User in $BannerList){
    try{$UID = (Get-AzureADUser -ObjectId $User.USERPRINCIPALNAME).objectid}
    Catch{
        Write-Host $User.USERPRINCIPALNAME "doesn't seem to exist in Azure"
        continue
    }
    try{Add-AzureADGroupMember -ObjectId $GroupID -RefObjectId $UID}
    Catch{Write-Host $User.USERPRINCIPALNAME "Is already a member"}
}
#>