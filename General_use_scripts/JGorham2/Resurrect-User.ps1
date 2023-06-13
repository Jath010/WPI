Import-Module SqlServer
Import-Module AzureAD

Function Restore-WPIUser ($username) {
    $StatusReply = $null
    $StudentStatus = $null
    $ADInfo = $null
    $SQLCheck = $null

    While (!$username) {
        Write-Host ''
        Write-Host "No username was entered" -ForegroundColor Red
        Write-Host ''
        Read-Host "Please enter a username"
    }

    ## Load Libraries
    if (!(Get-Module | Where-Object { $_.Name -match "ActiveDirectory" })) { Import-Module ActiveDirectory }
    If ($null -eq (Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ps.outlook.com" } -ErrorAction SilentlyContinue)) {
        $CloudCredential = Get-Credential -Credential "$($env:username)@wpi.edu"
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $CloudCredential -Authentication Basic -AllowRedirection
        Import-PSSession $ExchangeSession -Prefix Cloud
    }
    if (!(Get-MsolDomain -ErrorAction SilentlyContinue)) {
        if (!$CloudCredential) { $CloudCredential = Get-Credential -Credential "$($env:username)@wpi.edu" }
        Connect-AzureAD -Credential $CloudCredential
    }




    ## Test for Script Pre-Requisites
    if (Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ps.outlook.com" }) { $ExchangeOnlineCheck = $true }
    if (Get-Module | Where-Object { $_.Name -match "ActiveDirectory" }) { $ActiveDirectoryCheck = $true }
    if (Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query "SELECT TOP 10 * FROM Logins" -ErrorAction SilentlyContinue) { $SQLCheck = $true }
    if (Get-AzureADDomain -ErrorAction SilentlyContinue) { $MSOLCheck = $true }

    if (!$ExchangeOnlineCheck -or !$ActiveDirectoryCheck -or !$SQLCheck -or !$MSOLCheck) {
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "   One or more pre-requisites for this script is not available" -ForegroundColor Yellow
        write-host ""
        Write-Host "   Either run the script form a location that has the necessary" -ForegroundColor Yellow
        Write-Host "   tools and libraries, or update your computer." -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        if ($SQLCheck) { $StudentStatusColor = "Green" }else { $StudentStatusColor = "Red" }
        Write-Host "     SQL Status              : $SQLCheck" -ForegroundColor $StudentStatusColor
        if ($ExchangeOnlineCheck) { $StudentStatusColor = "Green" }else { $StudentStatusColor = "Red" }
        Write-Host "     Exchange Online Status  : $ExchangeOnlineCheck" -ForegroundColor $StudentStatusColor
        if ($ActiveDirectoryCheck) { $StudentStatusColor = "Green" }else { $StudentStatusColor = "Red" }
        Write-Host "     Active Directory Status : $ActiveDirectoryCheck" -ForegroundColor $StudentStatusColor
        if ($MSOLCheck) { $StudentStatusColor = "Green" }else { $StudentStatusColor = "Red" }
        Write-Host "     Azure AD Status         : $MSOLCheck" -ForegroundColor $StudentStatusColor
        Write-Host ""
        Write-Host "                                                             " -ForegroundColor Black -BackgroundColor Red
        Write-Host "   This script has terminated.  No changes have been made.   " -ForegroundColor Black -BackgroundColor Red
        Write-Host "                                                             " -ForegroundColor Black -BackgroundColor Red
        Write-Host ""
        break
    }
    else {
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host "   Pre-requisite check has passed" -ForegroundColor Green
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host "     SQL Status              : $SQLCheck" -ForegroundColor Green
        Write-Host "     Exchange Online Status  : $ExchangeOnlineCheck" -ForegroundColor Green
        Write-Host "     Active Directory Status : $ActiveDirectoryCheck" -ForegroundColor Green
        Write-Host "     Azure AD Status         : $MSOLCheck" -ForegroundColor Green
        write-host ""
    }

    try { $ADInfo = Get-ADUser $username -Properties DisplayName, DistinguishedName, UserPrincipalName, Title, Department, Description, EmployeeID, EmployeeNumber, WhenCreated }
    catch { }

    If (!$ADInfo) {
        Write-Host ''
        Write-Host "No user was found: $username" -ForegroundColor Red
        Write-Host 'Exiting script'
        break
    }


    ## Print user information
    write-host 'Name                :' $ADInfo.DisplayName
    write-host 'Distinguished Name  :' $ADInfo.DistinguishedName
    Write-Host ''
    write-host 'Email               :' $ADInfo.UserPrincipalName
    Write-Host ''
    write-host 'Title               :' $ADInfo.Title
    write-host 'Department          :' $ADInfo.Department
    write-host 'Description         :' $ADInfo.description
    write-host ''
    write-host 'WPI ID              :' $ADInfo.EmployeeID
    write-host 'PIDM                :' $ADInfo.EmployeeNumber
    write-host ''

    ## Show user in GAL
    Set-ADUser $username -Replace @{msExchHideFromAddressLists = "FALSE" }

    ## Process Exchange Forward
    $forward = (Get-CloudMailbox $username).ForwardingSmtpAddress
    if ($forward) {
        Write-Host ''
        Write-Host "The mailbox for [$username] has a forward set to [$forward]."
        $ChangeForward = Read-Host "Do you wish to remove the forward? (y/n)"
        if ($ChangeForward -eq "y") { Set-CloudMailbox $username -ForwardingSmtpAddress $null }
    }

    ## Clear account expiration and re-enable account
    Set-ADUser $username -Enabled $true
    Get-ADUser $username | Clear-ADAccountExpiration

    ## Remove account from "Deny Logon Interactively" group
    if ((Get-ADPrincipalGroupMembership $username) -match "Deny Logon Interactively") {
        Remove-ADGroupMember "Deny Logon Interactively" $username -ErrorAction SilentlyContinue -Confirm:$false
    }

    If (!$ADInfo.EmployeeID) {
        $id = Read-Host "Please specify the ID Number"
        Set-ADUser $username -EmployeeID $id
    }
    If (!$ADInfo.EmployeeNumber) {
        $pidm = Read-Host "Please specify the PIDM"
        Set-ADUser $username -EmployeeNumber $pidm
    }

    ## Validate user type and move account to correct OU
    While ($StatusReply -ne 'y' -and $StatusReply -ne 'n') {
        $StatusReply = (Read-Host "Is this a student? (y/n)").ToLower()
        if ($StatusReply -eq 'y') { $StudentStatus = $true }
    }

    If ($StudentStatus) { $TargetPath = "ou=Students,ou=Accounts,dc=admin,dc=wpi,dc=edu" }
    Else { $TargetPath = "ou=Employees,ou=Accounts,dc=admin,dc=wpi,dc=edu" }

    Move-ADObject -Identity $ADInfo.DistinguishedName -TargetPath $TargetPath

    ## Unlock the account if it is locked
    $lockout = (Get-ADUser $username -Properties accountlockouttime | Select-Object accountlockouttime).accountlockouttime
    if ($lockout) { Unlock-ADAccount -Identity $username }

    ## Add the user from Exodus.
    if ($SQLCheck) {
        $record = $null

        $date = Get-Date -f M/d/yyyy
        $time = Get-Date -f HH:mm:ss
        $localhost = $env:COMPUTERNAME
        $operator = $env:USERNAME

        $record = Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
            "SELECT * FROM Logins WHERE Login = '$username'"

        if (!$record) {
            Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
                "INSERT INTO Logins VALUES ('$PIDM','$ID','$Date','$username')"

            Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
                "INSERT INTO Log VALUES ('$date','$time','$localhost','Powershell','Powershell','','$ID','$PIDM','$username','User $username successfully ressurrected by $operator')"
        }
    }
    c:

    ## Update MSOL Licenses

    ## Declare Licenses for application and/or update
    $Lic_Faculty_Education = 'wpi0:STANDARDWOFFPACK_FACULTY'
    $Lic_Faculty_ProPlus = 'wpi0:OFFICESUBSCRIPTION_FACULTY'
    $Lic_Student_Education = 'wpi0:STANDARDWOFFPACK_STUDENT'
    $Lic_Student_EducationPlus = 'wpi0:STANDARDWOFFPACK_IW_STUDENT'
    $Lic_Alumni_ExchangeStandard = 'wpi0:EXCHANGE_STANDARD_ALUMNI'

    $LO_Student_EducationPlus = New-MsolLicenseOptions -AccountSkuId $Lic_Student_EducationPlus -DisabledPlans 'SCHOOL_DATA_SYNC_P1', 'STREAM_O365_E3', 'TEAMS1', 'Deskless', 'FLOW_O365_P2', 'POWERAPPS_O365_P2', 'RMS_S_ENTERPRISE', 'OFFICE_FORMS_PLAN_2', 'PROJECTWORKMANAGEMENT', 'SWAY', 'YAMMER_EDU', 'SHAREPOINTWAC_EDU', 'SHAREPOINTSTANDARD_EDU'
    #$LO_Student_Education = New-MsolLicenseOptions -AccountSkuId $Lic_Student_Education -DisabledPlans 'SCHOOL_DATA_SYNC_P1', 'STREAM_O365_E3', 'TEAMS1', 'Deskless', 'FLOW_O365_P2', 'POWERAPPS_O365_P2', 'RMS_S_ENTERPRISE', 'OFFICE_FORMS_PLAN_2', 'PROJECTWORKMANAGEMENT', 'SWAY', 'YAMMER_EDU', 'SHAREPOINTWAC_EDU', 'SHAREPOINTSTANDARD_EDU'
    $LO_Faculty_Education = New-MsolLicenseOptions -AccountSkuId $Lic_Faculty_Education -DisabledPlans 'SCHOOL_DATA_SYNC_P1', 'STREAM_O365_E3', 'TEAMS1', 'Deskless', 'FLOW_O365_P2', 'POWERAPPS_O365_P2', 'RMS_S_ENTERPRISE', 'OFFICE_FORMS_PLAN_2', 'PROJECTWORKMANAGEMENT', 'SWAY', 'YAMMER_EDU', 'SHAREPOINTWAC_EDU', 'SHAREPOINTSTANDARD_EDU'
    #$LO_Faculty_ProPlus = New-MsolLicenseOptions -AccountSkuId $Lic_Faculty_ProPlus -DisabledPlans 'SHAREPOINTWAC_EDU', 'OFFICE_FORMS_PLAN_2', 'SWAY', 'ONEDRIVESTANDARD'
    #$LO_Other_Education_Student_Other = New-MsolLicenseOptions -AccountSkuId $Lic_Student_Education -DisabledPlans 'EXCHANGE_S_STANDARD', 'SCHOOL_DATA_SYNC_P1', 'STREAM_O365_E3', 'TEAMS1', 'Deskless', 'FLOW_O365_P2', 'POWERAPPS_O365_P2', 'RMS_S_ENTERPRISE', 'OFFICE_FORMS_PLAN_2', 'PROJECTWORKMANAGEMENT', 'SWAY', 'INTUNE_O365', 'YAMMER_EDU', 'SHAREPOINTWAC_EDU', 'SHAREPOINTSTANDARD_EDU'


    ## Set variables for MSOL info
    $UPN = $null; $UserLicenses = $null

    $UPN = $ADInfo.UserPrincipalName
    $MSOLUser = Get-AzureADUser -UserPrincipalName $UPN
    $UserLicenses = $MSOLUser.Licenses

    if ($StudentStatus) {
        if (!($UserLicenses.AccountSkuID -match $Lic_Student_EducationPlus)) {
            Write-Host "Processing [$username] Adding Student license" -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $Lic_Student_EducationPlus -LicenseOptions $LO_Student_EducationPlus
        }
    }
    else {
        if (!($UserLicenses.AccountSkuID -match $Lic_Faculty_Education)) {
            Write-Host "Processing [$username] Adding Faculty license" -ForegroundColor Green
            Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $Lic_Faculty_Education -LicenseOptions $LO_Faculty_Education
        }
    }

    foreach ($LicenseItem in $UserLicenses) {
        $License = $null
        $License = $LicenseItem.AccountSkuID

        switch ($License) {
            $Lic_Faculty_Education {
                if ($StudentStatus) {
                    Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Faculty_Education
                }
            }
            $Lic_Faculty_ProPlus {
                if ($StudentStatus) {
                    Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Faculty_ProPlus
                }
            }
            $Lic_Faculty_Project {
                if ($StudentStatus) {
                    Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Faculty_Project
                }
            }

            $Lic_Student_Education {
                if (!$StudentStatus) {
                    Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Student_Education
                }
            }
            $Lic_Student_EducationPlus {
                if (!$StudentStatus) {
                    Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Student_EducationPlus
                }
            }
            $Lic_Alumni_ExchangeStandard {
                Write-Host "Processing [$username] Removing $License" -ForegroundColor Red
                Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $Lic_Alumni_ExchangeStandard
            }
            default {
                Write-Host "[$username] still has $License" -ForegroundColor Yellow
            }
        }
    }
}
