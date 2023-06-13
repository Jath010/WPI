Import-module SqlServer
Import-Module AzureAD

function Remove-WPIAccount {
    [CmdletBinding()]
    param (
        # Individual Accounts to destroy
        [Parameter(ValueFromPipeline = $true, ValueFromRemainingArguments = $true, ParameterSetName = "Pipe")]
        [String]
        $Identity,
        [Parameter(ParameterSetName = "CSV")]
        [String]
        $CSV
    )

    begin {
        $ErrorActionPreference = "SilentlyContinue"

        #Variable Declerations
        #$dc = $null
        #$dcs = $null
        $dchostname = $null
        $ExchangeCheck = $null
        $ExchangeOnlineCheck = $null
        $ActiveDirectoryCheck = $null
        $SQLCheck = $null
        $OldExchangeSessionPrefix = $null
        $out = $null
        $Global:userlist = @()
        $date = $null
        #$path = $null;
        $TerminatedUsers = $null
        $ADUsers = @();
        $AlumniUsers = @()
        $LOAUsers = @()
        $localhost = $env:COMPUTERNAME
        $operator = $env:USERNAME

        ## Get path and date
        $date = Get-Date
        #$path = (Get-Location).path

        #Set fileshare paths
        $FilePath_ScriptHome = "\\storage.wpi.edu\dept\Information Technology\CCC\Windows\fc_windows\account_removals\"
        $FilePath_ScriptHomeLogs = "$FilePath_ScriptHome\logs"

        #Set the path for specific OUs in the domain
        $OUPath_Disabled = "ou=Disabled,ou=Accounts,dc=admin,dc=wpi,dc=edu"    #This value needs to be in LDAP format

        #Import List of Terminated Users
        if ($null -eq $Identity -and $null -eq $CSV -and "$FilePath_ScriptHome\accountlist.txt" -ne "username`n") {
            $TerminatedUsers = Import-Csv "$FilePath_ScriptHome\accountlist.txt" | Where-Object { $_.Username -ne "" }
            $rename = "accountlist" + $date.ToString("MMddyyyy") + ".txt"
            Rename-Item "$FilePath_ScriptHome\accountlist.txt" -NewName $rename
            New-Item "$FilePath_ScriptHome\accountlist.txt" -Value "username`n"
        }
        if ($null -ne $CSV) {
            $TerminatedUsers = Import-Csv $CSV | Where-Object { $_.Username -ne "" }
        }
        if ($null -ne $Identity) {
            $TerminatedUsers = $Identity
        }

        #----------------------------------------------------------------------------
        #This code doesn't need to be changed
        #----------------------------------------------------------------------------

        ## Load Libraries
        if (!(Get-Module | Where-Object { $_.Name -match "ActiveDirectory" })) { Import-Module ActiveDirectory }
        if (!(Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "admin.wpi.edu" })) {
            if (!$Credentials) { $Credentials = Get-Credential -Credential "$($env:username)@wpi.edu" }
            $CASServer = $null; $ExchangeServers = 'EXCH-CAS-P-W01', 'EXCH-CAS-P-W02'
            foreach ($server in $ExchangeServers) { if (Test-Connection -Count 1 -BufferSize 15 -Delay 1 -ComputerName "$server.admin.wpi.edu") { $CASServer = "$server.admin.wpi.edu"; break } }
            if ($CASServer) { $ExchLocalSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$CASServer/PowerShell" -Credential $Credentials; Import-PSSession $ExchLocalSession -Prefix Local_ }
            else { Write-Verbose "No Exchange Server is available at this time.  Exchange Remote Shell cannot be loaded" -ForegroundColor Red }
        }
        If ((Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $null -eq $_.ComputerName -match "ps.outlook.com" } -ErrorAction SilentlyContinue)) {
            if (!$Credentials) { $Credentials = Get-Credential -Credential "$($env:username)@wpi.edu" }
            $ExchOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Credentials -Authentication Basic -AllowRedirection; Import-PSSession $ExchOnlineSession
        }
        If (!(Get-AzureADDomain)) { Connect-AzureAD -Credential $Credentials }

        ## Test for Script Pre-Requisites
        if (Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "admin.wpi.edu" }) { $ExchangeCheck = $true }
        if (Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ps.outlook.com" }) { $ExchangeOnlineCheck = $true }
        if (Get-Module | Where-Object { $_.Name -match "ActiveDirectory" }) { $ActiveDirectoryCheck = $true }
        if (Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query "SELECT TOP 10 * FROM Logins" -ErrorAction SilentlyContinue) { $SQLCheck = $true } ##Line requires module sqlserver
        if (Get-Command Get-CloudMailbox -ErrorAction SilentlyContinue) { $OldExchangeSessionPrefix = $true }

        if ($OldExchangeSessionPrefix) {
            Write-Verbose ""
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   Please update your profile to call Exchange without the   " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   'Cloud' prefix in order to use this script.               " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   This script has terminated.  No changes have been made.   " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose ""
            break
        }
        if (!$ExchangeCheck -or !$ExchangeOnlineCheck -or !$ActiveDirectoryCheck -or !$SQLCheck) {
            Write-Verbose "================================================================" -ForegroundColor Yellow
            Write-Verbose "   One or more pre-requisites for this script is not available" -ForegroundColor Yellow
            Write-Verbose ""
            Write-Verbose "   Either run the script form a location that has the necessary" -ForegroundColor Yellow
            Write-Verbose "   tools and libraries, or update your computer." -ForegroundColor Yellow
            Write-Verbose "================================================================" -ForegroundColor Yellow
            if ($SQLCheck) { $StatusColor = "Green" }else { $StatusColor = "Red" }
            Write-Verbose "     SQL Status              : $SQLCheck" -ForegroundColor $StatusColor
            if ($ExchangeCheck) { $StatusColor = "Green" }else { $StatusColor = "Red" }
            Write-Verbose "     Exchange Status         : $ExchangeCheck" -ForegroundColor $StatusColor
            if ($ExchangeOnlineCheck) { $StatusColor = "Green" }else { $StatusColor = "Red" }
            Write-Verbose "     Exchange Online Status  : $ExchangeOnlineCheck" -ForegroundColor $StatusColor
            if ($ActiveDirectoryCheck) { $StatusColor = "Green" }else { $StatusColor = "Red" }
            Write-Verbose "     Active Directory Status : $ActiveDirectoryCheck" -ForegroundColor $StatusColor
            Write-Verbose ""
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose "   This script has terminated.  No changes have been made.   " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose ""
            break
        }
        else {
            Write-Verbose "================================================================" -ForegroundColor Green
            Write-Verbose "   Pre-requisite check has passed" -ForegroundColor Green
            Write-Verbose "================================================================" -ForegroundColor Green
            Write-Verbose "     SQL Status              : $SQLCheck" -ForegroundColor Green
            Write-Verbose "     Exchange Status         : $ExchangeCheck" -ForegroundColor Green
            Write-Verbose "     Exchange Online Status  : $ExchangeOnlineCheck" -ForegroundColor Green
            Write-Verbose "     Active Directory Status : $ActiveDirectoryCheck" -ForegroundColor Green
            Write-Verbose ""
        }

        ## Get list of Domain Controller
        $dchostname = (Get-ADDomainController -Filter * | Where-Object { $_.OperationMasterRoles -like "*RIDMaster*" }).HostName

    }

    ########################################################################################
    #                      PROCESS BLOCK
    ########################################################################################

    process {

        ## Process Termination File
        foreach ($user in $TerminatedUsers) {
            $ADInfo = $null

            try { $ADInfo = Get-ADUser $user.Username -Properties Description, DisplayName, EmployeeNumber, EmployeeID, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12 -Server $dchostname }
            catch { }

            if ($ADInfo) {
                if ($ADInfo.extensionAttribute10 -match 'Alumni') { $AlumniUsers += $ADInfo; continue }             #If Alum skip to next user
                if ($ADInfo.extensionAttribute10 -match 'LeaveOfAbsence') { $LOAUsers += $ADInfo; continue }        #If LOA skip to next user
                $ADUsers += $ADInfo
            }
            else { Write-Verbose "There was a problem processing the entry '$($user.username)'." -ForegroundColor Red }
        }

        if ($ADUsers) {
            # Step 1 - Disable and move all accounts.  This is done as a first pass to quickly disable all accounts.  Additional cleanup is done later.  This is only an issue with larger terminations.
            Write-Verbose ""
            Write-Verbose "*** Processing Account Cleanup ***" -ForegroundColor Black -BackgroundColor White
            Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 1a: Setting account expiration for all userse on Termination List" -ForegroundColor Green
            # Set account expiration date,clearing the telephone number and disables the account
            $ADUsers | Set-ADUser -AccountExpirationDate $date -Clear HomeDirectory, HomeDrive, ScriptPath, Department, Title, telephonenumber, physicalDeliveryOfficeName -Enabled $false -Server $dchostname

            # Move all accounts to disabled OU
            Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 1b: Moving all disabled accounts to the Disabled OU" -ForegroundColor Green
            $ADUsers | Move-ADObject -TargetPath $OUPath_Disabled -Server $dchostname

            # Remove the user from Password Self Service Database (Exodus)
            foreach ($user in $ADUsers) {
                $username = $null; $displayName = $null

                $username = ($user.SamAccountName).ToUpper()
                $displayName = $user.DisplayName
                $ID = $user.EmployeeID
                $PIDM = $user.EmployeeNumber
                $SQLDate = $date.ToString('M/d/yyyy')
                $SQLTime = $date.ToString('HH:mm:ss')

                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 1c: Removing $username ($displayName) from SQL Tables" -ForegroundColor Green

                Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
                    "DELETE FROM Logins WHERE login = '$username'"

                Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
                    "INSERT INTO Log VALUES ('$SQLDate','$SQLTime','$localhost','Powershell','Powershell','','$ID','$PIDM','$username','User $username removed from EXODUS by $operator')"
            }


            # Step 2 - Clean Up AD Group Membership
            Write-Verbose ""
            Write-Verbose "*** Processing Group Cleanup ***" -ForegroundColor Black -BackgroundColor White
            foreach ($user in $ADUsers) {
                $username = $null; $displayName = $null
                $group = $null; $groups = $null; $grouplist = $null
                $AzureUser = $null; $AzureGroup = $null; $AzureGroups = $null; $AzureGroupList = $null

                $username = ($user.SamAccountName).ToUpper()
                $displayName = $user.DisplayName
                try { $AzureUser = Get-AzureADUser -ObjectId "$username@wpi.edu" }
                catch { }

                # Get List of Groups
                $groups = Get-ADPrincipalGroupMembership $username -Server $dchostname | Where-Object { $_.Name -ne "Domain Users" } | Sort-Object Name
                if ($AzureUser) { $AzureGroups = Get-AzureADUserMembership -ObjectId $AzureUser.ObjectId | Where-Object { $_.DirSyncEnabled -ne $true -and $_.DisplayName -ne 'WPI_All_Users_Dynamic' } }

                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 2a: Removing AD groups for $username ($displayName)" -ForegroundColor Green
                If ($groups) {
                    $grouplist = ($groups).Name -join ','
                    foreach ($group in $groups) { Remove-ADPrincipalGroupMembership $username $group.distinguishedName -Server $dchostname -Confirm:$false }
                }

                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 2a: Removing Azure groups for $username ($displayName)" -ForegroundColor Green
                if ($AzureGroups) {
                    $AzureGroupList = ($AzureGroups).DisplayName -join ','
                    foreach ($AzureGroup in $AzureGroups) {
                        $Recipient = $null
                        if ($AzureGroup.Mail) { $Recipient = Get-Recipient $AzureGroup.Mail -ErrorAction SilentlyContinue }
                        if ($Recipient.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') { Remove-DistributionGroupMember $Recipient.Name -Member $username -Confirm:$false }
                        else { Remove-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -MemberId $AzureUser.ObjectId }
                    }
                }

                $out = New-Object PSObject
                $out | add-member noteproperty Name $user.Name
                $out | add-member noteproperty SamAccountName $user.samAccountName
                $out | add-member noteproperty UPN $user.UserPrincipalName
                $out | add-member noteproperty Description $user.description
                $out | add-member noteproperty DistinguishedName $user.distinguishedname
                $out | add-member noteproperty GroupMembership $grouplist
                $out | add-member noteproperty AzureGroupMembership $AzureGroupList
                $Global:userlist += $out

                Write-Verbose "$((Get-Date).ToString('HH:mm'))  FINISH: Termination complete for $username ($displayName)"
                #End Import-CSV
            }

            # Step 3 - Clean up Exchange Mailboxes.
            Write-Verbose ""
            Write-Verbose "*** Processing Mailbox Cleanup ***" -ForegroundColor Black -BackgroundColor White
            foreach ($user in $ADUsers) {
                $username = $null; $displayName = $null

                $username = ($user.SamAccountName).ToUpper()
                $displayName = $user.DisplayName

                # Hide user from GAL
                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 3a: Hiding $username ($displayName) from GAL" -ForegroundColor Green
                Set-ADUser $username -Replace @{msExchHideFromAddressLists = "TRUE" }

                # Set mailbox to only allow messages from itself - This is necessary to have the mailbox bounce properly
                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 3b: Restricting access to receive messages to $username ($displayName) from GAL" -ForegroundColor Green
                $Mailbox = Get-Mailbox $username -ErrorAction SilentlyContinue
                if ($Mailbox) { Set-Mailbox $username -AcceptMessagesOnlyFrom $username }
                else { Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 3: ERROR - Mailbox does not exist for $username ($displayName)" -ForegroundColor Red }
            }

            # Step 4 - Update AzureAD Licenses
            Write-Verbose ""
            Write-Verbose "*** Update AzureAD Licenses ***" -ForegroundColor Black -BackgroundColor White
            $Lic_Alumni_ExchangeStandard = 'wpi0:EXCHANGE_STANDARD_ALUMNI'
            foreach ($user in $ADUsers) {
                $username = $null; $displayName = $null
                #$UPN = $null
                $AzureADUser = $null
                $LicensesForRemoval = $null
                $UserLicenses = $null

                $username = ($user.SamAccountName).ToUpper()
                $displayName = $user.DisplayName

                Write-Verbose "$((Get-Date).ToString('HH:mm'))  Step 4: Updating AzureAD Licenses for $username ($displayName)" -ForegroundColor Green
                #$UPN = $user.UserPrincipalName
                $AzureADUser = Get-AzureADUser -ObjectId $username -ErrorAction SilentlyContinue
                if (!$AzureADUser) { continue }
                $UserLicenses = $AzureADUser.Licenses
                if (!($UserLicenses.AccountSkuID -eq $Lic_Alumni_ExchangeStandard)) {
                    Write-Verbose "$((Get-Date).ToString('HH:mm'))  [GRANT ] $Lic_Alumni_ExchangeStandard for $username ($displayName)" -ForegroundColor Cyan
                    Set-AzureADUserLicense -ObjectId $username -AddLicenses $Lic_Alumni_ExchangeStandard
                }
                $LicensesForRemoval = $UserLicenses | Where-Object { $_.AccountSkuID -ne $Lic_Alumni_ExchangeStandard }
                if ($LicensesForRemoval) {
                    foreach ($LicenseItem in $LicensesForRemoval) {
                        $License = $null
                        $License = $LicenseItem.AccountSkuID
                        Write-Verbose "$((Get-Date).ToString('HH:mm'))  [REMOVE] $License for $username ($displayName)" -ForegroundColor Red
                        Set-AzureADUserLicense -ObjectId $username -RemoveLicenses $License
                    }
                }
            }


        }
        Else {
            Write-Verbose ""
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose "   There was a problem getting a list of users.              " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose "   This script has terminated.  No changes have been made.   " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose "                                                             " -ForegroundColor Black -BackgroundColor Red
            Write-Verbose ""
            #break
        }

        if ($AlumniUsers) {
            Write-Verbose ''
            Write-Verbose ''
            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   The following accounts were not terminated due to Alumni Status   " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   Please process the user(s) per the Alumni conversion script       " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow

            $AlumniUsers | Select-Object DisplayName, SamAccountName, extensionAttribute10 | Out-Default | Format-Table

            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose ''
            Write-Verbose ''
        }

        if ($LOAUsers) {
            Write-Verbose ''
            Write-Verbose ''
            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   The following accounts were not terminated due to LoA Status      " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "   Please process the user(s) per the LoA conversion script          " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow

            $LOAUsers | Select-Object DisplayName, SamAccountName, extensionAttribute10 | Out-Default | Format-Table

            Write-Verbose "                                                                     " -ForegroundColor Black -BackgroundColor Yellow
            Write-Verbose ''
            Write-Verbose ''
        }

        c:

        $Global:userlist | Export-Csv ("$FilePath_ScriptHomeLogs\output" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation

    }

    end {
    }
}
