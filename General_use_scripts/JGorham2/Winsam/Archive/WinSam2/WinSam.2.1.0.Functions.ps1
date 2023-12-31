function WinSam-Get-AccessLevel {
    <#
    .SYNOPSIS
    Get the access level for a user of Windows Samaritan Tool
    .DESCRIPTION
    Get the access level for a user of Windows Samaritan Tool.  Validates the user's group membership to identify what access they should get.
    .EXAMPLE
    WinSam-Get-AccessLevel username
    .PARAMETER username
    The username to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateLength(1,16)]
        [string]$username
        )
    #$ErrorActionPreference = "SilentlyContinue"
    $return = $null
    $usergroups = $null

    $usergroups = Get-ADPrincipalGroupMembership $username -Server $Global:DCServerName
    switch -wildcard ($usergroups) {
        "*CN=Hosting Services,*" {$return = 'SysAdmin'; break}
        "*CN=Windows Team,*" {$return = 'SysAdmin'; break}
        "*CN=G_WPI_Account_Maintenance,*" {$return = 'PasswordReset'; break}
        "*CN=U_WPI_Account_Maintenance_Unlock,*" {$return = 'Unlock'; break}
        "*CN=U_Helpdesk_Student_Staff,*" {$return = 'Unlock'; break}
        "*CN=U_WPI_Account_Maintenance_RO,*" {$return = 'ReadOnly'; break}
        }
    if (!$return) {$return="NoAccess"}
    $return
    # End of WinSam-Get-AccessLevel
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    

function WinSam-Get-LastLogonDate {
    <#
    .SYNOPSIS
    This function will survey all of the 2008 R2 Domain Controllers to find the most recent logon date.
    .DESCRIPTION
    This function will survey all of the 2008 R2 Domain Controllers to find the most recent logon date.
    .EXAMPLE
    WinSam-Get-LastLogonDate username
    .PARAMETER username
    The username to query. Just one.
    #>

    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateLength(1,16)]
        [string]$username
        )
    $dcs = $null
    $LastLogon = $null
    $LastLogonDate = $null
    $LastLogonServer = $null
   
    [hashtable]$Return = @{} 
    $DCs = (Get-ADDomainController -Filter {OperatingSystem -eq "Windows Server 2008 R2 Enterprise"} | Where {$_.Name -ne 'NEBULA'})
    foreach ($dc in $dcs) {
        $dclogon = $null
        $dclogon = (Get-ADUser $username -Properties LastLogon -Server $dc.hostname).LastLogon

        if ($DCLOGON -ne $null) {
            if ($LastLogon -lt $DCLOGON) {
                $LastLogon = $DCLOGON
                $LastLogonServer = $dc.name
                }
            }
        }

    if ($lastlogon -eq '0' -or $lastlogon -eq $null) {$lastlogondate = '';$LastLogonServer = ''}
    else {$lastlogondate = [DateTime]::FromFileTime($lastlogon)}
    
    $Return.samAccountName = $username
    $Return.LastLogon = $LastLogonDate
    $Return.LogonServer = $LastLogonServer

    Return $Return
    #End of WinSam-Get-LastLogonDate
    }   

#******************************************************************************************************************************
#******************************************************************************************************************************    

function WinSam-Get-PasswordExpiration {
    <#
    .SYNOPSIS
    This function will get the current status of an account password.
    .DESCRIPTION
    This function will get the current status of an account password.
    .EXAMPLE
    WinSam-Get-PasswordExpiration username
    .PARAMETER username
    The username to query. Just one.
    #>

    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateLength(1,16)]
        [string]$username
        )
    $ErrorActionPreference = "SilentlyContinue"
    $ADInfo = $null
    $MaxPasswordAge = $null
    $PasswordLastSet = $null
    $PasswordExpires = $null

    $ADInfo = Get-ADUser $username -Properties * -Server $Global:DCServerName
    if (!$ADInfo) {
        Write-Host "       WARNING : User does not exist" -ForegroundColor Yellow
        return
        }
    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -Server $Global:DCServerName| Select MaxPasswordAge).MaxPasswordAge
    $PasswordLastSet = $ADInfo.PasswordLastSet

    if ($PasswordLastSet) {$PasswordExpires = $PasswordLastSet.adddays($MaxPasswordAge.days)}
    else {$PasswordExpires = 0}
    
    if ($ADInfo.PasswordNeverExpires -eq 'True') {
        Write-Host "Password last set   : $PasswordLastSet"
        Write-Host "Password expiration : The password for $username is set to never expire" -ForegroundColor Red
        }
    elseif ($ADInfo.PasswordExpired -eq 'True' -and $PasswordExpires -ne '0') {
        Write-Host "Password last set   : $PasswordLastSet"
        Write-Host "Password expiration : The password for $username expired on $PasswordExpires." -ForegroundColor Red
        }
    elseif ($ADInfo.PasswordExpired -eq 'True' -and $PasswordExpires -eq '0') {
        Write-Host "Password last set   : Never Set" -ForegroundColor Red
        Write-Host 'Password expiration : This account has the "User must change password at next logon" checkbox enabled.' -ForegroundColor Red
        }
    else {
        Write-Host "Password last set   : $PasswordLastSet"
        Write-Host "Password expiration : $PasswordExpires"
        }

    #End WinSam-Get-PasswordExpiration Function
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    

function WinSam-Get-MailboxStats {
    <#
    .SYNOPSIS
    Get Mailbox Statistics for Info and WinSam
    .DESCRIPTION
    Get Mailbox Statistics for Info and WinSam
    .EXAMPLE
    WinSam-Get-MailboxStats username
    .PARAMETER username
    The username to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateLength(1,16)]
        [string]$username
        )
    $ErrorActionPreference = "SilentlyContinue"
    
    $mailbox = $null
    $mailboxStats = $null
    $mailboxSize = $null
    $MailboxQuota = $null
    $MailboxPercentUse = $null
    $MailboxStorageLimitStatus = $null

    $mailbox = Get-Mailbox $username -DomainController $Global:DCServerName
    if (!$mailbox){return}
    $MailboxStats = Get-MailboxStatistics $username -DomainController $Global:DCServerName
    if (!$mailboxstats){return}

    $MemberType = (($MailboxStats| Get-Member |Where {$_.Name -eq "TotalItemSize"}).TypeName).Split(".")
    If ($MemberType[0] -eq "Deserialized") {
        $MailboxSize = "{0:N2}" -f (($mailboxstats.TotalItemSize.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
        $DefaultQuota = "{0:N2}" -f (((Get-MailboxDatabase $mailbox.Database -DomainController $Global:DCServerName).ProhibitSendQuota.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
        $MailboxQuota = "{0:N2}" -f (($mailbox.ProhibitSendQuota.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
        }
    Else {
        $MailboxSize = "{0:N2}" -f ($mailboxstats.TotalItemSize.Value.Tobytes()/1gb)
        $DefaultQuota = "{0:N2}" -f ((Get-MailboxDatabase $mailbox.Database -DomainController $Global:DCServerName).ProhibitSendQuota.Value.Tobytes()/1gb)
        $MailboxQuota = "{0:N2}" -f ($mailbox.ProhibitSendQuota.Value.Tobytes()/1gb)
        }

    If ($mailbox.UseDatabaseQuotaDefaults -eq $true) {
        $MailboxQuota = $DefaultQuota
        }

    $MailboxPercentUse = "{0:P0}" -f($MailboxSize/$MailboxQuota)
    $MailboxStorageLimitStatus = $mailboxstats.StorageLimitStatus
    Switch ($MailboxStorageLimitStatus) {
        "ProhibitSend"    {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Black -BackgroundColor Red}
        "MailboxDisabled" {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Black -BackgroundColor Red}
        "IssueWarning"    {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Yellow -BackgroundColor Black}
        default           {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)"}
        }
    #End of WinSam-Get-MailboxStats
    }
    
#******************************************************************************************************************************
#******************************************************************************************************************************    

function WinSam-Get-AccountStatus {
    $Global:AccountStatus = $null
    $Global:AccountType = $null
    if ($Global:AccountEnabled -eq $false -and $Global:AccountExpirationDate -ne $null) {$Global:AccountStatus = "Disabled-Expired"}
    elseif ($Global:AccountEnabled -eq $false -and $Global:AccountExpirationDate -eq $null) {$Global:AccountStatus = "Disabled-Nonexpired"}
    elseif ($Global:AccountNEStatus) {$Global:AccountStatus = "NonEmployee"}
    else {$Global:AccountStatus = "Good"}

    Switch -wildcard ($global:ADinfo.canonicalname) {
        'admin.wpi.edu/Accounts/Students/*' {$Global:AccountType = "Student"}
        'admin.wpi.edu/Accounts/Work Study/*' {$Global:AccountType = "WorkStudy"}
        'admin.wpi.edu/Accounts/Retirees/*' {$Global:AccountType = "Retiree"}
        'admin.wpi.edu/Accounts/Vokes/*' {$Global:AccountType = "Contractor"}
        'admin.wpi.edu/Accounts/Employees/*' {$Global:AccountType = "Employee"}
        default {$Global:AccountType = "Other"}
        }
    #End of WinSam-Get-AccountStatus
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    

function WinSam-Get-InfoBanner {
    Switch ($Global:AccountStatus) {
        "Disabled-Expired"    {
            Write-Host ''
            $StatusText = "This account was disabled (terminated) on $global:AccountExpirationDate"
            While ($StatusText.Length -lt 85) {
                $StatusText += " "
                }
            Write-Host $StatusText -Foregroundcolor Black -BackgroundColor Red
            }
        "Disabled-Nonexpired" {Write-Host "This account is disabled (terminated). There is no expiration set for this account.  " -Foregroundcolor Black -BackgroundColor Red}
        "NonEmployee" {
            Write-Host ''
            Write-Host 'This user is flagged as a Non-Employee (NE) and is restricted from using the terminal' -Foregroundcolor Black -BackgroundColor Yellow
            Write-Host 'server and from accessing CLA Media.                                                 ' -Foregroundcolor Black -BackgroundColor Yellow
            }
        "Good" {
            Switch ($Global:AccountType) {
                "Student"    {Write-Host '     This is a Student Account.                                                      ' -Foregroundcolor Black -BackgroundColor Green}
                "WorkStudy"  {Write-Host '     This is a Student Workstudy Account.  Access is limited to specific computers.  ' -Foregroundcolor Black -BackgroundColor Yellow}
                "Retiree"    {Write-Host '     This is a Retired Employee, access is limited to OWA ONLY.                      ' -Foregroundcolor Black -BackgroundColor Yellow}
                "Contractor" {
                    Write-Host 'This is a limited contractor account, access is restricted to specific services only ' -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host '                                                                                     ' -Foregroundcolor Black -BackgroundColor Yellow
                    }
                "Employee"   {}
                "Other"      {}
                }
            }
        }
    #End of WinSam-Get-InfoBanner
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    


function WinSam-Get-LogonWorkstations {
    if ($Global:ADInfo.LogonWorkstations){
        $LogonWorkstations = ($Global:ADInfo.LogonWorkstations).Split(",")
        $Workstations=@()
        $temp="","",""
        $count=0
        $int=1
        foreach ($LogonWorkstation in $LogonWorkstations) {
            If ($count -eq 3) {$count=0}
            $temp[$count] = $LogonWorkstation
            If ($count -eq 2 -or $int -eq $LogonWorkstations.Count) {
            	$out = New-Object PSObject
            	$out | add-member noteproperty Column1 $temp[0]
            	$out | add-member noteproperty Column2 $temp[1]
            	$out | add-member noteproperty Column3 $temp[2]
            	$Workstations += $out
                $temp="","",""
                }
            $count++
            $int++
            }
        $Workstations
        }
    #End of WinSam-Get-LogonWorkstations
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    

function Write-ColorOutput($ForegroundColor)
{
    #Code from: http://stackoverflow.com/questions/4647756/is-there-a-way-to-specify-a-font-color-when-using-write-output
    # save the current color
    $fc = $host.UI.RawUI.ForegroundColor

    # set the new color
    $host.UI.RawUI.ForegroundColor = $ForegroundColor

    # output
    if ($args) {
        Write-Output $args
    }
    else {
        $input | Write-Output
    }

    # restore the original color
    $host.UI.RawUI.ForegroundColor = $fc
}