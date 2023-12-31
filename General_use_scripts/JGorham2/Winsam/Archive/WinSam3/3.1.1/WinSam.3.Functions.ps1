function WinSam-Get-UserMailboxInfo {
    <#
    .SYNOPSIS
    Returns mailbox data
    .DESCRIPTION
    Returns mailbox data and will determin if the mailbox is a local or cloud based mailbox.
    Written by Tom Collins.  Last Updated 1/26/2012
    .EXAMPLE
    WinSam-Get-UserMailboxInfo username
    .PARAMETER username
    The username to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$username
        )
    $ErrorActionPreference = "SilentlyContinue"

    $Global:RecipientInfo = Get-Recipient $username -ErrorAction SilentlyContinue
    if ($Global:RecipientInfo) {
        if ($Global:RecipientInfo.RecipientTypeDetails -match 'Remote') {$Mailbox = Get-CloudMailbox $username}
        else {$Mailbox = Get-Mailbox $username -DomainController $dchostname}
        }
    return $Mailbox
    }

function WinSam-Forward-Exchange ($alias) {
    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Mailbox Forward Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host "Email               : $alias@wpi.edu"
    Write-Host ''
    Write-Host "Current Forward     : $($Global:mailbox.ForwardingSmtpAddress)"

    Set-CloudMailbox -Identity $alias -ForwardingSmtpAddress $null
    $Global:mailbox = WinSam-Get-UserMailboxInfo $username
    
    Write-Host ''
    Write-Host "The forward for [$alias] has been removed."
    Write-Host "This mailbox will now retain messages on Exchange"
    Write-Host ''
    
    Write-Host "Press any key to continue ..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    WinSam-Menu-User
    }

function WinSam-Forward-UNIX ($alias) {
    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Mailbox Forward Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host "Email               : $alias@wpi.edu"
    Write-Host ''
    Write-Host "Current Forward     : $($Global:mailbox.ForwardingSmtpAddress)"

    $smtp = $alias + "@smtp.wpi.edu"
    Set-CloudMailbox -Identity $alias -ForwardingSmtpAddress $smtp
    $Global:mailbox = WinSam-Get-UserMailboxInfo $username

    Write-Host ''
    Write-Host "The forward for [$alias] has been set to [$smtp]"
    Write-Host "This mailbox will now retain messages on Exchange"
    Write-Host ''
    
    Write-Host "Press any key to continue ..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    WinSam-Menu-User
    }

function WinSam-Forward-Manual ($alias) {
    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Mailbox Forward Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host "Email               : $alias@wpi.edu"
    Write-Host ''
    Write-Host "Current Forward     : $($Global:mailbox.ForwardingSmtpAddress)"
    Write-Host ''
    Write-Host ''
    $choice = $null
    while ($choice -ne 'y') {
        $smtp = Read-Host "Please specify the forwarding address"
        Write-Host ''
        Write-Host "Setting the forwarding for [$alias] to be [$smtp]."
        $choice = Read-Host "Is this correct? (y/n)"
        }
    Set-CloudMailbox -Identity $alias -ForwardingSmtpAddress $smtp
    $Global:mailbox = WinSam-Get-UserMailboxInfo $username

    Write-Host ''
    Write-Host "The forward for $alias has been set to: $smtp"
    Write-Host "This mailbox will now retain messages on Exchange"
    Write-Host ''
    
    Write-Host "Press any key to continue ..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    WinSam-Menu-User
    }


function WinSam-Get-ObjectInfo ($DisplayName) {
    [hashtable]$Return = @{}
    $ADInfo = $null;$ObjectType = $null;$username = $null
    $return.Name = $null
    $return.Username = $null
    $return.Department = $null
    $return.DistinguishedName = $null
    $return.Enabled = $null

    if ($DisplayName) {
        if ($DisplayName -eq 'Default' -or $DisplayName -eq 'Anonymous') {$return.Name = $DisplayName;$return.Enabled = $true}
        else {
            if ($DisplayName -match "NT User:") {$username = $DisplayName.Replace("NT User:ADMIN\",""); $ObjectType = 'User'}
            elseif ($DisplayName.StartsWith("S-")) {$username = $DisplayName}
            else {
                $Recipient = Get-Recipient $DisplayName -ErrorAction SilentlyContinue
                if ($Recipient) {
                    if ($Recipient.RecipientType -match "Group") {$username = (Get-DistributionGroup $DisplayName -DomainController $dchostname -ErrorAction SilentlyContinue).SamAccountName; $ObjectType = 'Group'}
                    Else {$username = (Get-Recipient $DisplayName -DomainController $dchostname -ErrorAction SilentlyContinue).SamAccountName;$ObjectType = 'User'}
                    }
                elseif (Get-DistributionGroup $DisplayName -ErrorAction SilentlyContinue) {$username = (Get-DistributionGroup $DisplayName -DomainController $dchostname -ErrorAction SilentlyContinue).SamAccountName; $ObjectType = 'Group'}
                }
            if (!$username -and $DisplayName -like "ADMIN\*") {$username = $DisplayName.Replace("ADMIN\",""); $ObjectType = 'Group'}
            if ($username) {
                $ErrorActionPreference = 'SilentlyContinue'
                if ($ObjectType -eq 'User' -or !$ObjectType) {
                    $ADInfo = Get-ADUser $username -Properties DisplayName,Department -Server $dchostname
                    if ($ADInfo) {
                        $return.Name = $ADInfo.DisplayName
                        $return.Username = $ADInfo.SamAccountName
                        $return.Department = $ADInfo.Department
                        $return.DistinguishedName = $ADInfo.DistinguishedName
                        $return.Enabled = $ADInfo.Enabled
                        }
                    }
                if ($ObjectType -eq 'Group' -or !$ADInfo) {
                    $GroupInfo = Get-ADGroup $username -Server $dchostname
                    if ($GroupInfo) {
                        $return.Name = $GroupInfo.Name
                        $return.Username = $GroupInfo.SamAccountName
                        $return.Department = 'Distribution Group'
                        $return.DistinguishedName = $GroupInfo.DistinguishedName
                        $return.Enabled = $true
                        }
                    }
                if (!$ADInfo -and !$GroupInfo) {$return.Name = $DisplayName;$return.Enabled = $false}
                $ErrorActionPreference = 'Continue'
                }
            else {$return.Name = $DisplayName;$return.Enabled = $true}
            }
        }
    else {Write-Host "Entry does not exist. Please report error to System Administrator." -ForegroundColor Red;$return.Name = 'ERROR';$return.Enabled = $true}
    Return $return
    }

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
    $ErrorActionPreference = "SilentlyContinue"
    $return = $null

    $Global:UserAccessLevel = $null

    #Global Access
        $Access_SysAdmin = $null;
        $Access_SysAdmin = Get-ADGroupMember hosting -Recursive -Server $dchostname
        if ($Access_SysAdmin | Where {$_.SamAccountName -eq $username}) {
            $Global:UserAccessLevel = 'SysAdmin'
            return
            }
    
    #User Access
        $Global:UserAccessLevel = $null
        $Access_PasswordSet = $null;$Access_PasswordReset = $null;$Access_Unlock = $null;$Access_ReadOnly = $null

        $Access_PasswordSet = Get-ADGroupMember  G_WPI_Account_Maintenance_Lvl2 -Recursive -Server $dchostname
        $Access_PasswordReset = Get-ADGroupMember  G_WPI_Account_Maintenance -Recursive -Server $dchostname
        $Access_Unlock = Get-ADGroupMember U_WPI_Account_Maintenance_Unlock -Recursive -Server $dchostname
        $Access_ReadOnly = Get-ADGroupMember U_WPI_Account_Maintenance_RO -Recursive -Server $dchostname

        if ($Access_PasswordSet  | Where {$_.SamAccountName -eq $username}) {$Global:UserAccessLevel = 'PasswordReset_lvl2'}
        elseif ($Access_PasswordReset  | Where {$_.SamAccountName -eq $username}) {$Global:UserAccessLevel = 'PasswordReset'}
        elseif ($Access_Unlock  | Where {$_.SamAccountName -eq $username}) {$Global:UserAccessLevel = 'Unlock'}
        elseif ($Access_ReadOnly  | Where {$_.SamAccountName -eq $username}) {$Global:UserAccessLevel = 'ReadOnly'}
        else {$Global:UserAccessLevel = 'NoAccess'}

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

    [hashtable]$Return = @{} 

    $ADInfo = Get-ADUser $username -Properties PasswordLastSet,PasswordExpired,PasswordNeverExpires -Server $dchostname
    if (!$ADInfo) {Write-Host "       WARNING : User does not exist" -ForegroundColor Yellow;return}
    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -Server $dchostname| Select MaxPasswordAge).MaxPasswordAge
    $PasswordLastSet = $ADInfo.PasswordLastSet

    if ($PasswordLastSet) {$PasswordExpires = $PasswordLastSet.adddays($MaxPasswordAge.days)}
    else {$PasswordExpires = 0}

    $Return.PasswordLastSet = $PasswordLastSet
    $Return.PasswordExpiration = $PasswordExpires
    $Return.PasswordExpired = $ADInfo.PasswordExpired
    $Return.PasswordNeverExpires = $ADInfo.PasswordNeverExpires

    Return $Return
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

    $RecipientInfo = Get-Recipient $username
    if ($RecipientInfo.RecipientTypeDetails -match 'Remote') {$Mailbox = Get-CloudMailbox $username}
    else {$Mailbox = Get-Mailbox $username -DomainController $dchostname}
    if (!$mailbox){return}
    if ($RecipientInfo.RecipientTypeDetails -match 'Remote') {$MailboxStats = Get-CloudMailboxStatistics $username}
    else {$MailboxStats = Get-MailboxStatistics $username -DomainController $dchostname}
    if (!$mailboxstats){return}

    $MemberType = (($MailboxStats| Get-Member |Where {$_.Name -eq "TotalItemSize"}).TypeName).Split(".")

    if ($RecipientInfo.RecipientTypeDetails -match 'Remote') {
        $MailboxSize = "{0:N2}" -f (($mailboxstats.TotalItemSize.ToString().Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
        if ($mailbox.ProhibitSendQuota = 'Unlimited') {$MailboxQuota = 50}
        else {$MailboxQuota = "{0:N2}" -f (($mailbox.ProhibitSendQuota.ToString().Split("("))[1].Split(" ")[0].Replace(",","")/1gb)}
        }
    else {
        If ($MemberType[0] -eq "Deserialized") {
            $MailboxSize = "{0:N2}" -f (($mailboxstats.TotalItemSize.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
            $DefaultQuota = "{0:N2}" -f (((Get-MailboxDatabase $mailbox.Database -DomainController $dchostname).ProhibitSendQuota.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
            $MailboxQuota = "{0:N2}" -f (($mailbox.ProhibitSendQuota.Split("("))[1].Split(" ")[0].Replace(",","")/1gb)
            }
        Else {
            $MailboxSize = "{0:N2}" -f ($mailboxstats.TotalItemSize.Value.Tobytes()/1gb)
            $DefaultQuota = "{0:N2}" -f ((Get-MailboxDatabase $mailbox.Database -DomainController $dchostname).ProhibitSendQuota.Value.Tobytes()/1gb)
            $MailboxQuota = "{0:N2}" -f ($mailbox.ProhibitSendQuota.Value.Tobytes()/1gb)
            }

        If ($mailbox.UseDatabaseQuotaDefaults -eq $true) {
            $MailboxQuota = $DefaultQuota
            }
        }

    $MailboxPercentUse = "{0:P0}" -f($MailboxSize/$MailboxQuota)
    $MailboxStorageLimitStatus = $mailboxstats.StorageLimitStatus
    Switch ($MailboxStorageLimitStatus) {
        "BelowLimit"      {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)"}
        "ProhibitSend"    {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Black -BackgroundColor Red}
        "MailboxDisabled" {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Black -BackgroundColor Red}
        "IssueWarning"    {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - $MailboxStorageLimitStatus ($MailboxPercentUse)" -ForegroundColor Yellow -BackgroundColor Black}
        $null             {Write-Host "Mailbox Quota       : $MailboxSize [$MailboxQuota] GB - ($MailboxPercentUse)"}
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
        'admin.wpi.edu/Accounts/Alumni/*'     {$Global:AccountType = "Alumni"}
        'admin.wpi.edu/Accounts/Disabled/*'   {$Global:AccountType = "Disabled"}
        'admin.wpi.edu/Accounts/Employees/*'  {$Global:AccountType = "Employee"}
        'admin.wpi.edu/Accounts/Retirees/*'   {$Global:AccountType = "Retiree"}
        'admin.wpi.edu/Accounts/Students/*'   {$Global:AccountType = "Student"}
        'admin.wpi.edu/Accounts/Vokes/*'      {$Global:AccountType = "Contractor"}
        'admin.wpi.edu/Accounts/Work Study/*' {$Global:AccountType = "WorkStudy"}
        'admin.wpi.edu/Accounts/Leave Of Absence/*' {$Global:AccountType = "LOA"}
        'admin.wpi.edu/Accounts/Other Accounts/Resource Mailboxes/*' {$Global:AccountType = "ResourceMailbox"}
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
            Switch ($Global:AccountType) {
                'Disabled' {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "This account was disabled on $global:AccountExpirationDate" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "NOTE: This account has been processed for termination" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host ''
                    }
                'ResourceMailbox' {Write-Host (WinSam-Write-Header 'This is account is configured as a Resource Mailbox.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow;Write-Host ''}
                default {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "This account was disabled on $global:AccountExpirationDate" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "Please contact InfoSec and Hosting for further information about this account" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host ''
                    }
                }
            }
        "Disabled-Nonexpired" {
            Switch ($Global:AccountType) {
                'Disabled' {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header 'This account is disabled. There is no expiration set for this account.' $MenuLength 2) -Foregroundcolor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "NOTE: This account has been processed for termination" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host ''
                    }
                'ResourceMailbox' {Write-Host (WinSam-Write-Header 'This is account is configured as a Resource Mailbox.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow;Write-Host ''}
                default {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header 'This account is disabled. There is no expiration set for this account.' $MenuLength 2) -Foregroundcolor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header "Please contact InfoSec and Hosting for further information about this account" $MenuLength 2) -ForegroundColor Black -BackgroundColor Red
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Red
                    Write-Host ''
                    }
                }
            }
        "NonEmployee" {
            Write-Host ''
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '  This user is flagged as a Non-Employee (NE) and is restricted from using the' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '  terminal server and from accessing CLA Media' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host ''
            }
        "Good" {
            Switch ($Global:AccountType) {
                'Employee'   {}
                'Student'    {Write-Host (WinSam-Write-Header 'This is a Student Account.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Green;Write-Host ''}
                'Alumni'     {
                    Write-Host (WinSam-Write-Header 'This is a Alumni Pilot Account.  Access is limited to Exchange Online access.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'This user is not able to log in directly to campus computers.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '  ----------' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'NOTE: If this user should be an active student, please have them recreate their' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '      accounts in BannerWeb and notify SysOps to move the domain account status' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '      back to a standard student.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host ''
                    }
                'LOA'        {
                    Write-Host (WinSam-Write-Header 'This student is on a Leave of Absence.  Access is limited to Exchange Online access.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'This user is not able to log in directly to campus computers.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '  ----------' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'NOTE: If this user should be an active student, please have them recreate their' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '      accounts in BannerWeb and notify SysOps to move the domain account status' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '      back to a standard student.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host ''
                    }
                'WorkStudy'  {Write-Host (WinSam-Write-Header 'This is a Student work study Account.  Access is limited to specific computers.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow;Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow;Write-Host ''}
                'Retiree'    {Write-Host (WinSam-Write-Header 'This is a Retired Employee, access is limited to OWA ONLY.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow;Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow;Write-Host ''}
                'Contractor' {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'This is a limited contractor account, access is restricted to specific services only' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host ''
                    }
                'Other'      {
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'This is not a general user account and may be used for special purposes.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header 'Please contact Hosting Services for more information about this account.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
                    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
                    Write-Host ''
                    }
                }
            }
        }
    #End of WinSam-Get-InfoBanner
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    


function WinSam-Get-LogonWorkstations {
    if ($Global:ADInfo.LogonWorkstations){
        $LogonWorkstations = ($Global:ADInfo.LogonWorkstations).Split(",") | Sort
        $Workstations=@()
        $temp="","","",""
        $count=0
        $int=1
        foreach ($LogonWorkstation in $LogonWorkstations) {
            If ($count -eq 4) {$count=0}
            $temp[$count] = $LogonWorkstation.ToUpper()
            If ($count -eq 3 -or $int -eq $LogonWorkstations.Count) {
            	$out = New-Object PSObject
            	$out | add-member noteproperty Column1 $temp[0]
            	$out | add-member noteproperty Column2 $temp[1]
            	$out | add-member noteproperty Column3 $temp[2]
                $out | add-member noteproperty Column4 $temp[3]
            	$Workstations += $out
                $temp="","","",""
                }
            $count++
            $int++
            }
        $Workstations | ft -HideTableHeaders
        }
    #End of WinSam-Get-LogonWorkstations
    }

#******************************************************************************************************************************
#******************************************************************************************************************************    

function Write-ColorOutput($ForegroundColor) {
    #Code from: http://stackoverflow.com/questions/4647756/is-there-a-way-to-specify-a-font-color-when-using-write-output
    # save the current color
    $fc = $host.UI.RawUI.ForegroundColor

    # set the new color
    $host.UI.RawUI.ForegroundColor = $ForegroundColor

    # output
    if ($args) {
        Write-Output $args | Out-Default
    }
    else {
        $input | Write-Output
    }

    # restore the original color
    $host.UI.RawUI.ForegroundColor = $fc
}

#******************************************************************************************************************************
#******************************************************************************************************************************    
function WinSam-Write-Header {
    <#
    .SYNOPSIS
    Takes a string and pads with spaces to expand the length to fit a specific format.
    .DESCRIPTION
    Takes a string and pads with spaces to expand the length to fit a specific format.  Written by Tom Collins.  Last Updated 08/12/2014
    .EXAMPLE
    WinSam-Write-Header string
    .PARAMETER String
    The string that you want adjusted.
    .PARAMETER MaxLength
    The total length of the new header string
    .PARAMETER Indent
    The length of space you want to add to the start of the string
    .PARAMETER Center
    Call this value if you want the text of String to be centered within the header.
    #>
    param (
        [Parameter(Mandatory=$False,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=1)]
        [string]$string,
        [Parameter(Mandatory=$True,Position=2)]
        [int]$MaxLength,
        [Parameter(Mandatory=$False,Position=3)]
        [int]$Indent,
        [switch]$Center,
        [switch]$Line
        )
    $newstring = $null

    if ($center) {
        if ($Indent) {$indent = [math]::floor(($MaxLength-($string.Length+$indent))/2)}
        else {$indent = [math]::floor(($MaxLength-$string.Length)/2)}
        }

    if ($Line) {$pad = '-'}
    else {$pad = ' '}
    $newstring = $string.PadLeft($indent+$string.Length,$pad).PadRight($MaxLength,$pad)
    Return $newstring
    }

function WinSam-Get-ADGroups {
    Clear-Host
    $Global:UserGroups = $ADInfo | Get-ADPrincipalGroupMembership -Server $dchostname | Get-ADGroup -Properties Name,Description,GroupCategory -Server $dchostname

    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Distribution Groups  (These groups are also Exchange Mail Groups)' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    $UserGroups | Where {$_.GroupCategory -eq 'Distribution'} | Select Name,Description | Sort Name | ft -AutoSize -Wrap | Out-Default

    write-host ''
    Write-Host (WinSam-Write-Header 'Security Groups' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    $UserGroups | Where {$_.GroupCategory -eq 'Security'} | Select Name,Description | Sort Name | ft -AutoSize -Wrap | Out-Default
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Menu-User
    #End WinSam-Get-ADGroups
    }

function WinSam-Create-ComplexPassword {
    # *** Unable to generate complex password less then 5 chars ***
    # ASCII data taken from http://msdn2.microsoft.com/en-us/library/60ecse8t(VS.80).aspx

    Param (
     [int]$PassLength
     )

    # Let’s work out where our 3 complex characters will be inserted in the password…
    [int]$mark = ($PassLength/3)
    $ComplexChar = @(“marker”)
    $ComplexChar[0] = $mark
    $ComplexChar = $ComplexChar+($ComplexChar[0] + $mark)
    $ComplexChar = $ComplexChar+(($ComplexChar[1] + $mark) -1)

    $Password = $null
    $rnd = new-object random

    # “i” is our counter while we make the password, one char at a time.
    $i = $Password.length
    do {
        switch ($Password.length) {
            $ComplexChar[0] {$password = $password+([char]($rnd.next(48,57)))  <#Make this character a Numeric#>}
            $ComplexChar[1] {$password = $password+([char]($rnd.next(65,90)))  <#Make this character a LowerAlpha#>}
            $ComplexChar[2] {$password = $password+([char]($rnd.next(97,122))) <#Make this character a Upper Alpha#>}
            default {
                <# In case this is used in a DCPromo answer files, theres a few chars to avoid: Ampersand, Less than, double quote and back slash#>
                $NextChar = $rnd.next(33,123)
                switch ($nextChar) {
                    34 {break}
                    38 {break}
                    60 {break}
                    92 {break}
                    default {$Password = $Password+([char]$nextChar)}
                    }
                }
            }
        $i++
        }
    Until ($Password.length -eq $PassLength)
    return $Password
    }

function WinSam-Reset-Password ($PasswordType) {
    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
    Write-Host ''

    $PasswordStatus = WinSam-Get-PasswordExpiration $username

    if ($PasswordStatus.PasswordNeverExpires -eq 'True') {
        Write-Host "Password last set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
        Write-Host "Password expiration : The password for $username is set to never expire" -ForegroundColor Red
        }
    elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordStatus.PasswordExpiration -ne '0') {
        Write-Host "Password last set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
        Write-Host "Password expiration : The password for $username expired on $($PasswordStatus.PasswordExpiration.ToString("g"))." -ForegroundColor Red
        }
    elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordStatus.PasswordExpiration -eq '0') {
        Write-Host 'Password Warning     : This account has a temporary password set and must change their' -ForegroundColor Red
        Write-Host '                      password upon next logon' -ForegroundColor Red
        }
    else {
        Write-Host "Password last set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
        Write-Host "Password expiration : $($PasswordStatus.PasswordExpiration.ToString("g"))"
        }
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Password Processing' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)


    $resetoption = $null
    while ($resetoption -ne 'y' -and $resetoption -ne 'n') {
        Write-Host ''
        Write-Host "Are you sure you want to RESET the password for $($username)?" -ForegroundColor Yellow
        $resetoption = read-host '     [Y] Yes  [N] No  (default is "N")'
        Write-Host ''
        if (!$resetoption) {$resetoption = 'n'}
        if ($resetoption -ne 'y' -and $resetoption -ne 'n') {write-host 'Please specify one of the available options' -ForegroundColor Red}
        }
    if ($resetoption -eq 'y') {
        switch ($PasswordType) { 
            'ResetManualPerm' {
                Write-Host 'Please follow all password length, history, and complexity requirements.' -ForegroundColor Yellow
                Write-Host 'This will set a permanent password and the user will NOT be required to change upon login.' -ForegroundColor Yellow
                Write-Host ''
                Set-ADAccountPassword -Identity $username -Reset
                Write-Host ''
                }
            'ResetManualTemp' {
                Write-Host 'Please follow all password length, history, and complexity requirements.' -ForegroundColor Yellow
                Write-Host 'The user will be required to change their password upon login.' -ForegroundColor Yellow
                Write-Host ''
                Set-ADAccountPassword -Identity $username -Reset
                $ADInfo | Set-ADUser -ChangePasswordAtLogon $true
                Write-Host ''
                }
            'ResetRandomTemp' {
                $password = WinSam-Create-ComplexPassword 12
                $ldapPath =  "LDAP://"+$ADInfo.DistinguishedName
                $objUser = [ADSI] $ldapPath
                $objUser.SetPassword($password)
                $objUser.SetInfo()

                Write-Host "The password for $username has been successfully reset" -ForegroundColor Green
                Write-Host "Password: $password"
                $ADInfo | Set-ADUser -ChangePasswordAtLogon $true
                Write-Host ''
                }
            }
        }
    else {
        Write-Host "You have chosen NOT to reset the password for $username" -ForegroundColor Yellow
        sleep 3
        }
    WinSam-Menu-User
    }

function WinSam-User ($SamAccountName) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1

    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$SamAccountName) {
        Write-Host ''
        $username = read-host -Prompt 'Please enter a username'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-AccountInfo $username
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-User
    }

function WinSam-Group ($SamAccountName) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$SamAccountName) {
        Write-Host ''
        Write-Host 'Please enter a group.  You may need to use quotes if there is a space in the name.'
        Write-Host ''
        $GroupName = read-host -Prompt '   Group Name'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-GroupInfo $GroupName
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-Group
    }

function WinSam-PC ($SystemName) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$SystemName) {
        Write-Host ''
        $hostname = read-host -Prompt 'Please enter a hostname'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-ComputerInfo $hostname
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-PC
    }

function WinSam-Mailbox ($MBAlias) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$MBAlias) {
        Write-Host ''
        $alias = read-host -Prompt 'Please enter a mailbox alias'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-MailboxInfo $alias
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-Mailbox
    }

function WinSam-Print-DebugInfo {
    If ($debug) {
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Begin Debug Code' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Cyan
        Write-Host "Operator Username : $currentuser" -ForegroundColor Cyan
        Write-Host "ACL:User          : $UserAccessLevel" -ForegroundColor Cyan
        Write-Host "Domain Controller : $DCHostName" -ForegroundColor Cyan
        Write-Host "Exchange Server   : $CASServer" -ForegroundColor Cyan
        Write-Host "Username          : $username" -ForegroundColor Cyan
        Write-Host "Account Type      : $Global:AccountType" -ForegroundColor Cyan
        #Write-Host "Beta Access       : $BetaAccess" -ForegroundColor Yellow
        #Write-Host "Beta Access List  : $BetaAccessList" -ForegroundColor Yellow
        Write-Host "Path              : $ScriptPath" -ForegroundColor Cyan
        if ($ElapsedTime) {Write-Host "Total time        : "$ElapsedTime.Elapsed -ForegroundColor Cyan}
        Write-Host (WinSam-Write-Header 'End Debug Code' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Cyan
        $ElapsedTime = $null
        }
    }

Function WinSam-Get-LocalGroups ($hostname) {
    function List-LocalGroupMembers ($localGroupName) {
        #Variables
        $LocalUsers=@();$DomainUsers=@();$DomainGroups=@()

        $Group = $ComputerObj.psbase.children.find($localGroupName)
        $members= $Group.psbase.invoke(“Members”) #| %{$_.GetType().InvokeMember(“Name”, ‘GetProperty’, $null, $_, $null)}

        # Invoke the Members method and convert to an array of member objects.
        $Members= @($Group.psbase.Invoke("Members"))

        ForEach ($Member In $Members) {
            $Name = $Member.GetType().InvokeMember("Name", 'GetProperty', $Null, $Member, $Null)
            $Path = $Member.GetType().InvokeMember("ADsPath", 'GetProperty', $Null, $Member, $Null)
            $Class = $Member.GetType().InvokeMember("Class", 'GetProperty', $Null, $Member, $Null)
            if ($path -match $hostname) {$LocalUsers += $Name}
            Else {
                if ($class -eq "Group") {$DomainGroups += Get-ADGroup $Name -Properties DisplayName}
                else {$DomainUsers += Get-ADUser $Name -Properties DisplayName}
                }
            }

        Write-Host (WinSam-Write-Header $localGroupName $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Domain Groups' ($MenuLength*0.66))
        Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
        $DomainGroups | Select @{Name="Group Name";Expression={$_.SamAccountName}}  | Sort "Group Name" | Out-Default

        if ($DomainUsers) {
            Write-Host (WinSam-Write-Header 'Domain Users' ($MenuLength*0.66))
            Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
            $DomainUsers | Select @{Name="Name";Expression={$_.DisplayName}},@{Name="Username";Expression={$_.SamAccountName}} | Sort Name | Out-Default
            }

        Write-Host (WinSam-Write-Header 'Local Users' ($MenuLength*0.66))
        Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
        $LocalUsers | Sort
        Write-Host ""
        }

    #This script is designed to get a list of administrators and local accounts for the computer
    $ComputerObj=$null;$user=$null;$users=$null

    #Domain Controller Information
    $dcs = (Get-ADDomainController -Filter *)
    $dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
    $dchostname = $dc.HostName



    #Get Users on Computer
    
    $ComputerObj = [ADSI]"WinNT://$hostname,computer"
    if ($ComputerObj.Path) {
        List-LocalGroupMembers 'Administrators'
        List-LocalGroupMembers 'Remote Desktop Users'
        }
    else {Write-Host "Unable to retrieve Local Group Members.  You may not have the necessary permissions." -ForegroundColor Yellow}
    }

function WinSam-SearchByID ($WPI_ID) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1

    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''

    Write-Host ''
    Write-Host "NOTE: Searching for a user by ID number may take 2-4 minutes to complete a search." -ForegroundColor Yellow
    Write-Host ''
    if (!$WPI_ID) {$WPI_ID = Read-Host -Prompt 'Please enter the 9-digit WPI ID Number'}
    while ($WPI_ID.Length -ne 9 -or $WPI_ID -notmatch '^\d+$') {
        $ErrorCount++
        Write-Host "Your entry was incorrect." -ForegroundColor Red
        Write-Host ''
        $WPI_ID = Read-Host -Prompt 'Please enter the 9-digit WPI ID Number'

        if ($ErrorCount -gt 3) {
            Write-Host "You have entered incorrect information too many times.  Returning to main menu" -ForegroundColor Red
            WinSam-Menu-Main
            }
        }

    if (!$Global:UserIndex) {$Global:UserIndex = Get-ADUser -Filter * -Properties DisplayName,EmployeeID,EmployeeNumber -SearchBase 'OU=Accounts,DC=admin,DC=wpi,DC=edu'}
    $UserResult = $Global:UserIndex| Where {$_.EmployeeID -eq $WPI_ID} 

    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor



    If ($UserResult) {
        Write-Host ''
        Write-Host 'Name                :' $UserResult.DisplayName
        Write-Host 'Username            :' $UserResult.SamAccountName
        Write-Host 'WPI ID              :' $UserResult.EmployeeID
        Write-Host 'PIDM                :' $UserResult.EmployeeNumber
        }
    else {Write-Host "No result was returned" -ForegroundColor Yellow}

    WinSam-Menu-SearchByID ($UserResult.SamAccountName)
    }

function WinSam-PIN {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1

    $ValidatedID=$Null
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    $BadInputCount=$null
    while(!$ValidatedID) {
        ## Prompt user to enter ID or Username
        $ID = read-host -Prompt 'Please enter the WPI ID Number or username'
        
        ## Check to see if it is an ID number (all numbers and 9 digits) - if so, set as validated, loop will end.
        if ($ID -match '^\d+$' -and $ID.Length -eq 9) {
            $ValidatedID = $ID
            continue
            }
        
        ## Entry is not a valid ID.  Check to see if it is a valid username in AD
        try {$Global:ADInfo = Get-ADUser $ID -Properties AccountExpirationDate,AccountLockoutTime,badPwdCount,CanonicalName,Department,Description,DisplayName,DistinguishedName,EmployeeID,EmployeeNumber,LastLogon,LockedOut,LockoutTime,LogonWorkstations,Office,PasswordExpired,PasswordLastSet,TelephoneNumber,Title,UserPrincipalName,WhenCreated -Server $dchostname}
        catch {}
        
        ## If valid AD Account, set as validated, loop will end - Note this will allow service accounts or other non banner accounts to be validated.
        if ($ADInfo) {
            $username=$ADInfo.SamAccountName
            $ValidatedID=$ID
            continue
            }

        ## If not valid, determine error
        if ($ID -match '^\d+$' -and $ID.Length -ne 9) {$ErrorText = "The ID number you entered [$ID] is not 9 digits."}
        else {$ErrorText = "Your entry [$ID] is not a valid ID or username."}
        
        ## Refresh screen and prompt for new entry
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host ''
        Write-Host $ErrorText -ForegroundColor Red
        Write-Host ''

        $BadInputCount++

        ## Provide an escape to this loop if BadInputCount is too high
        if ($BadInputCount -ge 3) {
            Write-Host "You have made too many bad attempts.  You will be returned to the main menu." -ForegroundColor Red
            Write-Host ''
            Write-Host "Press any key to continue ..."
            $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            WinSam-Menu-Main
            }
        }
    WinSam-Reset-PIN ($ValidatedID)
    }

function WinSam-Reset-PIN ($ID) {
    #This function takes the value of $ID which can be either username or WPI ID
    Clear-Host
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Banner PIN Processing' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host ''
    Write-Host "Changing PIN for: $ID"
    Write-Host ''

    $BannerDataSource = 'prod'  ## Specify 'prod' for production or 'trng' for training.
    $BannerUserID     = 'wsamprox'
    $BannerPassword   = 'jdhr65_83bhdldh'

    # Set up connection to Oracle
    if ($BannerDataSource -eq 'trng') {Write-Host "This is processing PIN resets on TRNG" -ForegroundColor Yellow}
    $connectionString = "Data Source=$BannerDataSource;User ID=$BannerUserID;Password=$BannerPassword;Integrated Security=no"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient") | out-null
    $connection = New-Object System.Data.OracleClient.OracleConnection($connectionString)
    [System.Data.OracleClient.OracleCommand] $command = New-Object System.Data.OracleClient.OracleCommand
    $command.Connection = $connection

    # Set up command parameters for stored procedure
    $command.CommandType = [System.Data.CommandType]::StoredProcedure
    $command.CommandText = "oracle_mgr.gwpprst"
    $command.Parameters.Add("USER_NAME", [System.Data.OracleClient.OracleType]::VarChar2) | out-null
    $command.Parameters["USER_NAME"].Direction = [System.Data.ParameterDirection]::Input
    $command.Parameters["USER_NAME"].Value = $ID
    
    # Connect to Oracle
    $connection.Open()

    try {$command.ExecuteNonQuery() | out-null;Write-Host "PIN reset successfully."}
    catch {Write-Host "Error executing PIN reset: " $_}
    finally {$connection.Close();[GC]::Collect()}

    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-PIN
    }