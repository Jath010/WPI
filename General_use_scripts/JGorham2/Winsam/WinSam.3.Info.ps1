function WinSam-Get-AccountInfo {
    <#
    .SYNOPSIS
    Returns basic information about a user include AD info and Mailbox info
    .DESCRIPTION
    Returns basic information about a user include AD info and Mailbox info.  Written by Tom Collins.  Last Updated 1/26/2012
    .EXAMPLE
    WinSam-Get-AccountInfo username
    .PARAMETER username
    The username to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$username
        )
    $ErrorActionPreference = "SilentlyContinue"
    #*****************************************************************
    #Function Declarations
    #*****************************************************************
    #Required functions:
    #   - WPI-Get-LastLogonDate
    #   - WPI-Get-PasswordExpiration
    #   - WPI-Get-MailboxStats

    #*****************************************************************
    #Main Code
    #*****************************************************************

    $Global:mailbox = $null
    $Global:ADInfo = $null
    $Global:AccountLockoutStatus = $null
    $Global:AccountNEStatus = $Null
    $Global:AccountGroups = $null
    $Global:AccountExpirationDate = $null
    $mailboxStats = $null
    $mailboxSize = $null
    $MailboxQuota = $null
    $MailboxPercentUse = $null
    $MailboxStorageLimitStatus = $null
    $lastlogon = $null

    #Check to see if Account Exists
    $username = $username.Trim()
    $Global:ADInfo = Get-ADUser $username -Properties AccountExpirationDate,AccountLockoutTime,badPwdCount,CanonicalName,Department,Description,DisplayName,DistinguishedName,EmployeeID,EmployeeNumber,LastLogon,LockedOut,LockoutTime,LogonWorkstations,Office,PasswordExpired,PasswordLastSet,TelephoneNumber,Title,UserPrincipalName,WhenCreated -Server $dchostname -ErrorAction "SilentlyContinue"

    if (!$Global:ADInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : User '$username' does not exist" -ForegroundColor Yellow
        Write-Host ''
        return
        }
    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    $Global:AccountExpirationDate = $Global:ADInfo.AccountExpirationDate
    $Global:AccountEnabled = $global:ADinfo.enabled
    $Global:AccountGroups = Get-ADPrincipalGroupMembership $username -Server $dchostname
    $Global:AccountNEStatus  = ($Global:AccountGroups | Where {$_.Name -eq 'Nonemployees'}).Name
    WinSam-Get-AccountStatus
    $Global:AccountLockoutInfo = Get-ADUser $username -Properties AccountLockoutTime,badPwdCount,LockedOut,LockoutTime -Server $dchostname
    $Global:AccountLockoutStatus = $Global:AccountLockoutInfo.LockedOut

    if ($Global:AccountType -ne 'Disabled' -or $UserAccessLevel -eq "SysAdmin") {
        $Global:mailbox = Get-Mailbox $username
        }
    Write-Host ''
    Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
    Write-Host ''
    If ($Global:AccountType -ne 'Student' -and $Global:AccountType -ne 'ResourceMailbox' -and $Global:AccountType -ne 'Alumni' -and $Global:AccountType -ne 'LOA') {
        Write-Host 'Title               :' $Global:ADInfo.Title
        Write-Host 'Department          :' $Global:ADInfo.Department
        }
    
    if ($Global:AccountType -ne 'Disabled' -or $UserAccessLevel -eq "SysAdmin") {
        If ($Global:AccountType -ne 'Student' -and $Global:AccountType -ne 'Alumni' -and $Global:AccountType -ne 'LOA') {
            Write-Host 'Office              :' $Global:ADInfo.Office
            Write-Host 'Phone               :' $Global:ADInfo.telephoneNumber
            }
        Write-Host 'Description         :' $Global:ADInfo.description
        }
        Write-Host ''        
        If ($Global:AccountType -eq 'Student' -or $Global:AccountType -eq 'Employee' -or $Global:AccountType -eq 'Disabled' -or $Global:AccountType -eq 'Retiree' -or $Global:AccountType -eq 'Alumni' -or $Global:AccountType -eq 'LOA') {
            Write-Host 'WPI ID              :' $Global:ADinfo.EmployeeID
            Write-Host 'PIDM                :' $Global:ADinfo.EmployeeNumber
            Write-Host ''
            }

    if ($Global:AccountType -ne 'Disabled' -or $UserAccessLevel -eq "SysAdmin") {
        if ($Global:AccountType -ne 'ResourceMailbox') {
            Write-Host (WinSam-Write-Header 'Windows Account Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)

            Write-Host "Account Created     : $($Global:ADInfo.whenCreated)"

            $logonInfo = WinSam-Get-LastLogonDate $username
            If ($logonInfo.LastLogon){Write-Host "Last Login          : $($logonInfo.LastLogon.ToString("g")) ($($logonInfo.LogonServer))"}
            Else {Write-Host "Last Login          : Never logged in or too long since last login." -ForegroundColor Yellow}
            Write-Host ''

            $PasswordStatus = WinSam-Get-PasswordExpiration $username

            if (($PasswordStatus.PasswordExpiration -is [datetime]) -and $AccountExpirationDate -and $PasswordStatus.PasswordExpiration -gt $AccountExpirationDate) {$PasswordExpiration = $AccountExpirationDate}
            else {$PasswordExpiration = $PasswordStatus.PasswordExpiration}
            
            if ($PasswordStatus.PasswordNeverExpires -eq 'True') {
                Write-Host "Password Last Set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
                Write-Host "Password Expiration : The password for $username is set to never expire" -ForegroundColor Red
                }
            elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordExpiration -ne '0') {
                Write-Host "Password Last Set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
                Write-Host "Password Expiration : The password for $username expired on $($PasswordExpiration.ToString("g"))." -ForegroundColor Red
                }
            elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordExpiration -eq '0') {
                Write-Host 'Password Warning     : This user is required to change their password upon next logon' -ForegroundColor Red
                }
            else {
                Write-Host "Password Last Set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
                if ($today.AddDays(20) -gt $PasswordExpires) {Write-Host "Password Expiration : $($PasswordExpiration.ToString("g"))" -ForegroundColor Yellow}
                else {Write-Host "Password Expiration : $($PasswordExpiration.ToString("g"))"}
                }
            Write-Host ''
    	    if ($AccountExpirationDate -and $AccountExpirationDate -gt $Global:today) {Write-Host "Account  Expires On : $($AccountExpirationDate.ToString("g"))" -ForegroundColor Yellow;Write-Host ''}
            elseif ($AccountExpirationDate -and $AccountExpirationDate -lt $Global:today) {Write-Host (WinSam-Write-Header "This account expired on $($AccountExpirationDate.ToString("g"))" $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Red;Write-Host ''}

            if ($Global:AccountLockoutStatus) {
                if ($UserAccessLevel -eq "SysAdmin" -or $UserAccessLevel -eq 'PasswordReset_lvl2' -or $UserAccessLevel -eq "PasswordReset" -or $UserAccessLevel -eq "Unlock") {
                    Unlock-ADAccount -Identity $username
                    Write-Host "Account Lockout     : $username has been succesfully unlocked" -ForegroundColor green
                    $Global:AccountLockoutStatus = $null
                    }
                else {Write-Host 'Account  Lockout    : Locked Out' -ForegroundColor Red}
                }
            else{Write-Host 'Account  Lockout    : Not locked out'}
            Write-Host ''



            $LogonWorkstations = WinSam-Get-LogonWorkstations
            if ($LogonWorkstations) {
                Write-Host 'This user has restricted PC access.  They may only log onto the following computers: ' -ForegroundColor Black -BackgroundColor Yellow
                Write-Host (WinSam-Write-Header '' $MenuLength -Line) -ForegroundColor Yellow
                Write-ColorOutput Yellow ($LogonWorkstations)
                }
            }
        
        Write-Host (WinSam-Write-Header 'Mailbox Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor        
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        If ($Global:mailbox) {
            if ($Global:mailbox.ForwardingSmtpAddress) {Write-Host "Mail Delivery       : Mail is forwarded to: $($Global:mailbox.ForwardingSmtpAddress.Split(':')[1])" -ForegroundColor Yellow}
            WinSam-Get-MailboxStats $username

            $PrimarySMTPAddress = ($Global:Mailbox.EmailAddresses | Where {$_ -clike 'SMTP:*'}).Split(':')[1]
            $ProxySMTPAddresses = $Global:Mailbox.EmailAddresses | Where {$_ -like 'smtp:*' -and $_ -cnotlike 'SMTP:*'} | Sort
            $ProxySMTPAddresses = $ProxySMTPAddresses | Where {$_ -notmatch '@wpi0' -and $_ -notmatch '@exchange'}
            Write-Host ''
            Write-Host 'Proxy Address Information'
            Write-Host '============================'
            Write-Host '   Primary SMTP      :' $PrimarySMTPAddress
            if ($ProxySMTPAddresses) {
                Write-Host ''
                Write-Host '   Additional Addresses' -ForegroundColor Gray
                Write-Host '   --------------------'
                foreach ($address in $ProxySMTPAddresses) {Write-Host '      ' $address.Split(':')[1] -ForegroundColor Gray}
                }
            }
        else {write-host 'Mail Delivery       : No Exchange mailbox exists for this user.' -ForegroundColor Yellow}
    }
    #End of Info function
    }

function WinSam-Get-ADGroupInfo {
    <#
    .SYNOPSIS
    Returns basic information about a user include AD info and Mailbox info
    .DESCRIPTION
    Returns basic information about a user include AD info and Mailbox info.  Written by Tom Collins.  Last Updated 1/26/2012
    .EXAMPLE
    WinSam-Get-GroupInfo Group
    .PARAMETER GroupName
    The SAMAccountName for the group to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$GroupName
        )
    $ErrorActionPreference = "SilentlyContinue"
    #*****************************************************************
    #Function Declarations
    #*****************************************************************
    #Required functions:
    #   - WPI-Get-LastLogonDate
    #   - WPI-Get-PasswordExpiration
    #   - WPI-Get-MailboxStats

    #*****************************************************************
    #Main Code
    #*****************************************************************

    $GroupInfo = $null
    $GroupMembers = $null
    $GroupManagers = @()
    $GroupManagersList = @()
    $GroupMemberOf = $null

    #Check to see if Account Exists
    $GroupName = $GroupName.Trim()
    $GroupInfo = Get-ADGroup $GroupName -Properties DisplayName,DistinguishedName,Description,GroupCategory,GroupScope,ManagedBy,msExchCoManagedByLink,WhenCreated -Server $dchostname 
    $DistroInfo = Get-DistributionGroup $GroupName -ErrorAction SilentlyContinue

    if (!$GroupInfo) {
        if ($DistroInfo) {$GroupInfo = Get-ADGroup $DistroInfo.DistinguishedName -Properties DisplayName,DistinguishedName,Description,GroupCategory,GroupScope,ManagedBy,msExchCoManagedByLink,WhenCreated -Server $dchostname}
        else {
            Write-Host ''
            Write-Host ''
            Write-Host "       WARNING : Group '$GroupName' does not exist" -ForegroundColor Yellow
            write-host ''
            return
            }
        }

    $GroupMembers = $GroupInfo | Get-ADGroupMember -Server $dchostname -ErrorAction "SilentlyContinue"
    
    if ($DistroInfo) {foreach ($user in $GroupInfo.msExchCoManagedByLink) {$GroupManagers += $user}}
    $GroupManagers += $GroupInfo.ManagedBy   
    
    Foreach ($user in $GroupManagers) {$GroupManagersList += Get-ADUser $user | Select Name,SamAccountName}


    #$Global:GroupMemberOf = Get-ADPrincipalGroupMembership $GroupName -Server $dchostname -ErrorAction "SilentlyContinue"  ### Not using in code at this time.  May eventually add programming to show what groups a group is a member of.

    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    Write-Host ''
    Write-Host (WinSam-Write-Header 'General Group Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host 'Name                :' $GroupInfo.Name
    Write-Host 'Group ID            :' $GroupInfo.SamAccountName
    Write-Host 'Description         :' $GroupInfo.Description
    if ($DistroInfo) {Write-Host "Group Type          : $($GroupInfo.GroupCategory), Distribution Group" -ForegroundColor Green}
    else {Write-Host 'Group Type          :' $GroupInfo.GroupCategory}
    if ($UserAccessLevel -eq "SysAdmin") {Write-Host 'Created             :' $GroupInfo.WhenCreated -ForegroundColor Gray}
    if ($GroupManagersList) {
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Distribution List Managers' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $GroupManagersList | Select Name,@{Name="Username";Expression={$_.SamAccountName}} | Sort Name | Out-Default
        }
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Group Membership' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Green
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host ''
    Write-Host (WinSam-Write-Header 'User Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
    $GroupMembers | Where {$_.ObjectClass -eq 'user'} | Get-ADUser -Server $dchostname -Properties DisplayName,Department,Title | Select DisplayName,@{Name="Username";Expression={$_.SamAccountName}},Department,Title | Sort Name | FT | Out-Default
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Group Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66)  -Line)
    $GroupMembers | Where {$_.ObjectClass -eq 'group'} | Select Name,@{Name="Username";Expression={$_.SamAccountName}} | Sort Name | Out-Default
    Write-Host ''
    }

function WinSam-Get-UnifiedGroupInfo {
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$EmailAddress
        )
    $ErrorActionPreference = "SilentlyContinue"
    #*****************************************************************
    #Function Declarations
    #*****************************************************************
    #Required functions:
    #   - WPI-Get-LastLogonDate
    #   - WPI-Get-PasswordExpiration
    #   - WPI-Get-MailboxStats

    #*****************************************************************
    #Main Code
    #*****************************************************************

    $GroupInfo = $null
    $GroupMembers = $null
    $GroupManagers = @()
    $GroupManagersList = @()
    $GroupMemberOf = $null
    $AllowExternalEmail=$null

    #Check to see if Account Exists
    $EmailAddress = $EmailAddress.Trim()
    $GroupInfo = Get-UnifiedGroup $EmailAddress -ErrorAction SilentlyContinue

    if (!$GroupInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : O365 Group does not exist" -ForegroundColor Yellow
        Write-Host ''
        Write-Host "              Checked: $EmailAddress" 
        Write-Host ''
        return
        }
    
    #$Global:EmailAddress = $GroupInfo.PrimarySmtpAddress
    $GroupMembers  = $GroupInfo | Get-UnifiedGroupLinks -LinkType Members -ErrorAction "SilentlyContinue"
    $GroupManagers = $GroupInfo | Get-UnifiedGroupLinks -LinkType Owners -ErrorAction "SilentlyContinue"
    $SMTPAddresses = $GroupInfo.EmailAddresses | Where {$_ -like 'smtp:*wpi.edu'} | Sort
    $SMTPAddresses = $SMTPAddresses | Where {$_ -notmatch '@wpi0' -and $_ -notmatch '@exchange' -and $_ -cnotmatch 'SMTP'}
    #Foreach ($user in $GroupManagers) {$GroupManagersList += Get-ADUser $user | Select Name,SamAccountName}

    if ($GroupInfo.RequireSenderAuthenticationEnabled) {$AllowExternalEmail = $false}
    else {$AllowExternalEmail = $True}


    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Unified Group Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host 'DisplayName          :' $GroupInfo.DisplayName
    Write-Host 'Group ID             :' $GroupInfo.Alias
    Write-Host 'Description          :' $GroupInfo.Notes
    Write-Host ''
    Write-Host 'Hidden from GAL      :' $GroupInfo.HiddenFromAddressListsEnabled
    Write-Host 'Access Type          :' $GroupInfo.AccessType
    Write-Host 'Allow External Email :' $AllowExternalEmail
    Write-Host 'Moderation Enabled   :' $GroupInfo.ModerationEnabled
    Write-Host ''
    Write-Host 'Proxy Address Information'
    Write-Host '============================'
    Write-Host '   Primary SMTP      :' $GroupInfo.PrimarySmtpAddress
    if ($SMTPAddresses) {
        Write-Host ''
        Write-Host '   Additional Addresses' -ForegroundColor Gray
        Write-Host '   --------------------'
        foreach ($address in $SMTPAddresses) {
            Write-Host '      ' $address.Split(':')[1] -ForegroundColor Gray
            }
        }
    if ($UserAccessLevel -eq "SysAdmin") {Write-Host 'Created             :' $GroupInfo.WhenCreated -ForegroundColor Gray}
    if ($GroupManagers) {
        Write-Host ''
        Write-Host (WinSam-Write-Header 'O365 Group Managers' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $GroupManagers | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default
        }
    Write-Host ''
    Write-Host (WinSam-Write-Header 'O365 Group Membership' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Green
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host ''
    Write-Host (WinSam-Write-Header 'WPI Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
    $GroupMembers | Where {$_.RecipientTypeDetails -eq 'UserMailbox'} | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default

    Write-Host ''
    Write-Host (WinSam-Write-Header 'Guest Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66)  -Line)
    $GroupMembers | Where {$_.RecipientTypeDetails -ne 'UserMailbox'} | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default
    Write-Host ''
    }


function WinSam-Get-DistributionGroupInfo {
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$EmailAddress
        )
    $ErrorActionPreference = "SilentlyContinue"
    #*****************************************************************
    #Function Declarations
    #*****************************************************************
    #Required functions:
    #   - WPI-Get-LastLogonDate
    #   - WPI-Get-PasswordExpiration
    #   - WPI-Get-MailboxStats

    #*****************************************************************
    #Main Code
    #*****************************************************************

    $GroupInfo = $null
    $GroupMembers = $null
    $GroupManagers = @()
    $GroupManagersList = @()
    $GroupMemberOf = $null
    $AllowExternalEmail=$null

    #Check to see if Account Exists
    $EmailAddress = $EmailAddress.Trim()
    $GroupInfo = Get-DistributionGroup $EmailAddress -ErrorAction SilentlyContinue
    $RecipientInfo = Get-Recipient $EmailAddress

    if (!$GroupInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : Distribution Group does not exist" -ForegroundColor Yellow
        Write-Host ''
        Write-Host "              Checked: $EmailAddress" 
        Write-Host ''
        return
        }
    
    #$Global:EmailAddress = $GroupInfo.PrimarySmtpAddress
    $GroupMembers  = Get-DistributionGroupMember $EmailAddress -ResultSize Unlimited -ErrorAction "SilentlyContinue"
    $GroupMembersUsers = $GroupMembers | WHere {$_.recipientType -eq 'UserMailbox'} | Get-Mailbox
    $GroupMembersGroups = $GroupMembers | WHere {$_.recipientType -ne 'UserMailbox'} | Get-DistributionGroup
    $GroupManagers = $GroupInfo.ManagedBy | Get-Mailbox

    $SMTPAddresses = $GroupInfo.EmailAddresses | Where {$_ -like 'smtp:*wpi.edu'} | Sort
    $SMTPAddresses = $SMTPAddresses | Where {$_ -notmatch '@wpi0' -and $_ -notmatch '@exchange' -and $_ -cnotmatch 'SMTP'}
    #Foreach ($user in $GroupManagers) {$GroupManagersList += Get-ADUser $user | Select Name,SamAccountName}

    if ($GroupInfo.RequireSenderAuthenticationEnabled) {$AllowExternalEmail = $false}
    else {$AllowExternalEmail = $True}


    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Distribution Group Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host 'DisplayName          :' $GroupInfo.DisplayName
    Write-Host 'Group ID             :' $GroupInfo.Alias
    if ($RecipientInfo.Capabilities -contains 'MasteredOnPremise') {Write-Host 'Location             : This distribution group is managed on premise.'}
    else {Write-Host 'Location             : This distribution group is managed online.'}
    Write-Host ''
    Write-Host 'Hidden from GAL      :' $GroupInfo.HiddenFromAddressListsEnabled
    Write-Host 'Allow External Email :' $AllowExternalEmail
    Write-Host 'Moderation Enabled   :' $GroupInfo.ModerationEnabled
    Write-Host ''
    Write-Host 'Proxy Address Information'
    Write-Host '============================'
    Write-Host '   Primary SMTP      :' $GroupInfo.PrimarySmtpAddress
    if ($SMTPAddresses) {
        Write-Host ''
        Write-Host '   Additional Addresses' -ForegroundColor Gray
        Write-Host '   --------------------'
        foreach ($address in $SMTPAddresses) {
            Write-Host '      ' $address.Split(':')[1] -ForegroundColor Gray
            }
        }
    if ($UserAccessLevel -eq "SysAdmin") {Write-Host 'Created             :' $GroupInfo.WhenCreated -ForegroundColor Gray}
    if ($GroupManagers) {
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Distribution Group Managers' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $GroupManagers | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default
        }
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Distribution Group Membership' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Green
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Group Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
    $GroupMembersGroups | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default
    Write-Host ''
    Write-Host (WinSam-Write-Header 'User Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
    $GroupMembersUsers | Select DisplayName,@{Name="Email";Expression={$_.PrimarySmtpAddress}} | Sort DisplayName | Out-Default
    Write-Host ''
    }

function WinSam-Get-ComputerInfo {
    <#
    .SYNOPSIS
    Returns basic information about a computer include AD info and WMI info
    .DESCRIPTION
    Returns basic information about a computer include AD info and WMI info.  Written by Tom Collins.  Last Updated 1/26/2012
    .EXAMPLE
    WinSam-Get-AccountInfo username
    .PARAMETER computername
    The hostname to query. Just one.
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$computername
        )
    $ErrorActionPreference = "SilentlyContinue"

    $ComputerSystemInfo = $null;$ProcessorInfo = $null;$OSInfo = $null;$LogicalDiskInfo = $null;$MemInfo = $null
    $TotalMemory = $null;$CPUSpeed = $null;$ComputerInfo = $null;$IPAddresses=$null;$IPAddress=$null
    $Today = $null

    $Today = Get-Date

    #Check to see if Account Exists
    $computername = $computername.Trim()
    $ComputerInfo = Get-ADComputer $computername -Properties Created,Description,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate
    if (!$ComputerInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : Computer object '$computername' does not exist" -ForegroundColor Yellow
        write-host ''
        return
        }
    
    if (Test-Connection -Count 1 -BufferSize 15 -Delay 1 -ComputerName $computername) { 
        $ComputerSystemInfo = Get-WmiObject Win32_ComputerSystemProduct -ComputerName $computername
        $IPAddresses = [System.Net.Dns]::GetHostAddresses($computername) | Where {$_.IPAddressToString -match '130.215'} | Sort IPAddressToString | %{
            if (!$IPAddress) {$IPAddress = "$($_.IPAddressToString)"}
            Else {$IPAddress = "$IPAddress,$($_.IPAddressToString)"}
            }
        if ($ComputerSystemInfo) {
            $ProcessorInfo = Get-WmiObject Win32_Processor -ComputerName $computername
            $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $computername
            $MemInfo = Get-WmiObject Win32_PhysicalMemory -ComputerName $computername
            $LogicalDiskInfo = Get-WmiObject Win32_LogicalDisk -ComputerName $computername

            $MemInfo | %{$TotalMemory += $_.Capacity}
            if ($TotalMemory -lt 1gb) {$TotalMemory = "$($TotalMemory.ToString()/1mb) MB"}
            Else{$TotalMemory = "$($TotalMemory.ToString()/1gb) GB"}

            #Calculate Uptime
            $BootTime = $OSInfo.ConvertToDateTime($OSInfo.LastBootUpTime)
            $Uptime = ($Today - $BootTime)

            switch ($Uptime.Days) {
                0 {$UptimeDays = ""}
                1 {$UptimeDays = "1 Day "}
                default {$UptimeDays = "$($Uptime.Days) Days "}
                }

            if ($Uptime.Hours -eq 1) {$UptimeHours = "$($Uptime.Hours) Hour"}
            else {$UptimeHours = "$($Uptime.Hours) Hours"}
            $UptimeOutput = $UptimeDays + $UptimeHours

            # Display Computer Information
            Write-Host ''
            Write-Host (WinSam-Write-Header 'General Computer Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host "System Information  :" $OSInfo.CSName
            Write-Host "IP Address(es)      :" $IPAddress
            Write-Host "Description         :" $ComputerInfo.Description
            if ($UserAccessLevel -eq 'SysAdmin') {
                Write-Host "Distinguished Name  :" $ComputerInfo.DistinguishedName -ForegroundColor Gray
                Write-Host ""
                Write-Host "Date Created (AD)   :" $ComputerInfo.Created.ToString("g") -ForegroundColor Gray
                Write-Host "Last Logon Date     :" $ComputerInfo.LastLogonDate.ToString("g") -ForegroundColor Gray
                }
            Write-Host ''
            Write-Host (WinSam-Write-Header 'Operating System Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host "OS                  : $($OSInfo.Caption) ($($OSInfo.Version)) [$($OSInfo.OSArchitecture)]"
            Write-Host "Service Pack Level  : $($OSInfo.CSDVersion)"
            Write-Host "WindowsDirectory    : $($OSInfo.WindowsDirectory)"
            Write-Host "Uptime              : $UptimeOutput (Booted: $BootTime)"
            Write-Host ''
            Write-Host (WinSam-Write-Header 'Hardware Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host "PC Model            : $($ComputerSystemInfo.Name) ($($ComputerSystemInfo.Vendor))"
            Write-Host "System Product ID   : $($ComputerSystemInfo.IdentifyingNumber)"
            Write-Host ''
            foreach ($processor in $ProcessorInfo) {
                if ($processor.CurrentClockSpeed -ge 1000) {
                    $CPUSpeed = $processor.CurrentClockSpeed/1000
                    $CPUSpeed = $CPUSpeed.ToString().substring(0,3) + " Ghz"
                    }
                else {
                    $CPUSpeed = $CPUSpeed.ToString() + " Mhz"
                    }
                Write-Host "CPU                 : $CPUSpeed $($processor.Manufacturer)"
                }
            Write-Host "Total RAM           :" $TotalMemory
            Write-Host ''
            Write-Host "Logical Drives"
            Write-Host "-------------------------------------"
            foreach ($LogicalDisk in $LogicalDiskInfo) {
                if ($LogicalDisk.DriveType -eq "3") {
                    $DiskPercent = ($LogicalDisk.FreeSpace/$LogicalDisk.Size)*100
                    if ($LogicalDisk.FreeSpace -lt 1gb) {
                        $diskinfo =  "     $($LogicalDisk.Name)[$($LogicalDisk.VolumeName)]   $("{0:N2}" -f ($LogicalDisk.Size/1MB)) MB [$("{0:N2}" -f $DiskPercent) % ($("{0:N2}" -f ($LogicalDisk.FreeSpace/1MB)) MB) Available]"
                        }
                    else {
                        $diskinfo =  "     $($LogicalDisk.Name)[$($LogicalDisk.VolumeName)]   $("{0:N2}" -f ($LogicalDisk.Size/1GB)) GB [$("{0:N2}" -f $DiskPercent) % ($("{0:N2}" -f ($LogicalDisk.FreeSpace/1GB)) GB) Available]"
                        }
                    if ($DiskPercent -le '30' -and $DiskPercent -gt '10') {Write-Host $diskinfo -ForegroundColor Yellow}
                    elseif ($DiskPercent -le '10') {Write-Host $diskinfo -ForegroundColor Red}
                    else {Write-Host $diskinfo}
                    }
                }
            Write-Host ''
            WinSam-Get-LocalGroups $computername
            }
        else {
            Write-Host (WinSam-Write-Header 'General Computer Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host "System Information  :" $ComputerInfo.Name
            Write-Host "Description         :" $ComputerInfo.Description
            Write-Host "IP Address(es)      :" $IPAddress
            if ($UserAccessLevel -eq 'SysAdmin') {
                Write-Host "Distinguished Name  :" $ComputerInfo.DistinguishedName -ForegroundColor Gray
                Write-Host ""
                Write-Host "Date Created (AD)   :" $ComputerInfo.Created.ToString("g") -ForegroundColor Gray
                Write-Host "Last Logon Date     :" $ComputerInfo.LastLogonDate.ToString("g") -ForegroundColor Gray
                }
            Write-Host ''
            Write-Host (WinSam-Write-Header 'Operating System Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host "OS                  : $($ComputerInfo.OperatingSystem) ($($ComputerInfo.OperatingSystemVersion))"
            Write-Host "Service Pack Level  :" $ComputerInfo.OperatingSystemServicePack
            Write-Host ''
            WinSam-Get-LocalGroups $computername
            }
        }
    else {
        Write-Host ''
        Write-Host "The computer $computername is not reachable" -ForegroundColor Black -BackgroundColor Red
        Write-Host ''
        Write-Host (WinSam-Write-Header 'General Computer Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host "System Information  :" $ComputerInfo.Name
        Write-Host "Description         :" $ComputerInfo.Description
        if ($UserAccessLevel -eq 'SysAdmin') {
            Write-Host "Distinguished Name  :" $ComputerInfo.DistinguishedName -ForegroundColor Gray
            Write-Host ""
            Write-Host "Date Created (AD)   :" $ComputerInfo.Created.ToString("g") -ForegroundColor Gray
            Write-Host "Last Logon Date     :" $ComputerInfo.LastLogonDate.ToString("g") -ForegroundColor Gray
            }
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Operating System Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host "OS                  : $($ComputerInfo.OperatingSystem) ($($ComputerInfo.OperatingSystemVersion))"
        Write-Host "Service Pack Level  :" $ComputerInfo.OperatingSystemServicePack
        Write-Host ''
        }
    }

function WinSam-Get-MailboxInfo ($EmailAddress) {
    #*****************************************************************
    #Function Declarations
    #*****************************************************************
    #Required functions:
    #   - WinSam-Get-MailboxStats

    #*****************************************************************
    #Main Code
    #*****************************************************************
    $ErrorActionPreference = "SilentlyContinue"
        
    #Get Mailbox Information
    $Global:Mailbox = Get-Mailbox $EmailAddress

    #Process Mailbox
    If (!$mailbox) {Write-Host '';Write-Host '';Write-Host "     ERROR: The mailbox [$EmailAddress] doesn't exist" -ForegroundColor Red}
    Else {
        $out           = $null    
        
        $alias = $EmailAddress.Split('@')[0]

        $PrimarySMTPAddress = ($Global:Mailbox.EmailAddresses | Where {$_ -clike 'SMTP:*'}).Split(':')[1]
        $ProxySMTPAddresses = $Global:Mailbox.EmailAddresses | Where {$_ -like 'smtp:*' -and $_ -cnotlike 'SMTP:*'} | Sort
        $ProxySMTPAddresses = $ProxySMTPAddresses | Where {$_ -notmatch '@wpi0' -and $_ -notmatch '@exchange'}


        $MailboxUser      = $null;$MailboxUsers  = $null
        $InboxUser        = $null;$InboxUsers    = $null
        $CalendarUser     = $null;$CalendarUsers = $null
        $SendOnBehalfUser = $null;$SendOnBehalf  = $null
        $SendAsUser       = $null;$SendAs        = $null

        $MailboxUserlist        = @()
        $SendOnBehalfUserList   = @()
        $SendAsUserList         = @()
        $InboxUserlist          = @()
        $CalendarUserlist       = @()

        $Inbox = "$($mailbox.Name):\Inbox" 
        $Calendar = "$($mailbox.Name):\Calendar" 

        $AutoMappedUsers = Get-ADUser $alias -Properties msExchDelegateListLink | Select -ExpandProperty msExchDelegateListLink
    	$MailboxUsers    = Get-MailboxPermission $alias | Where {$_.IsInherited -ne $true -and $_.User -notlike "NT AUTHORITY\SELF" -and $_.User -notlike "Organization Management"}
        $SendOnBehalf    = $mailbox.GrantSendOnBehalfTo
        $SendAs          = Get-RecipientPermission $alias | where {($_.IsInherited -eq $false) -and -not ($_.Trustee -like "NT AUTHORITY\SELF")}
        $InboxUsers      = Get-MailboxFolderPermission -Identity $Inbox
        $CalendarUsers   = Get-MailboxFolderPermission -Identity $Calendar

        if ($MailboxUsers) {
            foreach($MailboxUser in $MailboxUsers) {
                $AutoMapped=$null;$ObjectInfo=$null
                $ObjectInfo = WinSam-Get-ObjectInfo $MailboxUser.User
                
                if ($AutoMappedUsers -contains $ObjectInfo.DistinguishedName) {$AutoMapped=$true}
           
                $out = New-Object PSObject
                $out | add-member noteproperty Name $ObjectInfo.Name
                $out | add-member noteproperty Username $ObjectInfo.Username
                $out | add-member noteproperty Department $ObjectInfo.Department
                $out | add-member noteproperty Enabled $ObjectInfo.Enabled
                $out | add-member noteproperty AccessRights $MailboxUser.AccessRights
                $out | add-member noteproperty AutoMapped $AutoMapped
                $MailboxUserlist += $out        
                }
            }
        
        if ($SendOnBehalf) {
            foreach ($SendOnBehalfUser in $SendOnBehalf) {
                $ObjectInfo = WinSam-Get-ObjectInfo $SendOnBehalfUser
            
                $out = New-Object PSObject
                $out | add-member noteproperty Name $ObjectInfo.Name
                $out | add-member noteproperty Username $ObjectInfo.Username
                $out | add-member noteproperty Department $ObjectInfo.Department
                $out | add-member noteproperty Enabled $ObjectInfo.Enabled
        	    $SendOnBehalfUserList += $out        
                }
            }

        if ($SendAs) {
            foreach ($SendAsUser in $SendAs) {
                $ObjectInfo = WinSam-Get-ObjectInfo $SendAsUser.Trustee.Split('@')[0]
            
                $out = New-Object PSObject
                $out | add-member noteproperty Name $ObjectInfo.Name
                $out | add-member noteproperty Username $ObjectInfo.Username
                $out | add-member noteproperty Department $ObjectInfo.Department
                $out | add-member noteproperty Enabled $ObjectInfo.Enabled
        	    $SendAsUserList += $out        
                }
            }

        if ($InboxUsers) {
            foreach($InboxUser in $InboxUsers) {
                $ObjectInfo = WinSam-Get-ObjectInfo $InboxUser.User.DisplayName
            
                $out = New-Object PSObject
                $out | add-member noteproperty Name $ObjectInfo.Name
                $out | add-member noteproperty Username $ObjectInfo.Username
                $out | add-member noteproperty Department $ObjectInfo.Department
                $out | add-member noteproperty Enabled $ObjectInfo.Enabled
                $out | add-member noteproperty AccessRights $InboxUser.AccessRights
                $InboxUserlist += $out        
                }
            }

        if ($CalendarUsers) {
            foreach($CalendarUser in $CalendarUsers) {
                $ObjectInfo = WinSam-Get-ObjectInfo $CalendarUser.User.DisplayName
            
                $out = New-Object PSObject
                $out | add-member noteproperty Name $ObjectInfo.Name
                $out | add-member noteproperty Username $ObjectInfo.Username
                $out | add-member noteproperty Department $ObjectInfo.Department
                $out | add-member noteproperty Enabled $ObjectInfo.Enabled
                $out | add-member noteproperty AccessRights $CalendarUser.AccessRights
                $CalendarUserlist += $out
                }
            }


        Write-Host ''
        Write-Host (WinSam-Write-Header 'General Mailbox Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host "Mailbox Display Name: $($mailbox.DisplayName)"
        Write-Host "Mailbox Alias       : $($mailbox.Alias)"
        WinSam-Get-MailboxStats $alias
        Write-Host ''
        Write-Host 'Proxy Address Information'
        Write-Host '============================'
        Write-Host '   Primary SMTP      :' $PrimarySMTPAddress
        if ($ProxySMTPAddresses) {
            Write-Host ''
            Write-Host '   Additional Addresses' -ForegroundColor Gray
            Write-Host '   --------------------'
            foreach ($address in $ProxySMTPAddresses) {Write-Host '      ' $address.Split(':')[1] -ForegroundColor Gray}
            }

        Write-Host ''
        Write-Host (WinSam-Write-Header 'Mailbox Access Rights (SysOps Managed)' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $MailboxUserlist | Where {$_.Enabled -eq $true -and $_.Department -eq "Distribution Group" -and $_.Username -ne $alias} | Select Name,Username,Department,AccessRights,AutoMapped | Sort Name | FT -AutoSize -Wrap | Out-Default
        $MailboxUserlist | Where {$_.Enabled -eq $true -and $_.Department -ne "Distribution Group" -and $_.Username -ne $alias} | Select Name,Username,Department,AccessRights,AutoMapped | Sort Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header "Mailbox 'Send On Behalf' Rights (SysOps Managed)" $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $SendOnBehalfUserList | Where {$_.Enabled -eq $true -and $_.Department -eq "Distribution Group"} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        $SendOnBehalfUserList | Where {$_.Enabled -eq $true -and $_.Department -ne "Distribution Group"} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header "Mailbox 'Send As' Rights (SysOps Managed)" $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Yellow
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $SendAsUserList | Where {$_.Enabled -eq $true -and $_.Department -eq "Distribution Group"} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        $SendAsUserList | Where {$_.Enabled -eq $true -and $_.Department -ne "Distribution Group"} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Inbox Permissions (User Managed)' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $InboxUserlist | Where {$_.Enabled -eq $true -and $_.Department -eq "Distribution Group"} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        $InboxUserlist | Where {$_.Enabled -eq $true -and $_.Department -ne "Distribution Group"} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Calendar Permissions (User Managed)' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $CalendarUserlist | Where {$_.Enabled -eq $true -and $_.Department -eq "Distribution Group"} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        $CalendarUserlist | Where {$_.Enabled -eq $true -and $_.Department -ne "Distribution Group"} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        }
    }


function WinSam-Get-MailResourceInfo ($EmailAddress) {
    $RecipientInfo=$null

    if ($EmailAddress -notmatch '@' -or $EmailAddress -notmatch 'wpi.edu') {Write-Host "Please enter a valid WPI email address";WinSam-CheckEmail}

    $RecipientInfo = Get-Recipient $EmailAddress -ErrorAction SilentlyContinue
    if (!$RecipientInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "     ERROR: [$EmailAddress] There is no such Exchange object." -ForegroundColor Red
        Write-Host "     Please contact a Mail Administrator for more information." -ForegroundColor Red
        Return
        }

    <#
    RecipientType
    -------------
    MailUser     : GuestMailUser, MailUser
    UserMailbox  : UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, DiscoveryMailbox

    MailContact  : MailContact

    MailUniversalDistributionGroup : GroupMailbox, MailUniversalDistributionGroup, RoomList
    MailUniversalSecurityGroup     : MailUniversalSecurityGroup
    DynamicDistributionGroup       : DynamicDistributionGroup
    #>

    switch ($RecipientInfo.RecipientType) {
        'UserMailbox' {WinSam-Get-MailboxInfo $EmailAddress}
        'MailContact' {
            Write-Host ''
            $Contact = Get-MailContact $EmailAddress
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header 'NOTE: This is a mail contact, not a mailbox.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host ''
            Write-Host (WinSam-Write-Header 'Mail Contact Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host ''
            Write-Host "Contact Display Name: $($Contact.DisplayName)"
            Write-Host "Contact Alias       : $($Contact.Alias)"
            Write-Host "Contact Target      : $($Contact.PrimarySmtpAddress)"
            Write-Host ''
            Return
            }
        'MailUniversalDistributionGroup' {
            if ($RecipientInfo.RecipientTypeDetails -eq 'GroupMailbox') {WinSam-Get-UnifiedGroupInfo $EmailAddress}
            if ($RecipientInfo.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') {WinSam-Get-DistributionGroupInfo $EmailAddress}
            }
        'MailUniversalSecurityGroup' {WinSam-Get-DistributionGroupInfo $EmailAddress}
        'DynamicDistributionGroup'   {WinSam-Get-DistributionGroupInfo $EmailAddress}
        default {
            ## Make a note of the type of recipient with a note to contact the mail administrator
            Write-Host ''
            Write-Host '   There was an error with WinSam.  Please contact the administrator' -ForegroundColor Red
            Write-Host ''
            }
        }
    }