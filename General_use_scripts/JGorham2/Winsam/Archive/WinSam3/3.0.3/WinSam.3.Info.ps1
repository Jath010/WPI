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
    $mailboxStats = $null
    $mailboxSize = $null
    $MailboxQuota = $null
    $MailboxPercentUse = $null
    $MailboxStorageLimitStatus = $null
    $lastlogon = $null

    #Check to see if Account Exists    
    $Global:ADInfo = Get-ADUser $username -Properties AccountExpirationDate,AccountLockoutTime,CanonicalName,Department,Description,DisplayName,DistinguishedName,EmployeeID,EmployeeNumber,LastLogon ,LockedOut,LockoutTime,LogonWorkstations,Office,PasswordExpired,PasswordLastSet,TelephoneNumber,Title,UserPrincipalName,WhenCreated -Server $dchostname -ErrorAction "SilentlyContinue"

    if (!$Global:ADInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : User '$username' does not exist" -ForegroundColor Yellow
        Write-Host ''
        return
        }
    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    $AccountExpirationDate = $Global:ADInfo.AccountExpirationDate
    $Global:AccountEnabled = $global:ADinfo.enabled
    $Global:AccountGroups = Get-ADPrincipalGroupMembership $username -Server $dchostname
    $Global:AccountNEStatus  = ($Global:AccountGroups | Where {$_.Name -eq 'Nonemployees'}).Name
    WinSam-Get-AccountStatus
    $Global:AccountLockoutStatus = (Get-ADUser $username -Properties AccountLockoutTime -Server $dchostname).AccountLockoutTime

    if (($Global:AccountType -ne 'Other' -and $Global:AccountEnabled) -or $UserAccessLevel -eq "SysAdmin") {$Global:mailbox = Get-Mailbox $username -ErrorAction "SilentlyContinue"}
    #else {Write-Host "Did not get mailbox info.  Account Status: $($Global:AccountEnabled)  Access Level: $UserAccessLevel" -ForegroundColor magenta}  ## Removed - this may be old debug code.

    Write-Host ''
    Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
    Write-Host ''
    If ($Global:AccountType -ne 'Student') {
        Write-Host 'Title               :' $Global:ADInfo.Title
        Write-Host 'Department          :' $Global:ADInfo.Department
        }
    
    if (($Global:AccountType -ne 'Other' -and $Global:AccountEnabled) -or $UserAccessLevel -eq "SysAdmin") {
        If ($Global:AccountType -ne 'Student') {
            Write-Host 'Office              :' $Global:ADInfo.Office
            Write-Host 'Phone               :' $Global:ADInfo.telephoneNumber
            }
        Write-Host 'Description         :' $Global:ADInfo.description
        }
        Write-Host ''        
        #If ($mailbox) {If ($Global:AccountType -eq 'Student') {Write-Host 'Student Status      :' $Global:mailbox.CustomAttribute13}}   ## Removed as the field is not consistantly populated.
        Write-Host 'WPI ID              :' $Global:ADinfo.EmployeeID
        Write-Host 'PIDM                :' $Global:ADinfo.EmployeeNumber
        Write-Host ''

    if (($Global:AccountType -ne 'Other' -and $Global:AccountEnabled) -or $UserAccessLevel -eq "SysAdmin") {
        Write-Host (WinSam-Write-Header 'Windows Account Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $logonInfo = WinSam-Get-LastLogonDate $username
        If ($logonInfo.LastLogon){Write-Host "Last Login          : $($logonInfo.LastLogon.ToString("g")) ($($logonInfo.LogonServer))"}
        Else {Write-Host "Last Login          : Never logged in or too long since last login." -ForegroundColor Yellow}

        if (!$Global:AccountLockoutStatus) {write-host 'Account Lockout     : Not locked out'}
        elseif ($UserAccessLevel -eq "SysAdmin" -or $UserAccessLevel -eq "PasswordReset" -or $UserAccessLevel -eq "Unlock") {Unlock-ADAccount -Identity $username; Write-Host "Account Lockout     : $username has been succesfully unlocked" -ForegroundColor green}
        else {Write-Host 'Account Lockout     : Locked Out' -ForegroundColor Red}

        if ($UserAccessLevel -eq 'SysAdmin'){Write-Host "Account Created     : $($Global:ADInfo.whenCreated)" -ForegroundColor Gray }
        Write-Host ''

    	if ($AccountExpirationDate -and $AccountExpirationDate -gt $Global:today) {Write-Host "Account Expires On  : $($AccountExpirationDate.ToString("g"))" -ForegroundColor Yellow}
        elseif ($AccountExpirationDate -and $AccountExpirationDate -lt $Global:today) {Write-Host (WinSam-Write-Header "This account expired on $($AccountExpirationDate.ToString("g"))" $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Red}

        $PasswordStatus = WinSam-Get-PasswordExpiration $username

        if (($PasswordStatus.PasswordExpiration -is [datetime]) -and $AccountExpirationDate -and $PasswordStatus.PasswordExpiration -gt $AccountExpirationDate) {$PasswordExpiration = $AccountExpirationDate}
        else {$PasswordExpiration = $PasswordStatus.PasswordExpiration}

        if ($PasswordStatus.PasswordNeverExpires -eq 'True') {
            Write-Host "Password last set   : $PasswordStatus.PasswordLastSet"
            Write-Host "Password expiration : The password for $username is set to never expire" -ForegroundColor Red
            }
        elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordExpiration -ne '0') {
            Write-Host "Password last set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
            Write-Host "Password expiration : The password for $username expired on $($PasswordExpiration.ToString("g"))." -ForegroundColor Red
            }
        elseif ($PasswordStatus.PasswordExpired -eq 'True' -and $PasswordExpiration -eq '0') {
            Write-Host 'Password Warning     : This account has a temporary password set and must change their' -ForegroundColor Red
            Write-Host '                      password upon next logon' -ForegroundColor Red
            }
        else {
            Write-Host "Password last set   : $($PasswordStatus.PasswordLastSet.ToString("g"))"
            Write-Host "Password expiration : $($PasswordExpiration.ToString("g"))"
            }
        Write-Host ''

        $LogonWorkstations = WinSam-Get-LogonWorkstations
        if ($LogonWorkstations) {
            Write-Host ''
            Write-Host 'This user has restricted PC access.  They may only log onto the following computers: ' -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '' $MenuLength -Line) -ForegroundColor Yellow
            Write-ColorOutput Yellow ($LogonWorkstations)
            }
        
        Write-Host (WinSam-Write-Header 'Mailbox Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor        
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        If ($Global:mailbox) {if (!$Global:mailbox.forwardingaddress) {write-host 'Mail Delivery       : Mail rests at Exchange'}
            else {Write-Host 'Mail Delivery       : Mail is forwarded to Unix'}
            WinSam-Get-MailboxStats $username
            if ($UserAccessLevel -eq 'SysAdmin'){Write-Host 'Mailbox Database    :' $Global:mailbox.Database -ForegroundColor Gray}
            }
        else {write-host 'Mail Delivery       : No Exchange mailbox exists for this user.' -ForegroundColor Yellow}
    }
    #End of Info function
    }

function WinSam-Get-GroupInfo {
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
    $GroupMemberOf = $null

    #Check to see if Account Exists    
    $GroupInfo = Get-ADGroup $GroupName -Properties DisplayName,DistinguishedName,Description,GroupCategory,GroupScope,ManagedBy,WhenCreated -Server $dchostname -ErrorAction "SilentlyContinue"
    
    if (!$GroupInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "       WARNING : Group '$GroupName' does not exist" -ForegroundColor Yellow
        write-host ''
        return
        }
    $GroupMembers = $GroupInfo | Get-ADGroupMember -Server $dchostname -ErrorAction "SilentlyContinue"
    #$Global:GroupMemberOf = Get-ADPrincipalGroupMembership $GroupName -Server $dchostname -ErrorAction "SilentlyContinue"  ### Not using in code at this time.  May eventually add programming to show what groups a group is a member of.

    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    Write-Host ''
    Write-Host (WinSam-Write-Header 'General Group Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host 'Name                :' $GroupInfo.Name
    Write-Host 'Group ID            :' $GroupInfo.SamAccountName
    Write-Host 'Description         :' $GroupInfo.Description
    Write-Host 'Group Type          :' $GroupInfo.GroupCategory
    if ($UserAccessLevel -eq "SysAdmin") {Write-Host 'Created             :' $GroupInfo.WhenCreated -ForegroundColor Gray}
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Group Membership' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host ''
    Write-Host (WinSam-Write-Header 'User Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66) -Line)
    $GroupMembers | Where {$_.ObjectClass -eq 'user'} | Select Name,@{Name="Username";Expression={$_.SamAccountName}} | Sort Name | Out-Default
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Group Members' ($MenuLength*0.66) -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' ($MenuLength*0.66)  -Line)
    $GroupMembers | Where {$_.ObjectClass -eq 'group'} | Select Name,@{Name="Username";Expression={$_.SamAccountName}} | Sort Name | Out-Default
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
            }
        else {
            Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
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
            }
        }
    else {
        Write-Host ''
        Write-Host "The computer $computername is not reachable" -ForegroundColor Black -BackgroundColor Red
        Write-Host ''
        Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
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

function WinSam-Get-MailboxInfo ($alias) {
    function WPI-Get-MailboxStats {
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
            [ValidateLength(1,32)]
            [string]$username
            )
        $ErrorActionPreference = "SilentlyContinue"
        
        $mailbox = $null
        $mailboxStats = $null
        $mailboxSize = $null
        $MailboxQuota = $null
        $MailboxPercentUse = $null
        $MailboxStorageLimitStatus = $null

        $mailbox = Get-Mailbox $username -DomainController $dchostname -ErrorAction SilentlyContinue
        if (!$mailbox){return}
        $MailboxStats = Get-MailboxStatistics $username -DomainController $dchostname
        if (!$mailboxstats){return}

        $MemberType = (($MailboxStats| Get-Member |Where {$_.Name -eq "TotalItemSize"}).TypeName).Split(".")
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
    $ErrorActionPreference = "SilentlyContinue"
        
    #Get Mailbox Information
    $mailbox = Get-Mailbox $alias -DomainController $dchostname -ErrorAction SilentlyContinue
    $Inbox = "$($mailbox.Name):\Inbox" 
    $Calendar = "$($mailbox.Name):\Calendar" 
    
    #Process Mailbox
    If (!$mailbox) {Write-Host '';Write-Host '';Write-Host "     ERROR: The mailbox '$alias' doesn't exist" -ForegroundColor Yellow}
    Else {
        $out           = $null    
        
        $MailboxUsers  = $null
        $InboxUsers    = $null
        $CalendarUsers = $null
        $Senders       = $null
        
        $MailboxUserlist        = @()
        $SendOnBehalfUserList   = @()
        $SendAsUserList         = @()
        $InboxUserlist          = @()
        $CalendarUserlist       = @()

    	$MailboxUsers  = Get-MailboxPermission $alias -DomainController $dchostname | Where {$_.IsInherited -ne $true -and $_.User -notlike "NT AUTHORITY\SELF"}
        $SendOnBehalf  = $mailbox.GrantSendOnBehalfTo
        $SendAs        = Get-Mailbox $alias | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")}
        $InboxUsers    = Get-MailboxFolderPermission -Identity $Inbox -DomainController $dchostname
        $CalendarUsers = Get-MailboxFolderPermission -Identity $Calendar -DomainController $dchostname
        

        foreach($user in $MailboxUsers) {
            $username = $null
            $ADInfo   = $null
            
            $username = ($user.User).ToString().Split("\")
            If ($username[0] -eq "ADMIN") {
                $ADInfo = Get-ADUser -Identity $username[1] -Properties Name,SamAccountName,Department,Enabled -Server $dchostname
                
            	$out = New-Object PSObject
            	$out | add-member noteproperty Name $ADInfo.Name
            	$out | add-member noteproperty Username $ADInfo.samAccountName
            	$out | add-member noteproperty Department $ADInfo.Department
                $out | add-member noteproperty AccessRights $user.AccessRights
                $out | add-member noteproperty Enabled $ADInfo.Enabled
            	$MailboxUserlist += $out        
                }
            }

        foreach ($user in $SendOnBehalf) {
            $username = $null
            $ADInfo   = $null
            
            $username = (Get-Mailbox $user -DomainController $dchostname -ErrorAction SilentlyContinue).Alias
            $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname
        	$out = New-Object PSObject
            $out | add-member noteproperty Name $ADInfo.Name
            $out | add-member noteproperty Username $ADInfo.samAccountName
            $out | add-member noteproperty Department $ADInfo.Department
            $out | add-member noteproperty Enabled $ADInfo.Enabled
        	$SendOnBehalfUserList += $out        
            }

        foreach ($user in $SendAs) {
            $username = $null
            $ADInfo   = $null

            $username = $user.User.Replace("ADMIN\","")
            $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname
        	$out = New-Object PSObject
            $out | add-member noteproperty Name $ADInfo.Name
            $out | add-member noteproperty Username $ADInfo.samAccountName
            $out | add-member noteproperty Department $ADInfo.Department
            $out | add-member noteproperty Enabled $ADInfo.Enabled
        	$SendAsUserList += $out        
            }

        foreach($user in $InboxUsers) {
            $username = $null
            $ADInfo   = $null
            
            $out = New-Object PSObject
            if ($user.User -eq 'Default' -or $user.User -eq 'Anonymous') {
                $out | add-member noteproperty Name $user.user
                $out | add-member noteproperty Username $null
                $out | add-member noteproperty Department $null
                $out | add-member noteproperty Enabled $true

                }
            else {
                if ($user.User -match "NT User:") {$username = $user.User.Replace("NT User:ADMIN\","")}
                else {$username = (Get-Mailbox $user.User -DomainController $dchostname -ErrorAction SilentlyContinue).alias}
                $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname

                $out | add-member noteproperty Name $ADInfo.Name
                $out | add-member noteproperty Username $ADInfo.samAccountName
                $out | add-member noteproperty Department $ADInfo.Department
                $out | add-member noteproperty Enabled $ADInfo.Enabled
            }
            $out | add-member noteproperty AccessRights $user.AccessRights

        	$InboxUserlist += $out        
            }

        foreach($user in $CalendarUsers) {
            $username = $null
            $ADInfo   = $null
            
            $out = New-Object PSObject
            if ($user.User -eq 'Default' -or $user.User -eq 'Anonymous') {
                $out | add-member noteproperty Name $user.user
                $out | add-member noteproperty Username $null
                $out | add-member noteproperty Department $null
                $out | add-member noteproperty Enabled $true

                }
            else {
                if ($user.User -match "NT User:") {$username = $user.User.Replace("NT User:ADMIN\","")}
                else {$username = (Get-Mailbox $user.User -DomainController $dchostname -ErrorAction SilentlyContinue).alias}
                $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname

                $out | add-member noteproperty Name $ADInfo.Name
                $out | add-member noteproperty Username $ADInfo.samAccountName
                $out | add-member noteproperty Department $ADInfo.Department
                $out | add-member noteproperty Enabled $ADInfo.Enabled
            }
            $out | add-member noteproperty AccessRights $user.AccessRights

        	$CalendarUserlist += $out        
            }

        Write-Host ''
        Write-Host (WinSam-Write-Header 'General Mailbox Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host "Mailbox Display Name: $($mailbox.Name)"
        Write-Host "Mailbox Alias       : $($mailbox.Alias)"
        if ($UserAccessLevel -eq 'SysAdmin'){Write-Host "Mailbox Database    : $($mailbox.Database)" -ForegroundColor Gray}
        WPI-Get-MailboxStats $alias
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Mailbox Access Rights' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $MailboxUserlist | Where {$_.Enabled -eq $true} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header "Mailbox 'Send On Behalf' Rights" $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $SendOnBehalfUserList | Where {$_.Enabled -eq $true} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header "Mailbox 'Send As' Rights" $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $SendAsUserList | Where {$_.Enabled -eq $true} | Select Name,Username,Department | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Inbox Permissions' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $InboxUserlist | Where {$_.Enabled -eq $true} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Calendar Permissions' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        $CalendarUserlist | Where {$_.Enabled -eq $true} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default
        }
    }