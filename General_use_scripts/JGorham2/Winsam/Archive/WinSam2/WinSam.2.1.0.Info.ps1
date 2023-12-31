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
        [ValidateLength(1,16)]
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
    $Global:ADInfo = Get-ADUser $username -Properties * -Server $Global:DCServerName -ErrorAction "SilentlyContinue"

    if (!$Global:ADInfo) {
        Write-Host "       WARNING : User does not exist" -ForegroundColor Yellow
        return
        }
    if ($Global:today.AddHours(1) -lt (Get-Date)) {$Global:today = Get-Date}
    $Global:AccountExpirationDate = $Global:ADInfo.AccountExpirationDate
    $Global:AccountEnabled = $global:ADinfo.enabled
    $Global:AccountGroups = Get-ADPrincipalGroupMembership $username -Server $Global:DCServerName
    $Global:AccountNEStatus  = ($Global:AccountGroups | Where {$_.Name -eq 'Nonemployees'}).Name
    WinSam-Get-AccountStatus
    $Global:AccountLockoutStatus = (Get-ADUser $username -Properties AccountLockoutTime -Server $Global:DCServerName).AccountLockoutTime

########TESTING#######
#$AccessLevel = "ReadOnly"
#Write-Host "Your access level has been reset to: $AccessLevel"
#$Global:AccountEnabled = $true
#Write-Host "The AccountEnabled value has been reset to: $Global:AccountEnabled"
########TESTING#######

    if ($Global:AccountEnabled -or $AccessLevel -eq "SysAdmin") {$Global:mailbox = Get-Mailbox $username -ErrorAction "SilentlyContinue"}
    else {Write-Host "Did not get mailbox info.  Account Status: $($Global:AccountEnabled)  Access Level: $AccessLevel" -ForegroundColor magenta}

    Write-Host ''
    Write-Host '     General User Information                                                        ' -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor 
    Write-Host '-------------------------------------------------------------------------------------'
    WinSam-Get-InfoBanner
    Write-Host 'Name                :' $Global:ADInfo.DisplayName
    Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
    
    if ($Global:AccountEnabled -or $AccessLevel -eq "SysAdmin") {
        Write-Host ''
        Write-Host 'Title               :' $Global:ADInfo.Title
        If ($Global:AccountType -ne 'Student') {
            Write-Host 'Department          :' $Global:ADInfo.Department
            Write-Host 'Office              :' $Global:ADInfo.Office
            Write-Host 'Phone               :' $Global:ADInfo.telephoneNumber
            }
        Write-Host 'Description         :' $Global:ADInfo.description
        Write-Host ''
        If ($mailbox) {
            If ($Global:AccountType -eq 'Student') {Write-Host 'Student Status      :' $Global:mailbox.CustomAttribute13}
            Write-Host 'WPI ID              :' $Global:ADinfo.EmployeeID
            Write-Host 'PIDM                :' $Global:ADinfo.EmployeeNumber
            Write-Host ''
        }    
        Write-Host '     Windows Account Information                                                     ' -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor 
        Write-Host '-------------------------------------------------------------------------------------'
        $logonInfo = WinSam-Get-LastLogonDate $username
        If ($logonInfo.LastLogon){Write-Host "Last Login          : $($logonInfo.LastLogon) ($($logonInfo.LogonServer))"}
        Else {Write-Host "Last Login          : Never logged in or too long since last login." -ForegroundColor Yellow}

        if (!$Global:AccountLockoutStatus) {write-host 'Account Lockout     : Not locked out'}
        elseif ($AccessLevel -eq "SysAdmin" -or $AccessLevel -eq "PasswordReset" -or $AccessLevel -eq "Unlock") {Unlock-ADAccount -Identity $username; Write-Host "Account Lockout     : $username has been succesfully unlocked" -ForegroundColor green}
        else {Write-Host 'Account Lockout     : Locked Out' -ForegroundColor Red}

        if ($AccessLevel -eq 'SysAdmin'){Write-Host "Account Created     : $($Global:ADInfo.whenCreated)"}
        Write-Host ''

        WinSam-Get-PasswordExpiration $username
        
    	if ($AccountExpirationDate -and $AccountExpirationDate -gt $Global:today) {Write-Host "Account Expires On  : $AccountExpirationDate" -ForegroundColor Yellow}
        elseif ($AccountExpirationDate -and $AccountExpirationDate -lt $Global:today) {Write-Host "This account expired on $AccountExpirationDate" -ForegroundColor Black -BackgroundColor Red}
        Write-Host ''

        $LogonWorkstations = WinSam-Get-LogonWorkstations
        if ($LogonWorkstations) {
            Write-Host ''
            Write-Host 'This user has restricted PC access.  They may only log onto the following computers: ' -ForegroundColor Black -BackgroundColor Yellow
            Write-Host '-------------------------------------------------------------------------------------' -ForegroundColor Yellow
            Write-ColorOutput Yellow ($LogonWorkstations)
            Write-Host ''
            }
        
        #Write-Host 'Mailbox Information'# -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor 
        Write-Host '     Mailbox Information                                                             ' -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor 
        Write-Host '-------------------------------------------------------------------------------------'
        If ($Global:mailbox) {if (!$Global:mailbox.forwardingaddress) {write-host 'Mail Delivery       : Mail rests at Exchange'}
            else {Write-Host 'Mail Delivery       : Mail is forwarded to Unix'}
            if ($AccessLevel -eq 'SysAdmin'){
                Write-Host 'Mailbox Database    :' $Global:mailbox.Database
                WinSam-Get-MailboxStats $username
                }
            }
        else {write-host 'Mail Delivery       : No Exchange mailbox exists for this user.'}
    }
    Write-Host ''
    Write-Host '-------------------------------------------------------------------------------------'
    #End of Info function
    }