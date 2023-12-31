##################################################################################################################
#FUNCTIONS
##################################################################################################################

#Checks to see if user or group exists in the domain path specified
function isindomain($ldappath,$samaccountname){
    try{$result = Get-ADUser -server $ldappath -Identity $samaccountname}
    catch{}
    $result
}

function checkpermissions{
$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
    $currentuser = $env:username
    $grp_read   = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'U_WPI_Account_Maintenance_RO'}).name
    $grp_unlock = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'U_WPI_Account_Maintenance_UNLOCK' -or $_.name -eq 'G_WPI_Account_Maintenance' -or $_.name -eq 'Windows Team' -or $_.name -eq 'U_Helpdesk_Student_Staff'}).name
    $grp_write  = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'G_WPI_Account_Maintenance' -or $_.name -eq 'Windows Team'}).name
    $permissions='None'
    If ($grp_read -ne $null) {$permissions='ReadOnly'}
    If ($grp_unlock -ne $null) {$permissions='Unlock'}
    If ($grp_write -ne $null) {$permissions='Reset'}
    $permissions
$ElapsedTime = $ElapsedTime.Elapsed
Write-Host "Index update complete.  Total time: $ElapsedTime"
}

#Lists information for account to be looked up
function accountlookup($username){
    $dcs = (Get-ADDomainController -Filter * | Select Hostname)
    $lastlogon = $null
    $today = Get-Date
    foreach ($dc in $dcs) {
        $dcos = (Get-ADDomainController $dc.hostname | Select OperatingSystem).OperatingSystem
        If ($dcos -eq 'Windows Server 2008 R2 Enterprise') {
            $dchostname = $dc.hostname
            $dclogon = (Get-ADUser $username -Properties * -Server $dchostname | Select LastLogon).LastLogon
            if ($DCLOGON -ne $null) {
                if ($lastlogon -lt $DCLOGON) {
                    $lastlogon = $DCLOGON
                    }
                }
            }        
        }
    if ($lastlogon -eq '0' -or $lastlogon -eq $null) {$lastlogondate = ''}
        else {$lastlogondate = [DateTime]::FromFileTime($lastlogon)}

    $displayname = (get-aduser -identity $username -properties * | select displayname).displayname    
    $userprincipalname = (get-aduser -identity $username -properties * | select userprincipalname).userprincipalname
    $canonicalname = (get-aduser -properties * $username | select canonicalname).canonicalname
    
    $grp_NE  = (get-adprincipalgroupmembership $username | select name | where {$_.name -eq 'Nonemployees'}).name
    $accountexpireddate = (Get-ADUser $username -Properties * | Select AccountExpirationDate).AccountExpirationDate
    $passwordlastset = (get-aduser $username -properties passwordlastset | select passwordlastset).passwordlastset
    $LogonWorkstations = (Get-ADUser $username -Properties * | Select LogonWorkstations).LogonWorkstations
    
    $enabled = (get-aduser -Identity $username -Properties * | select enabled).enabled
    if ($enabled -ne 'True' -and $accountexpireddate -ne $null) {write-host 'This account (' $displayname '-' $userprincipalname ') was disabled (terminated) on '$accountexpireddate -Foregroundcolor Black -BackgroundColor Red}
    elseif ($enabled -ne 'True' -and $accountexpireddate -eq $null) {write-host 'This account (' $displayname '-' $userprincipalname ') is disabled (terminated).  There is no expiration date set for this account' -Foregroundcolor Black -BackgroundColor Red}
    else {
        if ($grp_NE -ne $null) {
            write-host ''
    	    write-host '===================================================================' -Foregroundcolor Black -BackgroundColor Yellow
            write-host 'This user is flagged as a Non-Employee (NE) and is restricted' -Foregroundcolor Black -BackgroundColor Yellow
            write-host 'from using the terminal server and from accessing CLA Media.' -Foregroundcolor Black -BackgroundColor Yellow
            write-host '===================================================================' -Foregroundcolor Black -BackgroundColor Yellow
            }
        if ($canonicalname -like 'admin.wpi.edu/Accounts/Students/*') {
            write-host ''
            write-host 'This is a Student Account.' -Foregroundcolor Black -BackgroundColor Green
            }
        if ($canonicalname -like 'admin.wpi.edu/Accounts/Work Study/*') {
            write-host ''
            write-host 'This is a Student Workstudy Account.  Access is limited to specific computers' -Foregroundcolor Black -BackgroundColor Yellow
            }
        if ($canonicalname -like 'admin.wpi.edu/Accounts/Retirees/*') {
            write-host ''
            write-host 'This is a Retired Employee, access is limited to OWA ONLY' -Foregroundcolor Black -BackgroundColor Yellow
            }
        if ($canonicalname -like 'admin.wpi.edu/Accounts/Vokes/*') {
            write-host ''
            write-host 'This is a limited contractor account, access is restricted to specific services only' -Foregroundcolor Black -BackgroundColor Yellow
            }
        $title = (get-aduser -identity $username -properties * | select title).title
        $description = (get-aduser -identity $username -properties * | select description).description
        $department = (get-aduser -identity $username -properties * | select department).department
        $office = (get-aduser -identity $username -properties * | select office).office
        $phone = (get-aduser -identity $username -properties * | select telephonenumber).telephonenumber
        write-host ''
        write-host 'Name                :' $displayname
        write-host 'Email               :' $userprincipalname
        write-host 'Title               :' $title
        write-host 'Department          :' $department
        write-host 'Office              :' $office
        write-host 'Phone               :' $phone
        write-host 'Description         :' $description
        Write-Host ''
        write-host 'Last Login          :' $lastlogondate
        if ($accountexpireddate -ne $null -and $accountexpireddate -gt $today) {
            write-host 'Account Expires On  :' $accountexpireddate -Foregroundcolor Yellow
            passwordexpireson $username
            accountlockout $username
            }
        elseif ($accountexpireddate -ne $null -and $accountexpireddate -lt $today) {
            write-host ''
            write-host 'This account expired on '$accountexpireddate -Foregroundcolor Black -BackgroundColor Red
            }
        else {
            passwordexpireson $username
            accountlockout $username
            }
        write-host ''
        Get-ExchangeStatus $username
        
        if ($LogonWorkstations -ne $null){
            Write-Host ''
            write-host 'This user has restricted login access.  They may only log onto the following computers:' -ForegroundColor Yellow
            Write-Host '=======================================================================================' -ForegroundColor Yellow
            Write-Host $LogonWorkstations -ForegroundColor Yellow
            }
        }
}

#Resets admin password
function resetpassword($username){
    $resetoption = ''
    while ($resetoption -ne 'y' -and $resetoption -ne 'n') {
        write-host ''
        $resetoption = read-host 'Are you sure you want to RESET the password for ' $username ' ? (y/n)'
        write-host ''
        if ($resetoption -ne 'y' -and $resetoption -ne 'n') {write-host 'Please specify one of the available options' -ForegroundColor Red}
        }
    if ($resetoption -eq 'y') {
        write-host 'Please follow all password length, history, and complexity requirements' -ForegroundColor Green
        write-host ''
        set-adaccountpassword -Identity $username -reset
        Set-ADUser -Identity $username -ChangePasswordAtLogon $true
        }
    else {
        write-host 'You have chosen NOT to reset the password for '$username
        sleep 3
        }
}

function GetFileshareMemberships($username){
    $displayname = (get-aduser -identity $username -properties * | select displayname).displayname
    $userprincipalname = (get-aduser -identity $username -properties * | select userprincipalname).userprincipalname
    $title = (get-aduser -identity $username -properties * | select title).title
    $description = (get-aduser -identity $username -properties * | select description).description
    $department = (get-aduser -identity $username -properties * | select department).department
    $office = (get-aduser -identity $username -properties * | select office).office
    $phone = (get-aduser -identity $username -properties * | select telephonenumber).telephonenumber
    $Groups = Get-ADPrincipalGroupMembership $username | Get-ADGroup -Properties * | Select Name,Description,GroupCategory | Sort Name

    cls
    write-host '========================= WinSamaritan 2.08 ========================' -foregroundcolor Cyan
    write-host ''
    write-host 'Name                :' $displayname
    write-host 'Email               :' $userprincipalname
    write-host 'Title               :' $title
    write-host 'Department          :' $department
    write-host 'Office              :' $office
    write-host 'Phone               :' $phone
    write-host 'Description         :' $description
    write-host ''
    write-host 'Distribution Groups:' -foregroundcolor Green
    write-host '===================================================================' -foregroundcolor Green
    $Groups | Where {$_.GroupCategory -eq 'Distribution'} | Select Name,Description

    write-host ''
    write-host 'Security Groups:' -foregroundcolor Green
    write-host '===================================================================' -foregroundcolor Green
    $Groups | Where {$_.GroupCategory -eq 'Security'} | Select Name,Description
    write-host ''                  
    write-host '===================================================================='

}

#checks for account lockouts and unlocks if necessary
function accountlockout($username){
    $permissionsstatus = checkpermissions
    If ($permissionsstatus -eq 'Unlock' -or $permissionsstatus -eq 'Reset') {$permissionsstatus=$true} else {$permissionsstatus=$false}
    $lockout = (get-aduser $username -properties accountlockouttime | select accountlockouttime).accountlockouttime
    if (!$lockout) {write-host 'Account Lockout     : Not locked out'}
    elseif ($lockout -and $permissionsstatus -eq $false){
        write-host ''
        write-host 'Account Lockout     : Locked Out' -ForegroundColor Red
        write-host ''
    }
    else {
        Unlock-ADAccount -Identity $username
        write-host ''
        write-host 'Account Lockout     : ' $username ' has been succesfully unlocked' -ForegroundColor green
        write-host ''
    }
}

function passwordexpireson($username){
    $maxpassage = (Get-ADDefaultDomainPasswordPolicy | select maxpasswordage).maxpasswordage
    $passwordlastset = (get-aduser $username -properties passwordlastset | select passwordlastset).passwordlastset
    if ($passwordlastset -ne $null) {
        $passwordexpires = $passwordlastset.adddays($maxpassage.days)
        }
    else {
        $passwordexpires = 0
        }
    $passwordexpired = (get-aduser $username -properties passwordexpired | select passwordexpired).passwordexpired
    $passwordneverexpires = (get-aduser $username -properties passwordneverexpires | select passwordneverexpires).passwordneverexpires
    if ($passwordneverexpires -eq 'True') {
        write-host ''
        write-host 'Password expiration : The password for ' $username ' is set to never expire' -foregroundcolor Red
        write-host ''
        }
    elseif ($passwordexpired -eq 'True' -and $passwordexpires -ne '0') {
        write-host ''
        write-host 'Password last set   : ' $passwordlastset
        write-host 'Password expiration : The password for ' $username ' has expired.' -foregroundcolor Red
        write-host ''
        }
    elseif ($passwordexpired -eq 'True' -and $passwordexpires -eq '0') {
        write-host ''
        write-host 'Password expiration : This account has the "User must change password at next logon" checkbox enabled.' -ForegroundColor Red
        write-host ''
        }
    else {
        write-host 'Password last set   :' $passwordlastset
        write-host 'Password expiration :' $passwordexpires
        }
}

function Get-ExchangeStatus ($username){
    $mailbox = $null
    $mailboxStats = $null
    $mailboxSize = $null
    $MailboxQuota = $null
    $MailboxPercentUse = $null
    $MailboxStorageLimitStatus = $null
    
    $mailbox = Get-mailbox $username -ErrorAction silentlycontinue
    
    
    If ($mailbox) {
        $forward = (get-mailbox $username | select forwardingaddress).forwardingaddress
        if (!$forward) {
            write-host 'Mail Delivery       : Mail rests at Exchange'
            $MailboxStats = Get-MailboxStatistics $username -ErrorAction silentlycontinue
            if ($mailboxstats){
                $MailboxSize = "{0:N2}" -f ($mailboxstats.TotalItemSize.Value.Tobytes()/1gb)
                If ($mailbox.UseDatabaseQuotaDefaults -eq $true) {
                    $MailboxQuota = "{0:N2}" -f ((Get-MailboxDatabase $mailbox.Database).ProhibitSendQuota.Value.Tobytes()/1gb)
                    }
                Else {
                    $MailboxQuota = "{0:N2}" -f ($mailbox.ProhibitSendQuota.Value.Tobytes()/1gb)
                    }
                $MailboxPercentUse = "{0:P0}" -f($MailboxSize/$MailboxQuota)
                $MailboxStorageLimitStatus = $mailboxstats.StorageLimitStatus
                write-host ''
                write-host 'Exchange Mailbox Information'
                write-host '============================'
                Switch ($MailboxStorageLimitStatus) {
                    "BelowLimit"      {Write-Host '   Mailbox Status   :' $MailboxStorageLimitStatus "($MailboxPercentUse)"}
                    "ProhibitSend"    {Write-Host '   Mailbox Status   :' $MailboxStorageLimitStatus -ForegroundColor Black -BackgroundColor Red}
                    "MailboxDisabled" {Write-Host '   Mailbox Status   :' $MailboxStorageLimitStatus -ForegroundColor Black -BackgroundColor Red}
                    "IssueWarning"    {Write-Host '   Mailbox Status   :' $MailboxStorageLimitStatus "($MailboxPercentUse)" -ForegroundColor Yellow -BackgroundColor Black}
                    default           {Write-Host '   Mailbox Status   :' $MailboxStorageLimitStatus "($MailboxPercentUse)"}
                    }
                write-host '   Mailbox Size     :' $MailboxSize 'GB'
                write-host '   Mailbox Quota    :' $MailboxQuota 'GB'
                }
            }
        else {write-host 'Mail Delivery       : Mail is forwarded to Unix'}
        }
    else {write-host 'Mail Delivery       : No Exchange mailbox exists for this user.' -ForegroundColor Red}
}
