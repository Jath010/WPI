##################################################################################################################
# WinSam.ps1
# Windows Samaritan Tool
# Features
#    - reset passwords
#    - unlock accounts 
#    - view directory Information
#
# Version 2.0.4
# Last Updated: 2011-07-29 - tcollins
#
# Change Log at end of script
##################################################################################################################



##################################################################################################################
#Powershell Modules
##################################################################################################################
$module_ad = (Get-Module | select name | where {$_.name -eq 'activedirectory'}).name
if ($module_ad -ne 'activedirectory') {import-module activedirectory}

$pssession = Get-PSSession | Select ConfigurationName
if ($pssession -eq $null) {
    Add-pssnapin Microsoft.Exchange.Management.Powershell.E2010
    .$env:ExchangeInstallPath\bin\RemoteExchange.ps1
    Connect-exchangeserver –auto
}

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
    $currentuser = $env:username
    $grp_read   = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'U_WPI_Account_Maintenance_RO'}).name
    $grp_unlock = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'U_WPI_Account_Maintenance_UNLOCK' -or $_.name -eq 'G_WPI_Account_Maintenance' -or $_.name -eq 'Windows Team'}).name
    $grp_write  = (get-adprincipalgroupmembership $currentuser | select name | where {$_.name -eq 'G_WPI_Account_Maintenance' -or $_.name -eq 'Windows Team'}).name
    $permissions='None'
    If ($grp_read -ne $null) {$permissions='ReadOnly'}
    If ($grp_unlock -ne $null) {$permissions='Unlock'}
    If ($grp_write -ne $null) {$permissions='Reset'}
    $permissions
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
    write-host '========================= WinSamaritan 2.0.4 ========================' -foregroundcolor Cyan
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
    $MailboxStats = Get-MailboxStatistics $username -ErrorAction silentlycontinue
    
    $MailboxSize = "{0:N2}" -f ($mailboxstats.TotalItemSize.Value.Tobytes()/1gb)
    If ($mailbox.UseDatabaseQuotaDefaults -eq $true) {
        $MailboxQuota = "{0:N2}" -f ((Get-MailboxDatabase $mailbox.Database).ProhibitSendQuota.Value.Tobytes()/1gb)
        }
    Else {
        $MailboxQuota = "{0:N2}" -f ($mailbox.ProhibitSendQuota.Value.Tobytes()/1gb)
        }
    $MailboxPercentUse = "{0:P0}" -f($MailboxSize/$MailboxQuota)
    $MailboxStorageLimitStatus = $mailboxstats.StorageLimitStatus
    
    if ($mailbox) {
        $forward = (get-mailbox $username | select forwardingaddress).forwardingaddress
        if (!$forward) {
            write-host 'Mail Delivery       : Mail rests at Exchange'
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
        else {write-host 'Mail Delivery       : Mail is forwarded to Unix'}
        }
    else {write-host 'Mail Delivery       : No Exchange mailbox exists for this user.' -ForegroundColor Red}
}

##################################################################################################################
#MAIN
##################################################################################################################

#While loop for entire process to check if you want to look up another account
$anotheraccount = 'y'

while ($anotheraccount -eq 'y') {
    cls
    Write-Host '========================= WinSamaritan 2.0.4 ========================' -foregroundcolor Cyan

    $admindomain = 'admin.wpi.edu'
    $UserName = ''
    $option = $null 
    $optionaction=$null                                 
    $permissionsstatus = $null
    $permissionsstatus = checkpermissions
    
    if ($permissionsstatus -ne 'None'){
        while (!$username) {
            $username = read-host -Prompt 'Please enter a username'
            $adminresult = isindomain $admindomain $username
            if ($adminresult -eq $null) {
                write-host ''
                Write-Host 'The account ' $username ' does not exist.' -foregroundcolor Red
                write-host ''
                $UserName = '' #Username failed so clear field and restart while loop
                }
            }
        cls
        Write-Host '========================= WinSamaritan 2.0.4 ========================' -foregroundcolor Cyan
        accountlookup $username
        write-host ''                  
        Write-Host '===================================================================='
        }
    else {write-host 'Your account ' $env:username ' does not have permissions to run WinSamaritan' -ForegroundColor Red}
    if ($permissionsstatus -eq 'Reset') {
        while ($optionaction -eq $null) {
			write-host ''
            $option = read-Host 'Please choose one of the following options:
            (1) Show file share group memberships
            (2) Reset Password
            (3) Look up another account
            (4) EXIT
            '
            write-host ''
            switch ($option) { 
                1 {$optionaction='FileshareMemberships'} 
                2 {$optionaction='ResetPassword'} 
                3 {$optionaction='AnotherAccount'} 
                4 {$optionaction='Exit'} 
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    elseif ($permissionsstatus -eq 'ReadOnly' -or $permissionsstatus -eq 'Unlock' ) {
        while ($optionaction -eq $null) {
			write-host ''
            $option = read-Host 'Please choose one of the following options:
            (1) Show file share group memberships
            (2) Look up another account
            (3) EXIT
            '
            write-host ''
            switch ($option) { 
                1 {$optionaction='FileshareMemberships'} 
                2 {$optionaction='AnotherAccount'} 
                3 {$optionaction='Exit'} 
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    else {exit}

    $anotheraccount = ''

    switch ($optionaction) {
        'AnotherAccount' {$anotheraccount='y'}
        'FileshareMemberships' {GetFileshareMemberships $username}
        'ResetPassword' {resetpassword $username}
        'Exit' {exit}
        }

    while ($anotheraccount -ne 'y' -and $anotheraccount -ne 'n') {
        write-host ''
        $anotheraccount = read-host 'Would you like to look up another account? (y/n)'
        write-host ''
        if ($anotheraccount -ne 'y' -and $anotheraccount -ne 'n') {write-host 'Please specify one of the available options' -ForegroundColor Red}
        }        
}


<#
##################################################################################################################
Change Log

2.0.4 - tcollins - 2011-07-29
      - Added Mailbox Quota Information
            - Mailbox Status - BelowLimit, IssueWarning, ProhibitSend, MailboxDisabled
            - Mailbox Size
            - Mailbox Quota
      - Removed references to "ADMIN Domain".  System now prompts "Please enter a username:"
2.0.3 - tcollins - 2011-06-22
      - Fixed date calculation on determining if an account is expired.
2.0.4 - tcollins - 2011-05-06
      - Changed path for Work Study Accounts from \Students\ to \Accounts\Work Study\
      - Added identification code to highlight if an account is a student account.
2.0.0 - tcollins - 2011-02-08
      - Changed LastLogon code to look at all Server 2008 R2 DCs and get the most current LastLogon date.
2.0.RC.4 - tcollins - 2011-02-03
      - Cleaned up text output on some menu items.
      - Changed how the PasswordLastSet field works.  Won't show field if 'Change Password on Next Logon' is checked 
      or if it is set to 'Never Expire'
      - Changed all double quotes to single quotes.  This allows for encasulating double quotes into the string.
      - Changed Exchange snapin error checking to check if an exchange session is active.  This will prevent the 
      exchange.ps1 from loading if being run from the EMS.
2.0.RC.3 - tcollins - 2011-02-01
      - Added logic to check for account expiration.  Suppressed lockout and password expiration information if account 
      is expired.
      - Added flag for student work study account.
      - Added code to show logon restrictions
2.0.RC.2.002 - tcollins - 2011-02-01
      - Added error checking for AD Module
      - Fixed error checking for Exchange Snapin - used -Registered switch with Get-PSSnapin command
2.0.RC.2 - tcollins - 2011-02-01
      - Updated info screen to note if a user is a non-employee.
      - Added option to see group memberships - output sorted by GroupCategory (Distribution, Security)
2.0.RC.1 - tcollins - 2011-01-31
      - Published to SysAdmin Terminal Server for testing by Marie and Chris
      - Fixed Bug: Removed 'Password Last Set on *****' print out from password reset field.
##################################################################################################################
#>