<###################################################################################################################
Purpose       : Review all active accounts and audit account and password expiration.
                    Azure AD does not replicate password expiration or account expiration and as a result, accounts that should be
                    non-functional continue to work with an active password hash for Azure AD and any system that uses Azure SSO
                    Accounts that have a password that is expired more than 7 days will be scrambled to force a password change
                        - This can be done by SSPR
                    Accounts that have an account expiration date that has passed will be marked as disabled.
                        - This must be restored by SysOps
Target System : Production Domain (admin.wpi.edu)
--------------
Written by     : Tom Collins (tcollins)
Last Update    : 03/19/19
Updated by     : Tom Collins (tcollins)
###################################################################################################################>
Clear-Host
Start-Transcript -Path "D:\wpi\batch\IAM\AuditExpirations\Transcripts\Transcript-$((Get-Date).ToString('yyyy-MM-dd')).txt"
Write-Host "Script Begin $(get-date)"


########################
##  FUNCTIONS
########################

function Scramble-Password {
    Param (
    [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
    [ValidateLength(1,20)]
    [string]$username
    )


    function Create-ComplexPassword {
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
    
    $ADUser=$null;$password=$null;$dn=$null
    $ldapPath=$null;$objUser=$null

    $password = Create-ComplexPassword 16
    
    $ADUser = Get-ADUser $username
    if ($ADUser) {
        $dn = $ADUser.DistinguishedName

        $ldapPath =  "LDAP://"+$dn
        $objUser = [ADSI] $ldapPath
        $objUser.SetPassword($password)
        $objUser.SetInfo()
        }
    }

########################
##  Main Code
########################
Import-Module ActiveDirectory

## Domain Specific paths
$DNSDomain   = 'wpi.edu'
$DN_Domain   = 'DC=wpi,DC=edu'

## OU Locations
$OU_Accounts       = "OU=Accounts,DC=admin,$DN_Domain"
$OU_Alumni         = "OU=Alumni,OU=Accounts,DC=admin,$DN_Domain"
$OU_Disabled       = "OU=Disabled,OU=Accounts,DC=admin,$DN_Domain"
$OU_Employees      = "OU=Employees,OU=Accounts,DC=admin,$DN_Domain"
$OU_LeaveOfAbsence = "OU=Leave Of Absence,OU=Accounts,DC=admin,$DN_Domain"
$OU_NoOffice365Sync= "OU=No Office 365 Sync,OU=Accounts,DC=admin,$DN_Domain"
$OU_OtherAccounts  = "OU=Other Accounts,OU=Accounts,DC=admin,$DN_Domain"
$OU_Privileged     = "OU=Privileged,OU=Accounts,DC=admin,$DN_Domain"
$OU_Retirees       = "OU=Retirees,OU=Accounts,DC=admin,$DN_Domain"
$OU_ResourceMailbox= "OU=Resource Mailboxes,OU=Other Accounts,OU=Accounts,DC=admin,$DN_Domain"
$OU_Services       = "OU=Services,OU=Accounts,DC=admin,$DN_Domain"
$OU_Students       = "OU=Students,OU=Accounts,DC=admin,$DN_Domain"
$OU_TestAccounts   = "OU=Other Accounts,OU=Accounts,DC=admin,$DN_Domain"
$OU_Vokes          = "OU=Vokes,OU=Accounts,DC=admin,$DN_Domain"
$OU_WorkStudy      = "OU=Work Study,OU=Accounts,DC=admin,$DN_Domain"

## Get list of users from all active user OUs.
$ADUsers  = Get-ADUser -Filter * -Properties DisplayName,Department,Title,AccountExpirationDate,PasswordExpired,PasswordLastSet,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12 -SearchBase $OU_Employees
$ADUsers += Get-ADUser -Filter * -Properties DisplayName,Department,Title,AccountExpirationDate,PasswordExpired,PasswordLastSet,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12 -SearchBase $OU_Students
$ADUsers += Get-ADUser -Filter * -Properties DisplayName,Department,Title,AccountExpirationDate,PasswordExpired,PasswordLastSet,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12 -SearchBase $OU_Alumni
$ADUsers += Get-ADUser -Filter * -Properties DisplayName,Department,Title,AccountExpirationDate,PasswordExpired,PasswordLastSet,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12 -SearchBase $OU_LeaveOfAbsence
$ADUsers += Get-ADUser -Filter * -Properties DisplayName,Department,Title,AccountExpirationDate,PasswordExpired,PasswordLastSet,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12 -SearchBase $OU_Retirees

## Process raw data to identify accounts to target
$EnabledUsers = $ADUsers | Where {$_.Enabled}
$ExpiredUsers = $EnabledUsers | Where {$_.AccountExpirationDate -and $_.AccountExpirationDate -lt (get-date) -and $_.AccountExpirationDate -gt (get-date).AddDays(-14)} 
$PasswordExpired = $EnabledUsers | Where {$_.PasswordExpired -and $_.PasswordLastSet}

## Print list of Expired users for transcript
$ExpiredUsers | Select DisplayName, SamAccountName,Department,Title,AccountExpirationDate,DistinguishedName | FT | Out-Default

## Process account expirations.  
$ExpiredUsers | Set-ADUser -Enabled:$false

## Process password expirations.
$list=@()
Foreach ($user in $PasswordExpired) {
    $MaxPasswordAge=$null;$PasswordExpires=$null;$DaysExpired=$null
    
    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy | select MaxPasswordAge).MaxPasswordAge
    if ($user.PasswordLastSet) {$PasswordExpires = $user.PasswordLastSet.adddays($MaxPasswordAge.days)}
    $DaysExpired = (get-date) - $PasswordExpires

    if ($DaysExpired.TotalDays -gt 7) {
        Write-Host "The user [$($user.SamAccountName)] has a password that has been expired longer than 7 days [$($DaysExpired.Days)]"
        Scramble-Password $user.SamAccountName       
        Set-ADUser $user.SamAccountName -ChangePasswordAtLogon:$true
        }
    }

Stop-Transcript