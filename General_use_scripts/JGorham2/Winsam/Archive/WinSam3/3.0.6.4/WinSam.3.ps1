<##################################################################################################################
WinSam.ps1
Windows Samaritan Tool

Features
    - reset passwords
    - unlock accounts 
    - view directory Information

Version      : 3.0.6.4
Last Updated : 2014-11-21 - tcollins
Change Log   : Bottom of script

Libraries:
    - GlobalVariables: This contains variables to update the look and feel of the program as well as several global
        environment variables.
    - Functions: This contains all of the support functions that are used in the main functions
    - Info: This contains the primary info gathering functions
    - Menu: Cotnains the menu functions

##################################################################################################################>

##################################################################################################################
# Powershell Modules
##################################################################################################################
If ($Host.Name -ne "Windows PowerShell ISE Host"){
    $module_ad = (Get-Module | Where {$_.name -eq 'activedirectory'}).name
    if ($module_ad -ne 'activedirectory') {import-module activedirectory}

    if ((Get-PSSession | Where {$_.ConfigurationName -match "Microsoft.Exchange"} -ErrorAction SilentlyContinue) -eq $null ) {
        $Global:CASServer = $null
        $ExchangeServers = 'COUPE','FASTBACK','SN95','S281','S197','EXCHCAS-03'
        if (!($ExchangeServers -contains $localhost)) {
            foreach ($server in $ExchangeServers) {if (Test-Connection -Count 1 -BufferSize 15 -Delay 1 -ComputerName "$server.admin.wpi.edu") {$Global:CASServer = "$server.admin.wpi.edu";break}}

            if ($CASServer) {$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$CASServer/PowerShell";Import-PSSession $session}
            else {Write-Host "No Exchange Server is avaialble at this time.  Exchange Remote Shell cannot be loaded" -ForegroundColor Red}
            }
        }
    }

##################################################################################################################
# Includes
##################################################################################################################
$global:ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition

. $ScriptPath\WinSam.3.GlobalVariables.ps1
. $ScriptPath\WinSam.3.Functions.ps1
. $ScriptPath\WinSam.3.Info.ps1
. $ScriptPath\WinSam.3.Menu.ps1

<## DEBUG ##############
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
#######################>

##################################################################################################################
# Main Program
##################################################################################################################
Clear-Host

# Output Loading WinSam banner.  The Access level function takes about 10-15 seconds to process.
Write-Host ''
Write-Host ''
Write-Host (WinSam-Write-Header 'Loading Windows Samaritan...' $MenuLength -Center) -ForegroundColor Black -BackgroundColor Yellow

# Get the Access Levels for the current user.
WinSam-Get-AccessLevel $currentuser

#DEBUG Code - This section allows for forcing/overiding the Access Levels
#$UserAccessLevel = 'Unlock'  #DEBUG   

<#DEBUG CODE ##############
Write-Host '     Access Level Options:'
write-host '     (1) System Administrator'
write-host '     (2) Password Reset'
write-host '     (3) Unlock'
write-host '     (4) Read Only'
write-host '     (5) No Access'
$AccessOption = read-Host 'Choose an option'
Write-Host ''
switch ($AccessOption) { 
    1 {$AccessLevel ='SysAdmin'} 
    2 {$AccessLevel ='PasswordReset'}
    3 {$AccessLevel ='Unlock'} 
    4 {$AccessLevel ='ReadOnly'} 
    5 {$AccessLevel ='NoAccess'} 
    default {
        $AccessOption=$null
        Write-Host 'Please specify one of the available options' -foregroundcolor Red
        }
    }
## End DEBUG CODE ##############>


if ($UserAccessLevel -eq "NoAccess") {
    Write-Host "You are not authorized to access the Windows Samaritan interface" -ForegroundColor Black -BackgroundColor Red
    #Sleep 30
    Sleep 1
    return
    }

WinSam-Menu-Main


<#
    Other Notes
        * Tested with Roger's Account
            * Need to show if Service Accounts are enabled
            * Need to show if Service Accounts are locked.
            * Shows only Name and Email address for service account
    
    Requests
        * Oracle Add-ons
            * Lookup by ID Number
            * Reset PIN by ID Number?
        * Computer Info requests
            * Query Administrators on a system
                * Filter out default groups
            * Look at list of profiles on system (Directory query?)
            * List of updates on PC (WMI?)
        * Mailbox Info Requests
            * [COMPLETED] Mailbox - Send As
        * New Features
            * [COMPLETED]Group Memberships
            * Groups: Add GroupID to info
    Errors:
#>

<##################################################################################################################
Change Log

3.0.6.4 - 2015-04-13 - tcollins
    * Group Info Changes
        * Added code to detect and note if a security group is also a distribution list
        * Added code to show all managers of the distribution list
3.0.6.3 - 2014-11-21 - tcollins
    * Added new Password Change Menu to allow for different options for password changes
    * Updated Password Change options to include the following options
        * Set manual password and require change at login
        * Set manual password and do not require change at login
        * Set random password and require change at login
    * Fix bug in WinSam-Reset-Password (referenced $PasswordStatus.Expiration instead of $PasswordStatus.PasswordExpiration)
    * Fix bug in Mailing List Information - Was not showing groups previously
3.0.6.2 - 2014-11-21 - tcollins
    * Fixed bug in Group Info - unable to search by alias - added step to try Get-DistributionGroup if Get-ADGroup fails
3.0.6.1 - 2014-11-21 - tcollins
    * Updated Local Admin section to reformat how the groups and users are shown.
3.0.6 - 2014-11-18 - tcollins
    * Computer Option changes
        * Updated Computer Options to include the ability to get local admin information.
        * Changed it so that Computer Options are available for all users.  Will only show the information that is visible for the privs of the user.
    * Access Levels
        * Removed checking for ComputerAccess level now that the Computer Info system is redesigned.
    * Added new Beta script handling.  Now a switch can be turned on to allow users to move to a Beta tree to run the script without updating everyone.
3.0.5 - 2014-11-12 - tcollins
    * Updated Mailbox Info to correctly handle a $null set for "Send As" permissions (test case "archives" via tcollins2.test)
3.0.4 - 2014-11-10 - tcollins
    * Updated User Info function to correctly detect "Password Never Expires"
    * Fixed Mailbox Info to correctly show group permissions
    * Updated Mailbox Info to denote what privs are Hosting Managed vs User Managed
3.0.3 - 2014-09-15 - tcollins
    * Changed WinSam-Reset-Password to allow manually set passwords, instead of auto-generated.
3.0.2 - 2014-09-09 - tcollins
    * Fixed an error in WinSam-Get-AccountInfo.  The date comparison of Account Expiration and Password Expiration was failing if the password was set to 
      change upon next login.  It gave a value of '0' instead of a DateTime value.  Checking for this fixed the code.
    * Improved the warning text for when a password change is required
3.0.1 - 2014-09-08 - tcollins
    * Fixed WinSam-Get-PasswordExpiration to return if a password is expired.
    * Changed WinSam-Get-GroupInfo to pipe $GroupInfo to "Get-ADGroupMembers" instead of calling the direct function based on group name.  
      There was an issue that was causing it to have a null value on the initial pull of data.  I cannot replicate the issue so this many not fix it.
3.0.0 - 2014-09-04 - tcollins
    * Set PowerShell background color to black
    * Changed date format to use 12-hour format
3.0.B.2 - 2014-08-21 - tcollins
    * Rewrote Access Level function - Properly indicates access to Computer Info
    * Fixed Group Info function - was not getting user information the first time through.
    * Set user accounts that are not in Employees,Students,Workstudy,Contractor OUs to only show basic information and to contact Hosting for details.
3.0.B.1 - 2014-08-21 - tcollins
    * Released to HD Staff for testing
3.0.A.7 - 2014-08-20 - tcollins
    * Added Group Memberships
3.0.A.6 - 2014-08-20 - tcollins
    * Added Change Log Format
    * Fixed Lots of Stuff
3.0.Alpha
    * [Completed] Set menus and access accordingly
    * [Completed] Get General User Information
    * [Completed] Unlock Locked Account
    * [Completed] Get User Memberships
    * [Completed] Reset Password
#>