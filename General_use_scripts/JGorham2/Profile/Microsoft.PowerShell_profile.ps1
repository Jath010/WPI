##################################################################################################################
#Powershell Command Shell Customizations
##################################################################################################################
$ShellWidth = 116

$h = (Get-Host).UI.RawUI

$buffer = $h.BufferSize
$win = $h.WindowSize

$buffer.Height = 9999
if ($win.Height -le 60) {$win.Height = 60}

$buffer.Width = $ShellWidth
$win.Width = $ShellWidth

if ($ShellWidth -lt $h.BufferSize.Width) {
    $h.Set_WindowSize($win)
    $h.Set_BufferSize($buffer)
    }
Else {
    $h.Set_BufferSize($buffer)
    $h.Set_WindowSize($win)
    }
        
if (!(Test-Path c:\wpi)) {New-Item c:\wpi -ItemType Directory | Out-Null}
Set-Location c:\wpi

$DNSDomain   = 'wpi.edu'
$DN_Domain   = 'DC=wpi,DC=edu'

##################################################################################################################
# Powershell Variables
##################################################################################################################
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

#Domain Controller Information
$global:dchostname = (Get-ADDomainController -Filter * | Where-Object {$_.OperationMasterRoles -like "*RIDMaster*"}).HostName

#Environment Variables
$global:today = Get-Date
$global:currentuser = $env:username
$global:localhost = $env:COMPUTERNAME

#OU Locations
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


#Set Isilon FQDN
$StorageStatus=$true
if (Test-Path "\\storage.wpi.edu\dept") {$StorageFQDN = "storage.wpi.edu";$DRStorageFQDN = "drstorage.wpi.edu"}
elseif (Test-Path "\\storage\dept") {$StorageFQDN = "storage";$DRStorageFQDN = "drstorage"}
else {$StorageStatus=$false}

$ProfilePath = "$([Environment]::GetFolderPath("MyDocuments"))\WindowsPowerShell"

if ($StorageStatus) {
    $ScriptsPath = "\\$StorageFQDN\dept\Information Technology\CCC\Windows\fc_windows\Scripts"
    }

##################################################################################################################
#Powershell Load Modules
##################################################################################################################
Import-module ActiveDirectory
Import-Module ExchangeOnlineManagement

Connect-AzureAD -AccountId jmgorham2_prv@wpi.edu
Connect-ExchangeOnline -UserPrincipalName jmgorham2_prv@wpi.edu

##################################################################################################################
#Powershell Set Aliases
##################################################################################################################
if ($StorageStatus=$true) {
    Set-Alias WPI-Disable-Accounts "\\$StorageFQDN\dept\Information Technology\CCC\Windows\fc_windows\account_removals\accountdisablescript.ps1"
    Set-Alias WPI-Alumni-Conversion "\\$StorageFQDN\dept\Information Technology\SysOps\sysops_scripts\AlumniEmail\ConvertAlumniUsers.ps1"
    Set-Alias WPI-Retiree-Conversion "\\$StorageFQDN\dept\Information Technology\SysOps\sysops_scripts\Retiree\ConvertRetiree.ps1"
    Set-Alias WPI-LOA-Conversion "\\$StorageFQDN\dept\Information Technology\SysOps\sysops_scripts\LeaveOfAbsence\ConvertLOAUsers.ps1"
    #Set-Alias new_user "\\$StorageFQDN\dept\Information Technology\CCC\Windows\fc_windows\account_creations\Powershell\new_user.ps1"
    Set-Alias New-StudentWorker "\\$StorageFQDN\dept\Information Technology\CCC\Windows\fc_windows\account_creations\Powershell\New-StudentWorker.ps1"
    #Set-Alias WinSam '\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3\WinSam.3.ps1'
    #Set-Alias WinSamB '\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3B\WinSam.3.ps1'
    #Set-Alias WinSam-Dev '\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3C\WinSam.3.ps1'
    }

if (Test-Path "c:\Program Files\TextPad 8\TextPad.exe") {Set-Alias TextPad "c:\Program Files\TextPad 8\TextPad.exe"}
elseif (Test-Path "C:\Program Files (x86)\TextPad 7\TextPad.exe") {Set-Alias TextPad "C:\Program Files (x86)\TextPad 7\TextPad.exe"}

if(Test-Path "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE") {Set-Alias Excel "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE"}
elseif(Test-Path "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE") {Set-Alias Excel "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE"}

##################################################################################################################
#Load External Script Libraries
##################################################################################################################
if (!(Test-Path $ProfilePath\Libraries\)) {New-Item $ProfilePath\Libraries\ -ItemType Directory | Out-Null}

if ($StorageStatus=$true) {
    if (Test-Path "$ScriptsPath\Libraries\ADFunctions.ps1") {Copy-Item "$ScriptsPath\Libraries\ADFunctions.ps1" -Destination $ProfilePath\Libraries\}
    if (Test-Path "$ScriptsPath\Libraries\O365Functions.ps1") {Copy-Item "$ScriptsPath\Libraries\O365Functions.ps1" -Destination $ProfilePath\Libraries\}
    if (Test-Path "$ScriptsPath\Libraries\EnhancedFind.ps1") {Copy-Item "$ScriptsPath\Libraries\EnhancedFind.ps1" -Destination $ProfilePath\Libraries\}
    if (Test-Path "$ScriptsPath\Libraries\ExchangeAdminFunctions.ps1") {Copy-Item "$ScriptsPath\Libraries\ExchangeAdminFunctions.ps1" -Destination $ProfilePath\Libraries\}
    if (Test-Path "$ScriptsPath\Libraries\ExchangeGeneralFunctions.ps1") {Copy-Item "$ScriptsPath\Libraries\ExchangeGeneralFunctions.ps1" -Destination $ProfilePath\Libraries\}
    }

. $ProfilePath\Libraries\ADFunctions.ps1
. $ProfilePath\Libraries\EnhancedFind.ps1
. $ProfilePath\Libraries\ExchangeAdminFunctions.ps1
. $ProfilePath\Libraries\ExchangeGeneralFunctions.ps1
. $ProfilePath\Libraries\O365Functions.ps1

Import-Module C:\Users\jmgorham2_prv\WPI\General_use_scripts\JGorham2\Profile\ProfileInclusions.ps1

##################################################################################################################
#Powershell Custom Functions
##################################################################################################################
#Clear-Host