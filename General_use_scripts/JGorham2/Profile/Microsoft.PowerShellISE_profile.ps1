##################################################################################################################
#Powershell Command Shell Customizations
##################################################################################################################
$localhost = $env:COMPUTERNAME
$MaxHistoryCount = 1000
if (!(Test-Path c:\wpi)) {New-Item c:\wpi -ItemType Directory}
Set-Location c:\wpi

$DNSDomain   = 'wpi.edu'
$DN_Domain   = 'DC=wpi,DC=edu'

##################################################################################################################
#Powershell Load Modules
##################################################################################################################
Import-module ActiveDirectory
Import-module AzureAD
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline -UserPrincipalName jmgorham2_prv@wpi.edu
Connect-AzureAD -AccountId jmgorham2_prv@wpi.edu
##################################################################################################################
#Powershell Set Aliases/Variables
##################################################################################################################

#Domain Controller Information
$global:dchostname = (Get-ADDomainController -Filter * | Where {$_.OperationMasterRoles -like "*RIDMaster*"}).HostName

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

$ProfilePath = "$([Environment]::GetFolderPath("MyDocuments"))\WindowsPowerShell\"

if ($StorageStatus) {
    $ScriptsPath = "\\$StorageFQDN\dept\Information Technology\CCC\Windows\fc_windows\Scripts"
    $PersonalScriptsPath = "\\$StorageFQDN\dept\Information Technology\ccc\Windows\fc_windows\tcollins\scripts\PowerShell"
    }

##################################################################################################################
#Powershell Custom Functions
##################################################################################################################
Clear-Host


