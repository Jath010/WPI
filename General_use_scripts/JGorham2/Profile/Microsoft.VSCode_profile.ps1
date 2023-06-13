###############
# Profile for JMGorham2_PRV
#
#
###############


##################################################################################################################
#Powershell Command Shell Customizations
##################################################################################################################
#$localhost = $env:COMPUTERNAME
Set-Location c:\

#$DNSDomain   = 'wpi.edu'
#$DN_Domain   = 'DC=wpi,DC=edu'

##################################################################################################################
#Powershell Load Modules
##################################################################################################################
Import-module ActiveDirectory
Import-Module ExchangeOnlineManagement

Import-Module C:\Users\jmgorham2_prv\WPI\General_use_scripts\JGorham2\Profile\ProfileInclusions.ps1

##################################################################################################################
#Powershell Connect Sessions
##################################################################################################################

Connect-AzureAD -AccountId jmgorham2_prv@wpi.edu
Connect-ExchangeOnline -UserPrincipalName jmgorham2_prv@wpi.edu

##################################################################################################################
#Powershell Set Aliases/Variables
##################################################################################################################


##################################################################################################################
#Powershell Custom Functions
##################################################################################################################

#Aliases for Enabling and Disabling the WinRM Service on remote machines to allow for PSSession creation

Function setRM {Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StartService -Computername "$args" |Out-Null}
Set-Alias Enable-WinRM setRM

Function unsetRM {Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StopService -Computername "$args" |Out-Null}
Set-Alias Disable-WinRM unsetRM

Clear-Host