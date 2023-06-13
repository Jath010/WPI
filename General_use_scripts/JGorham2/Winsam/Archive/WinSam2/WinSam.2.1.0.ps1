##################################################################################################################
# WinSam.ps1
# Windows Samaritan Tool
# Features
#    - reset passwords
#    - unlock accounts 
#    - view directory Information
#
# Version 2.1.0
# Last Updated: 2012-01-27 - tcollins
#
# Change Log at end of script
##################################################################################################################
##################################################################################################################
#Powershell Command Shell Customizations
##################################################################################################################
If ($Host.Name -ne "Windows PowerShell ISE Host"){
    $h = (Get-Host).UI.RawUI
    $h.WindowTitle = "WPI Windows Samaritan 2.1.0"
    $h.BackgroundColor = "DarkBlue"
    $h.ForegroundColor = "White"
    $win = $h.WindowSize
    $win.Height = 50
    $win.Width = 86
    $h.Set_WindowSize($win)
    $buffer = $h.BufferSize
    $buffer.Height = 9999
    $buffer.Width = 86
    $h.Set_BufferSize($buffer)
    $MaxHistoryCount = 1000
    }    

##################################################################################################################
# Powershell Modules
##################################################################################################################
If ($Host.Name -ne "Windows PowerShell ISE Host"){
    $module_ad = (Get-Module | select name | where {$_.name -eq 'activedirectory'}).name
    if ($module_ad -ne 'activedirectory') {import-module activedirectory}

    $pssession = Get-PSSession | Select ConfigurationName
    if ($pssession -eq $null) {
        Add-pssnapin Microsoft.Exchange.Management.Powershell.E2010
        .$env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-exchangeserver –auto
        }
    }
$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
##################################################################################################################
# Global Variables
##################################################################################################################
$ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$global:ADInfo = $null
$global:mailbox = $null

$global:dcs = (Get-ADDomainController -Filter *)
$global:dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
$global:dchostname = $dc.HostName
$global:today = Get-Date
$global:currentuser = $env:username
$global:localhost = $env:COMPUTERNAME
If ($Host.Name -eq "Windows PowerShell ISE Host"){
    $global:fcolor = "White"
    $global:bcolor = "DarkBlue"
    }
Else {
    $global:fcolor = $host.UI.RawUI.ForegroundColor
    $global:bcolor = $host.UI.RawUI.BackgroundColor
    }

##################################################################################################################
# Includes
##################################################################################################################
. $ScriptPath\WinSam.2.1.0.Functions.ps1
. $ScriptPath\WinSam.2.1.0.Info.ps1

##################################################################################################################
# Main Program
##################################################################################################################
<#
    Step 0: Validate Rights of the person running Winsam
        * Set menus and access accordingly
    Step 1: Initial Menu - What do you want to do?
        * [Completed] Get General User Information
        * [Completed] Unlock Locked Account
        * Get User Memberships
        * Reset Password
        
    Other Notes
        * Tested with Roger's Account
            Need to show if Service Accounts are enabled
            Need to show if Service Accounts are locked.
            Shows only Name and Email address for service account
#>
#Sleep 5
Clear-Host
$Global:DCList = (Get-ADDomainController -Filter {OperatingSystem -eq "Windows Server 2008 R2 Enterprise"} | Where {$_.Name -ne 'NEBULA'})
$Global:DCServerName = $Global:DCList[0].HostName.ToString()

$AccessLevel = WinSam-Get-AccessLevel $currentuser
if ($AccessLevel -eq "NoAccess") {
    Write-Host "You are not authorized to access the Windows Samaritan interface" -ForegroundColor Black -BackgroundColor Red
    #Sleep 30
    Sleep 1
    return
    }


#WinSam-Get-MainMenu $AccessLevel
Write-Host "User         : $currentuser" -ForegroundColor Cyan
Write-Host "Access level : $AccessLevel" -ForegroundColor Cyan
Write-Host "DC           : $($Global:DCServerName)" -ForegroundColor Cyan
Write-Host "Path         : $ScriptPath" -ForegroundColor Cyan
Write-Host ""
$username = read-host -Prompt 'Please enter a username'
while ($true) {
    Write-Host '                             Windows Samaritan 3.0.0                                 ' -ForegroundColor Black -BackgroundColor White
    Write-Host '                                                                                     ' -ForegroundColor Black -BackgroundColor White
    WinSam-Get-AccountInfo $username
    Write-Host ""
    Write-Host "Script complete.  Total time: "$ElapsedTime.Elapsed -ForegroundColor Cyan
    $username = read-host -Prompt 'Please enter a username'
    Clear-Host
    . $ScriptPath\WinSam.2.1.0.Functions.ps1
    . $ScriptPath\WinSam.2.1.0.Info.ps1
    $ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "User         : $currentuser" -ForegroundColor Cyan
    Write-Host "Access level : $AccessLevel" -ForegroundColor Cyan
    Write-Host "DC           : $($Global:DCServerName)" -ForegroundColor Cyan
    Write-Host "Path         : $ScriptPath" -ForegroundColor Cyan
    Write-Host ""
    }