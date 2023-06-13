
Clear-Host
#Get list of adusers that have been modified in the last 24 hours and verify that there are no termination mail restrictions
$datestamp = Get-Date -Format "yyyyMMdd-HHmm"
$timestamp = Get-Date -Format "yyyy-MM-dd - HH:mm"
$log_Transcript = "D:\wpi\batch\Exchange\CheckMailRestrictions\Logs\Transcripts\TranscriptLog_$datestamp.txt"

Start-Transcript -Path $log_Transcript
Write-Host "Begin file"

##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
$credentials=$null;$credPath=$null
## The following credentials use the UPN "s_tcollins@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange Online
$credPath = 'D:\wpi\batch\Exchange\CheckMailRestrictions\exch_automation@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
#Import-PSSession $ExchangeOnlineSession
Connect-ExchangeOnline -Credential $credentials

## Load Active Directory Module
Import-Module ActiveDirectory

##################################################################################################################
## Powershell Modules Check
##################################################################################################################
$ExchangeOnlineCheck=$null;$ActiveDirectoryCheck=$null
if (Get-PSSession | Where {$_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ps.outlook.com"}) {$ExchangeOnlineCheck = $true}
if (Get-Module | Where {$_.Name -match "ActiveDirectory"}) {$ActiveDirectoryCheck = $true}

if (!$ExchangeOnlineCheck -or !$ActiveDirectoryCheck) {
    Write-Host ''
    Write-Host "Modules not correctly loading"
    Write-Host "Exchange Online         : $ExchangeOnlineCheck"
    Write-Host "Active Directory Status : $ActiveDirectoryCheck"
    break
    }
# Clear-Host
$OU_Employees = 'OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Students  = 'OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu'

$ADUsers  = Get-ADUser -Filter * -Properties Created,Modified,msExchHideFromAddressLists,AccountExpirationDate,LastLogonDate -SearchBase $OU_Employees
$ADUsers += Get-ADUser -Filter * -Properties Created,Modified,msExchHideFromAddressLists,AccountExpirationDate,LastLogonDate -SearchBase $OU_Students

$NewUsers = $ADUsers | Where {$_.Modified -gt (Get-Date).AddDays(-2)}  | Sort SamAccountName
$ProcessTime = Get-Date

$MailInfoChangedUsers = @()

$count = $null
$totalcount = ($NewUsers | Measure-Object).Count

foreach ($user in $NewUsers) {
    $mailbox=$null;$alias=$null
    $ExchForward=$null;$AcceptMessagesOnlyFrom=$null

    $Alias = $user.SamAccountName
    $mailbox = Get-Mailbox $alias -ErrorAction SilentlyContinue
    if (!$mailbox) {continue}

    $count++
    Write-Host "[$count of $totalcount] Processing $alias [$($mailbox.DisplayName)]"
    
    if ($mailbox.ForwardingSmtpAddress) {$ExchForward = $mailbox.ForwardingSmtpAddress.Split(':')[1]}
    if ($mailbox.AcceptMessagesOnlyFrom) {$AcceptMessagesOnlyFrom = $mailbox.AcceptMessagesOnlyFrom}

    $out = New-Object PSObject
    $out | Add-Member noteproperty Name $mailbox.DisplayName
    $out | Add-Member noteproperty Alias $alias
    $out | Add-Member noteproperty ExchangeForward $ExchForward
    $out | Add-Member noteproperty AcceptMessagesOnlyFrom $AcceptMessagesOnlyFrom
    $out | Add-Member noteproperty DistinguishedName $user.DistinguishedName
    $out | Add-Member noteproperty AccountExpirationDate $user.AccountExpirationDate
    $out | Add-Member noteproperty LastLogonDate $user.LastLogonDate
    $out | Add-Member noteproperty Enabled $user.Enabled
    $out | Add-Member noteproperty msExchHideFromAddressLists $user.msExchHideFromAddressLists
    $MailInfoChangedUsers += $out
    }

$BadForward=$null;$Hidden=$null;$ReceiveRestrictions=$null

Write-Host ''
Write-Host '** Forwarding Problems **' -ForegroundColor Green
$BadForward = $MailInfoChangedUsers | Where {$_.ExchangeForward -match "@smtp.wpi.edu"}
$BadForward | Select Alias,ExchangeForward | Out-Default

foreach ($account in $BadForward) {Set-Mailbox $account.Alias -ForwardingSmtpAddress $null}

Write-Host ''
Write-Host ''
Write-Host '** GAL Problems  **' -ForegroundColor Green

$Hidden = $MailInfoChangedUsers | Where {$_.msExchHideFromAddressLists -eq $true}
$Hidden | Select Alias,msExchHideFromAddressLists | Out-Default

foreach ($account in $Hidden) {Set-ADUser $account.Alias -Replace @{msExchHideFromAddressLists=$false}}

Write-Host ''
Write-Host ''
Write-Host '** Receive Restriction Problems  **' -ForegroundColor Green

$ReceiveRestricted = $MailInfoChangedUsers | Where {$_.AcceptMessagesOnlyFrom}
$ReceiveRestricted | Select Alias,AcceptMessagesOnlyFrom | Out-Default

foreach ($account in $ReceiveRestricted) {Set-Mailbox $account.Alias -AcceptMessagesOnlyFrom $null}

Write-Host ''
Write-Host "Processing started at $ProcessTime" -ForegroundColor Gray
Write-Host "Processing completed at $(Get-Date)" -ForegroundColor Gray
Write-Host ''

Remove-PSSession $ExchangeOnlineSession
Stop-Transcript
