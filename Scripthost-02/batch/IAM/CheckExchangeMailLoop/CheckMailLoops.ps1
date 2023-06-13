Clear-Host
#Get list of adusers that have been modified in the last 24 hours and verify that there are no mail loops
Start-Transcript -Path "d:\wpi\batch\IAM\CheckExchangeMailLoop\Transcripts\ScriptTranscript_$(Get-Date -Format 'yyyy-MM-dd_HHmm').txt"

##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
$credentials=$null;$credPath=$null
## The following credentials use the UPN "s_tcollins@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange Online
$credPath = 'D:\wpi\batch\IAM\CheckExchangeMailLoop\s_tcollins@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
$ExchangeLocalSession=$null;$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
Connect-ExchangeOnline $credentials
#Import-PSSession $ExchangeOnlineSession -Prefix Cloud

##################################################################################################################
## Powershell Modules Check
##################################################################################################################
$ExchangeOnlineCheck=$null;$ActiveDirectoryCheck=$null
if (Get-PSSession | Where {$_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ps.outlook.com"}) {$ExchangeOnlineCheck = $true}
if (!$ExchangeOnlineCheck) {
    Write-Host ''
    Write-Host "Modules not correctly loading"
    Write-Host "Exchange Online: $ExchangeOnlineSession"
    break
    }

$OU_Employees = 'OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Students = 'OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu'

$ADUsers = Get-ADUser -Filter * -Properties Created,Modified,msExchHideFromAddressLists,AccountExpirationDate,LastLogonDate -SearchBase $OU_Employees
$ADUsers += Get-ADUser -Filter * -Properties Created,Modified,msExchHideFromAddressLists,AccountExpirationDate,LastLogonDate -SearchBase $OU_Students

$NewUsers = $ADUsers | Where {$_.Modified -gt (Get-Date).AddDays(-1)}  | Sort SamAccountName
$ProcessTime = Get-Date

$MailForwardsNew = @()

$count = $null
$totalcount = @($NewUsers).Count

foreach ($user in $NewUsers) {
    $mailbox=$null
    $alias=$null;$ForwardFile=$null;$line=$null;$ForwardContent=$null;$UnixForward=$null;$ExchForward=$null

    $Alias = $user.SamAccountName
    $mailbox = Get-CloudMailbox $alias -ErrorAction SilentlyContinue
    if (!$mailbox) {continue}

    $count++
    Write-Host "[$count of $totalcount] Processing $alias [$($mailbox.DisplayName)]"
    
    If (Test-Path "\\storage\homes\$alias" -ErrorAction SilentlyContinue) {
        If (Test-Path "\\storage\homes\$alias\.forward" -ErrorAction SilentlyContinue) {
            $ForwardContent = Get-Content "\\storage\homes\$alias\.forward"
            if ($ForwardContent) {
                foreach ($line in $ForwardContent) {
                    if ($line -like '#*') {}
                    else {
                        if ($UnixForward) {$UnixForward = "$UnixForward,$line"}
                        else {$UnixForward = $line}
                        }
                    }
                }
            else {$UnixForward = '==EMPTY_FILE=='}
            }
        else {$UnixForward = '==FILE_MISSING=='}
        }
    else {$UnixForward = '==ACCOUNT_MISSING=='}

    if ($mailbox.ForwardingSmtpAddress) {$ExchForward = $mailbox.ForwardingSmtpAddress.Split(':')[1]}

    $out = New-Object PSObject
    $out | Add-Member noteproperty Name $mailbox.DisplayName
    $out | Add-Member noteproperty Alias $alias
    $out | Add-Member noteproperty RecipientType $Recipient.RecipientType
    $out | Add-Member noteproperty RecipientTypeDetails $Recipient.RecipientTypeDetails
    $out | Add-Member noteproperty ExchangeForward $ExchForward
    $out | Add-Member noteproperty UNIXForward $UnixForward
    $out | Add-Member noteproperty DistinguishedName $user.DistinguishedName
    $out | Add-Member noteproperty AccountExpirationDate $user.AccountExpirationDate
    $out | Add-Member noteproperty LastLogonDate $user.LastLogonDate
    $out | Add-Member noteproperty Enabled $user.Enabled
    $out | Add-Member noteproperty msExchHideFromAddressLists $user.msExchHideFromAddressLists

    $MailForwardsNew += $out
    }

$Loop=$null;$Hidden=$null

Write-Host ''
Write-Host '** Loop Problems **' -ForegroundColor Green
$Loop = $MailForwardsNew | Where {$_.UNIXForward -match "@exchange.wpi.edu" -and $_.ExchangeForward -match "@smtp.wpi.edu"}
$Loop | Select Alias,ExchangeForward,UNIXForward | Out-Default

foreach ($account in $loop) {
    $Alias = $account.Alias
    $ExchForward = $account.ExchangeForward.split('@')[0].ToLower()
    $UNIXForward = $account.UNIXForward.split('@')[0].ToLower()

    if ($ExchForward -eq $UNIXForward) {
        Set-CloudMailbox $Alias -ForwardingSmtpAddress $null
        Write-Host "Processing $alias"
        }
    }
Write-Host ''
Write-Host ''
Write-Host '** GAL Problems  **' -ForegroundColor Green

$Hidden = $MailForwardsNew | Where {$_.msExchHideFromAddressLists -eq $true}
$Hidden | Select Alias,msExchHideFromAddressLists | Out-Default

foreach ($account in $Hidden) {
    Set-ADUser $account.Alias -Replace @{msExchHideFromAddressLists=$false}
    }
Write-Host ''

Write-Host "Processing started at $ProcessTime" -ForegroundColor Gray
Write-Host "Processing completed at $(Get-Date)" -ForegroundColor Gray
Write-Host ''

Remove-PSSession $ExchangeOnlineSession
Stop-Transcript
