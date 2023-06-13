Clear-Host
$datestamp = Get-Date -Format "yyyyMMdd-HHmm"
$timestamp = Get-Date -Format "yyyy-MM-dd - HH:mm"

$log_Path   = "D:\wpi\batch\Exchange\ManagedBlockedUsers\Logs"
$log_Transcript = "$log_Path\Transcripts\TranscriptLog_$datestamp.txt"

Start-Transcript -Path $log_Transcript

##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
$credentials=$null;$credPath=$null

## The following credentials use the UPN "exch_automation@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange on-premise and Exchange Online
$credPath = 'd:\wpi\batch\ExchangeMigration\exch_automation@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
#try{Import-PSSession $ExchangeOnlineSession -Prefix Cloud}
#catch{$OnlineSessionError = $_}
Connect-ExchangeOnline -Credential $credentials
if ($OnlineSessionError) {
    "$(Get-Date -Format "yyyy-MM-dd - HH:mm:ss") - [Exchange Online Session Connect] - $($OnlineSessionError.Exception.Message)" | Out-File $log_error -Append
    break
    }
##################################################################################################################
## Process Blocked User List
##################################################################################################################
Clear-Host
$blockedusers=$Null
$BlockInfo=@()

$blockedusers = Get-CloudBlockedSenderAddress | Where {$_.SenderAddress -match '@wpi.edu'} | Sort CreatedDatetime
foreach ($user in $blockedusers) {
    $PasswordChanged=$false;$ClearToUnblock=$false
    $username=$null;$ADUser=$null;$out=$null
    $BlockedDate = $user.CreatedDatetime.AddHours(-5)
    $username = $user.SenderAddress.Split('@')[0]
    
    $ADUser = Get-ADUser $username -Properties PasswordLastSet
    $Mailbox = Get-CloudMailbox $username

    if ($BlockedDate -lt $ADUser.PasswordLastSet) {$PasswordChanged = $true}
    if (((get-date ) -gt $ADUser.PasswordLastSet.AddHours(6)) -and $PasswordChanged) {$ClearToUnblock = $true}
    
    $out = New-Object PSObject
    $out | Add-Member noteproperty Name $ADUser.Name
    $out | Add-Member noteproperty Username $ADUser.SamAccountName
    $out | Add-Member noteproperty SMTPAddress $user.SenderAddress
    $out | Add-Member noteproperty Forward  $Mailbox.ForwardingSMTPAddress
    $out | Add-Member noteproperty BlockDate $BlockedDate
    $out | Add-Member NoteProperty PasswordLastSet $ADUser.PasswordLastSet
    $out | Add-Member NoteProperty PasswordChanged $PasswordChanged
    $out | Add-Member NoteProperty ClearToUnblock $ClearToUnblock
    $BlockInfo += $out 
    }

foreach ($user in $BlockInfo) {
    If ($user.ClearToUnblock -eq $true) {
        Remove-CloudBlockedSenderAddress -SenderAddress $user.SMTPAddress -Reason "Password reset and account cleared"
        Sleep 5
        }
    }

Remove-PSSession $ExchangeOnlineSession
Stop-Transcript