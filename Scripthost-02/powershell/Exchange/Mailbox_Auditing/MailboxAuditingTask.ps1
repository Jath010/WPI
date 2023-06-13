#Handler for getting connected to exchange before running the script

#############################################
# Logging
# Set path for log files:
$logPath = "D:\wpi\Logs\MailboxAuditing"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_MailboxAuditing.log" -Force
#############################################

$creds = Import-Clixml D:\wpi\XML\exch_automation\exch_automation@wpi.edu.xml
Connect-ExchangeOnline -Credential $creds
D:\wpi\powershell\Exchange\Mailbox_Auditing\Configure-MailboxAuditLogging.ps1 -ConfigureAuditActions

Stop-Transcript