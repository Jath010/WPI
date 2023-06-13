Clear-Host
$datestamp = Get-Date -Format "yyyyMMdd-HHmm"
$timestamp = Get-Date -Format "yyyy-MM-dd - HH:mm"
$log_Transcript = "D:\wpi\batch\Exchange\AlumniPilotStats\Logs\TranscriptLog_$datestamp.txt"

Start-Transcript -Path $log_Transcript
Write-Host "Begin file"

##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
$credentials=$null;$RemoteCredentials=$null;$credPath=$null;$RemoteCredPath=$null
## The following credentials use the UPN "exch_automation@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange on-premise and Exchange Online
$credPath = 'D:\wpi\batch\Exchange\AlumniPilotStats\exch_automation@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
#Import-PSSession $ExchangeOnlineSession -Prefix Cloud
Connect-ExchangeOnline -Credential $credentials

##################################################################################################################
## Get List of Users
##################################################################################################################
Write-Host "Start $(get-date)" -ForegroundColor Gray
$Today = Get-Date

$Users_Employees=$null;$Users_Students=$null;$Users_Alumni=$null
$ExchangeUsers = @()

$OU_AccountsOU     = 'OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Alumni         = 'OU=Alumni,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Disabled       = 'OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Employees      = 'OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_NoOffice365Sync= 'OU=No Office 365 Sync,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_OtherAccounts  = 'OU=Other Accounts,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Privileged     = 'OU=Privileged,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Retirees       = 'OU=Retirees,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Services       = 'OU=Services,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Students       = 'OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_Vokes          = 'OU=Vokes,OU=Accounts,DC=admin,DC=wpi,DC=edu'
$OU_WorkStudy      = 'OU=Work Study,OU=Accounts,DC=admin,DC=wpi,DC=edu'

$ExchangeUsers = Get-ADUser -Filter * -SearchBase $OU_Employees -Properties DisplayName,GivenName,Surname,EmployeeID,EmployeeNumber,LastLogonDate,PasswordLastSet,PasswordExpired,extensionAttribute9,extensionAttribute10,WhenCreated,msExchWhenMailboxCreated | Where {$_.msExchWhenMailboxCreated}
$ExchangeUsers += Get-ADUser -Filter * -SearchBase $OU_Students  -Properties DisplayName,GivenName,Surname,EmployeeID,EmployeeNumber,LastLogonDate,PasswordLastSet,PasswordExpired,extensionAttribute9,extensionAttribute10,WhenCreated,msExchWhenMailboxCreated | Where {$_.msExchWhenMailboxCreated}
$ExchangeUsers += Get-ADUser -Filter * -SearchBase $OU_Alumni    -Properties DisplayName,GivenName,Surname,EmployeeID,EmployeeNumber,LastLogonDate,PasswordLastSet,PasswordExpired,extensionAttribute9,extensionAttribute10,WhenCreated,msExchWhenMailboxCreated | Where {$_.msExchWhenMailboxCreated}
$ExchangeUsers = $ExchangeUsers | Sort SamAccountName
Write-Host "Break 1 $(get-date)" -ForegroundColor Gray

if (!$ExchangeUsers) {break}

$ExchangeData  = @()

$count=$null
$TotalCount = $ExchangeUsers.Count
$TotalCount

##################################################################################################################
## Get Mailbox Information and Statistics
##################################################################################################################
Write-Host "Break 2 $(get-date)" -ForegroundColor Gray
# Break 2 10/04/2017 07:47:12
foreach ($user in $ExchangeUsers) {
    $username=$null;$mailbox=$null;$mbxstats=$null;$status=$null;$DN=$Null
    $count++
    $username = $user.SamAccountName
    if ($count%100 -eq 0) {Write-Host "[$(Get-Date -Format "yyyy-MM-dd HH:mm")] Processing [$count of $totalcount] $username"}
    $DN = $user.DistinguishedName
    switch -Wildcard ($DN) {
        '*OU=Employee*' {$Status = 'Employee'}
        '*OU=Student*'  {$Status = 'Student'}
        '*OU=Alumni*'   {$Status = 'Alumni'}
        default{$Status = 'Other'}
        }

    $mailbox = Get-CloudMailbox $username
    $mbxstats = Get-CloudMailboxStatistics $username

    $out = New-Object PSObject
    $out | Add-Member noteproperty Name $mailbox.DisplayName
    $out | Add-Member noteproperty FirstName $user.GivenName
    $out | Add-Member noteproperty LastName $user.Surname
    $out | Add-Member noteproperty Username $username
    $out | Add-Member noteproperty WPI_ID $user.EmployeeID
    $out | Add-Member noteproperty PIDM $user.EmployeeNumber
    $out | Add-Member noteproperty ADLastLogonDate $user.LastLogonDate
    $out | Add-Member noteproperty MBXLastLogonDate $mbxstats.LastLogonTime
    $out | Add-Member noteproperty PasswordLastSet $user.PasswordLastSet
    $out | Add-Member noteproperty PasswordExpired $user.PasswordExpired
    $out | Add-Member noteproperty ExchangeForward $mailbox.ForwardingSmtpAddress
    $out | Add-Member noteproperty MailboxSize $mbxstats.TotalItemSize.Value
    $out | Add-Member NoteProperty AlumniPilotGroup $user.extensionAttribute9
    $out | Add-Member NoteProperty AlumniPilotStatus $user.extensionAttribute10
    $out | Add-Member NoteProperty Status $Status
    $out | Add-Member NoteProperty DN $user.DistinguishedName
    $out | Add-Member noteproperty Created $mailbox.WhenCreated
    $ExchangeData += $out
    }

##################################################################################################################
## Break down statistics
##################################################################################################################
Write-Host "Break 3 $(get-date)" -ForegroundColor Gray

$ActiveUsers = $ExchangeData | Where {$_.Status -eq 'Employee' -or $_.Status -eq 'Student'}
$AlumniTotalUsers = $ExchangeData | Where {$_.Status -eq 'Alumni'}
$Alumni2016Users = $AlumniTotalUsers | Where {$_.AlumniPilotGroup -eq 'AlumniPilot2016'}
$Alumni2017Users = $AlumniTotalUsers | Where {$_.AlumniPilotGroup -eq 'AlumniPilot2017'}

function WPI-Get-ExchangeStats ($ExchData) {
    $ExchangeStats = @();$TotalMBX=$null;$Forwards=$null;$Active30=$null;$Active60=$null;$Active180=$null;$NeverUsed=$null

    ## Set a date to get stats since a specific date
    $ActivityDate = Get-Date('07/01/17')

    $TotalMBX = ($ExchData | Measure-Object).Count
    $Forwards = ($ExchData | Where {$_.ExchangeForward} | Measure-Object).Count
    $Active30  = ($ExchData | Where {$_.MBXLastLogonDate -gt $Today.AddDays(-30)} | Measure-Object).Count
    $Active60  = ($ExchData | Where {$_.MBXLastLogonDate -gt $Today.AddDays(-60)} | Measure-Object).Count
    $Active180  = ($ExchData | Where {$_.MBXLastLogonDate -gt $Today.AddDays(-180)} | Measure-Object).Count
    $ActiveDate  = ($ExchData | Where {$_.MBXLastLogonDate -gt $ActivityDate} | Measure-Object).Count
    $NeverUsed = ($ExchData | Where {!($_.MBXLastLogonDate) -and !($_.ExchangeForward)} | Measure-Object).Count


    #Active 30 Days
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Active: Last 30 Days"
    $out | Add-Member noteproperty Count $Active30
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($Active30/$TotalMBX))
    $ExchangeStats += $out

    #Active 60 Days
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Active: Last 60 Days"
    $out | Add-Member noteproperty Count $Active60
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($Active60/$TotalMBX))
    $ExchangeStats += $out

    #Active 180 Days
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Active: Last 180 Days"
    $out | Add-Member noteproperty Count $Active180
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($Active180/$TotalMBX))
    $ExchangeStats += $out

    #Active Since specified date
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Active: Since $($ActivityDate.ToShortDateString())"
    $out | Add-Member noteproperty Count $ActiveDate
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($ActiveDate/$TotalMBX))
    $ExchangeStats += $out

    #Never Used
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Never Used"
    $out | Add-Member noteproperty Count $NeverUsed
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($NeverUsed/$TotalMBX))
    $ExchangeStats += $out

    #Forward Externally
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Forward Externally"
    $out | Add-Member noteproperty Count $Forwards
    $out | Add-Member noteproperty Percentage $("{0:P2}" -f ($Forwards/$TotalMBX))
    $ExchangeStats += $out

    #Divider Line
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "----------"
    $out | Add-Member noteproperty Count '-----'
    $out | Add-Member noteproperty Percentage '----------'
    $ExchangeStats += $out
    
    #Total Count
    $out = New-Object PSObject
    $out | Add-Member noteproperty "Data Point" "Total Count"
    $out | Add-Member noteproperty Count $TotalMBX
    $out | Add-Member noteproperty Percentage $null
    $ExchangeStats += $out

    Return $ExchangeStats
    }

Write-Host "Break 4 $(get-date)" -ForegroundColor Gray

$ActiveUserStats = WPI-Get-ExchangeStats $ActiveUsers
$AlumniTotalUserStats = WPI-Get-ExchangeStats $AlumniTotalUsers
$Alumni2016UserStats = WPI-Get-ExchangeStats $Alumni2016Users
$Alumni2017UserStats = WPI-Get-ExchangeStats $Alumni2017Users

##################################################################################################################
## Output to Mail Message
##################################################################################################################
Write-Host "Break 5 $(get-date)" -ForegroundColor Gray
$messageParameters=$null;$Subject=$null;$Body=$null;$body2table=$null

$Date = Get-Date -Format "MM/dd/yyyy"

$body2table = $null
$body2table += "<table width='100%' border='1'><tbody>"
$body2table += "<tr bgcolor=#8fdb83><td colspan='3'>All Alumni Pilot Mailboxes ($date)</td></tr>"
$body2table += "<tr bgcolor=#CCCCCC><td align='Left'>Data Point</td><td align='Center'>Count</td><td align='Center'>Percentage</td></tr>"
foreach ($item in $AlumniTotalUserStats) {$body2table += "<tr><td>$($item.'Data Point')</td><td align='Right'>$($item.Count)</td><td align='Right'>$($item.Percentage)</td></tr>"}
$body2table += "<tr><td colspan='3'> </td></tr>"
$body2table += "<tr bgcolor=#80a1d6><td colspan='3'>Alumni Pilot 2016 Mailboxes ($date)</td></tr>"
$body2table += "<tr bgcolor=#CCCCCC><td align='Left'>Data Point</td><td align='Center'>Count</td><td align='Center'>Percentage</td></tr>"
foreach ($item in $Alumni2016UserStats) {$body2table += "<tr><td>$($item.'Data Point')</td><td align='Right'>$($item.Count)</td><td align='Right'>$($item.Percentage)</td></tr>"}
$body2table += "<tr><td colspan='3'> </td></tr>"
$body2table += "<tr bgcolor=#80a1d6><td colspan='3'>Alumni Pilot 2017 Mailboxes ($date)</td></tr>"
$body2table += "<tr bgcolor=#CCCCCC><td align='Left'>Data Point</td><td align='Center'>Count</td><td align='Center'>Percentage</td></tr>"
foreach ($item in $Alumni2017UserStats) {$body2table += "<tr><td>$($item.'Data Point')</td><td align='Right'>$($item.Count)</td><td align='Right'>$($item.Percentage)</td></tr>"}
$body2table += "<tr><td colspan='3'> </td></tr>"
$body2table += "<tr bgcolor=#e59292><td colspan='3'>Active User Mailboxes (Employees/Students) ($date)</td></tr>"
$body2table += "<tr bgcolor=#CCCCCC><td align='Left'>Data Point</td><td align='Center'>Count</td><td align='Center'>Percentage</td></tr>"
foreach ($item in $ActiveUserStats) {$body2table += "<tr><td>$($item.'Data Point')</td><td align='Right'>$($item.Count)</td><td align='Right'>$($item.Percentage)</td></tr>"}
$body2table += "<tr><td colspan='3'> </td></tr>"
$body2table += "</table>"

$Subject  = "Mail Usage Statistics - $date"
$Body     = "<p>The following table includes data and statistics on mailbox utilization for the Alumni Pilot program.  This data runs monthly to provide regular ongoing data.</p>
            $body2table"    

$messageParameters2 = @{
    Subject = "Mail Usage Statistics - $date"
    Body    = $Body
    From    = 'its@wpi.edu'
    To      = 'tcollins@wpi.edu'
    #To      = 'tcollins@wpi.edu','ccerny@wpi.edu','ajdealy@wpi.edu','amcmahon@wpi.edu'
    Bcc     = 's_tcollins@wpi.edu'
    SmtpServer = 'smtp.wpi.edu'
    }  
    
Send-MailMessage @messageParameters2 -BodyAsHtml
Write-Host "Message Sent"

Remove-PSSession $ExchangeOnlineSession
Stop-Transcript