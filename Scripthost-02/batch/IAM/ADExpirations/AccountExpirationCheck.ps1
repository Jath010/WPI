<#
AccountExpirationCheck.ps1
Written by: Tom Collins (tcollins)
glipin - changed links in email 020320
cidorr - added Exch_Automation Credentials 220302
glipin 050622

Last updated: 5/06/2022

Notes:
Script Processes in about 20 minutes
#>

#Exch_Automation Credentials
$credPath = "D:\wpi\XML\exch_automation.xml"
$credentials = Import-CliXml -Path $credPath

########################
##  FUNCTIONS
########################

Function Purge-Logs ($path) {
  $TargetFolder = $null; $Days = $null; $Now = $null; $File = $null; $Folders = $null; $SubFolder = $null; $LastWrite = $null

  $TargetFolder = $path
  $Days = 365
  $Now = Get-Date

  $Folders = Get-ChildItem $TargetFolder

  # Notice the minus sign before $days
  $LastWrite = $Now.AddDays(-$days)

  ForEach ($File in $Folders) { if ($file.LastWriteTime -lt $LastWrite) { Remove-Item $file.VersionInfo.FileName | out-null } }
}

Function Get-AccountExpirations {
  Import-Module ActiveDirectory

  #Declare Variables
  #-----------------------------------------------------------------------------------------------
  $users = $null; $out = $null; $PendingAccountExpirations = $null; $PendingPasswordExpirations = $null; $today = $null; $maxpassage = $null
  $AccountEmailLog = @(); $PasswordEmailLog = @(); $PendingAccountExpirationList = @(); $PendingPasswordExpirationList = @()

  #Set Variables
  #-----------------------------------------------------------------------------------------------
  $i = 1
  $today = Get-Date

  #Set email information.  For multiple values in the "To" field, seperate the values with a comma
  #-----------------------------------------------------------------------------------------------
  $HeaderFrom = 'its@wpi.edu'
  $SMTPServer = 'smtp.wpi.edu'

  $HTMLStyle = '
    <style>
      .headerText {
        color: #fff;
        font-family: Arial;
        font-weight: 800;
        font-size: 18px
      }

      .footerText {
        color: #fff;
        font-family: Arial;
      }
    </style>
    '
  $BodyHeader = '
    <!--header-->
    <table width="100%" cellpadding="0" cellspacing="0" style="min-width:100%;background:#c1272d; margin: 0;">
      <tr>
        <td style="padding-left: 50px;padding-top: 50px;padding-bottom: 10px;line-height: 20px;vertical-align:bottom">
          <div class="headerText">Worcester Polytechnic Institute</div>
          <div class="headerText" style="font-weight:500">Information Technology</div>
        </td>
      </tr>
    </table>
    '
  $BodyFooter = '
    <!--footer-->
    <table width="100%" cellpadding="0" cellspacing="0" style="min-width:100%;background:#c1272d; margin: 0;">
      <tr>
        <td><br></td>
      </tr>
    </table>
    <table width="100%" cellpadding="0" cellspacing="0" style="min-width:100%;background:#0d0d0d; margin: 0;">
      <tr>
        <td style="padding-left: 50px;padding-bottom: 30px;padding-top: 10px;font-size: 14px;line-height: 18px;vertical-align:bottom">
          <div class="footerText">Information Technology<br>
            Worcester Polytechnic Institute<br>
            <a href="mailto:ITS@wpi.edu" style="color:#c1272d">ITS@wpi.edu</a><br>
            +1 508-831-5888<br>
            100 Institute Road<br>
            Worcester, Massachusetts, 01609 United States
            </p>
        </td>
      </tr>
    </table>
    '


  $SendList_PendingAccountExpiration = 'windows@wpi.edu' #Domain Admins and Banner DBA
  $SendList_PendingAccountExpiration += 'kwheeler@wpi.edu', 'avalerio@wpi.edu' #Support Desk Staff
  $SendList_PendingAccountExpiration += 'lapierre@wpi.edu', 'maclancy@wpi.edu' # HR Team
  $SendList_PendingAccountExpiration += 'accountnotifications@wpi.edu' # Notification List managed by Deb Graves

  $SendList_PasswordNeverExpires = 'cdavidson@wpi.edu', 'windows@wpi.edu'
  $SendList_AccountExpired = 'cdavidson@wpi.edu', 'lapierre@wpi.edu', 'maclancy@wpi.edu', 'magillett@wpi.edu'
  $SendList_PasswordExpired = 'cdavidson@wpi.edu', 'windows@wpi.edu'


  #Set Log File Information
  #-----------------------------------------------------------------------------------------------
  $AccountLogFolderPath = '\\storage.wpi.edu\dept\Information Technology\CCC\CCC\fc_CCC_Temporary\Account Expirations'
  $PasswordLogFolderPath = '\\storage.wpi.edu\dept\Information Technology\CCC\CCC\fc_CCC_Temporary\Password Expirations'

  #Get and process list of users
  #-----------------------------------------------------------------------------------------------
  $users = Get-ADUser -Filter { (Enabled -eq $true) -and (EmailAddress -like "*") } -Properties Name, DisplayName, SamAccountName, Department, Title, OfficePhone, TelephoneNumber, EmployeeID, EmployeeNumber, Enabled, AccountExpirationDate, PasswordLastSet, PasswordExpired, PasswordNeverExpires, EmailAddress, LastLogonDate, msDS-UserPasswordExpiryTimeComputed, Created -SearchBase "OU=Accounts,DC=admin,DC=wpi,DC=edu" | Sort SamAccountName
  #$users = Get-ADUser -Filter * -Properties Name,DisplayName,SamAccountName,Department,Title,OfficePhone,TelephoneNumber,EmployeeID,EmployeeNumber,Enabled,AccountExpirationDate,PasswordLastSet,PasswordExpired,PasswordNeverExpires,EmailAddress,LastLogonDate,msDS-UserPasswordExpiryTimeComputed,Created -SearchBase "OU=Accounts,DC=admin,DC=wpi,DC=edu" | Where {$_.Enabled -eq $true -and $_.emailaddress -ne $null} | Sort SamAccountName
  #$users = Get-ADUser 'cidorr' -Properties Name,DisplayName,SamAccountName,Department,Title,OfficePhone,TelephoneNumber,EmployeeID,EmployeeNumber,Enabled,AccountExpirationDate,PasswordLastSet,PasswordExpired,PasswordNeverExpires,EmailAddress,LastLogonDate,msDS-UserPasswordExpiryTimeComputed,Created| Where {$_.Enabled -eq $true -and $_.emailaddress -ne $null} | Sort SamAccountName
  if ($users) {
    $PendingAccountExpirations = $users | Where { $_.AccountExpirationDate -ne $null -and $_.AccountExpirationDate -ge $today -and $_.AccountExpirationDate -lt $($today.AddDays(61)) } | Sort AccountExpirationDate
    $PendingPasswordExpirations = $users | Where { $_.PasswordNeverExpires -eq $false -and $_.PasswordExpired -eq $false -and $_.PasswordLastSet -lt $today.AddDays(-140) }
    $PasswordNeverExpires = $users | Where { $_.PasswordNeverExpires -eq $true -and $_.DistinguishedName -notmatch 'OU=Other Accounts' -and $_.DistinguishedName -notmatch 'OU=Services' } | Sort PasswordLastSet
    $AccountExpired = $users | Where { $_.AccountExpirationDate -ne $null -and $_.AccountExpirationDate -le $($today.AddDays(-90)) -and ($_.DistinguishedName -match 'OU=Employees' -or $_.DistinguishedName -match 'OU=Students') } | Sort AccountExpirationDate
    $PasswordExpired = $users | Where { $_.Enabled -and $_.PasswordExpired -and $_.AccountExpirationDate -eq $null -and $_.LastLogonDate -le $($today.AddDays(-90)) -and $_.Created -lt $today.AddDays(-30) -and $_.DistinguishedName -notmatch 'OU=Other Accounts' -and $_.DistinguishedName -notmatch 'OU=Services' } | Sort PasswordLastSet

    if ($PendingAccountExpirations) {
      foreach ($user in $PendingAccountExpirations) {
        $AccountExpiresIn = $null; $AccountExpirationDate = $null; $emailaddress = $null; $userclass = $null
   
        #Calculate the number of days that the account expires in
        $emailaddress = $user.emailaddress
        if ($user.AccountExpirationDate) {
          $AccountExpirationDate = $User.AccountExpirationDate.AddDays(-1)
          $AccountExpiresIn = ($AccountExpirationDate).subtract($today).days
        }

        #Determine if Student or Employee
        if ($user.distinguishedname -match "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu") { $userclass = "Student" }
        elseif ($user.distinguishedname -match "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu") { $userclass = "Employee" }
        else { $userclass = "Other" }

        #Collect information for logging
        #--------------------------------
        $out = New-Object PSObject
        $out | add-member noteproperty Name $user.DisplayName
        $out | add-member noteproperty Username $user.samAccountName
        $out | add-member noteproperty Email $emailaddress
        $out | add-member noteproperty UserClass $userclass
        $out | add-member noteproperty EmployeeID $user.EmployeeID
        $out | add-member noteproperty PIDM $user.EmployeeNumber
        $out | add-member noteproperty Department $user.department
        $out | add-member noteproperty Title $user.Title
        $out | add-member noteproperty Extension $user.OfficePhone
        $out | add-member noteproperty AccountExpires $AccountExpirationDate.ToShortDateString()
        $out | add-member noteproperty DaysRemaining $AccountExpiresIn
    
        $PendingAccountExpirationList += $out

        #if($AccountExpiresIn) {
        if ($AccountExpiresIn -eq 30 -or $AccountExpiresIn -eq 15 -or $AccountExpiresIn -eq 7 -or $AccountExpiresIn -eq 1) {
          $Subject = "Action Needed: WPI Account $emailaddress will expire on " + $AccountExpirationDate.ToLongDateString()
          $BodyMain = "
                    <!--main body-->
                    <div style=""font-family:Arial"">
                        <table width=""100%"" cellpadding=""0"" cellspacing=""0"" style=""min-width:100%;margin: 0;"">
                            <tr>
                                <td style=""padding-left: 50px;padding-top: 10px;padding-bottom: 10px;"">
                                    <p>Dear $($user.GivenName) $($user.Surname),</p>
                                    <p>Your acount $emailaddress will expire on <strong>$($AccountExpirationDate.ToLongDateString())</strong>.</p>
                                    <p>For security purposes, WPI sets an expiration for accounts of employees that are <strong>hourly, part-time, or temporary.</strong></p>
                    
                                    <p>Ongoing access requires regular review by department. Managers may request extensions for up to 1 year by emailing the IT Service Desk.</p>
                
                                    <p>If the account expiration date is correct, no further action is necessary. Your WPI Account <strong>$emailaddress</strong>
                                    will expire on this date. Access
                                    to the IT services will no longer be available after this date. 
                                    
                                    <p>If your role is continuing, or you believe the account expiration date is incorrect, please contact your
                                    <strong>manager or sponsor</strong> and have them contact the <strong>IT Service Desk</strong>
                                    with the updated position end date.</p>

                                    <p><strong>Non-faculty research appointees</strong> (<em>Post-Doctoral Fellows, Research
                                    Associates, Research Scientists, Research Engineering, International non-degree
                                    Students, and International Visiting Scholars</em>)
                                    with appointments nearing their scheduled appointment end date are highly encouraged to contact the
                                    <strong>department administrator</strong> and <strong>sponsoring faculty</strong> if an
                                    extension of appointment may be in order.
                                    Substantial time is required to process an appointment extension so it is recommended
                                    that you contact your department and sponsoring faculty <em>at least 30 days prior to</em>
                                    your current appointment end date</em>.
                                    Late requests may result in delayed start dates or a lapse in appointment. Any concerns
                                    regarding a <strong>visa status</strong> should be directed to the <strong>International
                                    House (hartvig@wpi.edu)</strong>.
                                    </p>

                                    <p><em><strong>NOTE:</strong> No one from Information Technology will <strong>ever</strong> ask for your password. You should <strong>never</strong> provide it.</em></p>

                                    <h3>Need Help?</h3>
                                    <p>If you need help, simply reply to this message or call the IT Service Desk.<br>
                                      Thank you in advance for your prompt attention.</p>

                                </td>
                            </tr>
                        </table>
                    </div>             
                
                    "

          $messageParameters = @{
            Subject    = $Subject
            Body       = $HTMLStyle + $BodyHeader + $BodyMain + $BodyFooter
            From       = $HeaderFrom
            To         = $emailaddress
            #                 TO = 'cidorr@wpi.edu'
            SmtpServer = $SMTPServer
            Priority   = "High"
            ###Send-MailMessage doesn't have any capacity for inserting headers.  Below are ones that should be added if a different method is used.
            ###$emailmessage.headers.add("X-Header","WPI-Password-Expiration"),$mailmessage.Headers.Add(“X-Auto-Response-Suppress“, “DR, OOF, AutoReply“)
          }                        

          if ($i -gt 100) { Start-Sleep -s 300; $i = 1 }
                
          Send-MailMessage @messageParameters -BodyAsHtml 

          $AccountEmailLog += $out
          $i = $i + 1
        }
      }
      if ($AccountEmailLog) {
        $AccountEmailLog | Sort AccountExpires | Export-Csv ("c:\wpi\logs\Email Logs\Account_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
        $AccountEmailLog | Sort AccountExpires | Export-Csv ($AccountLogFolderPath + "\Email Logs\Account_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
      }

      if ($PendingAccountExpirationList) {
        $PendingAccountExpirationList | Sort AccountExpires | Export-Csv ("c:\wpi\logs\List\PendingExpirations" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
        $PendingAccountExpirationList | Sort AccountExpires | Export-Csv ($AccountLogFolderPath + "\List\PendingExpirations" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
      }
        
      #If ($today) {
      If ($today.DayOfWeek -eq "Tuesday") {
        $messageParameters = $null; $Subject = $null; $Body = $null; $body2table = $null

        $body2table = $null
        $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Account Expiration</td><td align='center'>Days Remaining</td></tr>"

        foreach ($user in $PendingAccountExpirationList) { $body2table += "    <tr><td>$($user.Name)</td><td>$($user.Username)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.AccountExpires)</td><td>$($user.DaysRemaining)</td></tr>" }
        $body2table += "</table>"
        $Subject = "WPI Account Expirations through $($today.AddDays(60))"
        $Body = "<p>The following list of accounts are classified with HR as temporary or part time and have been set to expire based on the information provided at account creation. <p>
                        <p>If the job record is still active, all that is needed to extend the account is an email from the department to <strong>its@wpi.edu</strong> with the new end date.  
                        Please note the following information for each person that should be extended:</p>
                        <ul>
                            <li><strong>Username</strong></li>
                            <li><strong>ID Number</strong></li>
                            <li><strong>New end date</strong></li>
			            </ul>

                        <p>If the job record has changed and should now be full time, permanent employment, please confirm that the record in Banner is correct with HR.  Once confirmed, you can email the <strong>IT Service Desk</strong> and we can remove the 
                        account expiration</p>
                        <p>If no request to extend is received, the accounts will stop working at the <strong>end of the day</strong> on given account expiration date.</p>
                        <p><em>Please note that employee terminations supersede account expirations.</em></p> 

                        <p><em>Each of the users below will recieve an email at 30, 15, 7 and 1 day(s) prior to the account termination.
                        This report lists all users that will expire over the next 60 days and is sent every Tuesday to HR, selected Adminstrative and IT Staff.</em></p>
    
                        $body2table"    
    
        $messageParameters2 = @{
          Subject    = $Subject
          Body       = $Body
          From       = $HeaderFrom
          To         = $SendList_PendingAccountExpiration
          SmtpServer = $SMTPServer
          Priority   = "High"
        }  
    
        Send-MailMessage @messageParameters2 -BodyAsHtml
      }
    }

    If ($PasswordNeverExpires -and $today.DayOfWeek -eq "Monday") {
      $messageParameters = $null; $Subject = $null; $Body = $null; $body2table = $null
            
      $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Password Last Set</td><td align='center'>Last Logon Date</td><td align='center'>User Type</td></tr>"

      foreach ($user in $PasswordNeverExpires) {
        $UserType = $null
        switch -Regex ($user.DistinguishedName) {
          'OU=Alumni' { $UserType = 'Alumni' }
          'OU=Disabled' { $UserType = 'Disabled' }
          'OU=Employees' { $UserType = 'Employees' }
          'OU=Leave Of Absence' { $UserType = 'Leave Of Absence' }
          'OU=No Office 365 Sync' { $UserType = 'No Office 365 Sync' }
          'OU=Other Accounts' { $UserType = 'Other Accounts' }
          'OU=Privileged' { $UserType = 'Privileged' }
          'OU=Retirees' { $UserType = 'Retirees' }
          'OU=Services' { $UserType = 'Services' }
          'OU=Students' { $UserType = 'Students' }
          'OU=Vokes' { $UserType = 'Vokes' }
          'OU=Work Study' { $UserType = 'Work Study' }
          default { $UserType = 'Unknown User Type' }
        }
        $body2table += "    <tr><td>$($user.DisplayName)</td><td>$($user.SamAccountName)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.PasswordLastSet)</td><td>$($user.LastLogonDate)</td><td>$($UserType)</td></tr>"
      }
      $body2table += "</table>"
      $Subject = "WPI Accounts Set Never to Expire [Total: $(($PasswordNeverExpires | Measure-Object).Count)]"
      $Body = "<p>The following list of accounts are set to never expire. <p>
                    $body2table
                    "
    
      $messageParameters = @{
        Subject    = $Subject
        Body       = $Body
        From       = $HeaderFrom
        To         = $SendList_PasswordNeverExpires
        SmtpServer = $SMTPServer
        Priority   = "High"
      }  
    
      Send-MailMessage @messageParameters -BodyAsHtml
    }
    If ($AccountExpired -and $today.DayOfWeek -eq "Monday") {
      $messageParameters = $null; $Subject = $null; $Body = $null; $body2table = $null
            
      $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>PIDM</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Account Expiration Date</td><td align='center'>Last Logon Date</td></tr>"

      foreach ($user in $AccountExpired) { $body2table += "    <tr><td>$($user.DisplayName)</td><td>$($user.SamAccountName)</td><td>$($user.EmployeeID)</td><td>$($user.EmployeeNumber)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.AccountExpirationDate)</td><td>$($user.LastLogonDate)</td></tr>" }
      $body2table += "</table>"
      $Subject = "WPI Accounts - Enabled and Expired over 90 days [Total: $(($AccountExpired | Measure-Object).Count)]"
      $Body = "<p>The following list of accounts are enabled but have already expired. <p>
                    $body2table
                    "
    
      $messageParameters = @{
        Subject    = $Subject
        Body       = $Body
        From       = $HeaderFrom
        To         = $SendList_AccountExpired
        SmtpServer = $SMTPServer
        Priority   = "High"
      }  
    
      Send-MailMessage @messageParameters -BodyAsHtml
    }

    #reset counter
    #-------------
    $i = 1

    if ($PendingPasswordExpirations) {
      foreach ($user in $PendingPasswordExpirations) {
        #Declare Variables
        #----------------------------
        $PasswordExpiresIn = $null; $PasswordExpiresOn = $null; $emailaddress = $null; $userclass = $null; $AccountExpiresIn = $null; $AccountExpirationDate = $null
        $Subject = $null; $BodyPrefix = $null; $BodyAdmin = $null; $BodyStudent = $null; $BodySuffix = $null; $messageParameters = $null;
        
        #Set Variables
        #----------------------------
        $emailaddress = $user.EmailAddress
        $PasswordExpiresOn = [datetime]::fromfiletime($user."msDS-UserPasswordExpiryTimeComputed")
        if ($user.AccountExpirationDate) {
          $AccountExpirationDate = $User.AccountExpirationDate.AddDays(-1)
          $AccountExpiresIn = ($AccountExpirationDate).subtract($today).days
        }
            
        if ($null -ne $PasswordExpiresOn) { $PasswordExpiresIn = ($PasswordExpiresOn).subtract($today).days }
        else { $passwordexpiresin = 0 }
    
        if ($passwordexpiresin -lt 30 -and $passwordexpiresin -gt 0) {
          #Determine if Student or Employee
          if ($user.distinguishedname -match "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu") { $userclass = "Student" }
          elseif ($user.distinguishedname -match "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu") { $userclass = "Employee" }
          else { $userclass = "Other" }

          #Collect information for logging
          $out = New-Object PSObject
          $out | add-member noteproperty Name $user.DisplayName
          $out | add-member noteproperty Username $user.samAccountName
          $out | add-member noteproperty Email $emailaddress
          $out | add-member noteproperty UserClass $userclass
          $out | add-member noteproperty EmployeeID $user.EmployeeID
          $out | add-member noteproperty PIDM $user.EmployeeNumber
          $out | add-member noteproperty Department $user.department
          $out | add-member noteproperty Title $user.Title
          $out | add-member noteproperty Extension $user.OfficePhone
          $out | add-member noteproperty PasswordExpires $passwordexpireson
          $out | add-member noteproperty DaysRemaining $passwordexpiresin
          $out | add-member noteproperty AccountExpires $AccountExpirationDate

          $PendingPasswordExpirationList += $out

          if ($passwordexpiresin -eq 20 -or $passwordexpiresin -eq 15 -or $passwordexpiresin -le 7) {
            $messageParameters = $null; $Subject = $null; $Body = $null

            $Subject = "Action Needed: Your WPI Password for account $emailaddress expires on " + $PasswordExpiresOn.ToLongDateString()
            $BodyMain = "
                        <!--main body-->
                        <div style=""font-family:Arial"">
                          <table width=""100%"" cellpadding=""0"" cellspacing=""0"" style=""min-width:100%;margin: 0;"">
                            <tr>
                              <td style=""padding-left: 50px;padding-top: 10px;padding-bottom: 10px;"">
                                <p>Dear $($user.GivenName) $($user.Surname),</p>
                                  <p><strong>Your WPI Account Password for <font color=""red"">$emailaddress</font> expires on
                                      <font color=""red"">$PasswordExpiresOn.</font></strong></p>
                                  <hr>

                                  <h2>ACTION NEEDED: PASSWORD UPDATE</h2>

                                  <p>To ensure that there is no disruption of service, you should take time now to change the password. At least 10 characters are now required for WPI account passwords. To perform this remotely using a WPI-managed computer, please connect to the WPI VPN. If you use a WPI-managed Mac, please follow these instructions: <a href=""https://hub.wpi.edu/article/419/getting-started-with-nomad"">https://hub.wpi.edu/article/419/getting-started-with-nomad</a>.</p>

                                  <h3><em>If you have already configured WPI Self Service Password Reset (SSPR):</em></h3>
                                  <p>You can change your password through SSPR. You can reset your password online at:<br>
                                    <a href=""https://aka.ms/sspr"">https://aka.ms/sspr</a></p>

                                  <h3><em>If you have NOT set up SSPR, and do NOT KNOW your current password:</em></h3>
                                  <p>You must contact the IT Service Desk in person or over the phone.<br>
                                    Hours of operation can be found at:<br>
                                    <a href=""https://its.wpi.edu/Help"">https://its.wpi.edu/Help</a></p>

                                  <h3>Notes:</h3>
                                  <ul>
                                    <li>After setting a new password, please remember to change your password on all devices you use (laptops,
                                      phones, tablets, etc.) </li>
                                    <li>Please allow up to 30 minutes for your new password to be recognized across WPI systems.</li>
                                    <li>No one from Information Technology will <strong>ever</strong> ask for your password. You should never provide it. View our guidelines for creating secure passwords:<br>
                                      <a href=""https://hub.wpi.edu/article/454/wpiuser-account-responsibilities"">https://hub.wpi.edu/article/454/wpiuser-account-responsibilities</a></li>
                                  </ul>


                                  <h3>Need Help?</h3>
                                  <p>If you need help, simply reply to this message or call the IT Service Desk.<br>
                                    Thank you in advance for your prompt attention.</p>

                              </td>
                            </tr>
                          </table>
                        </div>
                        "

            $messageParameters = @{
              Subject    = $Subject
              Body       = $HTMLStyle + $BodyHeader + $BodyMain + $BodyFooter
              From       = $HeaderFrom
              To         = $emailaddress
              #          TO = 'cidorr@wpi.edu'
              SmtpServer = $SMTPServer
              Priority   = "High"
              ###Send-MailMessage doesn't have any capacity for inserting headers.  Below are ones that should be added if a different method is used.
              ###$emailmessage.headers.add("X-Header","WPI-Password-Expiration"),$mailmessage.Headers.Add(“X-Auto-Response-Suppress“, “DR, OOF, AutoReply“)
            }                        

            if ($i -gt 100) { Start-Sleep -s 300; $i = 1 }
            Send-MailMessage @messageParameters -BodyAsHtml
            
            $PasswordEmailLog += $out
        
            $i = $i + 1
          }
        }  
      }

      if ($PasswordEmailLog) {
        $PasswordEmailLog | Sort DaysRemaining | Export-Csv ("c:\wpi\logs\Email Logs\Password_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
        $PasswordEmailLog | Sort DaysRemaining | Export-Csv ($PasswordLogFolderPath + "\Email Logs\Password_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
      }
    
      if ($PendingPasswordExpirationList) {
        $PendingPasswordExpirationList | Sort DaysRemaining | Export-Csv ("c:\wpi\logs\List\Password_Expiration_List_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
        $PendingPasswordExpirationList | Sort DaysRemaining | Export-Csv ($PasswordLogFolderPath + "\List\Password_Expiration_List_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
      }
    }
  }
  else {
    #send TC an email that getting $users failed
  }

  #Purge-Logs "c:\wpi\logs\Email Logs\"
  #Purge-Logs "c:\wpi\logs\List\"
  #Purge-Logs "$AccountLogFolderPath\Email Logs\"
  #Purge-Logs "$AccountLogFolderPath\List\"
  #Purge-Logs "$PasswordLogFolderPath\Email Logs\"
  #Purge-Logs "$PasswordLogFolderPath\List\"

}

########################
##  Main Code
########################
Get-AccountExpirations
Exit