Clear-Host
Write-Host "Script Begin $(get-date)"
<#
EmailExpirations.ps1
Written by: Tom Collins (tcollins)
Last updated: 9/5/17

Notes:
Script Processes in about 20 minutes
#>


########################
##  FUNCTIONS
########################

Function Purge-Logs ($path) {
    $TargetFolder=$null;$Days=$null;$Now=$null;$File=$null;$Folders=$null;$SubFolder=$null;$LastWrite=$null

    $TargetFolder = $path
    $Days = 365
    $Now = Get-Date

    $Folders = Get-ChildItem $TargetFolder

    # Notice the minus sign before $days
    $LastWrite = $Now.AddDays(-$days)

    ForEach ($File in $Folders) {if ($file.LastWriteTime -lt $LastWrite) {Remove-Item $file.VersionInfo.FileName | out-null}}
    }

########################
##  Main Code
########################
Import-Module ActiveDirectory

#Declare Variables
#-----------------------------------------------------------------------------------------------
$users=$null;$out=$null;$PendingAccountExpirations=$null;$PendingPasswordExpirations=$null;$today=$null;$maxpassage=$null
$AccountEmailLog=@();$PasswordEmailLog=@();$PendingAccountExpirationList=@();$PendingPasswordExpirationList=@()

#Set Variables
#-----------------------------------------------------------------------------------------------
$i = 1
$today = Get-Date

#Set email information.  For multiple values in the "To" field, seperate the values with a comma
#-----------------------------------------------------------------------------------------------
$HeaderFrom = 'its@wpi.edu'
$SMTPServer = 'smtp.wpi.edu'

#$AccountSendList  = 's_tcollins@wpi.edu'
$AccountSendList  = "tcollins@wpi.edu","roger@wpi.edu","cdrenaud@wpi.edu","cidorr@wpi.edu","diruzza@wpi.edu","kwheeler@wpi.edu","debra@wpi.edu","tmlamantia@wpi.edu","caphelan@wpi.edu","lapierre@wpi.edu","accountnotifications@wpi.edu"

#Set Log File Information
#-----------------------------------------------------------------------------------------------
$AccountLogFolderPath  = '\\storage.wpi.edu\dept\Information Technology\CCC\CCC\fc_CCC_Temporary\Account Expirations'
$PasswordLogFolderPath = '\\storage.wpi.edu\dept\Information Technology\CCC\CCC\fc_CCC_Temporary\Password Expirations'

#Get and process list of users
#-----------------------------------------------------------------------------------------------
$users = Get-ADUser -Filter * -Properties Name,DisplayName,SamAccountName,Department,Title,OfficePhone,TelephoneNumber,EmployeeID,EmployeeNumber,Enabled,AccountExpirationDate,PasswordLastSet,PasswordExpired,PasswordNeverExpires,EmailAddress,LastLogonDate,msDS-UserPasswordExpiryTimeComputed -SearchBase "OU=Accounts,DC=admin,DC=wpi,DC=edu" | Where {$_.Enabled -eq $true -and $_.emailaddress -ne $null} | Sort SamAccountName

if ($users) {
    $PendingAccountExpirations  = $users | Where {$_.AccountExpirationDate -ne $null -and $_.AccountExpirationDate -ge $today -and $_.AccountExpirationDate -lt $($today.AddDays(61))} | Sort AccountExpirationDate
    $PendingPasswordExpirations = $users | Where {$_.PasswordNeverExpires -eq $false -and $_.PasswordExpired -eq $false -and $_.PasswordLastSet -lt $today.AddDays(-140)}
    $PasswordNeverExpires       = $users | Where {$_.PasswordNeverExpires -eq $true -and ($_.DistinguishedName -match 'OU=Employees' -or $_.DistinguishedName -match 'OU=Students')} | Sort PasswordLastSet
    $AccountExpired             = $users | Where {$_.AccountExpirationDate -ne $null -and $_.AccountExpirationDate -le $($today.AddDays(-30)) -and ($_.DistinguishedName -match 'OU=Employees' -or $_.DistinguishedName -match 'OU=Students')} | Sort AccountExpirationDate

    if ($PendingAccountExpirations) {
        foreach ($user in $PendingAccountExpirations) {
            $AccountExpiresIn=$null;$AccountExpirationDate=$null;$emailaddress=$null;$userclass=$null
   
            #Calculate the number of days that the account expires in
            $emailaddress = $user.emailaddress
            if ($user.AccountExpirationDate) {
                $AccountExpirationDate = $User.AccountExpirationDate.AddDays(-1)
                $AccountExpiresIn = ($AccountExpirationDate).subtract($today).days
                }

            #Determine if Student or Employee
            if ($user.distinguishedname -match "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu") {$userclass = "Student"}
            elseif ($user.distinguishedname -match "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu") {$userclass = "Employee"}
            else {$userclass = "Other"}

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
            if($AccountExpiresIn -eq 30 -or $AccountExpiresIn -eq 15 -or $AccountExpiresIn -eq 7 -or $AccountExpiresIn -eq 1) {
		        $Subject = "Action Needed: Your WPI Account [ADMIN\$($User.samaccountname)] will expire on " + $AccountExpirationDate.ToLongDateString()
                $Body    = "
                <p>Dear $($user.GivenName) $($user.Surname),
                <p>According to our records, your computer account status at WPI is <strong>temporary</strong> or <strong>part-time</strong>.  
                Your department has determined your computer account access at WPI will end on <strong>$($AccountExpirationDate.ToLongDateString())</strong>.<p>
                
                <p>If you believe this date is incorrect, please contact your 
                <strong>department</strong> and have them contact the <strong>ITS Service Desk</strong> with the updated position end date.</p>
                <p>If the end date is correct, no further action is necessary. Your WPI Account <strong>(ADMIN\$($User.samaccountname))</strong> will expire on this date. Access 
                to the ITS services will no longer be available after this date.  These include, but are not limited to, the following:</p>

                <ol>
                    <li><strong>WPI Web Information System</strong> (https://bannerweb.wpi.edu)</li>
                    <li><strong>Email</strong> (https://outlook.wpi.edu)</li>
                    <li><strong>Canvas</strong> (https://canvas.wpi.edu)</li>
                    <li><strong>VPN</strong> (https://vpn.wpi.edu)</li>
                    <li>Many other IT services</li>
			    </ol>
        
                <p><strong>Non-faculty research appointees</strong> (<em>Post-Doctoral Fellows, Research Associates, Research Scientists, Research Engineering, International non-degree Students, and International Visiting Scholars</em>) 
                with appointments nearing their scheduled end date are highly encouraged to contact the <strong>department administrator</strong> and <strong>sponsoring faculty</strong> if an extension of appointment may be in order. 
                Substantial time is required to process an appointment extension so it is recommended that you contact your department and sponsoring faculty at least <em>30 days before your current appointment end date</em>. 
                Late requests may result in delayed start dates or a lapse in appointment. Any concerns regarding a <strong>visa status</strong> should be directed to the <strong>International House (hartvig@wpi.edu)</strong>.  
                </p>
        
                <p>If you need help preparing for the account expiration, please reply to this message or call the <strong>ITS Service Desk</strong> at <strong>(508) 831-5888</strong>.</p> 

                <p>Thank you in advance for your prompt attention.</p>
        
                <p><strong>ITS Service Desk</strong><br>
                http://its.wpi.edu<br>
                (508) 831-5888 | Ext. 5888</p>

			    <p><strong>NOTICE:</strong> No one from Information Technology will ever ask for your 
			    password. You should never provide it.</p>                
                
                "

                $messageParameters = @{
                    Subject = $Subject
                    Body = $Body
                    From = $HeaderFrom
                    #To   = $emailaddress
                    To   = 's_tcollins@wpi.edu'
                    #Bcc  = 'acadtest@wpi.edu'
                    SmtpServer = $SMTPServer
                    Priority = "High"
                    ###Send-MailMessage doesn't have any capacity for inserting headers.  Below are ones that should be added if a different method is used.
                    ###$emailmessage.headers.add("X-Header","WPI-Password-Expiration"),$mailmessage.Headers.Add(“X-Auto-Response-Suppress“, “DR, OOF, AutoReply“)
                    }                        

                if ($i -gt 100) {Start-Sleep -s 300;$i = 1}

                Send-MailMessage @messageParameters -BodyAsHtml

    	        $AccountEmailLog += $out
                $i = $i + 1
                }
            }
        if ($AccountEmailLog) {
            $AccountEmailLog | Sort AccountExpires | Export-Csv ("logs\Email Logs\Account_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            $AccountEmailLog | Sort AccountExpires | Export-Csv ($AccountLogFolderPath + "\Email Logs\Account_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            }

        if ($PendingAccountExpirationList) {
            $PendingAccountExpirationList | Sort AccountExpires | Export-Csv ("logs\List\PendingExpirations" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            $PendingAccountExpirationList | Sort AccountExpires | Export-Csv ($AccountLogFolderPath + "\List\PendingExpirations" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            }
        
        #If ($today) {
        If ($today.DayOfWeek -eq "Tuesday") {
            $messageParameters=$null;$Subject=$null;$Body=$null;$body2table=$null

            $body2table = $null
            $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Account Expiration</td><td align='center'>Days Remaining</td></tr>"

            foreach ($user in $PendingAccountExpirationList) {$body2table += "    <tr><td>$($user.Name)</td><td>$($user.Username)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.AccountExpires)</td><td>$($user.DaysRemaining)</td></tr>"}
            $body2table += "</table>"
            $Subject  = "WPI Account Expirations through $($today.AddDays(60))"
            $Body     = "<p>The following list of accounts are classified with HR as temporary or part time and have been set to expire based on the information provided at account creation. <p>
                        <p>If the job record is still active, all that is needed to extend the account is an email from the department to <strong>its@wpi.edu</strong> with the new end date.  
                        Please note the following information for each person that should be extended:</p>
                        <ul>
                            <li><strong>Username</strong></li>
                            <li><strong>ID Number</strong></li>
                            <li><strong>New end date</strong></li>
			            </ul>

                        <p>If the job record has changed and should now be full time, permanent employment, please confirm that the record in Banner is correct with HR.  Once confirmed, you can email the <strong>ITS Service Desk</strong> and we can remove the 
                        account expiration</p>
                        <p>If no request to extend is received, the accounts will stop working at the <strong>end of the day</strong> on given account expiration date.</p>
                        <p><em>Please note that employee terminations supersede account expirations.</em></p> 

                        <p><em>Each of the users below will recieve an email at 30, 15, 7 and 1 day(s) prior to the account termination.
                        This report lists all users that will expire over the next 60 days and is sent every Tuesday to HR, selected Adminstrative and ITS Staff.</em></p>
    
                        $body2table"    
    
            $messageParameters2 = @{
                Subject = $Subject
                Body = $Body
                From = $HeaderFrom
                To = $AccountSendList
                Bcc = 's_tcollins@wpi.edu'
                SmtpServer = $SMTPServer
                Priority = "High"
                }  
    
            Send-MailMessage @messageParameters2 -BodyAsHtml
            }
        }

    If ($PasswordNeverExpires -and $today.DayOfWeek -eq "Tuesday") {
        $messageParameters=$null;$Subject=$null;$Body=$null;$body2table=$null
            
        $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Password Last Set</td></tr>"

        foreach ($user in $PasswordNeverExpires) {$body2table += "    <tr><td>$($user.DisplayName)</td><td>$($user.SamAccountName)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.PasswordLastSet)</td></tr>"}
        $body2table += "</table>"
        $Subject = "WPI Accounts Set Never to Expire [Total: $(($PasswordNeverExpires | Measure-Object).Count)]"
        $Body    =  "<p>The following list of accounts are set to never expire. <p>
                    $body2table
                    "
    
        $messageParameters = @{
            Subject = $Subject
            Body = $Body
            From = $HeaderFrom
            To = 'tcollins@wpi.edu','djjones@wpi.edu','ejmartin2@wpi.edu'
            Bcc = 's_tcollins@wpi.edu'
            SmtpServer = $SMTPServer
            Priority = "High"
            }  
    
        Send-MailMessage @messageParameters -BodyAsHtml
        }
    If ($AccountExpired -and $today.DayOfWeek -eq "Tuesday") {
        $messageParameters=$null;$Subject=$null;$Body=$null;$body2table=$null
            
        $body2table += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>ID Number</td><td align='center'>PIDM</td><td align='center'>Department</td><td align='center'>Title</td><td align='center'>Account Expiration Date</td><td align='center'>Last Logon Date</td></tr>"

        foreach ($user in $AccountExpired) {       $body2table += "    <tr><td>$($user.DisplayName)</td><td>$($user.SamAccountName)</td><td>$($user.EmployeeID)</td><td>$($user.EmployeeNumber)</td><td>$($user.Department)</td><td>$($user.Title)</td><td>$($user.AccountExpirationDate)</td><td>$($user.LastLogonDate)</td></tr>"}
        $body2table += "</table>"
        $Subject = "WPI Accounts - Enabled and Expired over 30 days [Total: $(($AccountExpired | Measure-Object).Count)]"
        $Body    =  "<p>The following list of accounts are enabled but have already expired. <p>
                    $body2table
                    "
    
        $messageParameters = @{
            Subject = $Subject
            Body = $Body
            From = $HeaderFrom
            To = 'tcollins@wpi.edu','djjones@wpi.edu','ejmartin2@wpi.edu'
            Bcc = 's_tcollins@wpi.edu'
            SmtpServer = $SMTPServer
            Priority = "High"
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
            $PasswordExpiresIn=$null;$PasswordExpiresOn=$null;$emailaddress=$null;$userclass=$null;$AccountExpiresIn=$null;$AccountExpirationDate=$null
            $Subject=$null;$BodyPrefix=$null;$BodyAdmin=$null;$BodyStudent=$null;$BodySuffix=$null;$messageParameters=$null;
        
            #Set Variables
            #----------------------------
            $emailaddress = $user.EmailAddress
            $PasswordExpiresOn = [datetime]::fromfiletime($user."msDS-UserPasswordExpiryTimeComputed")
            if ($user.AccountExpirationDate) {
                $AccountExpirationDate = $User.AccountExpirationDate.AddDays(-1)
                $AccountExpiresIn = ($AccountExpirationDate).subtract($today).days
                }
            
            if ($PasswordExpiresOn -ne $null) {$PasswordExpiresIn = ($PasswordExpiresOn).subtract($today).days}
            else {$passwordexpiresin = 0}
    
            if($passwordexpiresin -lt 30 -and $passwordexpiresin -gt 0) {
                #Determine if Student or Employee
                if ($user.distinguishedname -match "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu") {$userclass = "Student"}
                elseif ($user.distinguishedname -match "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu") {$userclass = "Employee"}
                else {$userclass = "Other"}

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

                if($passwordexpiresin -eq 20 -or $passwordexpiresin -eq 15 -or $passwordexpiresin -le 7) {
                    $messageParameters=$null;$Subject=$null;$Body=$null

                    $Subject = "Action Needed: Your WPI Password for account ADMIN\$($User.samaccountname) expires on " + $PasswordExpiresOn.ToLongDateString()
                    $Body = "
                    <p>Dear $($user.GivenName) $($user.Surname),
                    <p><strong>Your WPI Windows Password for account <font color=""red"">ADMIN\$($User.samaccountname)</font> expires on <font color=""red"">$PasswordExpiresOn.</font></strong></p>
                    <br><hr><br>

                    <h2>ACTION NEEDED: PASSWORD UPDATE</h2>
			        <p>To ensure that there is no disruption of service, you should take time now to change the password.</p>
			        <p><strong><em>If you know your current WPI password</em></strong>, change it by logging into WPI's VPN.</p>
			        <p><em>Please note, this only works on desktop browsers, not mobile devices</em></p>
                    <ol>
				        <li>Open any browser to https://vpn.wpi.edu</li>
				        <li>Click <strong>Proceed</strong></li>
				        <li>Sign in with <em>$($User.samaccountname)@wpi.edu </em>and your <em>current
				        </em>password</li>
				        <li>If you are not immediately prompted to change your password, click on the 
				        <strong>Preferences </strong>icon in the upper right</li>
				        <li>From within the Preferences window, select the <strong>General
				        </strong>tab</li>
				        <li>Create your new password in the provided fields</li>
				        <li>Click <strong>Sign Out </strong>in the upper right and you're done! 
				        <em>No need to install or start VPN</em></li>
			        </ol>
			        <p>OR</p>
			        <p><strong><em>If you do not know your current WPI password</em></strong>, reset it using the Account Maintenance Page.</p>
			        <ol>
				        <li>Open any browser to https://its.wpi.edu/accounts</li>
				        <li>Under the <strong>Windows Account</strong> section Select <strong>Change password.</strong></li>
				        <li>Click <strong>Begin Account Maintenance</strong></li>
				        <li>Select <strong>Change your Windows password </strong>and click 
				        <strong>Next</strong></li>
				        <li>Enter your <strong>WPI ID number </strong>and <strong>PIN
				        </strong>and click <strong>Next</strong></li>
				        <li>Type your new password in the boxes and click <strong>Next</strong><br />
				        </li>
			        </ol>

			        <p><strong>NOTE: </strong>After setting a new password, please remember to change your password on all mobile devices (phones, tablets)</p>

			        <br><hr><br>

                    <h2>HELP:</h2>
			        <p>If you need help, simply reply to this message or call the ITS Service Desk.</p>
			        <p>Thank you in advance for your prompt attention.</p>
			        <br><br>
                    <p><strong>ITS Service Desk</strong><br />
			        http://its.wpi.edu<br />
			        (508) 831-5888 | Ext. 5888</p>

			        <br><hr><br>

			        <h2>ADDITIONAL INFORMATION:</h2>
			        <p><strong>ACCOUNT:</strong> Your WPI Account provides access to services such as:</p>
			        <ul>
                        <li><strong>WPI Web Information System</strong> (https://bannerweb.wpi.edu)</li>
                        <li><strong>Email</strong> (https://outlook.wpi.edu)</li>
                        <li><strong>Canvas</strong> (https://canvas.wpi.edu)</li>
                        <li><strong>VPN</strong> (https://vpn.wpi.edu)</li>
				        <li>Many other IT services</li>
			        </ul>
			        <p><strong>GUIDELINES:</strong> No one from Information Technology will ever ask for your 
			        password. You should never provide it.  Guidelines for creating secure passwords are located here:</p>
			        <p>http://its.wpi.edu/Article/Password-Responsibilities</p>

                    "

                    $messageParameters = @{
                        Subject = $Subject
                        Body = $Body
                        From = $HeaderFrom
                        To   = $emailaddress
                        Bcc  = "s_tcollins@wpi.edu"
                        SmtpServer = $SMTPServer
                        Priority = "High"
                        ###Send-MailMessage doesn't have any capacity for inserting headers.  Below are ones that should be added if a different method is used.
                        ###$emailmessage.headers.add("X-Header","WPI-Password-Expiration"),$mailmessage.Headers.Add(“X-Auto-Response-Suppress“, “DR, OOF, AutoReply“)
                        }                        

                    if ($i -gt 100) {Start-Sleep -s 300;$i = 1}
                    Send-MailMessage @messageParameters -BodyAsHtml
            
                    $PasswordEmailLog += $out
        
                    $i = $i + 1
                    }
                }  
            }

        if ($PasswordEmailLog) {
            $PasswordEmailLog | Sort DaysRemaining | Export-Csv ("logs\Email Logs\Password_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            $PasswordEmailLog | Sort DaysRemaining | Export-Csv ($PasswordLogFolderPath + "\Email Logs\Password_Expiration_Email_Logs_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            }
    
        if ($PendingPasswordExpirationList) {
            $PendingPasswordExpirationList | Sort DaysRemaining | Export-Csv ("logs\List\Password_Expiration_List_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            $PendingPasswordExpirationList | Sort DaysRemaining | Export-Csv ($PasswordLogFolderPath + "\List\Password_Expiration_List_" + (get-date -uformat "%Y-%m-%d_%H%M") + ".csv") -NoTypeInformation
            }
        }
    }
else {
    #send TC an email that getting $users failed
    }

Purge-Logs "logs\Email Logs\"
Purge-Logs "logs\List\"
Purge-Logs "$AccountLogFolderPath\Email Logs\"
Purge-Logs "$AccountLogFolderPath\List\"
Purge-Logs "$PasswordLogFolderPath\Email Logs\"
Purge-Logs "$PasswordLogFolderPath\List\"

Write-Host "Script Complete $(get-date)"