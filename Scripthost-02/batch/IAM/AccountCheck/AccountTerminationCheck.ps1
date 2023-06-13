<# AccountTerminationCheck.ps1

.SYNOPSIS
    This script uses several methods to check whether a Linux account should or should not exist.

.DESCRIPTION
    This script uses a combination of LDAPSEARCH and Test-Path to determine if accounts which exist in AD should also exist in LDAP.
    The script also checks for valid accounts' home directories. Accounts that should exist and don't or do exist and shouldn't are
    reported to a select group of people via email.

.NOTES
    Author: tcollins
    Date: 10/24/2017

    Modified By: Stephen Gemme
    Modified On: 09/01/20
    Modifications:  A Lot. Completely overhauled the script so that it searches LDAP for users rather than assuming if they have a home directory
                    then they have a Linux account. It still searches for home directories and reports them separately. Also changed the search
                    and filter so that we're not basing our active account assumptions off which OU it's currently in, but rather if the account
                    is Enabled and has a proper extensionAttribute7 (staff/student/affiliate).

                    Other changes are logged in git.
#>

Param (
    [Parameter(Mandatory = $false)]
    [Switch]
    $testMode
)

$Global:LDAPConnection

function Test-LDAPConnection(){
    Write-Host "Testing LDAP Connection"
    # Setup our necesities.
    $ServerName = "ldapv2.wpi.edu"
    $Port = 636
    
    # Setup our connector.
    $Global:LDAPConnection = New-Object System.DirectoryServices.Protocols.LdapConnection ("$ServerName"+":"+"$Port")
    # Set the bind type.
    $Global:LDAPConnection.AuthType = [System.DirectoryServices.Protocols.AuthType]::Anonymous

    #Set session options (SSL + LDAP V3)
    $Global:LDAPConnection.SessionOptions.SecureSocketLayer = $false
    $Global:LDAPConnection.SessionOptions.ProtocolVersion = 3

    # Bind with no credentials since it's anonymous.
    try {
        $Global:LDAPConnection.Bind($null)
        Write-Host -foregroundColor GREEN "Successfully Connected to LDAP!"
        return $true
    }
    catch {
        $Global:LDAPConnection = $null
        Throw "Error connection to LDAP - $($_.Exception.Message)"
        return $false
    }
}

function ldapSearch(){
    # Setup our necesities.
    Param (
        # Make sure we get the username before setting the filter.
        [Parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$true)]
        [String]
        $userName
    )

    $basedn = "OU=People,DC=wpi,DC=edu"
    $scope = [System.DirectoryServices.Protocols.SearchScope]::Subtree
    #Null returns all available attributes
    $attrlist = $null
    $filter = "uid=$userName"

    # Compile our query.
    $ModelQuery = New-Object System.DirectoryServices.Protocols.SearchRequest -ArgumentList $basedn, $filter, $scope, $attrlist

    # Execute the query and save the results.
    try {
        $ModelRequest = $Global:LDAPConnection.SendRequest($ModelQuery)
    }
    catch {
        $ModelRequest = $null
        Throw "`nProblem looking up model account - $($_.Exception.Message)"
    }

    # If we found someone, return it!
    if ($ModelRequest.Entries.count -ge 1){
        return $true
    }
    else {
        return $false
    }

}

Clear-Host

# Get date for logging and file naming:
$date      = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Get our logging going.
Start-Transcript -Path "D:\wpi\Logs\IAM\AccountTerminaionCheck\AccountTermination-$datestamp.txt" -Append

########################
##  Main Code
########################
Import-Module ActiveDirectory

# Anyone who has an active AD account but no Linux Account
$MissingLinuxAccount = @()
# Anyone with an active AD account but no Home Dir.
$MissingHomeDir = @()
# Anyone in Disabled/LOA/Alumni and also has a Linux Account
$HasLinuxAccount = @()
# Anyone in Disabled/LOA/Alumni and also has a Home Dir
$HasHomeDir = @()

#Set email information.
$HeaderFrom = 'its@wpi.edu'
$SMTPServer = 'smtp.wpi.edu'

if ($testMode){
    #$Recipients = 'sgemme@wpi.edu'
    $Recipients = 'sgemme@wpi.edu', 'lleclerc@wpi.edu'
}
else {
    $Recipients = 'cdrenaud@wpi.edu','roger@wpi.edu','jpeismeier@wpi.edu','cidorr@wpi.edu','lleclerc@wpi.edu','sgemme@wpi.edu','mtaylor@wpi.edu','cwvangelder@wpi.edu'
}

Write-Host "Setting up config." -noNewLine

# Setup our base filters.
# This is gross but gets exactly who we're looking for: all enabled accounts that are either a Student OR (Are faculty, staff, or affiliate who aren't RE or WPI Affiliates).
$ExtensionAttributeFilter = '((ExtensionAttribute7 -eq "Student" -and ExtensionAttribute8 -ne "LA") -or ((Extensionattribute7 -eq "Faculty" -or Extensionattribute7 -eq "Staff" -or Extensionattribute7 -eq "Affiliate") -and (Extensionattribute1 -notlike "*" -or Extensionattribute1 -ne "RE") -and (Department -notlike "*" -or Department -ne "Wpi Affiliates")))'

# Specify our exceptions.
$IgnoreNamesFilter = '(SamAccountName -notlike "*_*" -and SamAccountName -ne "muditsharma" -and SamAccountName -ne "ssundaresan" -and SamAccountName -ne "mlerch" -and SamAccountName -ne "lhubenthal" -and SamAccountName -ne "ddenault" -and SamAccountName -ne "phoerter" -and SamAccountName -notlike "*guest*")'

# Get users in the specified group - saved here in case ever needed.
# $GroupFilter = 'memberOf -eq "CN=AdobeUserSync-Student,OU=Resources,OU=Groups,DC=admin,DC=wpi,DC=edu"'

# Get the attributes we care about most.
$Properties = @("Name", "SamAccountName", "Department", "Title", "EmployeeID", "EmployeeNumber", "Enabled", "extensionAttribute1", "extensionAttribute7", "whenCreated")

Write-Host -foregroundColor GREEN "Done."


# Before we gather anything, make sure we can connect to LDAP.
if (Test-LDAPConnection){
    # Build and execute our search for those who should have Linux accounts and Home Directories.
    Write-Host "Gathering active users..." -noNewLine

    $SearchFilter = 'Enabled -eq $true -and ' + $ExtensionAttributeFilter + ' -and ' + $IgnoreNamesFilter
    # We need to filter out LOAs, but we can't use a filter since DistinguishedName is calculated, so we have to where-object.
    $ShouldHaveLinuxandHome = (Get-ADUser -Filter $SearchFilter -Property $Properties | Where-Object DistinguishedName -notlike "*OU=Leave Of Absence*")

    Write-Host -foregroundColor GREEN "Done."

    # Setup things for the progress bar.
    $totalUsers = $ShouldHaveLinuxandHome.count
    $numSearched = 0

    # Search LDAP for each account and record the results.
    Write-Host "Parsing through active users for issues..." -noNewLine
    foreach ($user in $ShouldHaveLinuxandHome) {
        # Display our progress to the user.
        $percentComplete = ($numSearched / $totalUsers * 100)
        Write-Progress -Activity "Searching LDAP for Users" -Status "$numSearched / $totalUsers Complete:" -PercentComplete $percentComplete

        if (-NOT ($user.SamAccountName | ldapSearch)){
            $MissingLinuxAccount += $user
        }
        if (!(Test-Path "\\storage.wpi.edu\homes\$($user.SamAccountName)")) {
            $MissingHomeDir += $user
        }

        $numSearched++
    }
    Write-Host -foregroundColor GREEN "Done."

    # Build and execute our search for those who should NOT have Linux accounts and Home Directories.
    Write-Host "Gathering inactive users..." -noNewLine

    # We have to use the Where-Object clause because if you filter for extensionAttribute7 -ne "Resource Mailbox" it fails on all those users with null entries and skips them.
    $SearchFilter = '(Enabled -eq $false -or extensionattribute7 -eq "Alum") -and ' + $IgnoreNamesFilter
    $ShouldNotHaveLinuxandHome = Get-ADUser -Filter $SearchFilter -Property $Properties | Where-Object {$_.extensionAttribute7 -ne "Resource Mailbox"}

    Write-Host -foregroundColor GREEN "Done."

    # Setup things for the progress bar.
    $totalUsers = $ShouldNotHaveLinuxandHome.count
    $numSearched = 0

    # Search LDAP for each account and record the results.
    Write-Host "Parsing through inactive users for issues.." -noNewLine
    foreach ($user in $ShouldNotHaveLinuxandHome) {
        # Display our progress to the user.
        $percentComplete = ($numSearched / $totalUsers * 100)
        Write-Progress -Activity "Searching LDAP for Users" -Status "$numSearched / $totalUsers Complete:" -PercentComplete $percentComplete

        if ($user.SamAccountName | ldapSearch){
            $HasLinuxAccount += $user
        }
        if (Test-Path "\\storage.wpi.edu\homes\$($user.SamAccountName)") {
            $HasHomeDir += $user
        }

        $numSearched++
    }
    Write-Host -foregroundColor GREEN "Done."

    # Sort everything by username.
    $MissingLinuxAccount = ($MissingLinuxAccount | Sort-Object -Property samAccountName)
    $MissingHomeDir = ($MissingHomeDir | Sort-Object -Property samAccountName)
    $HasLinuxAccount = ($HasLinuxAccount | Sort-Object -Property samAccountName)
    $HasHomeDir = ($HasHomeDir | Sort-Object -Property samAccountName)

    # We're done, display our results for the viewer/log.
    Write-Host -foregroundColor CYAN "Missing Linux Accounts:"
    $MissingLinuxAccount | Select-Object samAccountName, extensionAttribute7, Enabled, whenCreated | Format-Table
    Write-Host "--------------------`n"

    Write-Host -foregroundColor CYAN "Missing Home Dir:"
    $MissingHomeDir | Select-Object samAccountName, extensionAttribute7, Enabled, whenCreated | Format-Table
    Write-Host "--------------------`n"

    Write-Host -foregroundColor CYAN "Has Linux Account and Shouldn't:"
    $HasLinuxAccount | Select-Object samAccountName, extensionAttribute7, Enabled, whenCreated | Format-Table
    Write-Host "--------------------`n"

    Write-Host -foregroundColor CYAN "Has Home Dir and Shouldn't:"
    $HasHomeDir | Select-Object samAccountName, extensionAttribute7, Enabled, whenCreated | Format-Table
    Write-Host "--------------------`n"

    # Display the total numbers.
    Write-Host "`n--------------------"
    Write-Host "Total Enabled Staff/Student/Affiliates Found: " -noNewLine
    Write-Host -foregroundColor CYAN $ShouldHaveLinuxandHome.count

    Write-Host "Total Disabled or Alumni Found: " -noNewLine
    Write-Host -foregroundColor CYAN $ShouldNotHaveLinuxandHome.count

    Write-Host "Total Enabled Without Linux Account in OU=People: " -noNewLine
    Write-Host -foregroundColor CYAN $MissingLinuxAccount.count

    Write-Host "Total Enabled Without Home Directory: " -noNewLine
    Write-Host -foregroundColor CYAN $MissingHomeDir.count

    Write-Host "Total Disabled or Alumni With Linux Account in OU=People: " -noNewLine
    Write-Host -foregroundColor CYAN $HasLinuxAccount.count

    Write-Host "Total Disabled or Alumni With Home Directory: " -noNewLine
    Write-Host -foregroundColor CYAN $HasHomeDir.count

    Write-Host "Prepping and sending emails if necessary."
    #-----------------------------------------------------------------------------------------------
    # Null out our message stuff before filling it.
    $messageParameters=$null
    $Subject=$null
    $Body=$null

    if ($null -ne $MissingLinuxAccount) {
        # Add a new paragraph to the email.
        $Body += "<p>The following accounts are <b>Enabled</b> and have a <i>Primary Affiliation</i> of <b>Staff, Student, or Affiliate</b> and an <i>ExtensionAttribute1</i> not equal to <b>RE</b> in AD but <u>do not</u> have Linux Accounts.</p>"
        # Start a new table.
        $Body += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>Enabled</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Primary Affiliation</td><td align='center'>Created</td></tr>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $MissingLinuxAccount) {
            $Body += "    <tr><td>$($user.Name)</td><td>$($user.SamAccountName)</td><td>$($user.Enabled)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td align='center'>$($user.extensionAttribute7)</td><td align='center'>$($user.whenCreated)</td></tr>"
        }

        # Close our table.
        $Body += "</table>"
        $Subject = "WPI Accounts - Missing Linux Accounts or Home Directories"
    }

    # Do the same as above but for the MissingHomeDir
    if ($null -ne $MissingHomeDir) {
        # Add a new paragraph to the email.
        $Body += "<p>The following accounts are <b>Enabled</b> and have a <i>Primary Affiliation</i> of <b>Staff, Student, or Affiliate</b> and an <i>ExtensionAttribute1</i> not equal to <b>RE</b> in AD but <u>do not</u> have Home Directories.</p>"
        # Start a new table.
        $Body += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>Enabled</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Primary Affiliation</td><td align='center'>Created</td></tr>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $MissingHomeDir) {
            $Body += "    <tr><td>$($user.Name)</td><td>$($user.SamAccountName)</td><td>$($user.Enabled)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td align='center'>$($user.extensionAttribute7)</td><td align='center'>$($user.whenCreated)</td></tr>"
        }

        # Close our table.
        $Body += "</table>"
        $Subject = "WPI Accounts - Missing Linux Accounts or Home Directories"
    }

    # If we have a message to send, send it!
    if ($null -ne $Subject -and $null -ne $Body){
        Write-Host "Sending email about missing Linux and/or Home Directories."
        $messageParameters = @{
            Subject = $Subject
            Body = $Body
            From = $HeaderFrom
            To = $Recipients
            SmtpServer = $SMTPServer
            Priority = "High"
        }

        Send-MailMessage @messageParameters -BodyAsHtml
    }

    # After we send the message, null everything out again.
    $messageParameters=$null
    $Subject=$null
    $Body=$null

    if ($null -ne $HasLinuxAccount) {
        # Add a new paragraph to the email.
        $Body = "<p>The following list of users have Linux accounts in <b>OU=People</b> despite being Disblaed in AD <b>or</b> being <i>Alumni</i>.<p>"
        # Start a new table.
        $Body += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>Enabled</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Primary Affiliation</td><td align='center'>Created</td></tr>"

        # Iterate through the list of HasLinuxAccount and add them all to a table.
        foreach ($user in $HasLinuxAccount) {
            $Body += "    <tr><td>$($user.Name)</td><td>$($user.SamAccountName)</td><td>$($user.Enabled)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td align='center'>$($user.extensionAttribute7)</td><td align='center'>$($user.whenCreated)</td></tr>"
        }

        # Close our table.
        $Body += "</table>"
        $Subject = "WPI Accounts - Superfluous Linux Accounts and Home Directories"

    }

    # Do the same as above but for HasHomeDir
    if ($null -ne $HasHomeDir) {

        # Add a new paragraph to the email.
        $Body += "<p>The following list of users have Home Directories despite being Disabled in AD.</p>"
        # Start a new table.
        $Body += "<table width='100%' border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>Name</td><td align='center'>Username</td><td align='center'>Enabled</td><td align='center'>ID Number</td><td align='center'>Department</td><td align='center'>Primary Affiliation</td><td align='center'>Created</td></tr>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $HasHomeDir) {
            $Body += "    <tr><td>$($user.Name)</td><td>$($user.SamAccountName)</td><td>$($user.Enabled)</td><td>$($user.EmployeeID)</td><td>$($user.Department)</td><td align='center'>$($user.extensionAttribute7)</td><td align='center'>$($user.whenCreated)</td></tr>"
        }

        # Close our table.
        $Body += "</table>"
        $Subject = "WPI Accounts - Superfluous Linux Accounts and Home Directories"

    }

    # If we have a message to send, send it!
    if ($null -ne $Subject -and $null -ne $Body){
        Write-Host "Sending email about superfluous Linux and/or Home Directories."
        $messageParameters = @{
            Subject = $Subject
            Body = $Body
            From = $HeaderFrom
            To = $Recipients
            SmtpServer = $SMTPServer
            Priority = "High"
        }

        Send-MailMessage @messageParameters -BodyAsHtml
    }
}

Stop-Transcript
