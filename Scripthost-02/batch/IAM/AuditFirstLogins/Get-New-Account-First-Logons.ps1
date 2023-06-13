<#
.SYNOPSIS
   Get a list of all students who have not begun their account claiming process (have not opened their school email yet).

.DESCRIPTION
   This script first gatheres a list of student accounts that have been created within a certain number of days. 
   It then takes that list and searches each of their mailboxes to see if they've been logged into. 
   If no logon timestamp exists (LastUserActionTime == null), we record who it is.
   A full list of user IDs who have not taken action is saved to the Informatica drive for Adam Epstein,
   while a full list of user information is sent to Sarah Miles.

.NOTES
    Created By: Stephen Gemme
    Created On: 04/30/2020

    Modified By: Stephen Gemme
    Modified On: 05/04/2020
    Modifications: Refined/corrected search and began setup for saving information where people requested.

    Modofied By: Stephen Gemme
    Modified On: 05/06/2020
    Modifications: Started using "LastUserActionTime" as the determining factor as that is behaving much more appropriately for our needs. 

    Modofied By: Stephen Gemme
    Modified On: 05/18/2020
    Modifications: Swapped the logic so that it outputs those who HAVE accessed, rather than have not, to Adam E.

#>

Param (
    [Parameter(Mandatory = $false)] 
    [Switch] 
    $testMode,

    [Parameter(Mandatory = $false)]
    [Switch]
    $manualMode,

    [Parameter(Mandatory = $false)] 
    [String] 
    $userName
)

if ($null -ne $userName -and $userName -ne ""){
   Write-Host "Only checking single user: $userName"
   $singleSearch = $true
   # If We're checking a single person, we must be doing it manually.
   $manualMode = $true
}
else {
   $singleSearch = $false
}

# Setup our variables.
$credential = $null
# Set our credentials if possible.
$credPath = "D:\wpi\XML\$($env:UserName)@wpi.edu.xml"

Write-Host "Using Creds from: $credPath" 

$dateString = Get-Date -format "MM-dd-yyyy_HHmm"
# We do this so we can test the path to make sure it's there while keeping all the variables up here in the script.
$logPath = "D:\wpi\Logs\IAM\AuditfirstLogons"
$logSavePath = "$logPath\First-Logon-Audit-$dateString.log"
$informaticaPath = "\\informatica.wpi.edu\UES_Flatfile\DataFiles"
$informaticaSavePath = "$informaticaPath\NewWPIAccounts.csv"

# How many days back will we search for new accounts.
$date = (Get-Date).addDays(-30)

if ($singleSearch ){
   $Global:studentSearch = @{
      Filter      = { SamAccountName -eq "$userName" }
      Property    = "whenCreated"
   }
}
elseif ($testMode){
   $Global:studentSearch = @{
      Filter      = { SamAccountName -eq "sgemme_test" -or SamAccountName -eq "sgemme_test1" -or SamAccountName -eq "sgemme_test2" }
      Property    = "whenCreated"
   }
}
else {
   $Global:studentSearch = @{
      Filter      = { whenCreated -gt $date }
      SearchBase  = "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
      Property    = "whenCreated"
   }
}

# Currently not sent anywhere.
[System.Collections.ArrayList]$Global:allAccessed = @()
# The ones that Sarah Miles cares about.
[System.Collections.ArrayList]$Global:allNotAccessed = @()
# Don't worry about accounts that are fairly new (within 72 hours).
[System.Collections.ArrayList]$Global:accountTooNew = @()
# This is for Adam Epstein
$Global:accessedIDsOnly = @()

# This is needed to check MFA status.
function loginToMsolService($credentials){
   Write-Host "Trying to login to MSol Service with credentials..." -noNewLine
   try {
      Connect-MSolService -credential $credentials
      Write-Host -foregroundColor GREEN "Done."
      performSearch
   }
   catch {
      Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate to MSol Service.`n"
      exit
   }
}

# The Import-PSSession method is being deprecated and needs to be looked into upgrading to the O-Auth2 method.
function loginToEchange($credentials){
   Write-Host "Trying to login to Exchange with credentials..."
   $ExchangeOnlineSession=$null

   try {
      # Load Exchange Online
      #$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
      Connect-ExchangeOnline -Credential $credentials
      #Import-PSSession $ExchangeOnlineSession
      Write-Host -foregroundColor GREEN "Done."
      
      loginToMsolService($credentials)
   }
   catch {
      Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate to Exchange or load modules.`n"
      Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
      exit
   }
}

function performSearch(){

   $students = Get-AdUser @Global:studentSearch

   Foreach ($student in $students){
      try {
         # First, check if they have MFA enabled, since that's a telltale sign that they've signed in. 
         $mfa = Get-MsolUser -userprincipalName $student.userPrincipalName | Select-Object -expandProperty StrongAuthenticationMethods
         $mfaStatus = ""

         # Try to parse out our student ID
         $idString = $student | Select-Object -expandProperty Name
         try {
            $idNumber = ($idString).split("(")[1].split(")")[0]
         }
         catch {
            $idNumber = "Unknown"
         }

         $entry = [PSCustomObject]@{
            ID                   = $idNumber
            First                = $student.GivenName
            Last                 = $student.Surname
            Username             = $student.SamAccountName
            whenCreated          = $student.whenCreated
            MFAStatus            = $mfaStatus
            LastUserActionTime	= $lastUserActionTime
         }

         # If they don't have MFA, it means they've either not logged in or chose to decline (for up to 14 days).
         if ($null -eq $mfa){
            $mfaStatus = "Not Configured"
            $entry.MFAStatus = $mfaStatus

            # We use LastUserActionTime because it's the most accurate. See: https://o365reports.com/2019/06/18/office-365-users-last-logon-time-incorrect/
            $lastUserActionTime = Get-MailboxStatistics -Identity $student.userPrincipalName | Select-Object -expandProperty lastUserActionTime

            # If they also have no action on their Exchange, they definitely haven't logged in to email.
            if ($null -eq $lastUserActionTime){
               $Global:allNotAccessed.Add($entry) | Out-Null
            }
            # Check if they were created more than 2 weeks ago (they have 14 days to defer MFA before it's forced)
            elseif ((Get-Date $student.whenCreated) -lt (Get-Date).AddDays(-14) ){
               # They definitely haven't interacted with their account yet. 
               $Global:allNotAccessed.Add($entry) | Out-Null

            }
            else {
               # If they've been created more than a few days ago, we check if they've been interacting without MFA.
               # Otherwise, we give them more time to login and whatnot, so we don't report them.
               if ((Get-Date $student.whenCreated) -lt (Get-Date).AddDays(-3)) {
                  # If they've interacted with their email within 48 hours, they've likely been suppressing MFA.
                  if ((Get-Date $lastUserActionTime) -lt (Get-Date).AddDays(-2)){
                     $Global:allAccessed.Add($entry) | Out-Null
                     $Global:accessedIDsOnly += [PSCustomObject]@{
                        ID                   = $idNumber
                        LastUserActionTime   = $LastUserActionTime
                     }
                  }
                  else {
                     $Global:allNotAccessed.Add($entry) | Out-Null
                  }
               }
               else {
                  $Global:accountTooNew.Add($entry) | Out-Null
               }
            }
         }
         else {
            # They have MFA, no action needed.
            $mfaStatus = "Enabled"
            $entry.MFAStatus = $mfaStatus

            $Global:accessedIDsOnly += [PSCustomObject]@{
               ID                   = $idNumber
               LastUserActionTime   = $LastUserActionTime
            }
            $Global:allAccessed.Add($entry) | Out-Null
         }
      }
      catch [Exception] {
         $_.Exception.Message

         Write-Host -foregroundColor RED "Error getting MSOL info for " -nonewLine
         Write-Host -foregroundColor CYAN $student.SamAccountName

         # If the inbox hasn't gotten any mail yet, it throws a warning (not error) that is impossible to catch without coming here. 
         # So, we make sure the mailbox exists at all.
         try {
            Write-Host -foregroundColor CYAN $student.SamAccountName
            $doesMailboxExist = Get-MailboxFolderStatistics -Identity $student.userPrincipalName -FolderScope "Inbox"
            if ($null -ne $doesMailboxExist){
               Get-Date $student.whenCreated
               # We know the mailbox exists, it's likely that they've never logged into it and haven't got any emails yet.
               if ((Get-Date $student.whenCreated) -lt (Get-Date).AddDays(-3)) {
                  $Global:allNotAccessed.Add($entry) | Out-Null
               }
               else {
                  $Global:accountTooNew.Add($entry) | Out-Null
               }
            }
         }
         catch [Exception]{
            $_.Message
            Write-Host "Error"
            # If we error out here, it means the full profile/inbox isn't even created yet, so obviously they won't have logged in yet.
            $Global:accountTooNew.Add($entry) | Out-Null
         }
      }
   }

   Write-Host -foregroundColor CYAN "`nStudents who have accessed: "
   $Global:allAccessed

   Write-Host -foregroundColor CYAN "`nStudents who have NOT accessed, but account is very new (72 hours): "
   $Global:accountTooNew

   Write-Host -foregroundColor CYAN "`nStudents who have NOT accessed: "
   $Global:allNotAccessed | Format-Table

   # Only save the data if we're not testing or doing a single case.
   if ((-NOT $testMode) -and (-NOT $singleSearch)){
      $Global:accessedIDsOnly | Export-CSV -Path $informaticaSavePath -NoTypeInformation
   }  
}

########################
#                      #
#     Begin Script     #
#                      #
########################

# Before we start, make sure everything is ready and exists.
# Make sure the above paths work.
if (-NOT (Test-Path $logPath)){
   New-Item -ItemType Directory -Force -Path $logPath

   Start-Transcript -Append -Path $logSavePath -Force
   Write-Host "Log Path did not exist, newly created at $logPath."
}
else {
   # Nothing broken, continue.
   Start-Transcript -Append -Path $logSavePath -Force
}

# If we don't have anywhere to put the data, just exit.
if (-NOT (Test-Path $informaticaPath)){
   Write-Host "Cannot find Informatica Save Path, please check and try again."
   Stop-Transcript
   exit
}

# If we are alreayd logged onto Exchange, no need to try again
try{ 
   # If we can get a random, known list, we can get them all.
   Write-Host "Testing connection to Exchange..." -nonewLine
   if(Get-DistributionGroup "dl-fuller"){
      # We're already in, begin everything.
      performSearch
   }
 }
 catch {
   # See if we have any credentials stored locally for O365
   if (Test-path $credPath){
      $credential = Import-CliXml -Path $credPath
      loginToEchange($credential)
   }
   # If we don't have any, prompt for some and then prompt if the user wants to save them.
   else {
      Write-Warning "Saved Exchange Credentials not Found. Please enter new credentials."
      $credential = Get-Credential

      $save = (Read-Host "Would you like to encrypt and save these credentials for future use? [Y|n]").toLower()

      if ($save -eq "y"){
         $credential | Export-CliXml -Path $credPath

         Write-Host "`nEncrypted credentials saved as: " -nonewLine 
         Write-Host -foregroundColor CYAN $credPath
      }
      else {
         Write-Host -foregroundColor YELLOW "`nCredentials not saved.`n"
      }

      # When we're ready to go, login to O365
      loginToEchange($credential)

   }
}

Stop-Transcript