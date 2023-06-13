Clear-Host
<#
Get_Software_Folder_Changes.ps1

Author : tcollins
Updated: 2018-01-10
Purpose: This script looks at the two software publishing locations and scans the folders for a list
     of changed folders.  This will notify the Service Desk to make changes to the various KB articles.

-------------------------------------#>

## Declare Array to contain list of changed folders
$SoftwareChanges=@()

## Set email information.  
$HeaderFrom = "its@wpi.edu"
$HeaderTo = "tcollins@wpi.edu" #For multiple values in the "To" field, seperate the values with a comma
$SMTPServer = "smtp.wpi.edu"

## Create list of paths to check
$SoftwarePath = '\\storage.wpi.edu\software','\\storage.wpi.edu\software_restricted'

## Set Target Date to check against.  Default is 30 days
$TargetDate = (get-date).AddDays(-180)

Clear-Host
## Check folders to a maximum depth of 2 levels
$BodyShareList = "<p>The list of folder paths being checked are:</p><ul>"
foreach ($path in $SoftwarePath) {
    $SoftwareChanges += Get-ChildItem $path -Directory -Recurse -Depth 2 | Where {$_.LastWriteTime -gt $TargetDate}  | Select FullName,LastWriteTime
    $BodyShareList += "<li>$path</li>"
    }
$BodyShareList += "</ul>" 

if ($SoftwareChanges) {$BodyChanges = $SoftwareChanges | ConvertTo-Html -As Table -Fragment | Out-String}
else {$BodyChanges = "<p>There have been no changes since $TargetDate</p>"}

$BodyHeader = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><head><title>Folder Changes</title></head><body>'
$BodyHeader += "<style>"
$BodyHeader += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$BodyHeader += "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$BodyHeader += "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$BodyHeader += "</style>"

$BodyFooter = '</body></html>'


$Body = $BodyHeader + $BodyShareList + $BodyChanges + $BodyFooter

$messageParameters = @{
    Subject    = "Changes to Software Installation Shares"
    Body       = $Body
    From       = $HeaderFrom
    To         = $HeaderTo
    SmtpServer = $SMTPServer
    }
Send-MailMessage @messageParameters -BodyAsHtml
Write-Host "Mail Sent"
