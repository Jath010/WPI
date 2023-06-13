Clear-Host
<###################################################################################################################
Script        : SIS_Import.ps1
Purpose       : Import data files from Banner to Canvas using the SIS integration API
Target System : PROD Instance (wpi.instructure.com)
--------------
Last Update   : 09/12/19
Update by     : cidorr
###################################################################################################################>
## GitHub samples: https://github.com/kajigga/canvas-contrib/tree/master/SIS_Integration
## Canvas CSV Field formatting: https://community.canvaslms.com/docs/DOC-3098
## Declare Functions
function Send-ErrorNotification ($subject,$body,$HeaderTo) {
    $HeaderFrom = 'lms-alert@wpi.edu'
    $SmtpServer = 'smtp.wpi.edu'

    $messageParameters = @{
        Subject = "[Canvas Alert] $subject"
        Body = $body
        From = $HeaderFrom
        To = $HeaderTo
        SmtpServer = $SMTPServer
        }
    Send-MailMessage @messageParameters -BodyAsHtml -Priority High
    }

function SIS-Process-Integration ($CanvasInstance,$APIToken,$DataFilePath,$LogFile,$BatchMode,$BatchModeTerm) {
    $POSTresults=$null;$POSTresults1=$null;$GETresults=$null;$GETresults1=$null

    $ContentType = "text/plain"
    $account_id = "1" #root account ID of Canvas, usually the number 1
    $headers = @{"Authorization"="Bearer "+$APIToken}
    
    if ($BatchMode) {$import_type = "instructure_csv&batch_mode=1&batch_mode_term_id=sis_term_id:"+$BatchModeTerm}
    else{$import_type = "instructure_csv"}
    
    $url = "https://$CanvasInstance/api/v1/accounts/"+$account_id+"/sis_imports.json?import_type="+$import_type

    if (Test-Path $DataFilePath) {
        $POSTResults1 = (Invoke-WebRequest -Method POST -ContentType $ContentType -Headers $headers -InFile $DataFilePath -Uri $url -Verbose)
        if (!$POSTresults1) {
            Write-Host ""
            Write-Host "Upload Failed:" -ForegroundColor Yellow
            Write-Host "--------------" -ForegroundColor Yellow
            Write-Host "Attempted command: Invoke-WebRequest -Method POST -ContentType $ContentType -Headers $headers -InFile $DataFilePath -Uri $url -Verbose" -ForegroundColor Yellow
            break
            }
        $POSTResults1.Content | Out-File -Append $LogFile
        $POSTResults = ($POSTResults1.Content | ConvertFrom-Json)
        $POSTResults | Out-File $LogFile -Append
        "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] ID:$($POSTResults.id) - Uploaded" | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Yellow

        $status_url = "https://$CanvasInstance/api/v1/accounts/"+$account_id+"/sis_imports/"+$POSTResults.id
        $GETResults1 = (Invoke-WebRequest -Headers $headers -Method GET -Uri $status_url)
        $GETResults1.Content | Out-File -Append $LogFile
        <#
        do {
            $GETResults1 = (Invoke-WebRequest -Headers $headers -Method GET -Uri $status_url)
            $GETResults1.Content | Out-File -Append $LogFile
            $GETResults = ($GETResults1.Content | ConvertFrom-Json)
            "ID:$($POSTResults.id) - Progress: $($GETResults.progress) - Workflow State: $($GETResults.workflow_state)" | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Yellow
            ""  | Out-File -Append -FilePath $LogFile
            Start-Sleep -s 5
            if($GETResults -eq $null) {break}
            }
        while($GETResults.progress -lt 100 -and $GETResults.workflow_state -ne "failed_with_messages")
        $GETResults1.Content | Out-File $LogFile -Append
        "ID:$($POSTResults.id) - Upload Processing Completed" | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Green
        ""  | Tee-Object -Append -FilePath $LogFile | Write-Host
        #>
        if ($GetResults.processing_warnings) {foreach ($item in $GETResults.processing_warnings) {$item | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Yellow}}
        if ($GetResults.processing_errors) {foreach ($item in $GETResults.processing_errors) {$item | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Red}}
        ""  | Tee-Object -Append -FilePath $LogFile | Write-Host
        }
    }

Function SIS-Archive-Files ($DataPath,$ArchiveRoot,$Date) {
    ## Verify Archive Directory Path
    $ArchivePath = "$ArchiveRoot\$($Date.Year)\$($Date.ToString("MM"))\$($Date.ToString("dd"))"
    if (!(Test-Path $ArchivePath)) {[VOID](New-Item $ArchivePath -Type Directory)}

    $datestamp    = get-date ($Date) -Format yyyy_MM_dd_HHmmss
    $Archive_Zip   = "$ArchivePath\$($datestamp)_SIS_Data.zip"
    
    ## Zip and Copy all files to archive
    [io.compression.zipfile]::CreateFromDirectory($DataPath, $Archive_Zip) 
    
    ## Remove files
    $files = Get-ChildItem $DataPath -File
    foreach ($file in $files) {Remove-Item -Path $file.FullName -Confirm:$false}
    }

## Declare Variables
$now=$null;$datestamp=$null;$Snapshottime=$null
$CanvasInstance=$Null;$APIToken=$null;
$ExtractPath=$null;$ScriptPath=$null;$DataPath=$null;$ArchiveRoot=$null;$extractLogFile=$null
$LogFile=$null;$ErrorLogFile=$null

Add-Type -assembly "system.io.compression.filesystem"

## Set Date/Time Formats
$now = Get-Date
if ($now.Hour -gt 1 -and $now.Hour -lt 6) {exit}
$SnapshotTime = get-date ("$($now.Month)/$($now.Day)/$($now.Year) $($now.Hour):00")
$datestamp    = get-date ($now) -Format yyyy_MM_dd_HHmmss

## Set Target Instance - Note this must be a name that is on the server certificate, otherwise the REST upload will fail.
$CanvasInstance = 'wpi.instructure.com'         #PROD Instance
#$CanvasInstance = 'wpi.test.instructure.com'    #TEST Instance
#$CanvasInstance = 'wpi.beta.instructure.com'    #BETA Instance

## Specify API Token
$APIToken = '7782~EGa3UwxhIlzPY9GODZNT1U6IIqp25OV6WEMa0LLSK3OsDQquz8XnTxLSzTdBb6e0' #Token is the same on all instances

## Specify data source path
$ExtractPath = '\\bannerprod.wpi.edu\canvas$'     ## Extract Path on Banner PROD
#$ExtractPath = '\\bannertest-02.wpi.edu\canvas$'  ## Extract Path on Banner PPRD

## Set Paths and other system variables
$ScriptPath = 'd:\wpi\batch\LMS\Canvas\SIS_Import_PROD' #PROD Instance
#$ScriptPath = 'd:\wpi\batch\LMS\Canvas\SIS_Import_TEST' #TEST Instance
#$ScriptPath = 'd:\wpi\batch\LMS\Canvas\SIS_Import_BETA' #BETA Instance

$DataPath  = "$ScriptPath\Data"
$ArchiveRoot = "$ScriptPath\Archive"
$extractLogFile = "$ExtractPath\canvas_extract.log"

$LogFile        = "$DataPath\$($datestamp)_Integration_Processing.log"
$ErrorLogFile   = "$DataPath\$($datestamp)_Integration_Error.log"

#<#
## Check status of existing imports
$account_id=$null;$headers=$null;$GETResultsTotal=$null;$GETResultsTotal1=$null;$GETResultsFiltered=$null
$BadResults=$null;$body=$null
$account_id = "1" #root account ID of Canvas, usually the number 1
$headers = @{"Authorization"="Bearer "+$APIToken}
$GETResultsTotal1 = (Invoke-WebRequest -Headers $headers -Method GET -Uri "https://$CanvasInstance/api/v1/accounts/$account_id/sis_imports/?per_page=100&page=1")
$GETResultsTotal = ($GETResultsTotal1.Content | ConvertFrom-Json).sis_imports
$GETResultsFiltered = $GETResultsTotal | Select -First 20
$BadResults =  $GETResultsFiltered | Where {$_.progress -lt 100 -and $_.workflow_state -notin ('aborted','failed','failed_with_messages','initializing')}
$ErrorResults = $GETResultsTotal | Where {$_.progress -lt 100 -and $_.workflow_state -notin ('created','importing','imported','initializing') -and $_.created_at -match (get-date -Format "yyyy-MM-dd")}
if ($ErrorResults) {
    $errors2table = $null
    $errors2table += "<table border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>ID</td><td align='center'>Progress</td><td align='center'>Workflow State</td><td align='center'>Created At (ET)</td><td align='center'>Started At (ET)</td><td align='center'>Ended At (ET)</td><td align='center'>errors_attachment</td><td align='center'>csv_attachments</td></tr>"
    
    foreach ($item in $ErrorResults) {
        $Started=$null;$Created=$null;$Ended=$null;$Duration=$null

        if ($item.created_at -ne $null) {$Created = Get-Date($item.created_at) -Format "hh:mm tt MM/dd/yyyy"}
        if ($item.ended_at -ne $null) {$Started = Get-Date($item.started_at) -Format "hh:mm tt MM/dd/yyyy"}
        if ($item.ended_at -ne $null) {$Ended = Get-Date($item.ended_at) -Format "hh:mm tt MM/dd/yyyy"}
        $errors2table += "    <tr><td>$($item.id)</td><td align='center'>$($item.progress)</td><td>$($item.workflow_state)</td><td>$Created</td><td>$Started</td><td>$Ended</td><td>$($item.errors_attachment)</td><td>$($item.csv_attachments)</td></tr>"
        }
    $errors2table += "</table>"

    #Build body text
    $errorbody    = "These are all of the errors from today's imports."
    $errorbody    = $errorbody + "<br><br>$errors2table"
    }

if (($BadResults | Measure-Object).Count -ge 3) {

    "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] ERROR: Existing imports are still processing.  Skipping scheduled import." | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Yellow

    $body2table = $null
    $body2table += "<table border='1'><tbody><tr bgcolor=#CCCCCC><td align='center'>ID</td><td align='center'>Progress</td><td align='center'>Workflow State</td><td align='center'>Created At (ET)</td><td align='center'>Started At (ET)</td><td align='center'>Ended At (ET)</td><td align='center'>Duration (minutes)</td></tr>"
    
    $MostRecentUpload = Get-Date($GETResultsTotal[0].created_at)
    foreach ($item in $GETResultsFiltered) {
        $Started=$null;$Created=$null;$Ended=$null;$Duration=$null

        if ((Get-Date($item.created_at)).Hour -ne $MostRecentUpload.Hour) {continue}

        if ($item.created_at -ne $null) {$Created = Get-Date($item.created_at) -Format "hh:mm tt MM/dd/yyyy"}
        if ($item.ended_at -ne $null) {$Started = Get-Date($item.started_at) -Format "hh:mm tt MM/dd/yyyy"}
        if ($item.ended_at -ne $null) {$Ended = Get-Date($item.ended_at) -Format "hh:mm tt MM/dd/yyyy"}
        if ($Created -and $Ended) {$Duration = (Get-Date($Ended)) - (Get-Date($Started))}
        $body2table += "    <tr><td>$($item.id)</td><td align='center'>$($item.progress)</td><td>$($item.workflow_state)</td><td>$Created</td><td>$Started</td><td>$Ended</td><td align='center'>$($Duration.TotalMinutes)</td></tr>"
        }
    $body2table += "</table>"

    ## Send Error Notifiction if failure occurs.
    $HeaderTo   = "sysops@wpi.edu","kwrigley@wpi.edu","lftapper@wpi.edu"
    $subject = "Canvas has not completed processing previous import jobs $(get-date $now -format g)"
    $body    = 'The SIS Import script for Canvas has terminated. <br><br> Canvas was still processing import jobs.  This script has terminated and will attempt an import again at the next hour.'
    $body    = $body + "<br><br>$body2table"
    $body    = $body + "<br><br>$errorbody"
    Send-ErrorNotification $subject $body $HeaderTo
    
    break
    }

if ($ErrorResults) {
    $HeaderTo   = "sysops@wpi.edu","kwrigley@wpi.edu","lftapper@wpi.edu"
    $subject = "Canvas non-fatal errors from prior to $(get-date $now -format g)"
    Send-ErrorNotification $subject $errorbody $HeaderTo
    }





## Copy files from Banner Production
$file=$null;$files=$null;$count=$null

if (!(Test-Path $ExtractPath)) {
    "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] Extract Data Path unavailable" | Tee-Object -Append -FilePath $ErrorLogFile | Write-Host -ForegroundColor Red
    SIS-Archive-Files $DataPath $ArchiveRoot $now
    exit
    }

$files = Get-ChildItem $ExtractPath -File | Sort LastWriteTime -Descending

while (!(Select-String -Path $extractLogFile -Pattern 'FileFooter') -or !((Get-Item $extractLogFile).LastWriteTime -gt $SnapshotTime)) {
    $count++
    Write-Host "Sleeping ($count)...";sleep 40
    "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] WARNING: Script is waiting for extract generation to complete ." | Tee-Object -Append -FilePath $ErrorLogFile | Write-Host -ForegroundColor Yellow

    if ($count -ge 60) {
        $now | Out-File $ErrorLogFile -Append
        "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] ERROR: There was a problem with the extract process and file copying was aborted." | Tee-Object -Append -FilePath $ErrorLogFile | Write-Host -ForegroundColor Red
        $count=$null

        ## Send Error Notifiction if failure occurs.
        $HeaderTo   = "sysops@wpi.edu","kwrigley@wpi.edu","lftapper@wpi.edu","dwgalvin@wpi.edu","roger@wpi.edu","jplunkett@wpi.edu"
        $subject = "Feed files failed to copy $(get-date $now -format g)"
        $body    = 'The SIS Import script for Canvas has terminated. <br><br> The feed files available were either out of date or had not finished being created.  Please check the feed files located at \\bannerprod\canvas$'
        Send-ErrorNotification $subject $body $HeaderTo

        exit
        }
    }

Write-Host "While check passed - Moving to copy files." -ForegroundColor Green
foreach ($file in $files){Copy-Item -Path $file.VersionInfo.FileName -Destination $DataPath}

## Process files and import to Canvas
$DataFiles=$null;$UserFilePath=$null;$CourseFiles=$null;$EnrollmentFiles=$null
$DataFiles = Get-ChildItem $DataPath
$UserFilePath = "$DataPath\users.csv"
$CourseFiles = $DataFiles | Where {$_.Name -match 'courses'}
$EnrollmentFiles = $DataFiles | Where {$_.Name -match 'enrollment'}

"[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] FILE PROCESSING: users.csv"  | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Green
SIS-Process-Integration $CanvasInstance $APIToken $UserFilePath $LogFile $false

if ($CourseFiles) {
    foreach ($file in $CourseFiles) {
        $TermCode = $file.Name.Split('.')[0].Split('-')[1]
        ""  | Tee-Object -Append -FilePath $LogFile | Write-Host
        "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] FILE PROCESSING: $($file.Name)"  | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Green
        SIS-Process-Integration $CanvasInstance $APIToken $file.FullName $LogFile $false
        }
    }

if ($EnrollmentFiles) {
    $SemesterFiles = $EnrollmentFiles | Where {$_.Name -like '*-F*.csv' -or $_.Name -like '*-S*.csv'}
    $TermFiles = $EnrollmentFiles | Where {$_.Name -like '*-A*.csv' -or $_.Name -like '*-B*.csv' -or $_.Name -like '*-C*.csv' -or $_.Name -like '*-D*.csv' -or $_.Name -like '*-E*.csv'}

    foreach ($file in $SemesterFiles) {
        $TermCode = $file.Name.Split('.')[0].Split('-')[1]
        ""  | Tee-Object -Append -FilePath $LogFile | Write-Host
        "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] FILE PROCESSING: $($file.Name)"  | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Green
        SIS-Process-Integration $CanvasInstance $APIToken $file.FullName $LogFile $true $TermCode
        }

    foreach ($file in $TermFiles) {
        $TermCode = $file.Name.Split('.')[0].Split('-')[1]
        ""  | Tee-Object -Append -FilePath $LogFile | Write-Host
        "[$(get-date -Format "yyyy/MM/dd HH:mm:ss")] FILE PROCESSING: $($file.Name)"  | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Green
        SIS-Process-Integration $CanvasInstance $APIToken $file.FullName $LogFile $true $TermCode
        }
    }

## Check status of existing imports
$GETResultsTotal1 = (Invoke-WebRequest -Headers $headers -Method GET -Uri "https://$CanvasInstance/api/v1/accounts/$account_id/sis_imports/")
$GETResultsTotal = ($GETResultsTotal1.Content | ConvertFrom-Json).sis_imports 
$GETResultsTotal | Where {$_.progress -lt 100} | Select ID,Workflow_state,Progress,Data  | Tee-Object -Append -FilePath $LogFile | Write-Host -ForegroundColor Yellow

## Archve Files
SIS-Archive-Files $DataPath $ArchiveRoot $now
