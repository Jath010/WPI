Clear-Host
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


## Check status of existing imports
$account_id=$null;$headers=$null;$GETResultsTotal=$null;$GETResultsTotal1=$null
$account_id = "1" #root account ID of Canvas, usually the number 1
$headers = @{"Authorization"="Bearer "+$APIToken}

$GETResultsTotal1 = (Invoke-WebRequest -Headers $headers -Method GET `
-Uri "https://$CanvasInstance/api/v1/accounts/$account_id/sis_imports/?per_page=100&page=1")

$GETResultsTotal = ($GETResultsTotal1.Content | ConvertFrom-Json).sis_imports 

$BadResults =  $GETResultsTotal | Select -First 20 | Where {$_.progress -lt 100 -and $_.workflow_state -notin ('aborted','failed','failed_with_messages','initializing')}
if (($BadResults | Measure-Object).Count -ge 3) {
    Write-Host ''
    Write-Host "The Import Script would fail" -ForegroundColor Red
    Write-Host ''
    }
else {
    Write-Host ''
    Write-Host "The Import Script would successfully run" -ForegroundColor Green
    Write-Host ''
    }

$GETResultsTotal | Where {$_.progress -lt 100} | Select ID,Progress,Workflow_state | ft -autoSize
Write-Host ""
#$GETResultsTotal | Select ID,Progress,Workflow_state,created_at,Ended_at -First 50 | ft -autoSize
$GETResultsTotal | Select ID,Progress,Workflow_state,created_at,Ended_at -First 20 | ft -autoSize
Write-Host ""

$GETResultsTotal.count


<#

   id progress workflow_state
   -- -------- --------------
47502        0 created       
47501        0 created       
47500        0 created       
47499        0 created       
47498        0 created       
47497        0 created       
47496       16 importing     




   id progress workflow_state
   -- -------- --------------
47502        0 created       
47501        0 created       
47500        0 created       
47499        0 created       
47498        0 created       
47497        0 created       
47496       16 importing     
#>