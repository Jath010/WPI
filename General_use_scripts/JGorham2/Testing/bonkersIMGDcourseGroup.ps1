# Set path for log files:
$logPath = "D:\wpi\Logs\IMGDcourseGroup"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_IMGDcourseGroup.log" -Force

#Creds for automagic login on ScriptHost-02
$Credentials = $null
if ($env:COMPUTERNAME -eq "SCRIPTHOST-02") {
    $Credentials = Import-Clixml -Path 'D:\wpi\XML\exch_automation@wpi.edu.xml'
    Connect-ExchangeOnline -Credential $Credentials
}

$courselist = Get-ADGroupMember Imgd_courses | Select-Object name
$sectionList = @()
foreach ($course in $courselist) {
    $sectionlist += Get-ADGroupMember $course.name | Where-Object { $_.Name -notmatch ".*_Faculty" } | Select-Object name
}
$sectionList = $sectionlist | sort-object -Property name -Unique
$studentList = @()
foreach ($section in $sectionList) {
    $studentList += Get-ADGroupMember $section.name | select-object samaccountname
}
$studentList = $studentList | Sort-Object -Property samaccountname -Unique

$comparisons = $null
$AddMembers = $null
$RemoveMembers = $null

#Section Group Sync
$currentMembers = Get-DistributionGroupMember -Identity "dl-imgd-courses@wpi.edu" | Select-Object @{Name="samaccountname";Expression={$_.WindowsLiveID.split("@")[0]}} | Sort-Object -Property samaccountname -Unique
if ($null -eq $currentMembers) {
    Write-Host "Group is empty. Initializing."
    $AddMembers = $studentList
}
else {
    Write-Host "Running Comparison"
    $comparisons = Compare-Object $currentMembers $studentList -Property samaccountname
    Write-Host "$($comparisons.count) differences."
    $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } #| Select-Object samaccountname
    Write-Host "$($AddMembers.count) Additions"
    $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } #| Select-Object samaccountname
    Write-Host "$($RemoveMembers.count) Removals"
}
foreach ($removal in $RemoveMembers.samaccountname) {
    Write-Host "Removing $removal"
    Remove-DistributionGroupMember -Identity "dl-imgd-courses@wpi.edu" -Member "$removal@wpi.edu" -Confirm:$false
}
foreach ($addition in $AddMembers.samaccountname) {
    Write-Host "Adding $addition"
    Add-DistributionGroupMember -Identity "dl-imgd-courses@wpi.edu" -Member "$addition@wpi.edu" -Confirm:$false
}

Stop-Transcript