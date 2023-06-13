Import-Module "D:\wpi\powershell\CourseShareGeneration\Get-CourseList.ps1"
# -NoProfile -WindowStyle Hidden -File "D:\wpi\powershell\CourseShareGeneration\CourseShare_Syncer.ps1" -autosync

# Set path for log files:
$logPath = "D:\wpi\Logs\CourseShare"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_CourseShare.log" -Force

$courseList = Get-TermCourses

Set-CourseGroups $courseList

Stop-Transcript