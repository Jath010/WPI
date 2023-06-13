############################################
#   Global Address List visibility maintenance
#   Joshua Gorham 6/14/2022
#   This script searches for all users who are currently hidden from the GAL who should not be and makes them visible
############################################

# Set path for log files:
$logPath = "D:\wpi\Logs\GALmaint"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_GALmaint.log" -Force

$hiddenUsers = get-aduser -Filter {msExchHideFromAddressLists -eq $true -and Enabled -eq $true -and (extensionattribute7 -eq "Student" -or extensionattribute7 -eq "Staff" -or extensionattribute7 -eq "Alum" -or extensionattribute7 -eq "Affiliate" -or extensionattribute7 -eq "Faculty")} -Properties msExchHideFromAddressLists, extensionattribute7

foreach ($user in $hiddenUsers) {
    Write-Host "Unhiding $($user.SamAccountName) from the GAL"
    set-ADobject $user.distinguishedname -replace @{msExchHideFromAddressLists = $false }
}

Stop-Transcript