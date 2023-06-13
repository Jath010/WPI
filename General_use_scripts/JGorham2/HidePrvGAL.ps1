<#
    Remove Prv Accounts from the GAL
#>

$logPath = "D:\wpi\Logs\HidePrvGAL"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_HidePrvGAL.log" -Force

$unhiddenPRV = get-aduser -Filter {Enabled -eq $true} -Properties msExchHideFromAddressLists -SearchBase "OU=Privileged,OU=Accounts,DC=admin,DC=wpi,DC=edu" | Where-Object {$_.msExchHideFromAddressLists -ne $true}

foreach ($user in $unhiddenPRV) {
    Write-Host "Hiding $($user.SamAccountName) from the GAL"
    set-ADobject $user.distinguishedname -replace @{msExchHideFromAddressLists = $True }
}

Stop-Transcript