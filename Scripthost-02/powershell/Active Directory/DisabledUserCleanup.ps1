#$var = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate
#$noNumbers = ($var| where {$_.name -notmatch ".*(\(\d.*\))"})                           # This finds accounts that lack numbers in their name to find non-standard accounts in disabled
#$noNumbers = ($var| where {$_.name -notmatch ".*\(\d{9}\)"})                           # This finds accounts that lack 9 digits between parens in their name to find non-standard accounts in disabled
#$olderThanThree = ($var| where {$_.lastlogondate -lt (get-date).addyears(-3)}).count    # this finds accounts that haven't been logged into in 3 years
# specific search ($_.name -match ".*\(\d{9}\)" -or $_.name -match ".*\(\d{6}\)" -or $_.name -match ".*\(\d{7}\)" -or $_.name -match ".*\(\d{5}\)"")

# Set path for log files:
$logPath = "D:\wpi\Logs\DisabledUserCleanup"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force -Recurse
}

function Move-OldDisabled {
    [CmdletBinding()]
    param (
        $NumberOfYearsBack = 1
    )
    $TargetOU = "RemovalStaging"
    Write-Host "Gathering Staging List"
    $oldUsers = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate | where-object {$_.name -match ".*\(\d{5,9}\)" -and $_.lastlogondate -lt (get-date).addyears("-$($NumberOfYearsBack)")}
    Write-Host "Staging List Gathered"
    foreach($user in $oldUsers){
        Write-Host "Moving $($user.samaccountname) to $TargetOU OU"
        Move-ADObject -TargetPath "OU=$($TargetOU),OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Identity $user.DistinguishedName
    }
}

function Move-AbnormalDisabled {
    [CmdletBinding()]
    param (
        
    )
    $TargetOU = "AbnormalDisabled"
    Write-Host "Gathering Abnormal List"
    $abnormalUsers = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate | where-object {$_.name -notmatch ".*\(\d{5,9}\)"}
    Write-Host "Abnormal List Gathered"
    foreach($user in $abnormalUsers){
        Write-Host "Moving $($user.samaccountname) to $targetOU OU"
        Move-ADObject -TargetPath "OU=$($TargetOU),OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Identity $user.DistinguishedName
    }
}

function Remove-OldDisabled {
    [CmdletBinding()]
    param (
        $NumberOfYearsBack = 3
    )
    $TargetOU = "RemovalStaging"
    Write-Host "Gathering Deletion List"
    $deletionUsers = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=$($TargetOU),OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate | where-object { $_.lastlogondate -lt (get-date).addyears("-$($NumberOfYearsBack)")}
    Write-Host "Deletion List Gathered"
    foreach($user in $deletionUsers){
        Write-Host "Removing $($user.samaccountname)"
        Move-ADObject -TargetPath "OU=$($TargetOU),OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Identity $user.DistinguishedName
    }
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_DisabledUserCleanup.log" -Force

Move-OldDisabled
Move-AbnormalDisabled

Stop-Transcript