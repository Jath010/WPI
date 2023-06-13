
# Set path for log files:
$logPath = "D:\wpi\Logs\DL-Student-Upkeep"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Clean out logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_DL-Upkeep.log" -Force


function get-DLStudentStatus {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $Student = Get-DistributionGroup dl-students
        $undergrad = Get-DistributionGroup dl-undergraduates
        $other = Get-DistributionGroup dl-other-students
        $part = Get-DistributionGroup dl-grads-parttime
        $full = Get-DistributionGroup dl-grads-fulltime
        $lists = $Student, $undergrad, $other, $part, $full
    }
    
    process {
        foreach ($List in $lists) {
            Write-host $list.Identity
            if ($list.ModerationEnabled) {
                Write-Host "Moderation Enabled" -BackgroundColor Green -ForegroundColor Black
            }
            else {
                Write-Host "Moderation Disabled" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -ModerationEnabled:$true
                Get-DLModerationOffAuditLog -DL $list.Name
            }
            if ("$($list.Name)@wpi.edu" -eq $List.PrimarySmtpAddress) {
                Write-Host "Primary SMTP Correct" -BackgroundColor Green -ForegroundColor Black
            }
            else {
                Write-Host "Primary SMTP Wrong: $($list.PrimarySmtpAddress)" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -PrimarySmtpAddress "$($list.Name)@wpi.edu"
                Get-DLWrongAddressAuditLog -DL $list.Name
            }
            if ($list.BypassNestedModerationEnabled) {
                Write-Host "Nested Moderation Enabled" -BackgroundColor Green -ForegroundColor Black
            }
            else {
                Write-Host "Nested Moderation Disabled" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -BypassNestedModerationEnabled:$true
            }
        }
    }
    
    end {
        
    }
}

function Get-DLWrongAddressAuditLog {
    [CmdletBinding()]
    param (
        $DL
    )

    begin { # Get the most recent log where the primary smtp wasn't what it was supposed to be
        $regex = "SMTP:" + $DL + "@wpi\.edu.*"
        $log = (Search-AdminAuditLog -Cmdlets set-distributiongroup -EndDate (get-date) -StartDate (get-date).AddDays(-7) -IsSuccess $true | Where-Object { $_.ObjectModified -eq $DL -and $_.cmdletparameters.emailaddresses -notmatch $regex -and $_.cmdletparameters.name -contains "emailaddresses" })[0]
    }
    
    process {
        
        Write-Host "Rundate: " $log.rundate -BackgroundColor Green -ForegroundColor Black
        write-host "Caller: " $log.Caller
        Write-Host "Cmdlet Parameters:"
        $log.cmdletparameters | format-table
    }
    
    end {
        
    }
}

function Get-DLModerationOffAuditLog {
    [CmdletBinding()]
    param (
        $DL
    )

    begin { # This is a bad way to do this, it gets the most recent log with a False value in the params, this generally seems to be the correct log
        $log = (Search-AdminAuditLog -Cmdlets set-distributiongroup -parameters ModerationEnabled -EndDate (get-date) -StartDate (get-date).AddDays(-7) -IsSuccess $true | Where-Object { $_.ObjectModified -eq $DL -and $_.cmdletparameters.Value -contains "False"})[0]
    }
    
    process {
        
        Write-Host "Rundate: " $log.rundate -BackgroundColor Green -ForegroundColor Black
        write-host "Caller: " $log.Caller
        Write-Host "Cmdlet Parameters:"
        $log.cmdletparameters | format-table
    }
    
    end {
        
    }
}

#Main

$credpath = "D:\wpi\powershell\Exchange\DL-Student-Upkeep\exch_automation.xml" # Get the normal XML with exchange creds

$credential = Import-CliXml -Path $credPath

Connect-ExchangeOnline -Credential $credential -ShowBanner:$false # keep from cluttering the transcript

get-DLStudentStatus

Stop-Transcript

## TODO Add an emailer function that just attaches the transcript