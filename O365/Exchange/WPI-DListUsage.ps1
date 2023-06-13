# Script intended to allow for verification of if a mailbox has been sent to within the last 10 days

#Example
#Get-MessageTrace -RecipientAddress dl-cs-grads@wpi.edu -StartDate "$((Get-Date).AddDays(-10).ToString("MM/dd/yyyy"))" -EndDate "$((Get-Date).ToString("MM/dd/yyyy"))"

function Get-WPIDLUsage {
    param (
        $DistributionList,
        [switch]
        $Inactive
    )
    # Grab dates for search period
    $EndTime = (Get-Date).ToString("MM/dd/yyyy")
    $StartTime = (Get-Date).AddDays(-10).ToString("MM/dd/yyyy")
    
    # Fix address if just the username is entered
    if($DistributionList -notlike "*@wpi.edu") {
        $DistributionList = $DistributionList+"@wpi.edu"
    }

    #The actual working command for getting all the emails received
    $Trace = Get-MessageTrace -RecipientAddress $DistributionList -StartDate $StartTime -EndDate $EndTime

    $Count = $Trace.count #Loaded into a variable here because it was breaking the Write-host when I was doing it there
    if($Inactive) {
        if($Count -eq 0) {
            Write-Host "${DistributionList} is inactive" -BackgroundColor DarkGreen -ForegroundColor Black
        }
        else {
            if($var.count -lt 1) {
                Write-Host "${DistributionList} is active	                                    *%* Messages received by ${DistributionList} in last 10 days: 1" -BackgroundColor DarkRed
            }
            else {
                Write-Host "${DistributionList} is active	                                    *%* Messages received by ${DistributionList} in last 10 days: ${Count}" -BackgroundColor DarkRed
            }
        }
    }
    else {
        $Trace.count
    }
}

function Get-WPIDLUsageCSV {
    param (
        $CSV
    )
    $List = Get-Content $CSV
    foreach($DL in $List) {
        Get-WPIDLUsage $DL -Inactive
    }
}