function Remove-AdvisingLists {
    [CmdletBinding()]
    param (
        [switch] $WhatIf
    )
    
    begin {
        
        # Set path for log files:
        $logPath = "D:\wpi\Logs\DYNManagement\Advising"

        # Get date for logging and file naming:
        $date = Get-Date
        $datestamp = $date.ToString("yyyyMMdd-HHmm")

        Start-Transcript -Append -Path "$($logPath)\$($datestamp)_AdvisingRemoval.log" -Force
        #connect
        Connect-AzureAD
        Connect-ExchangeOnline -ShowBanner:$false
        #It's probably a better idea to instead just grab all the lists starting with ADV- and trim the adv to get our list of advisor lists to delete
        $advisors = get-azureadgroup -SearchString "ADV-" | Select-Object Displayname | ForEach-Object {$_.displayname.split("-")[1]}
        #$advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student"} -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
    }
    
    process {
        foreach ($Advisor in $advisors) {
            $advisor = $advisor.trim()
            Write-Host -backgroundColor GRAY -foregroundColor BLACK "`nChecking List for Advisor: $($advisor)"

            #Get the Lists
            $DynList = get-DistributionGroup -Identity "ADV-$advisor" -ErrorAction SilentlyContinue
            $DynGroup = get-AzureADMSGroup -SearchString "Dyn-$advisor" -ErrorAction SilentlyContinue | where-object { $_.Displayname -eq "Dyn-$advisor" }

            if ($whatif) {
                try {
                    remove-distributiongroup $DynList.Name -WhatIf
                    write-host "remove-azureadmsgroup $($DynGroup.ID)"
                }
                catch {
                    #Write-Host "Lists do not exist for $Advisor"
                } 
            }
            else {
                try {
                    remove-distributiongroup $DynList.Name -Confirm:$false
                    remove-azureadmsgroup -Id $DynGroup.ID
                }
                catch {
                    #Write-Host "Lists do not exist for $Advisor"
                }
            }
            
        }
    }
    
    end {
        Stop-Transcript
    }
}