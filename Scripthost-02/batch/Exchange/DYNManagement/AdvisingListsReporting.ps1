function Get-AdvisingListMembers {
    [CmdletBinding()]
    param (
        [switch] $WhatIf
    )
    
    begin {
        
        # Set path for log files:
        $logPath = "D:\wpi\Logs\DYNManagement\AdvisingPopulation"

        # Get date for logging and file naming:
        $date = Get-Date
        $datestamp = $date.ToString("yyyyMMdd-HHmm")

        Start-Transcript -Append -Path "$($logPath)\$($datestamp)_AdvisingPopulation.log" -Force
        #connect
        Connect-AzureAD
        Connect-ExchangeOnline
        $advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student"} -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
        #$advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -match "PADV-(.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.replace("PADV-", "").replace("OADV-", "").split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
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
                    Get-DistributionGroupMember -Identity $DynList.Name | Select-Object DisplayName, PrimarySmtpAddress | Out-Null
                }
                catch {
                    Write-Host "ADVlist $($DynList.Name) not found"
                }
                try {
                    Get-AzureADGroupMember -ObjectId $DynGroup.ID | select-object DisplayName, UserPrincipalName | Out-Null
                }
                catch {
                    Write-Host "Dynlist $($DynGroup.ID) not found"
                } 
            }
            else {
                try {
                    Get-DistributionGroupMember -Identity $DynList.Name | Select-Object DisplayName, PrimarySmtpAddress | Export-Csv -NoTypeInformation -Path "$($logPath)\ADV-$($advisor)_population.csv"
                }
                catch {
                    Write-Host "ADVlist $($DynList.Name) not found"
                }
                try {
                    Get-AzureADGroupMember -ObjectId $DynGroup.ID | select-object DisplayName, UserPrincipalName | Export-Csv -NoTypeInformation -Path "$($logPath)\dyn-$($advisor)_population.csv"
                }
                catch {
                    Write-Host "Dynlist $($DynGroup.ID) not found"
                }
            }
            
        }
    }
    
    end {
        Stop-Transcript
    }
}