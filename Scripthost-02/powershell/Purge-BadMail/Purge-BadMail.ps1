<#
    Script purges all searches named Phishing -* and changes their name to append "Completed"
#>

#These lines were nabbed from the Get-AlumniPilotStats.ps1 script on Scripthost-02
##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
<# $credentials = $null;
#$RemoteCredentials=$null;
$credPath = $null;
#$RemoteCredPath=$null
## The following credentials use the UPN "exch_automation@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange on-premise and Exchange Online
$credPath = 'D:\wpi\batch\Exchange\PurgeBadMail\exch_automation@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath #>

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
##$ExchangeOnlineSession = $null

## Load Exchange Online
##$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
##Import-PSSession $ExchangeOnlineSession -Prefix Cloud

function Remove-BadMail {
    [cmdletbinding()]
    param(
        [switch]
        $noblock
    )
    # Get login credentials 
    # if (test-path C:\tmp\jmgorham2_prvCredential.xml) { $Credentials = Import-Clixml C:\tmp\jmgorham2_prvCredential.xml }
    # if (!(Get-PSSession | Where-Object { $_.ComputerName -match 'protection.outlook.com' })) {
    #     if (!$Credentials) { $Credentials = Get-Credential -UserName "$($env:username)@wpi.edu" -Message 'Please enter your email address and password' }
    #     $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Credentials -Authentication Basic -AllowRedirection 
    #     Import-PSSession $Session -AllowClobber -DisableNameChecking 
    #     #$Host.UI.RawUI.WindowTitle = $UserCredential.UserName + " (Office 365 Security & Compliance Center)" 
    # } 
    $credPath = "D:\wpi\powershell\Purge-BadMail\exch_automation"
    $credentials = Import-CliXml -Path $credPath
    #Import-Module (Import-PSSession $(New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection) -DisableNameChecking) -Global -DisableNameChecking

    $ComplianceOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $credentials -Authentication Basic -AllowRedirection
    #Import-PSSession $ComplianceOnlineSession
    #Export-PSSession -Session $ExchangeOnlineSession -CommandName "*-ComplianceSearch*" -OutputModule SpamClean -AllowClobber
    $ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
    #Import-PSSession $ExchangeOnlineSession
    Connect-AzureAD -Credential $credentials
 

    #Assemble Search List
    try{
        #$Searches = Get-ComplianceSearch | Sort-Object JobEndTime -Descending
        $Searches = Invoke-command -ScriptBlock {Get-ComplianceSearch} -Session $ComplianceOnlineSession |Sort-Object JobEndTime -Descending
    }
    catch{
        $_|Out-File "D:\wpi\powershell\Purge-BadMail\Log\ErrorLog-$(get-date -f yyyyMMdd-HHmm)"
    }

    #Display Search List
    #$Searches | Select-Object Name,Status,JobEndTime  | Out-Default

    #$TargetLog = $null
    #Loop through each item
    foreach ($search in $Searches) {
        #Only run on searches beginning with "Phishing - "
        if ($search.name -like "Phishing -*" -and $Search.status -eq "Completed") {
            $SearchDetails = $Search   
            $SearchName = $search.name
        
            #Display Search being acted on
            #Get-ComplianceSearch $SearchName | Select Name,Status,JobEndTime | Out-Default
            Write-Verbose $searchname

            #Stall if there are 4 jobs already running
            While ((Invoke-command -ScriptBlock {Get-ComplianceSearchAction} -Session $ComplianceOnlineSession | Where-Object { $_.Action -eq 'Purge' -and $_.Status -eq 'InProgress' } | Measure-Object).Count -ge 4) {
                Start-Sleep 10
            }


            #Purge current Search
            Write-Verbose "Creating Purge" 
            try { Invoke-command -ScriptBlock {New-ComplianceSearchAction -SearchName $args[0] -Purge -PurgeType SoftDelete -Confirm:$false} -ArgumentList $SearchName -Session $ComplianceOnlineSession }
            catch { continue }

            if (!$noblock) { 
                #Extract Email Address from the Search
                
                $regex = [regex] ".*\(senderauthor=(.*?)\)"
                
                $match = $regex.match($SearchDetails.contentmatchquery)
                $BlockList_Email = $match.groups[1].value
                $BlockList_Email | Out-File -FilePath "D:\wpi\powershell\Purge-BadMail\Log\PurgeLog-$(get-date -f yyyyMMdd-HHmm).txt" -Append

                if ($BlockList_Email -notlike "*@wpi.edu") {
                    #Blocklist the Email Address the Search was run against
                    Write-Verbose "Blocking Sender"
                    Invoke-command -ScriptBlock {Set-HostedContentFilterPolicy -Identity 'Default' -BlockedSenders @{add = $args[0] }} -ArgumentList $BlockList_Email -Session $ExchangeOnlineSession
                }
            }
            #Rename the Search so it will not be acted on by this script in the future
            $rename = 'Completed - ' + $SearchName
            Write-Verbose "Renaming Search"
            Invoke-command -ScriptBlock {Set-ComplianceSearch $args[0] -Name $args[1]} -ArgumentList $SearchName, $rename -Session $ComplianceOnlineSession
        }
    }

    Disconnect-AzureAD
    Get-PSSession | Remove-PSSession

}
Remove-BadMail