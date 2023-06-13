<#
    Script purges all searches named Phishing -* and changes their name to append "Completed"
#>

#These lines were nabbed from the Get-AlumniPilotStats.ps1 script on Scripthost-02
##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
$credentials=$null;
#$RemoteCredentials=$null;
$credPath=$null;
#$RemoteCredPath=$null
## The following credentials use the UPN "exch_automation@wpi.edu" for the username.  This is used to connect to remote PSSessions for Exchange on-premise and Exchange Online
$credPath = 'D:\wpi\batch\Exchange\PurgeBadMail\exch_automation@wpi.edu.xml'
$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
#Import-PSSession $ExchangeOnlineSession -Prefix Cloud
Connect-ExchangeOnline -Credential $credentials

<# 
# Get login credentials 
if (!(Get-PSSession | Where-Object {$_.ComputerName -match 'protection.outlook.com'}))
{
    if (!$Credentials) {$Credentials = Get-Credential -UserName "$($env:username)@wpi.edu" -Message 'Please enter your email address and password'}
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Credentials -Authentication Basic -AllowRedirection 
    Import-PSSession $Session -AllowClobber -DisableNameChecking 
    #$Host.UI.RawUI.WindowTitle = $UserCredential.UserName + " (Office 365 Security & Compliance Center)" 
} 
#>

#Assemble Search List
$Searches = Get-ComplianceSearch | Sort-Object JobEndTime -Descending

#Display Search List
#$Searches | Select-Object Name,Status,JobEndTime  | Out-Default

#Loop through each item
foreach($search in $Searches)
{
    #Only run on searches beginning with "Phishing - "
    if($search.name -like "Phishing -*")
    {
        $SearchDetails = $Search   
        $SearchName = $search.name
        
        #Display Search being acted on
        #Get-ComplianceSearch $SearchName | Select Name,Status,JobEndTime | Out-Default
        Write-host $searchname

        #Stall if there are 4 jobs already running
        While ((Get-ComplianceSearchAction | Where-Object {$_.Action -eq 'Purge' -and $_.Status -eq 'InProgress'} | Measure-Object).Count -ge 4) 
        {
            Start-Sleep 10
        }


        #Purge current Search 
        New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete -Confirm:$false

        #Extract Email Address from the Search
        $regex = [regex] ".*\(senderauthor=(.*?)\)"
        $match = $regex.match($SearchDetails.contentmatchquery)
        $BlockList_Email = $match.groups[1].value

        #Blocklist the Email Address the Search was run against
        Set-HostedContentFilterPolicy -Identity 'Default' -BlockedSenders @{add=$BlockList_Email}
        
        #Rename the Search so it will not be acted on by this script in the future
        $rename = 'Completed - '+$SearchName
        Set-ComplianceSearch $SearchName -Name $rename
    }
}