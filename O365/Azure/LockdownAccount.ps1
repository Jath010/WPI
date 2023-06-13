Import-Module ExchangeOnlineManagement

function Block-WPIUser {
    [CmdletBinding()]
    param (
        $EmailAddress
    )
    
    begin {
        $RunningUser = ([Environment]::UserDomainName + "\" + [Environment]::UserName)
        Connect-ExchangeOnline -UserPrincipalName $RunningUser
        Connect-AzureAD -AccountId $RunningUser
    }
    
    process {
        $OID = get-azureADUser -ObjectId $EmailAddress
        Set-AzureADUser -objectID $OID -AccountEnabled $false
        Revoke-AzureADUserAllRefreshToken -ObjectId $OID
        Clear-MobileDevice -AccountOnly -Identity $EmailAddress -NotificationEmailAddresses ([Environment]::UserName + "@wpi.edu")
    }
    
    end {
        
    }
}


