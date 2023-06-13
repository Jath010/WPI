Connect-AzureAD
#Get-AzureADUser -All $True | Where-object {$_.ExtensionProperty.onPremisesDistinguishedName -like "*,OU=Resource Mailboxes,OU=Other Accounts,OU=Accounts,DC=admin,DC=wpi,DC=edu"}
#Get-AzureADUser -objectid calendar_GWP1050@wpi.edu | Where-object {$_.ExtensionProperty.onPremisesDistinguishedName -like "*,OU=Resource Mailboxes,OU=Other Accounts,OU=Accounts,DC=admin,DC=wpi,DC=edu"}

$UserList = Get-AzureADUser -All $True

foreach($User in $userlist){
    if($User.ExtensionProperty.onPremisesDistinguishedName -like "*,OU=Resource Mailboxes,OU=Other Accounts,OU=Accounts,DC=admin,DC=wpi,DC=edu"){
        Write-host $user.UserPrincipalName
        $user.UserPrincipalName >> C:\tmp\temp\ResourceMailboxes.txt
    }
}