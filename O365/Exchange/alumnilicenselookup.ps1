#$userList = Import-csv C:\tmp\listBustedAccounts.txt
$userlist = get-aduser -filter * -SearchBase "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
foreach($user in $userlist){
    $UUID = $user.samaccountname+"@wpi.edu"
    if((get-azureaduser -ObjectId $UUID).assignedlicenses.skuid -eq "28db6bcc-8442-405b-9ebb-e2f4da7355ed"){
        Write-Host $user.samaccountname "has Exchange Plan 1 for Alumni assigned"
    }
    # else{
    #     Write-host $User.samaccountname "Does not have Exchange Plan 1 for Alumni assigned" -BackgroundColor DarkGreen
    # }
}