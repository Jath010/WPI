$Users = Import-Csv C:\tmp\MissingMFA.txt
foreach($user in $users){
    $samaccountname = (($User.UserName).Split("@"))[0]
    
}