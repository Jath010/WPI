# Created by Jmgorham2
# 4/10/2019
# Purpose is to run through AD and edit any user who has a CN that's their EmployeeID and change it to be their PIDM

#EmployeeID is their actual employee ID
#EmployeeNumber is their PIDM

$Userlist = Get-ADUser -Properties *
$Studentlist = $UserList | where-object {$_.DistinguishedName -like "*OU=Students*"} |select-object samaccountname

foreach($user in $Studentlist){
    $EmployeeID = $user.EmployeeID
    if($user.CN -like "*(${EmployeeID})"){
        if($null -ne $user.initials){
            $CorrectName = $User.sn+", "+$user.GivenName+" "+$user.initials+" "+"("+$user.EmployeeNumber+")"
        }
        else{
            $CorrectName = $User.sn+", "+$user.GivenName+" "+"("+$user.EmployeeNumber+")"
        }
        Rename-ADObject -Identity $user.DistinguishedName -NewName $CorrectName                                 #Replaces CN with new one
    }
}