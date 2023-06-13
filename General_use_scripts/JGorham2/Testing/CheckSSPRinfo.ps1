connect-azuread

<#
Easier Version
$users |ForEach-Object {Get-AzureADUser -ObjectID $_@wpi.edu | select UserPrincipalName,otherMails,mobile,telephonenumber}
#>

$array = @()
$users = Get-Content C:\tmp\missingSSPR
$users |ForEach-Object {Get-AzureADUser -ObjectID $_@wpi.edu | select UserPrincipalName,otherMails,mobile,telephonenumber}
<#
foreach ($user in $Users) {
    $Result = "" | Select user,altmail,mobile,TelephoneNumber
    $result.user = $user
    $result.altMail = (Get-AzureADUser -ObjectID ${user}@wpi.edu | select otherMails).otherMails
    $result.Mobile = (Get-AzureADUser -ObjectID ${user}@wpi.edu | select Mobile).mobile
    $result.TelephoneNumber = (Get-AzureADUser -ObjectID ${user}@wpi.edu | select TelephoneNumber).TelephoneNumber
    $array += $result
}
$array|format-table
#>

<#
Connect-MsolService
$users |ForEach-Object {Get-MsolUser -UserPrincipalName $_@wpi.edu | select UserPrincipalName,AlternateEmailAddresses,MobilePhone,PhoneNumber}
#>
<#
$RepairUsers = import-csv -Path C:\tmp\frontier_students_072020.xlsx
foreach($user in $RepairUsers){
    $email = $user.USERNAME
    $altmail =  $User.ATAT_EMAIL
    Set-AzureADUser -ObjectId ${email}@wpi.edu -OtherMails @("${altmail}")
}
#>