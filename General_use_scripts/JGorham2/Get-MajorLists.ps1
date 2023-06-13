function Get-CEE_EVE_AREN_Undergrads {
    [CmdletBinding()]
    param (
        
    )
    
    $students = Get-ADUser -Filter 'Enabled -eq $True -And ExtensionAttribute7 -eq "Student"' -Properties extensionattribute7, extensionattribute4 |Where-Object {$_.extensionattribute4 -match "(MJ-|.*;)(EV|CE|AREN);.*MN-.*"}
    if (Test-Path -Path "C:\Tmp") {
        $students | select-object Name, UserPrincipalName, ExtensionAttribute4 | Export-Csv -Path "c:\tmp\CEE_EVE_AREN_Undergrads.csv" -NoTypeInformation
    }else {
        mkdir C:\tmp
        $students | select-object Name, UserPrincipalName, ExtensionAttribute4 | Export-Csv -Path "c:\tmp\CEE_EVE_AREN_Undergrads.csv" -NoTypeInformation
    }
}