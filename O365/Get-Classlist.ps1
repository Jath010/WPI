function Get-WPIClassList {
    [CmdletBinding()]
    param (
        $ClassNumber,
        $Term,
        $Year
    )

    $ClassGroups = get-adgroup -filter "name -like ""${ClassNumber}_*""" | Where-Object {$_.name -like "*_${Term}*" -and $_.name -like "*_${Year}*"} | Select-Object samaccountname
    
    foreach($Group in $ClassGroups.samaccountname){
        Get-ADGroupMember $Group | Select-Object samaccountname
    }
}