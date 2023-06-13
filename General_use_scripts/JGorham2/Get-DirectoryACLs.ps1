function Get-DirectoryACLs {
    param (
        $path,
        $reportpath
    )
    #Logic courtesy of Michael Fimin, some guy on spiceworks
    #$path = "\\storage\ifs$\dept\Human Resources" #define path to the shared folder
    #$reportpath = "C:\tmp\ACL.csv" #define path to export permissions report
    #script scans for directories under shared folder and gets acl(permissions) for all of them
    #Get-ChildItem -Recurse $path | Where-Object { $_.PsIsContainer } | ForEach-Object { $path1 = $_.fullname; Get-Acl $_.Fullname | ForEach-Object { $_.access | Add-Member -MemberType NoteProperty '.\Application Data' -Value $path1 -passthru }} | Export-Csv $reportpath -NoTypeInformation
    Get-ChildItem $path | Where-Object { $_.PsIsContainer } | ForEach-Object { $path1 = $_.fullname; Get-Acl $_.Fullname | ForEach-Object { $_.access | Add-Member -MemberType NoteProperty '.\Application Data' -Value $path1 -passthru } } | Export-Csv $reportpath -NoTypeInformation
}

function Get-DirectoryGroupMembers {
    param (
        $path,
        $reportpath
    )
    $var = Get-ChildItem $path | Where-Object { $_.PsIsContainer } | ForEach-Object { $path1 = $_.fullname; Get-Acl $_.Fullname | ForEach-Object { $_.access | Add-Member -MemberType NoteProperty '.\Application Data' -Value $path1 -passthru } }
    
    foreach ($path in $var) {
        $GroupName = (($path.IdentityReference).split("\"))[1]
        $GroupName | Out-File C:\tmp\$reportpath -Append
        Get-ADGroupMember $GroupName | Select-Object samaccountname | Out-File C:\tmp\$reportpath -Append
    }
}