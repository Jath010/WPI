foreach ($User in $ZoomUsers) {
    $samaccountname = $user.Email.Split("@")[0]
    get-aduser $samaccountname -Properties extensionattribute7 | Export-Csv -Path C:\tmp\ZoomADUser.csv -NoTypeInformation -Append -Force
}

#junk