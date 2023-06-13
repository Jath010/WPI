Write-host "Generating a secure credential XML"
$Credentials = get-credential
$Path = read-host "Please enter the Target Path"

$script={$args[0]|export-clixml -path "$($args[1])\credentials.xml"}

$script={
    $var = new-item "D:\tmp\testing.txt"
    set-content $var "$path"
}
# , "-args $Credentials, $Path"
Start-Process powershell.exe -ArgumentList "-NoProfile -nologo -command $script -args $Credentials, $Path" -credential $Credentials -WindowStyle Hidden