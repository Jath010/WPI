
$computers = Get-ADComputer -Filter * -SearchBase "OU=PUBLIC,OU=WPIWorkstations,DC=admin,DC=wpi,DC=edu" -ResultSetSize $null

foreach($computer in $computers){
    try{
        $result = (Get-CimInstance -ComputerName $Computer.Name -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard).SecurityServicesRunning
    }
    catch{

    }
    if($result -eq '1'){
        $computer.Name
    }
}