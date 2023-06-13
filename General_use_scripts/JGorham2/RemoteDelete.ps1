

function Enable-WinRMRemoting {
    [CmdletBinding()]
    param (
        $ComputerName
    )
    Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StartService -Computername $ComputerName |Out-Null
}

function Remove-RemoteItem {
    [CmdletBinding()]
    param (
        $Target    
    )
    
}

$credentials = Get-Credential


$targets = Get-Content C:\tmp\xilinxbadinstalls.txt
$targets | ForEach-Object {$targets += "$_.wpi.edu"}

foreach($Computer in $targets){
    Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StartService -Computername $Computer |Out-Null
    Invoke-Command {Remove-Item -Path C:\xilinx -Recurse} -computername $computer -credential $credentials
}