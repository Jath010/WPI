$computers = (get-adcomputer -Filter "Name -like 'FI103*'").Name

foreach($computer in $computers){
    #Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StartService -Computername $computer |Out-Null
    #Invoke-Command -ComputerName $computer -ScriptBlock {[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\RBE\sloeber\jre\bin\java.exe", "Machine")}
    #Invoke-Command -ComputerName $computer -ScriptBlock {setx PATH "$env:path;C:\RBE\sloeber\jre\bin\java.exe" -m}
    Write-Host $computer
    Invoke-Command -ComputerName $computer -ScriptBlock {$env:path}
}

setx /s FI103-03 PATH "C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\System32\Wbem;C:\WINDOWS\System32\WindowsPowerShell\v1.0\;C:\WINDOWS\System32\OpenSSH\;C:\Program Files\doxygen\bin;C:\Program Files (x86)\Graphviz2.38\bin;C:\Program Files\dotnet\;C:\Program Files\Microsoft SQL Server\130\Tools\Binn\;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\;C:\Program Files (x86)\IncrediBuild;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\;C:\Program Files (x86)\Microsoft SQL Server\120\Tools\Binn\;C:\Program Files\Microsoft SQL Server\120\Tools\Binn\;C:\Program Files\Microsoft SQL Server\120\DTS\Binn\;C:\Program Files (x86)\Windows Kits\8.1\Windows Performance Toolkit\;C:\Program Files\PuTTY\;C:\RBE\sloeber\jre\bin\java.exe" /m

[Environment]::SetEnvironmentVariable("Path", "C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\System32\Wbem;C:\WINDOWS\System32\WindowsPowerShell\v1.0\;C:\WINDOWS\System32\OpenSSH\;C:\Program Files\doxygen\bin;C:\Program Files (x86)\Graphviz2.38\bin;C:\Program Files\dotnet\;C:\Program Files\Microsoft SQL Server\130\Tools\Binn\;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\;C:\Program Files (x86)\IncrediBuild;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\;C:\Program Files (x86)\Microsoft SQL Server\120\Tools\Binn\;C:\Program Files\Microsoft SQL Server\120\Tools\Binn\;C:\Program Files\Microsoft SQL Server\120\DTS\Binn\;C:\Program Files (x86)\Windows Kits\8.1\Windows Performance Toolkit\;C:\Program Files\PuTTY\;C:\RBE\sloeber\jre\bin\java.exe", [System.EnvironmentVariableTarget]::Machine)
[Environment]::SetEnvironmentVariable("Path", "", [System.EnvironmentVariableTarget]::Machine)