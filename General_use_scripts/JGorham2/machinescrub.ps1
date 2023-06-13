$machines = get-content C:\tmp\xilinxfails.txt
workflow clear-xilinx {
    param (
        $Path
    )
    $machines = get-content $path
    foreach -Parallel ($machine in $machines) {
        #Invoke-WmiMethod -Path "Win32_Service.Name='WinRM'" -Name StartService -Computername $machine
        if (test-path \\$machine\c$\Xilinx) {
            Remove-Item \\$machine\c$\Xilinx -Recurse -Force
        }
    }  
}