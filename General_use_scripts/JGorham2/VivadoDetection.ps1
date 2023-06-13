if (Test-Path C:\Xilinx\Vivado\2019.1\bin\vivado) {
    Write-Host "Installed"
}
elseif (Test-Path C:\Xilinx) {
    if (Test-Path C:\Xilinx\Vivado) {
        cmd /c rmdir C:\Xilinx\Vivado /Q/S
    }
    Remove-Item C:\Xilinx -Recurse -Force
}