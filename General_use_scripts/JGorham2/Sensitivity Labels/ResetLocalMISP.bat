REM ***Delete the RMS registry settings for the user. 

reg delete HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\DRMC /v "DefaultServer" /f
reg delete HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\DRMC /v "DefaultServerUrl" /f


REM ***Find and delete RMS Server/Cluster

for /f "delims=" %%a in (' 
    reg query "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\MSIPC" ^|
    find "rms.eu.aadrm.com"
') do (
     set "regs=%%a"
)

reg delete "%regs%" /f


REM ***Clear the existing licenses, GICs (RACs), and etc.

REM rmdir /S /Q %localappdata%\microsoft\MSIPC 
md %localappdata%\temp_RMS
robocopy %localappdata%\Temp_RMS %localappdata%\Microsoft\MSIPC /mir