@echo off

"C:\Program Files (x86)\WinSCP\WinSCP.com" ^
  /log="D:\wpi\batch\SFTP\Workday\MMS\mms-quest-WinSCP.log" /ini=nul /rawconfig Logging\LogFileAppend=0 ^
  /command ^
    "open sftp://quest@sftp.wpi.edu/ -hostkey=""ssh-ed25519 255 8QusvSdM0zwmb6MjYBYsGBoGZqdyFy2R9OOT0atPlII="" -privatekey=""D:\wpi\batch\SFTP\ssh\quest-sftp.ppk""" ^
    "cd /quest2" ^
    "lcd ""//storage.wpi.edu/dept/Workday Integrations/COVID/SFTP""" ^
    "get *.*" ^
    "exit"

set WINSCP_RESULT=%ERRORLEVEL%
if %WINSCP_RESULT% equ 0 (
  echo Success
) else (
  echo Error
)

exit /b %WINSCP_RESULT%
