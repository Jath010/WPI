REM get the file and place it somewhere
REM rotate logging
REM which user should move stuff

@echo off

"C:\Program Files (x86)\WinSCP\WinSCP.com" ^
  /log="D:\wpi\batch\SFTP\handshake\handshake-WinSCP.log" /ini=nul /rawconfig Logging\LogFileAppend=0 ^
  /command ^
    "open sftp://handshake@sftp.wpi.edu/ -hostkey=""ssh-ed25519 255 8QusvSdM0zwmb6MjYBYsGBoGZqdyFy2R9OOT0atPlII="" -privatekey=""D:\wpi\batch\SFTP\ssh\quest-sftp.ppk""" ^
    "cd /handshake2" ^
    "get *.csv ""\\cher-app-p-w02\files\handshake\staging\""" ^
    "get -delete *.csv ""\\storage\dept\Information Technology\CCC\Helpdesk\fc_hdstaff\CherwellImports\CDC\staging\""" ^
	"exit"

set WINSCP_RESULT=%ERRORLEVEL%
if %WINSCP_RESULT% equ 0 (
  echo Success
) else (
  echo Error
)



exit /b %WINSCP_RESULT%



