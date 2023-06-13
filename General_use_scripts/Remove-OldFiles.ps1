##Path where files to clean reside
$LogPath = "E:\inetpub\logs\LogFiles\" 

##Set File age to be removed
$maxDaystoKeep = -7 

##Set variable for old files to remove
$itemsToDelete = dir $LogPath -Recurse -File *.log | Where LastWriteTime -lt ((get-date).AddDays($maxDaystoKeep)) 

##Remove files
ForEach ($item in $itemsToDelete){ Get-item $item.PSPath | Remove-Item -Verbose }