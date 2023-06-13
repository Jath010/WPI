# Fix ID Photos
# jmac@wpi.edu 10/5/21
#$c = Get-Credential -UserName "ADMIN\jmac_prv"
#$photoDir = "\\storage.wpi.edu\services\id_photos\Photos"
$photoDir = "C:\Users\jmac\Desktop\ID_test"
$server = "ONGRD-DB-T-W01\LENEL"
$time_since = (Get-date).AddDays(((-24).ToString("yyyy-MM-dd HH:mm:ss")))
cd $photoDir
foreach ($file in get-childitem) {
    #if photo.date newer than 1 day
    if ($file.CreationTime -gt 0) {
        $id = $file.BaseName #extract 9-digit ID from file name of photo
        $empid = Invoke-Sqlcmd -ServerInstance $server -Query "SELECT EMPID FROM [AccessControl].[dbo].[BADGE] WHERE ID like '$id'" #get the empid off a badge
        $empid = $empid.Item(0) #convert empid into useable value
        $photo = [System.IO.File]::ReadAllBytes($file.FullName) #read in file as raw bytes
        $lastChanged = Invoke-Sqlcmd -ServerInstance $server -Query "SELECT LASTCHANGED FROM [AccessControl].[dbo].[MMOBJS] WHERE EMPID like '$empid'" #last changed time of photo in Lenel DB
        $lastChanged = $lastChanged.Item(0)
        $photoCheck = Invoke-Sqlcmd -ServerInstance $server -Query "IF EXISTS (SELECT LNL_BLOB FROM Accesscontrol.dbo.MMOBJS WHERE EMPID ='$empid') BEGIN SELECT 1 END ELSE BEGIN SELECT 0 END " | Select-Object -ExpandProperty Column1 #see if photo exists in Lenel DB
        $lastAccess = $file.LastAccessTime
        if ($photoCheck -eq 0) {
            #photo doesn't exist in DB yet.
            Invoke-Sqlcmd -ServerInstance $server -Query "INSERT INTO [AccessControl].[dbo].[MMOBJS] (empid,object,type,lnl_blob) Values($empid,1,0,$photo)"
        }
        else {
            #(($photoCheck -eq 1) -and ($lastChanged -lt $lastAccess)) { #photo does exist, but it's older.
            Invoke-Sqlcmd -ServerInstance $server -Query "UPDATE [AccessControl].[dbo].[MMOBJS] SET LNL_BLOB = '$photo' where empid = $empid and object = 1 and type = 0" #update photo in DB
        }
        Write-Host $id
        Write-Host $empid
    }
}


$file = "D:\tmp\ImageFormatTesting\coolSkeleton.jpg"
$photo = [System.IO.File]::ReadAllBytes($file)

([System.IO.File]::ReadAllBytes("D:\tmp\ImageFormatTesting\Toast.jpg") | Format-Hex | Select-Object -Expand Bytes | ForEach-Object { '{0:x2}' -f $_ }) -join '' | Out-File D:\tmp\ImageFormatTesting\HexDump.txt
(Get-Content .\coolSkeleton.jpg | Out-String | Format-Hex -encoding utf7 | Select-Object -Expand Bytes | ForEach-Object { '{0:x2}' -f $_ }) -join '' | Out-File .\HexDump.txt
