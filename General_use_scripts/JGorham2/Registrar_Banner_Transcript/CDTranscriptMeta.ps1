Connect-PnPOnline -url "https://wpi0.sharepoint.com/sites/RegistrarTranscriptArchive" -Interactive

$rootFolder = (get-pnplist -Identity "Record Library").RootFolder.Name
$StoragePathRoot = "D:\tmp\Registrar Banner Transcript\copyToSharePoint\CD Contents"
$RootPath = "$rootfolder/Powershell CD Upload"

$CDIndex = Import-Csv "D:\tmp\Registrar Banner Transcript\CDFiles.csv"

$counter = 0

foreach ($User in $CDIndex) {
    $counter++
    Write-Progress -Activity "Processing Transcripts" -CurrentOperation $User.lastname -PercentComplete (($counter / $CDIndex.count) * 100) -Id 0
    $Files = $user.Files.split(";")
    $fullname = $User.lastname+" "+$user.firstname

    $counter2 = 0
    foreach ($File in $Files) {
        $counter2++
        Write-Progress -Activity "Processing Subfiles" -CurrentOperation $File -PercentComplete (($counter2 / $Files.count) * 100) -Id 1 -ParentId 0
        $OriginalPath = $StoragePathRoot+"\"+$File.split("/")[0]+"\"+$File.split("/")[1]
        $filepath = $RootPath+"/"+$File.split("/")[0]
        $filename = $User.firstname+"_"+$User.lastname+"_"+$File.split("/")[1]
        
        $var = Add-PnPFile -path $OriginalPath -Folder $filepath -Values @{DOB = $User.DOB; Names = $fullname} -NewFileName $filename
        Set-PnPFileCheckedIn -Url $var.ServerRelativeUrl -ErrorAction SilentlyContinue
    }
}