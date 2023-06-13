function Add-IndividualBannerTranscript {
    [CmdletBinding()]
    param (
        $WPI_ID,
        $Transcript_Type,
        $bannerIndex,
        [switch] $transcript
    )
    
    
    if ($transcript) {
        $logpath = 'D:\tmp\Registrar Banner Transcript'
        $date = Get-Date
        $datestamp = $date.ToString("yyyyMMdd-HHmm")
        Start-Transcript -Append -Path "$($logPath)\$($datestamp)_IndividualBannerUpload.log" -Force
    }

    Connect-PnPOnline -url "https://wpi0.sharepoint.com/sites/RegistrarTranscriptArchive" -Interactive

    $rootFolder = (get-pnplist -Identity "Record Library").RootFolder.Name
    $StoragePathRoot = "\\storage\dept\Enrollment Management\Registrar\reg_banner_transcripts"
    $RootPath = "$rootfolder/Powershell Upload"

    if ($null -eq $bannerIndex) {
        Write-Host "Getting Banner Index"
        $BannerIndex = Import-Csv "D:\tmp\Registrar Banner Transcript\wpi_transcript_index.csv"
    }

    $entry = $BannerIndex | Where-Object { $_.WPI_ID -eq $WPI_ID -and $_.TRANSCRIPT_TYPE -eq $Transcript_Type }

    if ($null -ne $entry) {
        
        $PathSplit = $Entry.PDF_File_Name.split("\")[0] + "/" + $Entry.PDF_File_Name.split("\")[1]
        $uploadPath = "$RootPath/$pathsplit"
        #$storagePath = $StoragePathRoot+"\"+$Entry.PDF_FILE_NAME
        $storagePath = $StoragePathRoot + "\" + $Entry.PDF_File_Name.split("\")[0] + "\" + $Entry.PDF_File_Name.split("\")[1] + "\Transcript_Fix_JimM\" + $Entry.PDF_File_Name.split("\")[2]

        Write-Host "Uploading $storagepath"
        try {
            $var = Add-PnPFile -path $storagePath -Folder $uploadPath -Values @{DOB = $Entry.DOB; WPI_x0020_ID = $Entry.WPI_ID; Names = $Entry.NAMES; Entry_x0020_Date = $Entry.ENTRY_DATE; Last_x0020_Term_x0020_Date = $Entry.LAST_TERM_DATE; SSN = $Entry.SSN }
        }
        catch {
            Write-Host "Failed to upload file $storagepath" -BackgroundColor Red -ForegroundColor Black
        }

        Write-Host "Checking in $storagepath"
        try {
            Set-PnPFileCheckedIn -Url $var.ServerRelativeUrl -ErrorAction SilentlyContinue
        }
        catch {
            Write-host "Failed to Check-In $storagepath" -BackgroundColor Red -ForegroundColor Black
        }
    }
    

    
    if ($transcript) {
        Stop-Transcript
    }
    
}

