
# #get dropped csv files

# $csvcontent = get-childitem -Filter "*.csv" -Path "\\storage\dept\Information Technology\CCC\Helpdesk\fc_hdstaff\CherwellImports\CDC\staging\" 

# if ($csvcontent -ne $null){

# # get newest file 
# $newfile = $csvcontent | sort-object -property "LastWriteTime" -Descending | select-object -First 1

# # copy with static name and overwrite existing csv
# Copy-Item $newfile.VersionInfo.FileName -Destination "\\storage\dept\Information Technology\CCC\Helpdesk\fc_hdstaff\CherwellImports\CDC\handshake-export.csv" -force

# #delete files in staging directory older than 5 days
# $csvcontent | Where-Object LastWriteTime -lt (Get-Date).AddDays(-5) | remove-item -Force
# }



function Set-CopyCleanupCSV ($source,$destination) {

    $csvcontent = get-childitem -Filter "*.csv" -Path $source

    if ($csvcontent -ne $null){
    
    # get newest file 
    $newfile = $csvcontent | sort-object -property "LastWriteTime" -Descending | select-object -First 1
    
    # copy with static name and overwrite existing csv
    Copy-Item $newfile.VersionInfo.FileName -Destination $destination -force
    
    #delete files in staging directory older than 5 days
    $csvcontent | Where-Object LastWriteTime -lt (Get-Date).AddDays(-5) | remove-item -Force

    # grow up and add some actual logging you professional you 

    }

}

Set-CopyCleanupCSV "\\storage\dept\Information Technology\CCC\Helpdesk\fc_hdstaff\CherwellImports\CDC\staging\" "\\storage\dept\Information Technology\CCC\Helpdesk\fc_hdstaff\CherwellImports\CDC\handshake-export.csv"
Set-CopyCleanupCSV "\\cher-app-p-w02\files\handshake\staging\" "\\cher-app-p-w02\files\handshake\handshake-export.csv"