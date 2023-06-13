#dump files into C:\WPIAPPS on windows terminal servers
#List of terminal servers
# win-ts-p-w07 through 22

function Copy-ToWindowsTS {
    param (
        $SourcePath,
        [string]$DestinationPath
    )
    $TerminalServers = get-adcomputer -filter "Name -like ""win-ts-p-w*""" | Select-Object Name

    if ($null -eq $DestinationPath) {
        foreach ($server in $TerminalServers) {
            copy-item -path $SourcePath -Recurse -Destination "\\$($server.Name)\C$\WPIAPPS" -Container
        }
    }
    else {
        if ($DestinationPath.StartsWith("C")) {
            $destinationpath = $destinationpath.TrimStart("C:\")
        }
        elseif ($DestinationPath.StartsWith("c")) {
            $destinationpath = $destinationpath.TrimStart("c:\")
        }
        foreach ($server in $TerminalServers) {
            copy-item -path $SourcePath -Recurse -Destination "\\$($server.Name)\C$\$($DestinationPath)" -Container
        }
    }
}