#Module Skype for Business
# set-csOnlineliswirelessAccessPoint
#https://docs.microsoft.com/en-us/powershell/module/skype/set-csonlinelisswitch?view=skype-ps

function Set-WPIWirelessAccessPointLocation {
    [CmdletBinding()]
    param (
        $CSV
    )
    
    begin {
        $ImportList = import-csv $CSV
    }
    
    process {
        $counter = 0
        foreach ($AccessPoint in $ImportList) {
            $counter++
            Write-Progress -Activity "Importing Wireless Access Points" -Status "Importing $($AccessPoint.BSSID)" -PercentComplete (($Counter/$ImportList.Count) * 100)
            Set-CsOnlineLisWirelessAccessPoint -BSSID $AccessPoint.BSSID.replace(":","-") -Description $AccessPoint.Description -LocationId $AccessPoint.LocationId
        }
    }
    
    end {
        
    }
}

function Set-WPISwitchLocation {
    [CmdletBinding()]
    param (
        $CSV
    )
    
    begin {
        $ImportList = import-csv $CSV
    }
    
    process {
        $counter = 0
        foreach ($Switch in $ImportList) {
            $counter++
            Write-Progress -Activity "Importing Wireless Access Points" -Status "Importing $($Switch.ChassisID)" -PercentComplete (($Counter/$ImportList.Count) * 100)
            Set-CsOnlineLisSwitch -ChassisID $Switch.ChassisID.replace(":","-") -Description $Switch.Description -LocationId $Switch.LocationId
        }
    }
    
    end {
        
    }
}