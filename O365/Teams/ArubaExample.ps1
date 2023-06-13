Function WriteAPLocations ($airwavehost, $username, [SecureString] $password, [int]$campusindex) {
    # #Set up login and url, build web client
    $webClient = New-Object system.Net.Webclient
    $urlstring = “https://” + $airwavehost + ”/visualrf/campus.xml?buildings=1&sites=1&aps=1”
    Write-host $urlstring
    $url = $urlstring
    $webClient.credentials = new-object System.Net.NetworkCredential($username, $password)
    $webClient.Headers.Add(“user-agent”, “PowerShell”)
    # #Query campus information, get down to buildings
    $data = $webClient.OpenRead($url)
    $reader = new-object system.io.StreamReader $data
    [string] $s = $reader.ReadToEnd()
    $data.Close()
    $reader.Close()
    $xmls = [xml]($s)
    ##Now we have a big xml document with campuses, buildings, sites, aps & radios…
    $ctrcampus = 0
    foreach ($campus in $xmls.campuses.campus) {
        if ($ctrcampus -eq $campusindex) {
            Write-host “`nCampusindex_ “ $campusindex “Campus name_” $campus.name
            $ctrbuilding = 0
            foreach ($building in $campus.building) {
                write-host “`nBuilding index_ ” $ctrbuilding “Building name_“ $building.name
                #foreach($building in $xmls.campuses.campus[0].building){
                #write-host “Campus name_“ $xmls.campuses.campus[0].name
                ##Parsing the building address string in JSON
                #$formatbldgaddress = $campus.name + $building.name
                $formatbldgaddress = “ “
                $a = $building.address
                #Write-host “`nBuilding address string_ “ $a
                $b = $a.Split(“,”)
                #Write-host $b
                foreach ($i in $b) {
                    $c = $i.split(“:”)
                    $t0 = $c[0]
                    #$inda = $t0.IndexOf(“”””)
                    #$indb = $t0.IndexOf(“”””, $inda + 1)
                    #$property = $t0.substring($inda + 1, $indb - 1)
                    $t0 = $t0.trim()
                    $property = $t0.substring(1, $t0.length – 2)
                    $t1 = $c[1]
                    #$indc = $t1.IndexOf(“`””)
                    #$indd = $t1.IndexOf(“`””, $indc + 1)
                    #$value = $t1.substring($indc+ 1, $indd - 1)
                    $t1 = $t1.Trim()
                    $value = $t1.substring(1, $t1.length – 2)
                    Write-host $property “:” $value
                    $formatbldgaddress = $formatbldgaddress + “ -$property ” + $value
                }
                Write-host “`nBuilding address_ “ $formatbldgaddress
                ##Site level extraction
                $ctrsite = 0
                foreach ($site in $building.site) {
                    #write-host “Building index_“ $ctrbuilding “ Site index_“ $ctrsite
                    #write-host “Building name_“$building.name “ Building address_“ $building.address
                    #write-host “Floor_“ $site.floor
                    #write-host “ “
                    ## Add floor number to formatted address as location
                    $formatbldgsiteaddress = $formatbldgaddress + “ –Location Floor “ + $site.floor
                    Write-host “`nSite address_ “$formatbldgsiteaddress
                    $ctrap = 0
                    foreach ($ap in $site.ap) {
                        #check whether the AP supports 8 or 16 SSIDs per radio
                        $ctrradio = 0
                        $bssidsperradio = 16
                        [string] $apmodel = $ap.model
                        [string] $apmanufacturer = $ap.manufacturer
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“65”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“80”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“52”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“1200”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“DUO”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“85”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“102”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“1250”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“RAP-2”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“RAP 2”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“101”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“105”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“60P”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“185”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“44”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“68”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“175”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“930”)) { $bssidsperradio = 8 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“653”)) { $bssidsperradio = 16 }
                        If ( $apmanufacturer.contains(“Aruba”) –and $apmodel.contains(“653”)) { $bssidsperradio = 16 }
                        foreach ($radio in $ap.radio) {
                            #write-host “ radio index_ “ $ctrradio “ MAC_ “ $radio.mac
                            Write-host “`nAP_Name_” $ap.name “ Radio_“ $ctrradio “ Manufacturer_” $apmanufacturer “Model_” $apmodel “ SSIDs_Per_Radio_” $bssidsperradio “ Base_MAC_ “ $radio.mac
                            If ( $null –ne $radio.mac) {
                                [string]$basebssid = $radio.mac
                                $stringpairs = $basebssid.split(“:”)
                                #Write-host “non-null mac_” $radio.mac
                                #Write-host $stringpairs[5]
                                $lsb = [Convert]::ToInt32($stringpairs[5], 16)
                                for ($i = 0; $i –le $bssidsperradio - 1; $i++) {
                                    $stringpairs[5] = “{0:X2}” –f $lsb
                                    $bssid = [string]::join(“:”, $stringpairs)
                                    #Create the LIS configuration cmdlet
                                    $formatbldgsitebssidaddress = “ –BSSID “ + $bssid + “ –Description x:” + $ap.x + “,y:” +
                                    $ap.y + $formatbldgsiteaddress
                                    Write-host “Set-CsLisWirelessAccessPoint “ $formatbldgsitebssidaddress
                                    $lsb++
                                }
                            }
                            $ctrradio ++
                        }
                        $ctrap ++
                    }
                    write-host “----------------------------“
                    $ctrsite ++
                }
                write-host “____________________”
                $ctrbuilding ++
            }
        }
        $ctrcampus ++
    }
}