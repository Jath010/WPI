#Goal is to generically get all the elements found in an extension attribute

function Get-AttributeUniques {
    [CmdletBinding()]
    param (
        $AttributeName
    )
    
    $Uniques = get-aduser -Filter * -Properties $AttributeName | Select-Object $AttributeName | Sort-Object -Property $AttributeName -Unique
    
    return $Uniques
    
}

function Get-ExtensionAttribute5BuildingUniques {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $uniquesUnparsed = Get-AttributeUniques extensionAttribute5 | Where-Object {$null -ne $_.extensionAttribute5}
    }
    
    process {
        $buildings = $uniquesUnparsed | ForEach-Object {if($null -ne $_.extensionattribute5){$_.extensionattribute5.split(";")[0]}} | get-unique | ForEach-Object {$_.split("-")[1]} | Where-Object {$_ -ne $null -and $_ -ne " "}
    }
    
    end {
        return $buildings
    }
}

function Get-AdvisingUniques {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        
    }
    
    process {
        $advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique

    }
    
    end {
        return $advisors
    }
}