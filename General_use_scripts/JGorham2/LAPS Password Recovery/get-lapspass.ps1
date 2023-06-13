#PS Microsoft.PowerShell.Core\FileSystem::\\scripthost-02\d$\wpi\powershell\workstations> cat .\ancient-log.txt | select-string me-nb04

function Get-LAPSExpired {
    [CmdletBinding()]
    param (
        $ComputerName
    )
    
    begin {
        $targetComputer = "scripthost-02"
    }
    
    process {
        $results = invoke-command -ComputerName $targetcomputer -scriptblock {
            $file = get-content D:\wpi\powershell\workstations\ancient-log.txt
            foreach($computer in $Using:ComputerName){
                ($file | select-string $computer | Sort-Object -Unique | Out-String).trim()
            }
        }
    }
    
    end {
        $results
    }
}

function Get-LAPSPassword {
    [CmdletBinding()]
    param(
        $ComputerName
    )

    (get-adcomputer $computername -Properties ms-Mcs-admPwd)."ms-Mcs-admPwd"
}