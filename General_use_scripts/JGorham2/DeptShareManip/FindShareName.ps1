<#

The point of this script is to be able to take the provided visible directory name and spit out the name of the associated AD group samaccountname

#>

function Get-ShareGroupID {
    [CmdletBinding()]
    param (
        $FolderName
    )
    
    begin {
        
    }
    
    process {
        if ($FolderName.StartsWith('fc_')) {
            try {
                $ShareGroupSam = (Get-ADGroup "l_$($FolderName)").samaccountname
            }
            catch {
                Write-Host "Folder name did not conform to standard l_fc_ format"
            }
        }else{
            #$DeptCode = $FolderName.split("_")[0]
            try {
                $ShareGroupSam = (Get-ADGroup "shares_$($FolderName)").samaccountname
            }
            catch {
                Write-Host "Folder name did not conform to standard shares_ format"
            }
        }
    }
    
    end {
        Return $ShareGroupSam
    }
}