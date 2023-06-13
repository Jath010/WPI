function Get-AdvisingListVerify {
    [CmdletBinding()]
    param (
        $List
    )
    
    begin {
        $ListID = (get-AzureADGroup -Filter "DisplayName eq '$List'").objectid
        $ListPop = Get-AzureADGroupMember -ObjectId $ListID -All $true
        $advisor = $list.split("-")[1]
    }
    
    process {
        foreach ($member in $ListPop) {
            $alias = $member.UserPrincipalName.Split("@")[0]
            $ext6 = (get-aduser $alias -Properties extensionattribute6).extensionattribute6
            if($ext6 -notmatch "PADV-(.*;)*$advisor;.*OADV-.*"){
                Write-Host "User $alias should not be in this list" -BackgroundColor Red -ForegroundColor Black
            }else{
                Write-Verbose "User $alias should be in this list"
            }
        }
    }
    
    end {
        
    }
}

function Get-AdvisingListVerifyHidden {
    [CmdletBinding()]
    param (
        $List
    )
    
    begin {
        $ListID = (get-AzureADGroup -Filter "DisplayName eq '$List'").objectid
        $ListPop = Get-AzureADGroupMember -ObjectId $ListID -All $true
    }
    
    process {
        foreach ($member in $ListPop) {
            $alias = $member.UserPrincipalName.Split("@")[0]
            if((get-aduser $alias -Properties msExchHideFromAddressLists).msExchHideFromAddressLists -eq $true){
                Write-Host "$alias is hidden" -BackgroundColor Red -ForegroundColor Black
            }
        }
    }
    
    end {
        
    }
}