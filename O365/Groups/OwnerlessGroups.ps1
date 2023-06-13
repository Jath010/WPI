
function Get-OwnerlessGroups {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $AllOwnerlessGroups = Get-UnifiedGroup -Filter {ManagedBy -eq '$null'} -ResultSize Unlimited
    }
    
    process {
        
    }
    
    end {
        return $AllOwnerlessGroups
    }
}

#This is for grabbing groups that have owners, but they're disabled
function Get-OrphanGroups {
    [CmdletBinding()]
    param (
        $Allgroups = (Get-UnifiedGroup -Filter * -ResultSize Unlimited)
    )
    
    begin {
    $OrphanedGroups = New-Object System.Collections.ArrayList       
    }
    
    process {
        foreach($group in $AllGroups){
            $ActiveOwner = $false
            foreach($user in $group.ManagedBy){
                if((Get-Mailbox -Filter "Name -eq ""$($user)""").Exchangeuseraccountcontrol -eq "None"){
                    $ActiveOwner = $true
                    continue
                    Write-host "That break didn't work"
                }
            }
            if(!$ActiveOwner){
                Write-Host "$($group.PrimarySmtpAddress) has no active owners"
                $OrphanedGroups += $group
            }
        }
    }
    
    end {
        return $OrphanedGroups
    }
}