function Switch-GroupMembershipToHidden {
    [CmdletBinding()]
    param (
        $TargetGroup,
        $NewGroup
    )
    
    begin {
        $Owners = Get-UnifiedGroupLinks -Identity $TargetGroup -LinkType Owner
        $Members = Get-UnifiedGroupLinks -Identity $TargetGroup -LinkType Member
        
        $OldGr = get-unifiedgroup -identity $TargetGroup
        $NewGr = New-UnifiedGroup -displayName $NewGroup -HiddenGroupMembershipEnabled:$true -PrimarySmtpAddress $newGroup+"@wpi.edu"
        if ($oldGr.WelcomeMessageEnabled -eq $false) {
            set-unifiedgroup -Identity $Newgr.name -UnifiedGroupWelcomeMessageEnabled:$false
        }
    }
    
    process {
        foreach ($User in $Owners) {
            add-unifiedgrouplinks -identity $NewGr.Name -LinkType Owner -Links $User.Alias
        }
        foreach ($User in $Members) {
            add-unifiedgrouplinks -identity $NewGr.Name -LinkType Member -Links $User.Alias
        }
    }
    
    end {
        
    }
}