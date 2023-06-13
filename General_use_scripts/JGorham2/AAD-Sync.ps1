function Start-AADSyncRemote {
    [CmdletBinding()]
    param (
        [switch]$Initial
    )
    
    begin {
        $AADC = New-PSSession aadc-utl-p-w03
    }
    
    process {
        Invoke-Command -Session $AADC -ScriptBlock { Import-module adsync }
        if ($Initial) {
            Invoke-Command -Session $AADC -ScriptBlock { start-adsyncsynccycle -policytype Initial }
        }
        else {
            Invoke-Command -Session $AADC -ScriptBlock { start-adsyncsynccycle -policytype Delta }
        }
    }
    
    end {
        Remove-PSSession $AADC
    }
}

function Get-AADSyncRemoteStatus {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $AADC = New-PSSession aadc-utl-p-w03
    }
    
    process {
        Invoke-Command -Session $AADC -ScriptBlock { Import-module adsync }
        Invoke-Command -Session $AADC -ScriptBlock { Get-ADSyncScheduler }
    }
    
    end {
        Remove-PSSession $AADC
    }
}