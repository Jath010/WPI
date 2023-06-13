# Script to trigger a sync of the intune configuration.

function Sync-Intune {
    [CmdletBinding()]
    param (
        $Hostname
    )
    
    begin {
        $s = New-Pssession -ComputerName $Hostname
    }
    
    process {
        Invoke-Command -Session $s -ScriptBlock {
            Get-ScheduledTask | Where-Object {$_.TaskName -eq 'PushLaunch'} | Start-ScheduledTask
        }
    }
    
    end {
        Remove-PSSession -Session $s
    }
}