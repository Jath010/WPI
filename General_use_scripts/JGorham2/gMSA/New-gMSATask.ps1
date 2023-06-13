
function Install-WPIgMSA {
    [CmdletBinding()]
    param (
        $gMSA,
        $targetMachine
    )
    
    begin {
        try {
            $Samaccountname = (Get-ADServiceAccount $gMSA).Samaccountname
        }
        catch {
            Write-Host "Could not find gMSA $gMSA" -ForegroundColor Red
            return
        }
    }
    #Install-WindowsFeature -Name “RSAT-AD-PowerShell” -IncludeAllSubFeature
    process {
        if (!($targetMachine -ne $env:COMPUTERNAME)) {
            if (!(Get-WindowsFeature -name "RSAT-AD-Powershell").installed) {
                Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature
            }
            Install-ADServiceAccount $Samaccountname
            if (Test-AdServiceAccount $Samaccountname) {
                Write-Host "gMSA has been successfully installed" -ForegroundColor Green
            }
            else {
                Write-Host "gMSA was NOT successfully installed" -ForegroundColor Red
            }
            
        }
        else {
            <#  
            # It looks like to do this you would need to use credSSP, not sure if it's worth not logging into the server
            ## Enable-WSManCredSSP -role Client -DelegateComputer $using:service_server
            ## Invoke-Command -cn $service_server -Credential $user  -Authentication Credssp {Install-ADServiceAccount -identity $using:service_acct1 -Verbose}
            Write-Verbose "Creating Session"
            $session = New-PSSession -computerName $targetMachine
            Invoke-Command -Session $session -scriptblock {
                if (!(Get-WindowsFeature -name "RSAT-AD-Powershell").installed) {
                    Write-verbose "AD tools not present: Installing"
                    Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature
                }
                Write-Verbose "Installing gMSA Account"
                Install-ADServiceAccount $Using:Samaccountname
                if (Test-AdServiceAccount $Using:Samaccountname) {
                    Write-Host "gMSA has been successfully installed" -ForegroundColor Green
                }
                else {
                    Write-Host "gMSA was NOT successfully installed" -ForegroundColor Red
                }
            }
            #>
            Write-Host "You need to run this locally" -ForegroundColor Red
        }
    }
    
    end {
        
    }
}

function New-gMSATask {
    [CmdletBinding()]
    param (
        $TaskName,
        $gMSA,
        $Action,
        $trigger
    )
    <#
    Action looks like 
        $action = New-ScheduledTaskAction  "c:\scripts\backup.cmd"
    Trigger looks like
        $trigger = New-ScheduledTaskTrigger -At 23:00 -Daily
    #>
    begin {
        try {
            $Samaccountname = (Get-ADServiceAccount -Filter $gMSA).$Samaccountname
            $UserID = "Admin\" + $Samaccountname
        }
        catch {
            Write-Host "Could not find gMSA $gMSA" -ForegroundColor Red
            Return
        }
        if (!(Test-AdServiceAccount $Samaccountname)) {
            Write-Host "gMSA $samaccountname is not installed on this computer" -ForegroundColor Red
            return
        }
    }
    
    process {
        $principal = New-ScheduledTaskPrincipal -UserID $UserID -LogonType Password
        Register-ScheduledTask $TaskName -Action $Action -Trigger $trigger -Principal $principal
    }
    
    end {
        
    }
}

function New-PowershellScriptTaskAction {
    [CmdletBinding()]
    param (
        $ScriptPath,
        $Switch
    )
    
    begin {
        if (!(Test-Path $ScriptPath)) {
            Write-Host "Path seems to be incorrect" -ForegroundColor Red
            Return
        }
        if ($null -eq $args) {
            $argument = "-NoProfile -WindowStyle Hidden -File " + $ScriptPath + '"'
        }
        else {
            $argument = "-NoProfile -WindowStyle Hidden -File " + $ScriptPath + " " + $Switch + '"'
        }
    }
    
    process {
        New-ScheduledTaskAction  -Execute "Powershell.exe" -Argument $argument
    }
    end {
        
    }
}

<#
$tenmins = (Get-Date).AddMinutes(10)
$Action = New-ScheduledTaskAction -Execute "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe" -Argument "/applyUpdates -autoSuspendBitLocker=enable -silent -reboot=enable"
$Trigger = New-ScheduledTaskTrigger -Once -At $tenmins
$Settings = New-ScheduledTaskSettingsSet
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
$user = "NT AUTHORITY\SYSTEM"
Register-ScheduledTask -TaskName 'Initial DCU Check' -InputObject $Task -User $user
#>