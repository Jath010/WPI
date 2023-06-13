Clear-Host
#Domain Controller Information
$global:dcs = (Get-ADDomainController -Filter *)
$global:dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
$global:dchostname = $dc.HostName

#Environment Variables
$global:today = Get-Date
$global:currentuser = $env:username
$global:localhost = $env:COMPUTERNAME


function WinSam-Get-ObjectInfo ($DisplayName) {
    [hashtable]$Return = @{}

    $return.Name = $null
    $return.Username = $null
    $return.Department = $null
    $return.Enabled = $null

    if ($DisplayName) {
        if ($DisplayName -eq 'Default' -or $DisplayName -eq 'Anonymous') {$return.Name = $DisplayName;$return.Enabled = $true}
        else {
            if ($DisplayName -match "NT User:") {$username = $DisplayName.Replace("NT User:ADMIN\",""); $ObjectType = 'User'}
            else {
                if (Get-Mailbox $DisplayName -ErrorAction SilentlyContinue) {$username = (Get-Mailbox $DisplayName -DomainController $dchostname -ErrorAction SilentlyContinue).SamAccountName; $ObjectType = 'User'}
                elseif (Get-DistributionGroup $DisplayName -ErrorAction SilentlyContinue) {$username = (Get-DistributionGroup $DisplayName -DomainController $dchostname -ErrorAction SilentlyContinue).SamAccountName; $ObjectType = 'Group'}
                }
            if ($username) {
                if ($ObjectType -eq 'User') {
                    $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname
                    if ($ADInfo) {
                        $return.Name = $ADInfo.Name
                        $return.Username = $ADInfo.SamAccountName
                        $return.Department = $ADInfo.Department
                        $return.Enabled = $ADInfo.Enabled
                        }
                    }
                if ($ObjectType -eq 'Group') {
                    $GroupInfo = Get-ADGroup $username -Server $dchostname
                    if ($GroupInfo) {
                        $return.Name = $GroupInfo.Name
                        $return.Username = $GroupInfo.SamAccountName
                        $return.Department = 'Distribution Group'
                        $return.Enabled = $true
                        }
                    }
                }
            else {$return.Name = $DisplayName;$return.Enabled = $true}
            }
        }
    else {Write-Host "Entry does not exist. Please report error to Tom Collins (tcollins@wpi.edu)." -ForegroundColor Red;$return.Name = 'ERROR';$return.Enabled = $true}
    Return $return
    }

WinSam-Get-ObjectInfo 'LaMotte, Jeanne K.'

Write-Host ''

WinSam-Get-ObjectInfo 'University Marketing'
