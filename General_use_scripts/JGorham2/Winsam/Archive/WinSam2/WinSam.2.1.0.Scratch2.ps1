Clear-Host
$dcs = (Get-ADDomainController -Filter *)
$dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
$dchostname = $dc.HostName
$today = Get-Date
$currentuser = $env:username
$localhost = $env:COMPUTERNAME
$username = "dmurphy"
#$username = read-host -Prompt 'Please enter a username'
$Global:ADInfo = Get-ADUser $username -Properties * -Server $global:dc.hostname

#$user = Get-ADUser $username -Server $global:dc.hostname
##################







##################
#$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
#Write-Host "Script complete.  Total time: "$ElapsedTime.Elapsed