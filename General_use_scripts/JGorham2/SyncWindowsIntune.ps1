###Magic internet code
$IntuneModule = Get-Module -Name "Microsoft.Graph.Intune" -ListAvailable
if (!$IntuneModule){
 
write-host "Microsoft.Graph.Intune Powershell module not installed..." -f Red
write-host "Install by running 'Install-Module Microsoft.Graph.Intune' from an elevated PowerShell prompt" -f Yellow
write-host "Script can't continue..." -f Red
write-host
exit
}
####################################################
# Importing the SDK Module
Import-Module -Name Microsoft.Graph.Intune
 
if(!(Connect-MSGraph)){
Connect-MSGraph
}
####################################################
 
workflow Sync-WindowsMachines {
    param (
        $devices
    )

    foreach -parallel ($Device in $Devices) { 
        Invoke-IntuneManagedDeviceSyncDevice -managedDeviceId $Device.managedDeviceId
    }
}

###########################
#   Main
###########################

#### Gets all devices running Windows
$Devices = Get-IntuneManagedDevice -Filter "contains(operatingsystem,'Windows')" | Get-MSGraphAllPages

Write-Host "Device List Gathered" -ForegroundColor Yellow

#Sync-WindowsMachines $devices

Foreach ($Device in $Devices)
{
 
Invoke-IntuneManagedDeviceSyncDevice -managedDeviceId $Device.managedDeviceId
Write-Host "Sending Sync request to Device with DeviceID $($Device.managedDeviceId)" -ForegroundColor Yellow
 
}

####################################################