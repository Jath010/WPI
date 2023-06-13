Clear-Host

$PathDevelopment = "\\drstorage\systools\Devel\WinSam - Powershell"

$PathProduction = "\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3"
$PathBeta = "\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3B"
$PathDev = "\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3C"

#Archive working copy to backup location
<#
Copy-Item "$PathProduction\WinSam.3.ps1" "$PathProduction\Old\WinSam.3.ps1"
Copy-Item "$PathProduction\WinSam.3.Functions.ps1" "$PathProduction\Old\WinSam.3.Functions.ps1"
Copy-Item "$PathProduction\WinSam.3.Info.ps1" "$PathProduction\Old\WinSam.3.Info.ps1"
Copy-Item "$PathProduction\WinSam.3.Menu.ps1" "$PathProduction\Old\WinSam.3.Menu.ps1"
Copy-Item "$PathProduction\WinSam.3.GlobalVariables.ps1" "$PathProduction\Old\WinSam.3.GlobalVariables.ps1"
Write-Host "Archive of Current Production Code Complete" -ForegroundColor Yellow
#>

#Copy new version to production location
<#
Copy-Item "$PathDevelopment\WinSam.3.ps1" "$PathProduction\WinSam.3.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Functions.ps1" "$PathProduction\WinSam.3.Functions.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Info.ps1" "$PathProduction\WinSam.3.Info.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Menu.ps1" "$PathProduction\WinSam.3.Menu.ps1"
Copy-Item "$PathDevelopment\WinSam.3.GlobalVariables.ps1" "$PathProduction\WinSam.3.GlobalVariables.ps1"
Write-Host "Update to Production Code Complete" -ForegroundColor Green
#>

#Copy new version to Beta location
<#
Copy-Item "$PathDevelopment\WinSam.3.ps1" "$PathBeta\WinSam.3.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Functions.ps1" "$PathBeta\WinSam.3.Functions.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Info.ps1" "$PathBeta\WinSam.3.Info.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Menu.ps1" "$PathBeta\WinSam.3.Menu.ps1"
Copy-Item "$PathDevelopment\WinSam.3.GlobalVariables.ps1" "$PathBeta\WinSam.3.GlobalVariables.ps1"
Write-Host "Update to Beta Code Complete" -ForegroundColor Yellow
#>

#Copy new version to development location
#<#
Copy-Item "$PathDevelopment\WinSam.3.ps1" "$PathDev\WinSam.3.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Functions.ps1" "$PathDev\WinSam.3.Functions.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Info.ps1" "$PathDev\WinSam.3.Info.ps1"
Copy-Item "$PathDevelopment\WinSam.3.Menu.ps1" "$PathDev\WinSam.3.Menu.ps1"
Copy-Item "$PathDevelopment\WinSam.3.GlobalVariables.ps1" "$PathDev\WinSam.3.GlobalVariables.ps1"
Write-Host "Update to Dev Code Complete" -ForegroundColor Yellow
#>

Get-Date