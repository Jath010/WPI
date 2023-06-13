# Add the service principal application ID and secret here
$servicePrincipalClientId="5c3e8745-eaeb-4c88-b639-456ed660e3f0"
$servicePrincipalSecret="uql8Q~mLxEUq_XdCkDWq.5mmfhZeb9oJV9u3UdpR"

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Download the installation package
Invoke-WebRequest -Uri "https://aka.ms/azcmagent-windows" -TimeoutSec 30 -OutFile "$env:TEMP\install_windows_azcmagent.ps1"

# Install the hybrid agent
& "$env:TEMP\install_windows_azcmagent.ps1"
if($LASTEXITCODE -ne 0) {
    throw "Failed to install the hybrid agent"
}

# Run connect command
& "$env:ProgramW6432\AzureConnectedMachineAgent\azcmagent.exe" connect --service-principal-id "$servicePrincipalClientId" --service-principal-secret "$servicePrincipalSecret" --resource-group "AzureArc" --tenant-id "589c76f5-ca15-41f9-884b-55ec15a0672a" --location "eastus" --subscription-id "ff9f05f1-f9f4-4f11-b0e1-948bd9561f17" --cloud "AzureCloud" --tags "Datacenter=OnPrem,City=Worcester,StateOrDistrict=Mass,CountryOrRegion=US" --correlation-id "dcdfa329-4204-427d-a427-fc6a7b29f030"

if($LastExitCode -eq 0){Write-Host -ForegroundColor yellow "To view your onboarded server(s), navigate to https://portal.azure.com/#blade/HubsExtension/BrowseResource/resourceType/Microsoft.HybridCompute%2Fmachines"}
