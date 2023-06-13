#------------------------------------------------------------
# Function Get-AccessToken taken from https://blogs.technet.microsoft.com/cloudlojik/2018/06/29/connecting-to-microsoft-graph-with-a-native-app-using-powershell/
# (Paul Kotylo)
#
#
# 20200223
# Stephan Wälde
#
# Get-WHfBDeviceKeys-from-AzureAD extracts the Windows Hello for Business device keys from Azure AD for a given user
# Needs Tenant ID and User UPN as input
#------------------------------------------------------------


# Replace with own Tenant ID and UPN
$tenantId = "589c76f5-ca15-41f9-884b-55ec15a0672a"
$userUPN = "adbotelhofilho@wpi.edu"

# Getting the keys from Azure AD Graph API
$url = "https://graph.windows.net/$tenantId/users/" + $userUPN + "?api-version=1.6-internal&`$select=searchableDeviceKey"




Function Get-AccessToken ($TenantName, $ClientID, $redirectUri, $resourceAppIdURI, $CredPrompt){
    Write-Host "Checking for AzureAD module..."
    if (!$CredPrompt){$CredPrompt = 'Auto'}

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    if ($AadModule -eq $null) {$AadModule = Get-Module -Name "AzureADPreview" -ListAvailable}
    if ($AadModule -eq $null) {write-host "AzureAD Powershell module is not installed. The module can be installed by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt. Stopping." -f Yellow;exit}
    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
        $aadModule      = $AadModule | ? { $_.version -eq $Latest_Version.version }
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    else {
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    $authority          = "https://login.microsoftonline.com/$TenantName"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
    $authResult         = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters).Result
    return $authResult
    }






# First, let's authenticate
$clientId = "1b730954-1685-4b74-9bfd-dac224a7b894" # PowerShell
$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
$MSResourceURI = "https://graph.windows.net/"
$CredPrompt = "Always"

$AccessToken = Get-AccessToken -TenantName $tenantId -ClientID $clientId -redirectUri $redirectUri -resourceAppIdURI $MSResourceURI -CredPrompt $CredPrompt






# Next, let's get the data
$headers = @{
        'Content-Type'  = 'application\json'
        'Authorization' = $AccessToken.CreateAuthorizationHeader()
        }

$result = (Invoke-WebRequest -UseBasicParsing -Headers $headers -Uri $url)
$keylist = $result.Content | ConvertFrom-Json | select -expand searchableDeviceKey | where {$_.usage -eq "NGC"}
$keylist
