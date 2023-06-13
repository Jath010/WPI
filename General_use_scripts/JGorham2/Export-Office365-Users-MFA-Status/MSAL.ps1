
Function DelegatedAuthN {
    <#
    .SYNOPSIS
Authenticate to Azure AD (using Delegated Auth) and receieve Access and Refresh Tokens.
.DESCRIPTION
Authenticate to Azure AD (using Delegated Auth) and receieve Access and Refresh Tokens.
.PARAMETER tenantID
(required) Azure AD TenantID.
.PARAMETER clientID
(required) ClientID of the Azure AD registered application with the necessary permissions.
.EXAMPLE
AuthN -tenantID '74ea519d-9792-4aa9-86d9-abcdefgaaa' -clientID '1122334c-bad7e-4eef-8a06-1234567890'
.LINK
http://darrenjrobinson.com/
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$tenantID,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$clientID,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
        [string]$redirectUri = "https://localhost"
    )

    if (!(get-command Get-MsalToken)) {
        Install-Module -name MSAL.PS -Force -AcceptLicense
    }

    try {
        # Authenticate and Get Tokens
        $token = Get-MsalToken -ClientId $clientID -TenantId $tenantID -RedirectUri $redirectUri -Authority "https://login.microsoftonline.com/$($tenantID)" -Silent 
        return $token 
    }
    catch {
        $_
    }
}

Function GetAADUsersAuthRegisteredByMethod {
    <#
    .SYNOPSIS
Get AAD Users Authentication Registered by Method.
.DESCRIPTION
Get AAD Users Authentication Registered by Method.
.EXAMPLE
GetAADUsersAuthRegisteredByMethod 
.LINK
http://darrenjrobinson.com/
    #>
    [cmdletbinding()]
    param()

    # Refresh Access Token
    $global:myToken = DelegatedAuthN -tenantID $global:myTenantId -clientID $global:myClientID 

    try {
        $aadOrgUsersRegisteredByMethod = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
            -Uri  "https://graph.microsoft.com/beta/reports/authenticationMethods/usersRegisteredByMethod" `
            -Method Get

        return $aadOrgUsersRegisteredByMethod
    }
    catch {
        $_
    }
}

Function GetAADUsersAuthRegisteredByFeature {
    <#
    .SYNOPSIS
Get AAD Users Authentication Registered by Feature.
.DESCRIPTION
Get AAD Users Authentication Registered by Feature.
.EXAMPLE
GetAADUsersAuthRegisteredByFeature 
.LINK
http://darrenjrobinson.com/
    #>
    [cmdletbinding()]
    param()

    # Refresh Access Token
    $global:myToken = DelegatedAuthN -tenantID $global:myTenantId -clientID $global:myClientID 

    try {
        $aadOrgUsersRegisteredByFeature = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
            -Uri  "https://graph.microsoft.com/beta/reports/authenticationMethods/usersRegisteredByFeature" `
            -Method Get

        return $aadOrgUsersRegisteredByFeature
    }
    catch {
        $_
    }
}

Function GetAADUsersCredentialUserRegistrationCount {
    <#
    .SYNOPSIS
Get AAD Users Credential User Registration Count Summary.
.DESCRIPTION
Get AAD Users Credential User Registration Count Summary.
.EXAMPLE
GetAADUsersCredentialUserRegistrationCount 
.LINK
http://darrenjrobinson.com/
    #>
    [cmdletbinding()]
    param()

    # Refresh Access Token
    $global:myToken = DelegatedAuthN -tenantID $global:myTenantId -clientID $global:myClientID 

    try {
        $aadOrgUsersUserRegistrationCount = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
            -Uri  "https://graph.microsoft.com/beta/reports/getCredentialUserRegistrationCount" `
            -Method Get

        return $aadOrgUsersUserRegistrationCount.value 
    }
    catch {
        $_
    }
}

Function GetAADUsersCredentialUsageSummary {
    <#
    .SYNOPSIS
Get AAD Users Credential Usage Summary.
.DESCRIPTION
Get AAD Users Credential Usage Summary.
.PARAMETER period
(optional) Period for Summary Report. Valid options are 1, 7 and 30 (days)
.EXAMPLE
GetAADUsersCredentialUsageSummary -period '7'
.LINK
http://darrenjrobinson.com/
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateSet("1", "7", "30")]
        [string]$period
    )

    # Refresh Access Token
    $global:myToken = DelegatedAuthN -tenantID $global:myTenantId -clientID $global:myClientID 

    try {
        $aadOrgUsersCredentialUsageSummary = Invoke-RestMethod -Headers @{Authorization = "Bearer $($myToken.AccessToken)" } `
            -Uri  "https://graph.microsoft.com/beta/reports/getCredentialUsageSummary(period='D$($period)')" `
            -Method Get

        return $aadOrgUsersCredentialUsageSummary.value 
    }
    catch {
        $_
    }
}

# Globals
# Tenant ID 
$global:myTenantId = '<Your AAD Tenant ID>'
# Registered AAD App ID
$global:myClientID = '<Your AAD Registered App Client ID>'

# One Time only to authorize access and get an Access and Refresh Token into the local MSAL Cache
# Comment out after the one time use (per user profile on the local host running the script)
$global:myToken = Get-MsalToken -DeviceCode -ClientId $myClientID -TenantId $myTenantId -RedirectUri "https://localhost" -Authority "https://login.microsoftonline.com/$($global:myTenantId)"

$AuthRegisteredByMethod = GetAADUsersAuthRegisteredByMethod 
$AuthRegisteredByMethod.userRegistrationMethodCounts

$AuthRegisteredByFeature = GetAADUsersAuthRegisteredByFeature 
$AuthRegisteredByFeature
$AuthRegisteredByFeature.userRegistrationFeatureCounts

$CredentialUserRegistrationCount = GetAADUsersCredentialUserRegistrationCount 
$CredentialUserRegistrationCount 
$CredentialUserRegistrationCount.userRegistrationCounts 

$CredentialUsageSummary = GetAADUsersCredentialUsageSummary -period '30' 
$CredentialUsageSummary