
#Connection String for Global Reader on Scripthost-02

Connect-AzureAD -TenantId 589c76f5-ca15-41f9-884b-55ec15a0672a -ApplicationId 07925265-94a7-43b7-add7-861ce0c72597 -CertificateThumbprint 12D6B2D64A2531A32BD070340B57693C5342CE4F

<#
#Creation:
Service Principle ID: ec930a1e-28fb-4840-b405-68abb4c92e2a
# Login to Azure AD PowerShell With Admin Account
Connect-AzureAD 

# Create the self signed cert
$currentDate = Get-Date
$endDate = $currentDate.AddYears(1)
$notAfter = $endDate.AddYears(1)
$pwd = "<password>"
$thumb = (New-SelfSignedCertificate -CertStoreLocation cert:\localmachine\my -DnsName admin.wpi.edu -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter).Thumbprint
$pwd = ConvertTo-SecureString -String $pwd -Force -AsPlainText
Export-PfxCertificate -cert "cert:\localmachine\my\$thumb" -FilePath c:\temp\exch_autocert.pfx -Password $pwd

# Load the certificate
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate("C:\temp\examplecert.pfx", $pwd)
$keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())


# Create the Azure Active Directory Application
$application = New-AzureADApplication -DisplayName "Exchange_Automation Service" -IdentifierUris "https://exch_automation.wpi.edu"
New-AzureADApplicationKeyCredential -ObjectId $application.ObjectId -CustomKeyIdentifier "Exchange_Automation Service" -StartDate $currentDate -EndDate $endDate -Type AsymmetricX509Cert -Usage Verify -Value $keyValue

# Create the Service Principal and connect it to the Application
$sp=New-AzureADServicePrincipal -AppId $application.AppId

# Give the Service Principal Reader access to the current tenant (Get-AzureADDirectoryRole)
Add-AzureADDirectoryRoleMember -ObjectId 552f89f6-784c-495d-acdf-d1ff441d5e53 -RefObjectId $sp.ObjectId

# Get Tenant Detail
$tenant=Get-AzureADTenantDetail
# Now you can login to Azure PowerShell with your Service Principal and Certificate
Connect-AzureAD -TenantId $tenant.ObjectId -ApplicationId  $sp.AppId -CertificateThumbprint $thumb
#>