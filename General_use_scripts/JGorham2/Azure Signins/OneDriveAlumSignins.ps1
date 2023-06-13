#Onedrive appid ab9b8c07-8f02-4f72-87fa-80105867a763

if (!((Get-Module).name -contains "azureadpreview")) {
    if (!((Get-installedModule).name -contains "azureadpreview")) {
        Install-Module azureadpreview
    }
    import-module azureadpreview -force
}

$filter = "appID eq 'ab9b8c07-8f02-4f72-87fa-80105867a763'"
$alumniList = get-aduser -filter "extensionattribute7 -eq 'alum'" -Properties extensionattribute7 | select-object samaccountname

$hits = Get-AzureADAuditSignInLogs -Filter $filter

$alumHits = $hits | where-object {$alumniList.samaccountname -contains $_.UserPrincipalName.split("@")[0] }

$alumHits | Select-object UserPrincipalName, CreatedDateTime