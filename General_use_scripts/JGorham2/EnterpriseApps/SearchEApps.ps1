

$EntApps = Get-AzureADServicePrincipal -All $true | Where-Object {($_.Tags -contains "WindowsAzureActiveDirectoryGalleryApplicationNonPrimaryV1") -or ($_.Tags -contains "WindowsAzureActiveDirectoryCustomSingleSignOnApplication")} | Select-Object DisplayName,ObjectID

foreach ($app in $EntApps) {
    <# $app is the current item #>
    $groupAssignments = Get-AzureADServiceAppRoleAssignment -ObjectId $EntApps[136].ObjectId
    foreach ($group in $groupAssignments) {
        <# $group is the current item #>
        if ($group.PrincipalDisplayName) {
            currentItemName
        }
    }
}