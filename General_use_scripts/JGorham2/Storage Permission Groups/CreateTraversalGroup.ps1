function Set-TraversalGroup {
    [CmdletBinding()]
    param (
        $Group,
        $path
    )
    
    $newACL = get-acl -path $path

    #Set Properties
    $Identity = "ADMIN\$group"
    $fileSystemRights = "ReadAndExecute, Synchronize"
    $type = "Allow"

    #Create new rule
    $fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
    $fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
    # Apply new rule
    $NewAcl.SetAccessRule($fileSystemAccessRule)
    Set-Acl -Path $path -AclObject $NewAcl

}

function Set-WDITraversal{
    [CmdletBinding()]
    param (
        $path
    )

    $path = "\\storage.wpi.edu\dept\Workday Integrations"
    $directories = get-childitem -Path $path -Directory

    $searchbase = "OU=Workday Integrations,OU=Isilon Share Access,OU=Groups,DC=admin,DC=wpi,DC=edu"

    foreach ($directory in $directories) {
        $directoryPath = $directory.fullname
        try{
            $search = "OU=$($directory.name),"+$searchbase #needed to put this here for some reason
            $group = get-adgroup -filter "Name -like 'wdi_traversal_*'" -SearchBase $search
        }
        catch{
            "Skipping $($directory.name)"
        }
        Set-TraversalGroup -path $directoryPath -Group $group.name
    }
}