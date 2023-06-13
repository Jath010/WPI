function Add-OptGroups {
    [CmdletBinding()]
    param (
        
    )
    
    # This function adds all the opt groups in the helpdesk catalog to an access package for assignment later

    begin {
        $accesspackageID = "25e278e5-1c75-439a-bbcb-ab48ed36dc99"
        #$HelpdeskManagementID = "fc1d3f6d-8854-4436-abf2-4db5d16e25ac"
        $optGroups = Get-MgEntitlementManagementAccessPackageCatalogAccessPackageResource -AccessPackageCatalogId "50629f7f-f029-4922-9161-221915e00101"|Where-Object {$_.displayname -like "Opt*"}

    }
    
    process {
        foreach ($Group in $optGroups) {
            <# $Group is the current item #>
            $accessPackageResource = @{
                "id"           = $group.id
                "resourceType" = 'Security Group'
                "originId"     = $group.originId
                "originSystem" = 'AadGroup'
            }
    
            $accessPackageResourceRole = @{
                "originId"              = "Owner_"+$group.originId  #Change Owner to Member here for switching.
                "displayName"           = 'Owner'
                "originSystem"          = 'AadGroup'
                "accessPackageResource" = $accessPackageResource
            }
    
            $accessPackageResourceScope = @{
                "originId"     = $group.originId
                "originSystem" = 'AadGroup'
            }
    
            New-MgEntitlementManagementAccessPackageResourceRoleScope -AccessPackageId $accesspackageID -AccessPackageResourceRole $accessPackageResourceRole -AccessPackageResourceScope $accessPackageResourceScope | Format-List
        }
        
    }
    
    end {
        
    }
}


function Add-GroupToAccessPackage {
    [CmdletBinding()]
    param (
        $AccessPackage,
        $Group
    )
    
    # This function adds all the opt groups in the helpdesk catalog to an access package for assignment later

    begin {
        $accesspackageID = (Get-MgEntitlementManagementAccessPackageCatalog -filter "DisplayName eq $AccessPackage").id
        #$HelpdeskManagementID = "fc1d3f6d-8854-4436-abf2-4db5d16e25ac"
        $Group = get-mggroup -filter "DisplayName eq $group"

    }
    
    process {
        
            <# $Group is the current item #>
            $accessPackageResource = @{
                "id"           = $group.id
                "resourceType" = 'Security Group'
                "originId"     = $group.originId
                "originSystem" = 'AadGroup'
            }
    
            $accessPackageResourceRole = @{
                "originId"              = "Owner_"+$group.originId  #Change Owner to Member here for switching.
                "displayName"           = 'Owner'
                "originSystem"          = 'AadGroup'
                "accessPackageResource" = $accessPackageResource
            }
    
            $accessPackageResourceScope = @{
                "originId"     = $group.originId
                "originSystem" = 'AadGroup'
            }
    
            New-MgEntitlementManagementAccessPackageResourceRoleScope -AccessPackageId $accesspackageID -AccessPackageResourceRole $accessPackageResourceRole -AccessPackageResourceScope $accessPackageResourceScope | Format-List
        
        
    }
    
    end {
        
    }
}