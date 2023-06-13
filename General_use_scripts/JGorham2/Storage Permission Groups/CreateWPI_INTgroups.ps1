# Given a path to a workday integrations top level directory, create groups for each of the "WPI_INT" directories, create the permissions, assign them, add traversal.

function Set-WPI_INTaccess {
    [CmdletBinding()]
    param (
        $path
    )
    
    begin {
        $searchbase = "OU=Workday Integrations,OU=Isilon Share Access,OU=Groups,DC=admin,DC=wpi,DC=edu"
        $directories = get-childitem -Path $path -Directory | where-object { $_.name -like "WPI_INT*" }

        $regex = "(WPI_INT\d*)\D.*"
    }
    
    process {
        #in order to spend less time waiting later, create all the groups immediately
        foreach ($directory in $directories) {
            <# $directory is the current item #>
            $directory.name -match $regex | Out-Null
            $groupname = "wdi_" + $matches[1]

            #Parent to get the correct OU
            $ParentOU = "OU=$($directory.parent.name)," + $searchbase

            Write-Host "Searching for existing group $groupname"

            if ($null -eq (get-adgroup -filter 'Name -eq $groupname')) { 
                #create new access group
                try {
                    new-adgroup -name $groupname -path $ParentOU -groupscope DomainLocal -GroupCategory Security
                    Write-Host "Group Created"
                }
                catch {
                    Write-Host "Failed to create group $groupname"
                }
            }
        }


        foreach ($directory in $directories) {
            <# $directory is the current item #>
            $directory.name -match $regex | Out-Null
            $groupname = "wdi_" + $matches[1]

            #Parent to get the correct OU
            $ParentOU = "OU=$($directory.parent.name)," + $searchbase

            Write-Host "Getting group $groupname"
            #get group to work with
            $newGroup = get-adgroup $groupname
            #get travesal group
            $traversalGroup = get-adgroup -filter "Name -like 'wdi_traversal_*'" -SearchBase $ParentOU

            if ((Get-ADGroupMember $traversalGroup).name -notcontains (get-adgroup -filter 'Name -eq $groupname').name) { 
                #add new group to traversal
                try {
                    Add-ADGroupMember -Identity $traversalGroup.name -Members $newGroup.Name
                }
                catch {
                    Write-Host "Failed to add group $groupname to it's traversal group $($traversalGroup.name)"
                }

                #Now to configure and set acls
                $newACL = get-acl -path $directory.fullname
                #Set Properties
                $Identity = "ADMIN\" + $NewGroup.Name
                $fileSystemRights = "Modify, Synchronize"
                $InheritanceFlags = "ContainerInherit, ObjectInherit"
                $PropagationFlags = "None"
                $type = "Allow"

                #Create new rule
                $fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $InheritanceFlags, $PropagationFlags, $type
                $fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
                # Apply new rule
                while (1) {
                    Try {
                        $NewAcl.SetAccessRule($fileSystemAccessRule)
                        Write-Host "ACL successfully set"
                        Break
                    }
                    catch {
                        Write-Host "Waiting for group to propagate"
                        Start-Sleep -seconds 5
                    }
                }
                Set-Acl -Path $directory.fullname -AclObject $NewAcl
            }
            else {
                Write-Host "Group $groupName already exists."
            }
        }
    }
    
    end {
        
    }
}
function Set-AllIntegrationAccess {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $path = "\\storage\dept\Workday Integrations"
        $directories = Get-ChildItem -Path $path -Directory
    }
    
    process {
        foreach($directory in $directories){
            Set-WPI_INTaccess -path $directory.FullName
        }
    }
    
    end {
        
    }
}