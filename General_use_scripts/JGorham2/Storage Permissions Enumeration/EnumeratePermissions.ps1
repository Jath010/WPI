<#

Given some arbitrary location produce a list of all the users who can access the location


#>

function Get-StorageAccessPermissions {
    [CmdletBinding()]
    param (
        $path = (get-location).path
    )
    
    begin {
        $path = get-item $path
        $ACLs = @{}
    }
    
    process {
        get-acl $path | ForEach-Object {$_.access} #gets the ACLs of the directory
        get-acl $path.parent.fullname | ForEach-Object {$_.access} #gets the acls of the parent, need to sort for the ones that matter
    
        get-childitem -Recurse $path | Where-Object { $.PsIsContainer } | ForEach-Object { $path1 = $.fullname; Get-Acl $.Fullname | ForEach-Object { $.access | Add-Member -MemberType NoteProperty '.\Application Data' -Value $path1 -passthru }}
    }
    
    end {
        
    }
}