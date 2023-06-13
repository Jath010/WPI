
#So it's called as Add-CourseSharePermimssions ES ES3323 D21 ES_3323_Graders ReadAndExecute
function Add-CourseSharePermissions {
    [CmdletBinding()]
    param (
        $Course,
        $Section,
        $Term,
        $AddedObject,
        $Permission
    )
    
    begin {

        $path = "\\storage\academics\courses\$Course\$Section\$term\submissions"

        try {
            test-path $path | Out-Null
        }
        catch {
            Write-Host "$path was not a valid path, please check your inputs"
            exit
        }
        $folders = Get-ChildItem $path
    }
    
    process {
        foreach($folder in $folders){
            $acl = Get-Acl $folder.PSPath
            $newRule = "Admin\$addedObject",$permission,"ContainerInherit,ObjectInherit","None","Allow"
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $newRule
            $acl.AddAccessRule($AccessRule)
            $acl | Set-Acl $folder.PSPath

        }
    }
    
    end {
        
    }
}