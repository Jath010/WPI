<#

    Idea is to create a script that goes through all the users in a department group and collects their group membership, then sorts for uniques

#>

function Get-DepartmentMemberships {
    [CmdletBinding()]
    param (
        $DepartmentGroup
    )
    
    begin {
        $OutputPath = "D:\tmp\Entitlements\CurrentDepartmentMemberships"+"\$departmentGroup.csv"
        $groupMembers = Get-AzureADGroupMember -ObjectId (get-azureadgroup -SearchString $DepartmentGroup).ObjectId
        $GroupMemberships = @()
    }
    
    process {
        foreach ($Member in $groupMembers) {
            <# $Member is the current item #>
            $GroupMemberships += Get-AzureADUserMembership -ObjectId $Member.ObjectId
        }
        $GroupMemberships = $GroupMemberships | Sort-Object -Property DisplayName -Unique
    }
    
    end {
        $GroupMemberships | Export-Csv -Path $OutputPath -NoTypeInformation
    }
}

function Get-AllDepartmentMemberships {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $DepartmentGroups = Get-AzureADGroupMember -ObjectId "47861a6c-c7a6-4c29-8f64-3c0b8c721ba1"
    }
    
    process {
        foreach ($Department in $DepartmentGroups) {
            <# $Department is the current item #>
            Get-DepartmentMemberships -DepartmentGroup $Department.DisplayName
        }
    }
    
    end {
        
    }
}