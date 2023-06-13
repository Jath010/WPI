
function New-DepartmentGroup {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        # Need AzureADPreview module for New-AzureADMSGroup to have dynamic list capabilities
        import-module azureadpreview -force
        $departments = get-aduser -filter "enabled -eq 'True'" -Properties extensionattribute8, department | Where-Object { $_.extensionattribute8 -match ".*Staff;.*" } | Select-Object department | sort-object -Property department -unique | Where-Object { $null -ne $_.department }

    }
    
    process {
        foreach ($department in $departments) {
            <# $department is the current item #>

            $name = $department.department
            $DisplayName = "Department_" + $name
            $Rule = "(user.department -eq `"$name`") and (user.accountEnabled -eq True) and (user.extensionAttribute8 -match `".*Staff;.*`")"
            $mailnickname = $DisplayName.replace(" ", "").replace(",","").replace("(","-").replace(")","")

            $filter = "DisplayName eq '"+$DisplayName+"'"
            if ($null -eq (Get-AzureADMSGroup -Filter "$filter")) {
                <# Action to perform if the condition is true #>
                try {
                    # Try to create a group with the correct name
                    # Example Line
                    # New-AzureADMSGroup -DisplayName "Dyn-$name" -Description "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$name" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
                    New-AzureADMSGroup -DisplayName $DisplayName -Description "Dynamic department group created from PS for entitlements" -MailNickname $mailnickname -MailEnabled $false -SecurityEnabled $true -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
                }
                catch {
                    <#Do this if a terminating exception happens#>
                    Write-Host "Failed to create $($department.department)"
                    Write-Warning $error[0]
                }
            }
        }
    }
    
    end {
        
    }
}