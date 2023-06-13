function Set-ExtAffiliationReplacement {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        if(!(get-module azureadpreview)){
            Import-Module azureadpreview -Force # Need to import version 2.0.2.149 to get the rules in get-azureadmsgroup
        }
        $dynamicLists = get-azureadmsgroup -SearchString "Dyn-" -All:$true | Where-Object {$_.MembershipRule -match '\(user\.extensionattribute7 -eq "Student"\)'}
    }
    
    process {
        foreach($list in $dynamicLists){
            #change the rule to (user.extensionattribute8 match "(.*;)*Student;.*")
            $newRule = $list.MembershipRule.replace('(user.extensionattribute7 -eq "student")','(user.extensionattribute8 -match "(.*;)*Student;.*")')
            Set-AzureADMSGroup -id $list.id -membershiprule $newRule
        }
    }
    
    end {
        
    }
}