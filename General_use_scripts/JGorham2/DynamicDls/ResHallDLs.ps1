<#

Creating the script for Resedent Hall lists

Copied from DynamicList code
So to start
we need to detect if the groups exist for each one we want to create
if they don't exist we create them

so the list will be
    The DL
        The Dynamic Group containing all the users fitting a particular description
        The Opt-in group
        The Opt-out group
    
    so the only important aspect of the opt group creation is making sure they match the naming scheme for later searching

the logic will be to act on a list of descriptions, get the list, check existence, gather the lists of dyn and opt in, subtract opt out then sync the DL



#>

#Create a dynamic list
#needs to take a simple description 
function New-WPIDynamicGroup {
    [CmdletBinding()]
    param (
        $Name,
        $Rule
    )
    
    #     begin {
    #         switch ($ExtensionAttribute)
    #         {
    #             #Extension Attribute 1 is Employee Code
    #             1 { $Rule = "(user.extensionattribute1 -match ""$condition"")"}
    # #            2 { $Rule = "(user.extensionattribute2 -match ""$condition"")" }
    # #            3 { $Rule = "(user.extensionattribute3 -match ""$condition"")" }

    #             #Extension Attribute 4 is Major/minor
    #             4 { $Rule = "(user.extensionattribute4 -match ""$condition"")" }
    #             #Extension Attribute 5 is Dorm and Visa
    #             5 { $Rule = "(user.extensionattribute5 -match ""$condition"")" }
    #             #Extension Attribute 6 is Advisor
    #             6 { $Rule = "(user.extensionattribute6 -match ""$condition"")" }
    #             #Extension Attribute 7 is Affiliation
    #             7 { $Rule = "(user.extensionattribute7 -match ""$condition"")" }
            
    # #            8 { $Rule = "(user.extensionattribute8 -match ""$condition"")" }
    # #            9 { $Rule = "(user.extensionattribute9 -match ""$condition"")" }
    # #            10 { $Rule = "(user.extensionattribute10 -match ""$condition"")" }
    # #            11 { $Rule = "(user.extensionattribute11 -match ""$condition"")" }
    # #            12 { $Rule = "(user.extensionattribute12 -match ""$condition"")" }
    # #            13 { $Rule = "(user.extensionattribute13 -match ""$condition"")" }
    # #            14 { $Rule = "(user.extensionattribute14 -match ""$condition"")" }
    # #            15 { $Rule = "(user.extensionattribute15 -match ""$condition"")" }

    #             Default {}
    #         }

    #         $Rule = "(user.extensionAttribute${$ExtensionAttribute} -)"
    #     }
    
    process {
        New-AzureADMSGroup -DisplayName "Dyn-$name" -Description "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$name" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
    }
    
    end {
        
    }
}

function New-WPIOptInList {
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        New-AzureADGroup -DisplayName "OptIn-${Name}" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
    }
}
function New-WPIOptOutList {
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        New-AzureADGroup -DisplayName "OptOut-${Name}" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
    }
}
function New-WPIMailEnabledSyncGroup {
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        New-DistributionGroup -Name "DL-${Name}" -Alias "DL-${Name}" -Type Distribution -PrimarySmtpAddress "Dl-${Name}@wpi.edu"
    }
}


#Example New-WPIDynlistComponents -name Test-DynPScreation -Rule "(user.extensionattribute14 -match ""02"")"
function New-WPIDynlistComponents {
    [CmdletBinding()]
    param (
        $Name,
        $Rule
    )
    New-WPIOptInList $Name
    New-WPIOptOutList $Name
    New-WPIMailEnabledSyncGroup $Name
    New-WPIDynamicGroup -Name $Name -Rule $Rule
    
}


function Sync-WPIDynlist {
    [CmdletBinding()]
    param (
        $name
    )
    $OptIn = get-AzureADGroup -SearchString "OptIn-${Name}"
    $OptOut = get-AzureADGroup -SearchString "OptOut-${Name}"
    $DynList = get-DistributionGroup -Identity "DL-${Name}"
    $DynGroup = get-AzureADMSGroup -SearchString "Dyn-$name"

    $OptInMembers = Get-AzureADGroupMember -ObjectId $OptIn.ObjectId | Select-Object UserPrincipalName
    $OptOutMembers = Get-AzureADGroupMember -ObjectId $OptOut.ObjectId | Select-Object UserPrincipalName
    $DynGroupMembers = Get-AzureADGroupMember -ObjectId $DynGroup.Id | Select-Object UserPrincipalName

    $regex = '(?i)^(' + (($OptOutMembers | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$'

    $CorrectMembers = ($DynGroupMembers + $OptInMembers) -notmatch $regex

    $CurrentMembers = Get-DistributionGroupMember -Identity $DynList.DisplayName | Select-Object @{N = 'UserPrincipalName'; E = { $_.primarysmtpaddress } }

    if ($null -eq $CurrentMembers) {
        ForEach ($member in $CorrectMembers) {
            Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $member.UserPrincipalName
        }
    }
    else {
        # #reconcile lists 
        $comparisons = Compare-Object $CurrentMembers $CorrectMembers -Property UserPrincipalName
                    
        $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object userprincipalname
        $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object userprincipalname
            
        ForEach ($Removal in $RemoveMembers.userprincipalname) {                
            Remove-DistributionGroupMember -Identity $DynList.DisplayName -Member $removal -Confirm:$false
        }
                        
        ForEach ($Addition in $AddMembers.userprincipalname) {
            Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $Addition                                
        }
    }
}


function New-WPIAdvisingListComponents {
    [CmdletBinding()]
    param (
        $Name
    )
    New-WPIOptInList $Name
    New-WPIOptOutList $Name
    New-WPIMailEnabledSyncGroup $Name
    New-WPIDynamicGroup -Name $Name -Rule (user.extensionAttribute6 -match ".*${name}.*")
    
}

workflow New-InitialWPIAdvisingLists {
    param (
        
    )
    $advisors = Get-aduser -filter { Enabled -eq $true -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6 | Where-Object { $_.extensionAttribute6 -match "PADV-[^ ](.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";")[0].split("-")[1] } | Sort-Object | Get-Unique
    foreach -parallel ($Advisor in $advisors) {
        if ($null -eq (Get-AzureADGroup -SearchString "optin-${name}")) {
            New-WPIAdvisingListComponents $Advisor
        }
        else {
            Sync-WPIDynlist $Advisor
        }
    }
}