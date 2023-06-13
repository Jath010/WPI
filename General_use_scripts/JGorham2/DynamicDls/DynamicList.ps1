<#
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
    Import-Module azureadpreview

    New-AzureADMSGroup -DisplayName "Dyn-$name" -Description "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$name" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"

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

function New-WPIMailEnabledADVSyncGroup {
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        New-DistributionGroup -Name "ADV-$Name" -Alias "ADV-$Name" -Type Distribution -PrimarySmtpAddress "ADV-$Name@wpi.edu"
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
    #-filter "DisplayName eq 'optin-${name}'"
    $OptIn = get-AzureADGroup -filter "DisplayName eq 'optin-${name}'"
    $OptOut = get-AzureADGroup -filter "DisplayName eq 'OptOut-${name}'"
    $DynList = get-DistributionGroup -Identity "DL-${Name}"
    $DynGroup = get-AzureADMSGroup -filter "DisplayName eq 'Dyn-${name}'"

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

function Sync-WPIADVlist {
    [CmdletBinding()]
    param (
        $name
    )

    $DynList = get-DistributionGroup -Identity "ADV-${Name}"
    $DynGroup = get-AzureADMSGroup -filter "DisplayName eq 'Dyn-${name}'"

    $DynGroupMembers = Get-AzureADGroupMember -ObjectId $DynGroup.Id | Select-Object UserPrincipalName

    $CorrectMembers = $DynGroupMembers

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
    New-WPIMailEnabledADVSyncGroup $Name
    New-WPIDynamicGroup -Name $Name -Rule '(user.extensionAttribute6 -match ".*$name.*") and (user.accountEnabled -eq true)'
    
}

function New-InitialWPIAdvisingLists {
    param (
        
    )
    $advisors = Get-aduser -filter { Enabled -eq $true -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6 | Where-Object { $_.extensionAttribute6 -match "PADV-[^ ](.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";")[0].split("-")[1] } | Sort-Object | Get-Unique
    foreach ($Advisor in $advisors) {
        $advisor = $advisor.trim()
        if ($null -eq (Get-AzureADGroup -filter "DisplayName eq 'Dyn-$Advisor'")) {
            Write-Host "$Advisor detected as missing"
            New-WPIAdvisingListComponents $Advisor
            Sync-WPIADVlist $Advisor
        }
        else {
            Sync-WPIADVlist $Advisor
        }
    }
}