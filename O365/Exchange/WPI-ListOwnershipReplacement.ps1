# WPI-ListOwnershipReplacement
# Functions intended to handle the process of cleaning out a user from ownership of groups and distribution lists

# These get-s spit out a full object, it requires a little further handling to work with, but not enough that it can't be passed along painlessly
# This doesn't technically require a full email address as an input, but I know emails will always work
# Output can be used in a foreach, but you need to pull out the data you want otherwise it all gets spit out at the end of the loop
function Get-WPIDLManagedBy {
    param (
        $EmailAddress
    )
    # ManagedBy requires a full DN from Exchange to work with, the easiest way I could produce one is to just pull from it's identity first.
    $DN = (Get-Recipient -Identity "${EmailAddress}").DistinguishedName
    #So in the process of writing this I learned that you need to swap quote types as you go to include a variable in a filter, I think the variable needs to be enclosed in singles
    Get-DistributionGroup -Filter ("ManagedBy -eq '${DN}'")
}

function Get-WPIDLModeratedBy {
    param (
        $EmailAddress
    )
    # ManagedBy requires a full DN from Exchange to work with, the easiest way I could produce one is to just pull from it's identity first.
    $DN = (Get-Recipient -Identity "${EmailAddress}").DistinguishedName
    #So in the process of writing this I learned that you need to swap quote types as you go to include a variable in a filter, I think the variable needs to be enclosed in singles
    Get-DistributionGroup -Filter ("ModeratedBy -eq '${DN}'")
}

# This doesn't technically require a full email address as an input, but I know emails will always work
# Output can be used in a foreach, but you need to pull out the data you want otherwise it all gets spit out at the end of the loop
function Get-WPIGRManagedBy {
    param (
        $EmailAddress
    )
    $DN = (Get-Recipient -Identity "${EmailAddress}").DistinguishedName
    Get-UnifiedGroup -Filter ("ManagedBy -eq '${DN}'")
}


function Get-WPIGRModeratedBy {
    param (
        $EmailAddress
    )
    $DN = (Get-Recipient -Identity "${EmailAddress}").DistinguishedName
    Get-UnifiedGroup -Filter ("ModeratedBy -eq '${DN}'")
}

function Switch-WPIDLManagedBy {
    param (
        $TargetEmail,
        $ReplacementEmail,
        [switch]
        $WhatIf
    )
    $DLlists = Get-WPIDLManagedBy -EmailAddress $TargetEmail
    # If you don't specify .name here it won't iterate correctly
    foreach($List in $DLlists.name){
        Write-Verbose $List
        if($WhatIf -eq $true){
            Set-DistributionGroup -Identity $List -ManagedBy @{Add="$ReplacementEmail" ; Remove="$TargetEmail"} -WhatIf
        }
        else{
            Set-DistributionGroup -Identity $List -ManagedBy @{Add="$ReplacementEmail" ; Remove="$TargetEmail"}
        }
    }
}

function Switch-WPIDLModeratedBy {
    param (
        $TargetEmail,
        $ReplacementEmail,
        [switch]
        $WhatIf,
        [switch]
        $Duplicate
    )
    $DLlists = Get-WPIDLModeratedBy -EmailAddress $TargetEmail
    # If you don't specify .name here it won't iterate correctly
    foreach($List in $DLlists.name){
        Write-Verbose $List
        if($WhatIf -eq $true){
            if($Duplicate -eq $true){
                Set-DistributionGroup -Identity $List -ModeratedBy @{Add="$ReplacementEmail"} -WhatIf
            }
            else {
                Set-DistributionGroup -Identity $List -ModeratedBy @{Add="$ReplacementEmail" ; Remove="$TargetEmail"} -WhatIf
            }
        }
        else{
            if($Duplicate -eq $true){
                Set-DistributionGroup -Identity $List -ModeratedBy @{Add="$ReplacementEmail" ; Remove="$TargetEmail"}
            }
            else {
                Set-DistributionGroup -Identity $List -ModeratedBy @{Add="$ReplacementEmail" ; Remove="$TargetEmail"}
            }
        }
    }
}

function Remove-WPIDLManagedBy {
    param (
        $TargetEmail,
        [switch]
        $WhatIf
    )
    $DLlists = Get-WPIDLManagedBy -EmailAddress $TargetEmail
    foreach($List in $DLlists.name){
        Write-Verbose $List
        if($WhatIf -eq $true){
            Set-DistributionGroup -Identity $List -ManagedBy @{Remove="$TargetEmail"} -WhatIf
        }
        else{
            Set-DistributionGroup -Identity $List -ManagedBy @{Remove="$TargetEmail"}
        }
    }
}

function Switch-WPIGRManagedBy {
    [cmdletbinding()]
    param (
        $TargetEmail,
        $ReplacementEmail,
        [switch]
        $WhatIf
    )
    $GRlists = Get-WPIGRManagedBy -EmailAddress $TargetEmail
    foreach($Group in $GRlists.name){
        Write-Verbose $Group
        #I needed to throw the email address into quotes in order to have it work correctly, it apparently cast strangely when passed in
        $NewOwnerIsMember = (Get-UnifiedGroupLinks -Identity $Group -LinkType Members).Name -notcontains (Get-Recipient "${ReplacementEmail}").Name
        if($WhatIf -eq $true){
            Write-Verbose "Executing new owner membership check"
            if($NewOwnerIsMember){
                Write-Verbose "New Owner not a member of the Group: Adding"
                Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $ReplacementEmail -WhatIf
            }
            Add-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $ReplacementEmail -WhatIf
            Remove-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $TargetEmail -Confirm:$false -WhatIf
        }
        else{
            Write-Verbose "Executing new owner membership check"
            if($NewOwnerIsMember){
                Write-Verbose "New Owner not a member of the Group: Adding"
                Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $ReplacementEmail
            }
            Add-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $ReplacementEmail
            Remove-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $TargetEmail -Confirm:$false

        }
    }
}


function Switch-WPIGRModeratedBy {
    [cmdletbinding()]
    param (
        $TargetEmail,
        $ReplacementEmail,
        [switch]
        $WhatIf,
        [switch]
        $Duplicate
    )
    $GRlists = Get-WPIGRModeratedBy -EmailAddress $TargetEmail
    foreach($Group in $GRlists.name){
        Write-Verbose $Group
        #I needed to throw the email address into quotes in order to have it work correctly, it apparently cast strangely when passed in
        $NewOwnerIsMember = (Get-UnifiedGroupLinks -Identity $Group -LinkType Members).Name -notcontains (Get-Recipient "${ReplacementEmail}").Name
        if($WhatIf -eq $true){
            Write-Verbose "Executing new owner membership check"
            if($NewOwnerIsMember){
                Write-Verbose "New Owner not a member of the Group: Adding"
                Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $ReplacementEmail -WhatIf
            }
            Set-UnifiedGroup -Identity $Group -ModeratedBy @{add="$ReplacementEmail"} -WhatIf
            if($Duplicate -eq $false){
                Set-UnifiedGroup -Identity $Group -ModeratedBy @{remove="$TargetEmail"} -WhatIf
            }
        }
        else{
            Write-Verbose "Executing new owner membership check"
            if($NewOwnerIsMember){
                Write-Verbose "New Owner not a member of the Group: Adding"
                Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $ReplacementEmail
            }
            Set-UnifiedGroup -Identity $Group -ModeratedBy @{add="$ReplacementEmail"}
            if($Duplicate -eq $false){
                Set-UnifiedGroup -Identity $Group -ModeratedBy @{remove="$TargetEmail"}
            }

        }
    }
}


function Remove-WPIGRManagedBy {
    param (
        $TargetEmail,
        [switch]
        $WhatIf
    )
    $GRlists = Get-WPIGRManagedBy -EmailAddress $TargetEmail
    foreach($Group in $GRlists.name){
        Write-Verbose $Group
        if($WhatIf -eq $true){
            Remove-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $TargetEmail -Confirm:$false -WhatIf
        }
        else{
            Remove-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $TargetEmail -Confirm:$false
        }
    }
}

# Combined version for easy usage; If you feed it a single email it cleans that email out of the system. If you feed it two you replace them.
function Switch-WPIListOwnership {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string[]]
        $TargetEmail,
        [string[]]
        $ReplacementEmail,
        [switch]
        $WhatIf,
        [switch]
        $Duplicate
    )

    # If not fed a person to swap in the function will just remove them from everything they control
    if($null -eq $ReplacementEmail){
        if($WhatIf -eq $true){
            Write-Verbose "Executing in WhatIf mode"
            if($PSBoundParameters['Verbose']){
                Remove-WPIGRManagedBy -TargetEmail $TargetEmail -Verbose -WhatIf
                Remove-WPIDLManagedBy -TargetEmail $TargetEmail -Verbose -WhatIf
            }
            else{
                Remove-WPIGRManagedBy -TargetEmail $TargetEmail -WhatIf
                Remove-WPIDLManagedBy -TargetEmail $TargetEmail -WhatIf
            }
        }
        else{
            if($PSBoundParameters['Verbose']){
                Remove-WPIGRManagedBy -TargetEmail $TargetEmail -Verbose
                Remove-WPIDLManagedBy -TargetEmail $TargetEmail -Verbose
            }
            else{
                Remove-WPIGRManagedBy -TargetEmail $TargetEmail
                Remove-WPIDLManagedBy -TargetEmail $TargetEmail
            }
        }
    }

    # If given both targets the function will replace all instances of ownership
    elseif($Duplicate -eq $true){
        if($WhatIf -eq $true){
            Write-Verbose "Executing in WhatIf mode"
            if($PSBoundParameters['Verbose']){
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -Verbose -WhatIf
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -Verbose -WhatIf
            }
            else{
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -WhatIf
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -WhatIf
            }
        }
        else {
            if($PSBoundParameters['Verbose']){
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -Verbose
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate -Verbose
            }
            else{
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Duplicate
            }
        }
    }
    else{
        if($WhatIf -eq $true){
            Write-Verbose "Executing in WhatIf mode"
            if($PSBoundParameters['Verbose']){
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Verbose -WhatIf
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Verbose -WhatIf
            }
            else{
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -WhatIf
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -WhatIf
            }
        }
        else {
            if($PSBoundParameters['Verbose']){
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Verbose
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail -Verbose
            }
            else{
                Switch-WPIGRManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail
                Switch-WPIDLManagedBy -TargetEmail $TargetEmail -ReplacementEmail $ReplacementEmail
            }
        }
    }
}