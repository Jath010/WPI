#Intent is to copy a list on users fom an existing list to a gr

$dynamicGroupTypeString = "DynamicMembership"

function Disable-WPIGroupWelcome {
    param(
        $Group
    )
    Set-UnifiedGroup -Identity $Group -UnifiedGroupWelcomeMessageEnabled:$false
}

function New-WPIGroupHiddenMembershipVisibility {
    param (
        $Group
    )
    $Object = New-UnifiedGroup -DisplayName $group -PrimarySmtpAddress "$Group@wpi.edu" -HiddenGroupMembershipEnabled:$true -UnifiedGroupWelcomeMessageEnabled:$false -AccessType Private
    Set-UnifiedGroup -Identity $Object.id -UnifiedGroupWelcomeMessageEnabled:$false
}

function New-WPIGroupHiddenExchange {
    param (
        $Group
    )
    $object = New-UnifiedGroup -DisplayName $group -PrimarySmtpAddress "$Group@wpi.edu" -AutoSubscribeNewMembers:$true -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromExchangeClientsEnabled:$true -AccessType Private
    
    $dynamicMembershipRule = "(user.extensionAttribute15 -match ""PADV-xkong;.*"") or (user.extensionAttribute15 -match "".*OADV-.*xkong.*"")"

    #existing group types
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

    if($null -ne $groupTypes -and $groupTypes.Contains($dynamicGroupTypeString))
    {
        throw "This group is already a dynamic group. Aborting conversion.";
    }
    #add the dynamic group type to existing types
    $groupTypes.Add($dynamicGroupTypeString)
    
    Set-AzureADMSGroup -Id $object.ObjectId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
}

function New-WPIAdvisingList {
    param (
        $AdvisorCode
    )
    #Create Group and save info for further use.
    #Setting Name, Email, Owner, Disabling WelcomeMessage, Enable Hidden From Client and Autosubscribe
    $object = New-UnifiedGroup -DisplayName "adv-$AdvisorCode" -PrimarySmtpAddress "adv-$AdvisorCode@wpi.edu" -AutoSubscribeNewMembers:$true -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromExchangeClientsEnabled:$true -Owner "$AdvisorCode@wpi.edu" -AccessType Private
    #Add the Advisor as the only User who can send to the group
    Set-UnifiedGroup -Identity $object.ObjectId -AcceptMessagesOnlyFromSendersOrMembers "$AdvisorCode@wpi.edu"
    #Create rule for membership
    $dynamicMembershipRule = "(user.extensionAttribute15 -match ""PADV-$AdvisorCode;.*"") or (user.extensionAttribute15 -match "".*OADV-.*$AdvisorCode.*"")"

    ###########
    #This block of code is swiped from Microsoft. It modifies an existing static membership group into a dynamic membership one
    #existing group types
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

    if($null -ne $groupTypes -and $groupTypes.Contains($dynamicGroupTypeString))
    {
        throw "This group is already a dynamic group. Aborting conversion.";
    }
    #add the dynamic group type to existing types
    $groupTypes.Add($dynamicGroupTypeString)
    #Set-AzureADMsGroup is from the azureadpreview module
    Set-AzureADMSGroup -Id $object.ObjectId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
    ###########
}

function ConvertStaticGroupToDynamic
{
    Param([string]$groupId, [string]$dynamicMembershipRule)

    #existing group types
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $groupId).GroupTypes

    if($null -ne $groupTypes -and $groupTypes.Contains($dynamicGroupTypeString))
    {
        throw "This group is already a dynamic group. Aborting conversion.";
    }
    #add the dynamic group type to existing types
    $groupTypes.Add($dynamicGroupTypeString)

    #modify the group properties to make it a static group: i) change GroupTypes to add the dynamic type, ii) start execution of the rule, iii) set the rule
    Set-AzureAdMsGroup -Id $groupId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
}

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

function Copy-UsersToGr {
    param (
        $DL,
        $GR
    )
    $originList = Get-AdGroupMember $DL

    foreach($member in $originList)
    {
        #TODO Add a check to prevent wasting time on stuff already in the group
        Add-UnifiedGroupLinks -Identity $GR -LinkType Members -Links ($member.samaccountname)
    }
}

function Copy-UsersToSG {
    param (
        $DL,
        $SG
    )

    $originList = Get-AdGroupMember $DL

    foreach($member in $originList)
    {
        #TODO Add a check to prevent wasting time on stuff already in the group

        #Mail-Enabled Security Groups require the -BypassSecurityGroupManagerCheck switch to be modified by someone not in their owner list.
        Add-DistributionGroupMember -Identity $SG -Member ($member.samaccountname) -BypassSecurityGroupManagerCheck
    }
}

function Add-GroupModerator {
    param (
        $Group,
        $Moderator
    )
    if((Get-UnifiedGroup $group).ModerationEnabled -eq $false) {
        Set-UnifiedGroup $group -ModerationEnabled $true
    }
    set-UnifiedGroup $Group -ModeratedBy @{Add="$Moderator"}
}

#Takes a CSV that lists each of the samaccountnames to add
function Add-GroupMemberFromCSV {
    [cmdletbinding()]
    param (
        $Group,
        $file
    )
    $List = Get-Content $file

    Import-Module azureadpreview
    Connect-AzureAD

    Write-Verbose "Adding users from ${File} to the group ${group}"
    foreach($samaccountname in $List){
        if(!($samaccountname -like "*@wpi.edu") -and !($samaccountname -like "*@cs.wpi.edu"))
        {
            if(!(Get-AzureADUser -Filter "UserPrincipalName eq '${samaccountname}'"))
            {
                Write-Verbose "${samaccountname} is not a WPI user, creating a contact object for them."
                New-AzureADMSInvitation -ErrorAction continue -InvitedUserEmailAddress $samaccountname -InvitedUserDisplayName $samaccountname -SendInvitationMessage:$false -InviteRedirectUrl "https://myapps.microsoft.com" | Out-Null
            }
        }
    }

    Write-Verbose "Disabling Welcome Message on ${Group}"
    Set-UnifiedGroup $Group -UnifiedGroupWelcomeMessageEnabled:$false
    foreach($samaccountname in $List){
        Write-Verbose "Adding ${samaccountname} to the group"
        Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $samaccountname
    }
    #Write-Verbose "Re-Enabling Welcome Message on ${Group}"
    # Set-UnifiedGroup $Group -UnifiedGroupWelcomeMessageEnabled:$true # Honestly nobody seems to want this
}

function Add-WPIGROwnerFromCSV {
    [cmdletbinding()]
    param (
        $CSV,
        $Owner
    )
    $File = Get-Content $CSV
    ForEach($Group in $File){
        Write-Verbose "Adding ${Owner} to ${Group} as Owner"
        Add-UnifiedGroupLinks -Identity $Group -LinkType Members -Links $Owner
        Add-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $Owner
    }
}

function Remove-WPIGROwnerFromCSV {
    [cmdletbinding()]
    param (
        $CSV,
        $Owner
    )
    $File = Get-Content $CSV
    ForEach($Group in $File){
        Write-Verbose "Removing ${Owner} as Owner from ${Group}"
        Remove-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $Owner -Confirm:$false
    }
}

function Export-WPIGRtoCSV {
    param (
        $Group
    )
    Get-UnifiedGroupLinks -Identity $Group -LinkType Members | Select-Object Name, PrimarySmtpAddress | Export-Csv -Path "C:\tmp\temp\${Group}-members.csv"
}

function Get-WPIGRMembership {
    param (
        $EmailAddress
    )
    Get-AzureADUser -ObjectID $EmailAddress | Get-AzureADUserMembership | Where-Object {$_.ObjectType -ne "Role"}  | ForEach-Object {Get-UnifiedGroup -Identity $_.ObjectId -ErrorAction Ignore}|Format-Table
}

#If a person still has groups after running these then they are probably the sole owner of those groups
function Remove-WPIGRMembership {
    param (
        $EmailAddress
    )
    Get-AzureADUser -ObjectID $EmailAddress | Get-AzureADUserMembership | Where-Object {$_.ObjectType -ne "Role"}  | ForEach-Object {Remove-UnifiedGroupLinks -Identity $_.ObjectId -Links $EmailAddress -LinkType members -Confirm:$False -ErrorAction Ignore}
}

function Remove-WPIGROwnership {
    param (
        $EmailAddress
    )
    Get-AzureADUser -ObjectID $EmailAddress | Get-AzureADUserMembership | Where-Object {$_.ObjectType -ne "Role"}  | ForEach-Object {Remove-UnifiedGroupLinks -Identity $_.ObjectId -Links $EmailAddress -LinkType Owners -Confirm:$False -ErrorAction Ignore}
}

function Set-SwapWPIGROwnership {
    param (
        $NewOwner,
        $OldOwner,
        $Group
    )
    Add-UnifiedGroupLinks -Identity $Group -LinkType Owners -Links $NewOwner
    Remove-UnifiedGroupLinks -Identity $Group -Links $OldOwner -LinkType Owners -Confirm:$False -ErrorAction Ignore
}

function Remove-GrPrefix {
    [CmdletBinding()]
    param (
        $EmailAddress
    )
    
    try {
        $group = Get-UnifiedGroup $EmailAddress
        $guid = $group.guid.tostring()
    }
    catch {
        Write-host "Please retry with a valid email address" -ForegroundColor Red
        return
    }

    if ($group.DisplayName.StartsWith("gr-")) {
        $DisplayName = $group.DisplayName.Trim("gr-")
    }else {
        $DisplayName = $group.DisplayName
    }
    if ($group.PrimarySmtpAddress.StartsWith("gr-")) {
        $PrimarySmtpAddress = $group.PrimarySmtpAddress.Trim("gr-")
    }else {
        $PrimarySmtpAddress = $group.PrimarySmtpAddress
    }

    set-UnifiedGroup -Identity $guid -DisplayName $DisplayName -PrimarySmtpAddress $PrimarySmtpAddress

    if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
        Get-UnifiedGroup -Identity $guid|Select-Object DisplayName, PrimarySmtpAddress
    }
}

function Add-GroupSendAs {
    [CmdletBinding()]
    param (
        $group,
        $users
    )
    
    $group = Get-Recipient -RecipientTypeDetails groupmailbox -Identity $group

    foreach($user in $users){
        Add-RecipientPermission -Identity $group.name -Trustee $user -AccessRights SendAs -Confirm:$false
    }
}

function Set-GroupMembersToSubscribe {
    param (
        $Group
    )
    $GroupObject = Get-UnifiedGroup -Identity $Group
    $Members = Get-UnifiedGroupLinks -Identity $GroupObject.name -LinkType Members
    $Subscribers = Get-unifiedgrouplinks -Identity $GroupObject.name -LinkType Subscribers
    foreach($Member in $Members){
        if($member.name -notin $Subscribers.name){
            Add-UnifiedGroupLinks -Identity $GroupObject.name -LinkType Subscribers -Links $member.name
        }
    }
}