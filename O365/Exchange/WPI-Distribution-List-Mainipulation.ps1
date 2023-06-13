<#
    Manipulation of Distribution List AllowSenders
#>

<#
Example use
$HallLists = @("campushouses@wpi.edu","danielshall@wpi.edu","easthall@wpi.edu","ellsworth@wpi.edu","faradayhall@wpi.edu","foundershall@wpi.edu","fuller@wpi.edu","institutehall@wpi.edu","messengerhall@wpi.edu","morganhall@wpi.edu","sanfordrileyhall@wpi.edu","stoddardcomplex@wpi.edu")
foreach($hall in $HallLists){
    $hall = "dl-$hall"
    Add-WPIDLModerator -DL $hall -Moderator mdumke@wpi.edu
    Remove-WPIDLModerator -DL $hall -Moderator jglinos@wpi.edu
}
#>
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
    $DN = (Get-Recipient -Identity "${EmailAddress}").DistinguishedName
    Get-DistributionGroup -Filter ("ModeratedBy -eq '${DN}'")
}

function Get-WPIDLModerators {
    param (
        $DL
    )
    $List = (Get-DistributionGroup $DL).ModeratedBy
    return($List)
}

function Get-WPIDLSenders {
    param (
        $DL
    )
    (Get-DistributionGroup $DL).AcceptMessagesOnlyFromSendersOrMembers
}

function Add-WPIUserToDLSenders {
    param (
        $members,
        $DL
    )

    #Takes a list of users
    foreach($member in $members)
    {
        set-DistributionGroup $DL -AcceptMessagesOnlyFromSendersOrMembers @{Add=$member}
    }
}

function Copy-WPIDLModeratedToSenders {
    param (
        $DL
    )
    $Moderators = Get-WPIDLModeratedList $DL
    Add-WPIUserToDLSender $Moderators $DL
}

function Add-WPIDLToOwnSenders {
    param (
        $DL
    )dkl 
    set-DistributionGroup $DL -AcceptMessagesOnlyFromSendersOrMembers @{Add=$DL}
}

function Copy-WPIDLEverythingIntoSenders {
    param (
        $DL
    )
    Add-WPIDLToOwnSenders $DL
    Copy-WPIDLModeratedToSenders $DL
}

function Add-WPIDLModerator {
    param (
        $DL,
        $Moderator,
        [switch]
        $Check
    )
    if($Moderator -notlike "*@wpi.edu") {
        $Moderator = "${Moderator}@wpi.edu"
    }
    Set-DistributionGroup $DL -BypassModerationFromSendersOrMembers @{Add=$Moderator} -ModeratedBy @{Add=$Moderator}

    if($Check) {
        $DL
        (Get-DistributionGroup $DL)|Select-Object ModeratedBy, BypassModerationFromSendersOrMembers|Format-List
    }
}
function Remove-WPIDLModerator {
    param (
        $DL,
        $Moderator,
        [switch]
        $Check
    )
    if($Moderator -notlike "*@wpi.edu") {
        $Moderator = "${Moderator}@wpi.edu"
    }
    Set-DistributionGroup $DL -BypassModerationFromSendersOrMembers @{Remove=$Moderator} -ModeratedBy @{Remove=$Moderator}

    if($Check) {
        $DL
        (Get-DistributionGroup $DL)|Select-Object ModeratedBy, BypassModerationFromSendersOrMembers|Format-List
    }
}

<#
Author: Joshua Gorham
Date: 10/18/2018
Add user as Sender on lists without moderation, and as a bypass sender on lists with moderation
#>

Function Add-WPIDLSender ($DistributionList, $EmailAddress, $Path)
{
    if($DistributionList -like "*.txt")
    {
        $Path = $DistributionList
        $DistributionList = $null
    }
    if($Path)
    {
        $List = Get-Content $Path

        foreach($dl in $List)
        {
            $workingList = Get-DistributionGroup $dl

            if($workingList.ModerationEnabled)
            {
                Set-DistributionGroup $dl -BypassModerationFromSendersOrMembers @{Add=$EmailAddress}
            }
            else
            {
                Set-DistributionGroup $dl -AcceptMessagesOnlyFromSendersOrMembers @{Add=$EmailAddress}
            }
      }
    }
    elseif($DistributionList)
    {
        $workingList = Get-DistributionGroup $DistributionList

        if($workingList.ModerationEnabled)
        {
            Set-DistributionGroup $DistributionList -BypassModerationFromSendersOrMembers @{Add=$EmailAddress}
        }
        else
        {
            Set-DistributionGroup $DistributionList -AcceptMessagesOnlyFromSendersOrMembers @{Add=$EmailAddress}
        }
    }
}

function Add-DLMemberFromCSV {
    [cmdletbinding()]
    param (
        $DL,
        $file
    )
    $List = Get-Content $file

    Import-Module azureadpreview
    Connect-AzureAD

    Write-Verbose "Adding users from ${File} to the group ${DL}"
    foreach($samaccountname in $List){
        if(!($samaccountname -like "*@wpi.edu") -and !($samaccountname -like "*@cs.wpi.edu"))
        {
            if(!(Get-AzureADUser -Filter "UserPrincipalName eq '${samaccountname}'"))
            {
                Write-Verbose "${samaccountname} is not a WPI user, creating a contact object for them."
                New-MailContact -Name $samaccountname -ExternalEmailAddress $samaccountname
            }
        }
    }
    foreach($samaccountname in $List){
        Write-Verbose "Adding ${samaccountname} to the DL"
        Add-DistributionGroupMember -Identity $DL -Member $samaccountname
    }
    #Write-Verbose "Re-Enabling Welcome Message on ${Group}"
    # Set-UnifiedGroup $Group -UnifiedGroupWelcomeMessageEnabled:$true # Honestly nobody seems to want this
}

function Remove-DLMemberFromCSV {
    [CmdletBinding()]
    param (
        $DL,
        $file
    )
    $List = Get-Content $file

    foreach($samaccountname in $list){
    Remove-DistributionGroupMember -identity $DL -Member $samaccountname -Confirm:$false
    }
}

function Get-DLListMembership {
    [CmdletBinding()]
    param (
        $Username
    )
    
    begin {
        
    }
    
    process {
        Get-DistributionGroup | Where-Object { (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$Username"}
    }
    
    end {
        
    }
}

function Add-DLOwner {
    [CmdletBinding()]
    param (
        $Owner
    )
    
    begin {
        $OptIn = Get-AzureADGroup -SearchString "OptIn-" -All:$true
        $OptOut = Get-AzureADGroup -SearchString "OptOut-" -All:$true
        $targets = $OptIn + $OptOut
    }
    
    process {
        $UID = (Get-AzureADUser -SearchString $owner).ObjectId
        foreach($list in $targets){
            Add-AzureADGroupOwner -ObjectId $list.objectID -RefObjectId $UID
        }
    }
    
    end {
        
    }
}

function Remove-DLOwner {
    [CmdletBinding()]
    param (
        $Owner
    )
    
    begin {
        $OptIn = Get-AzureADGroup -SearchString "OptIn-" -All:$true
        $OptOut = Get-AzureADGroup -SearchString "OptOut-" -All:$true
        $targets = $OptIn + $OptOut
    }
    
    process {
        $UID = (Get-AzureADUser -SearchString $owner).ObjectId
        foreach($list in $targets){
            Remove-AzureADGroupOwner -ObjectId $list.objectID -RefObjectId $UID
        }
    }
    
    end {
        
    }
}

function Add-DLBypassModerator {
    [CmdletBinding()]
    param (
        $DistributionList,
        $UserEmail
    )
    set-DistributionGroup -Identity $DistributionList -BypassModerationFromSendersOrMembers @{Add=$UserEmail}
}

function Remove-DLSubscription{
    [CmdletBinding()]
    param (
        $EmailAddress,
        $List
    )
    $User = Get-AzureADUser -ObjectId $EmailAddress
    $UID = $User.ObjectId

    if ($list.startswith("dl-") ) {
        $ListCore = $list.split("-")[1].split("@")[0]
    }
    else {
        $ListCore = $list.Split("@")[0]
    }

    if ($ListCore -eq "allemployees") {
        $ListCore = "employees"        
    }

    #$DL_ID =(Get-DistributionGroup dl-$ListCore).ExternalDirectoryObjectId
    $OptIn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Optin-$ListCore'").ObjectId
    $OptOut_ID = (Get-AzureADGroup -filter "DisplayName eq 'OptOut-$ListCore'").ObjectId
    $Dyn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Dyn-$ListCore'").ObjectId

    $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

    if($membership.objectid -contains $Dyn_ID){
        if($membership.objectid -contains $OptOut_ID){
            Write-Host "User is already opted out of $ListCore" -ForegroundColor Yellow
        }
        elseif($membership.objectid -contains $OptIn_ID){
            Remove-AzureADGroupMember -ObjectId $OptIn_ID -MemberId $UID
            Add-AzureADGroupMember -ObjectId $OptOut_ID -RefObjectId $UID
            Write-Host "User removed from $ListCore-OptIn Group" -ForegroundColor Yellow
            Write-Host "User added to $ListCore-OptOut Group" -ForegroundColor Green
        }
        else{
            Add-AzureADGroupMember -ObjectId $OptOut_ID -RefObjectId $UID
            Write-Host "User added to $ListCore-OptOut Group" -ForegroundColor Green
        }
    }
    else{
        Write-Host "User is not a member of $ListCore" -ForegroundColor Yellow
    }
}

function Add-DLSubscription{
    [CmdletBinding()]
    param (
        $EmailAddress,
        $List
    )
    $User = Get-AzureADUser -ObjectId $EmailAddress
    $UID = $User.ObjectId

    if ($list.startswith("dl-") ) {
        $ListCore = $list.split("-")[1].split("@")[0]
    }
    else {
        $ListCore = $list.Split("@")[0]
    }

    if ($ListCore -eq "allemployees") {
        $ListCore = "employees"        
    }

    $DL_ID =(Get-DistributionGroup dl-$ListCore).ExternalDirectoryObjectId
    $OptIn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Optin-$ListCore'").ObjectId
    $OptOut_ID = (Get-AzureADGroup -filter "DisplayName eq 'OptOut-$ListCore'").ObjectId
    $Dyn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Dyn-$ListCore'").ObjectId

    $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

    if(!($membership.objectid -contains $DL_ID)){
        if($membership.objectid -contains $OptOut_ID){
            Remove-AzureADGroupMember -ObjectId $OptOut_ID -MemberId $UID
            if(!($membership.objectid -contains $Dyn_ID)){
                Add-AzureADGroupMember -ObjectId $OptIn_ID -RefObjectId $UID
            }
            Write-Host "User removed from $Listcore-OptOut Group" -ForegroundColor Yellow
            Write-Host "User added to $ListCore-OptIn Group" -ForegroundColor Green
        }
        else{
            Add-AzureADGroupMember -ObjectId $OptIn_ID -RefObjectId $UID
            Write-Host "User added to $ListCore-OptIn Group" -ForegroundColor Green
        }
    }
    else{
        Write-Host "User is a member of $ListCore" -ForegroundColor Yellow
    }
}


function get-OptInGroups {
    param (
        
    )
    get-AzureADGroup -SearchString "OptIn-" -All:$true
}
function get-OptOutGroups {
    param (
        
    )
    get-AzureADGroup -SearchString "OptOut-" -All:$true
}

function Update-OptListOwners {
    [CmdletBinding()]
    param (
        
    )
    $optin_Lists = get-AzureADGroup -SearchString "OptIn-" -All:$true
    $optout_Lists = get-AzureADGroup -SearchString "OptOut-" -All:$true

    #$ManagerGroupMembers = Get-AzureADGroupMember -ObjectId "bf4bdb7c-ede6-4e54-949c-f84ecf3f0c14"

    foreach($list in $optin_Lists){
        Sync-OptOwner -TargetGroupID $list.ObjectId -ReferenceGroupID "bf4bdb7c-ede6-4e54-949c-f84ecf3f0c14"
    }

    foreach($list in $optout_Lists){
        Sync-OptOwner -TargetGroupID $list.ObjectId -ReferenceGroupID "bf4bdb7c-ede6-4e54-949c-f84ecf3f0c14"
    }
    
}

function Sync-OptOwner {
    [CmdletBinding()]
    param (
        $TargetGroupID,
        $ReferenceGroupID
    )
    
    begin {
        $TargetMembers = Get-AzureADGroupOwner -ObjectId $TargetGroupID -All:$true
        $ReferenceMembers = Get-AzureADGroupMember -ObjectId $ReferenceGroupID -All:$true
    }
    
    process {
        if ($null -eq $TargetMembers) {
            $AddMembers = $ReferenceMembers
        }
        elseif ($null -eq $ReferenceMembers) {
            $RemoveMembers = $TargetMembers
        }
        else {
            #reconscile lists
            $comparisons = Compare-Object  $TargetMembers $ReferenceMembers -Property ObjectId
                
            # Store the users who should and shouldn't be in the lists in variables.
            $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object ObjectId
            $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object ObjectId
        }
    }
    
    end {
        # Iterate through the add/remove lists and do what's necessary.
        ForEach ($Removal in $RemoveMembers.ObjectId) {
            Write-Verbose "Removing $removal from $targetgroupid"
            Remove-AzureADGroupOwner -ObjectId $TargetGroupID -OwnerId $Removal
        }
                        
        ForEach ($Addition in $AddMembers.ObjectId) {
            Write-Verbose "Adding $Addition to $targetgroupid"
            Add-AzureADGroupOwner -ObjectId $TargetGroupID -RefObjectId $Addition
                        
        }
    }
}


function Update-DynRule {
    param (
        
    )
    Import-Module AzureADPreview -Force

    $dynlists = Get-AzureADMSGroup -SearchString "dyn-" -All:$true
    foreach ($List in $dynlists) {
        $RuleAddition = ' and (user.UserType -eq "Member")'
        $completeRule = $list.MembershipRule + $RuleAddition
        Set-AzureADMSGroup -Id $List.ID -MembershipRule $completeRule
    }
}

function Update-ADVRule {
    param (
        
    )
    Import-Module AzureADPreview -Force

    $advlists = Get-AzureADMSGroup -SearchString "adv-" -All:$true
    $counter = 0
    foreach ($List in $advlists) {
        $counter++
        Write-Progress -Activity "Processing Lists" -CurrentOperation "dyn-$($list.DisplayName.Split("-")[1])" -PercentComplete (($counter / $advlists.count) * 100)
        $dynlist = Get-AzureADMSGroup -filter "Displayname eq 'dyn-$($list.DisplayName.Split("-")[1])'"
        
        #$completeRule = $dynlist.MembershipRule -replace "$($dynlist.displayname.split("-")[1]);", "$($dynlist.displayname.split("-")[1]);.*OADV-"               # this rule omits OADV from the population <+++++++++++++++
        #$completeRule = $dynlist.MembershipRule -replace "$($dynlist.displayname.split("-")[1]);", "(OADV-)*(.*;)*$($dynlist.displayname.split("-")[1]);"        # This rule makes the list reliably capture OADV
        
        Set-AzureADMSGroup -Id $dynList.ID -MembershipRule $completeRule
    }
}

function Add-UserToOnPremMailEnabledSecurityGroup {
    [CmdletBinding()]
    param (
        $Group,
        $user
    )
    
    begin {
        $user = get-aduser $user.split("@")[0]
        $group = $group.split("@")[0]
    }
    
    process {
        set-adgroup $group -add @{authorig="$($user.DistinguishedName)"}
    }
    
    end {
        
    }
}

function Set-StudentListModerator {
    param (
        $NewModerator
    )
    $Lists = @("dl-students", "dl-undergraduates", "dl-seniors", "dl-juniors", "dl-sophomores", "dl-firstyears")
    foreach ($List in $Lists) {
        Set-DistributionGroup $list -ModeratedBy $NewModerator
    }
    foreach ($List in $Lists) {
        $currentlist = Get-DistributionGroup $list
        if($currentlist.moderatedby -contains (get-mailbox $NewModerator).name){
            Write-Host -BackgroundColor Green "User $newmoderator was Successfully Made Moderator of $list"
        }
        else {
            Write-Host -BackgroundColor Red "User $newmoderator was Not Successfully Made Moderator of $list"
        }
    }
}