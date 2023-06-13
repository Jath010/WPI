function Repair-MailboxAdminAccess {
    param(
        $Mailbox
    )
    Add-MailboxPermission -identity $Mailbox -User "Organization Management" -AccessRights fullaccess -InheritanceType all -AutoMapping $false
}

function Repair-AllMailboxAdminAccess {
    $BoxList = Get-Mailbox -ResultSize unlimited -Filter { (RecipientTypeDetails -eq 'UserMailbox') } 
    foreach ($Box in $BoxList) {
        if ($null -eq (Get-MailboxPermission -Identity $Box.alias -user "Organization Management")) {

            Add-MailboxPermission -Identity $Box.alias -User "Organization Management" -AccessRights fullaccess -InheritanceType all -AutoMapping $False        
        }
    } 
}


<# workflow Repair-AllMailboxAdminAccessParallel {
    $BoxList = Get-Mailbox -ResultSize unlimited -Filter { (RecipientTypeDetails -eq 'UserMailbox') } 
    foreach -parallel ($Box in $BoxList) {
        if ($null -eq (Get-MailboxPermission -Identity $Box.alias -user "Organization Management")) {

            Add-MailboxPermission -Identity $Box.alias -User "Organization Management" -AccessRights fullaccess -InheritanceType all -AutoMapping $False        
        }
    } 
} #>

# Trivia: This line gets every person who is hidden in the GAL who is currently in the students OU
# get-aduser -Filter {msExchHideFromAddressLists -eq $true -and Enabled -eq $true} -Properties msExchHideFromAddressLists -SearchBase "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"

function Repair-WPIGALvisibility {
    param(
        $User
    )
    $DN = (get-aduser $user).distinguishedname
    set-ADobject $dn -replace @{msExchHideFromAddressLists = $false }
}

function Remove-WPIGALvisibility {
    param(
        $User
    )
    $DN = (get-aduser $user).distinguishedname
    set-ADobject $dn -replace @{msExchHideFromAddressLists = $true }
    # set-ADobject $dn -replace @{mailnickname = "" }           # Apparently you need to add a mail nickname for it to work?
}

function Add-Infoed {
    [CmdletBinding()]
    param (
        $user
    )
    
    set-aduser $user -add @{proxyAddresses = "smtp:${user}@infoed.wpi.edu" }
}

<# Adds users as full access on boxes
$users = get-content -path "C:\tmp\temp\MailboxAdd.txt"
$boxes = "Library-Staff-Events@wpi.edu", "library-staff-vacation@wpi.edu", "cal_library_208B@wpi.edu"
foreach($User in $Users){
    foreach($box in $Boxes){
        Add-MailboxPermission -Identity $box -User $User -AccessRights fullaccess -InheritanceType all -AutoMapping $False
    }
    
}
#>

function Add-MailboxAccess {
    [CmdletBinding()]
    param (
        $Users,
        $Mailboxes,
        [switch]$OnBehalf
    )
    foreach ($User in $Users) {
        if ($user -notmatch ".*@wpi.edu") {
            $user = $user + "@wpi.edu"
        }
        foreach ($box in $mailboxes) {
            Add-MailboxPermission -Identity $box -User $User -AccessRights fullaccess -InheritanceType all -AutoMapping $False
            Add-RecipientPermission -Identity $box -Trustee $User -AccessRights SendAs -Confirm:$false
            if ($OnBehalf) {
                Set-Mailbox -Identity $box -GrantSendOnBehalfTo @{add = "$User" }
                if ((Get-Mailbox -Identity $box).GrantSendOnBehalfTo -contains (get-mailbox $user).name) {
                    Write-Host -BackgroundColor Green -ForegroundColor Black "User $User successfully granted On-Behalf Rights."
                }
                else {
                    Write-Host -BackgroundColor Red -ForegroundColor Black "User $User NOT successfully granted On-Behalf Rights."
                }
            }
            Get-UserMailboxAccess -User $User -Mailbox $box
        }
    
    }
}
function Add-MailboxOnBehalf {
    param (
        $Users,
        $Mailboxes
    )
    foreach ($User in $Users) {
        if ($user -notmatch ".*@wpi.edu") {
            $user = $user + "@wpi.edu"
        }
        foreach ($box in $mailboxes) {
            Set-Mailbox -Identity $box -GrantSendOnBehalfTo @{add = "$User" }
            if ((Get-Mailbox -Identity $box).GrantSendOnBehalfTo -contains (get-mailbox $user).name) {
                Write-Host -BackgroundColor Green -ForegroundColor Black "User $User successfully granted On-Behalf Rights."
            }
            else {
                Write-Host -BackgroundColor Red -ForegroundColor Black "User $User NOT successfully granted On-Behalf Rights."
            }
        }
    
    }
}
function Remove-MailboxAccess {
    [CmdletBinding()]
    param (
        $Users,
        $Mailboxes
    )
    foreach ($User in $Users) {
        foreach ($box in $mailboxes) {
            Remove-MailboxPermission -Identity $box -User $User -AccessRights fullaccess -InheritanceType all -Confirm:$False
            Remove-RecipientPermission -Identity $box -Trustee $User -AccessRights SendAs -Confirm:$false
            Get-UserMailboxAccess -User $User -Mailbox $box
        }
    
    }
}
function Get-UserMailboxAccess {
    [CmdletBinding()]
    param (
        [string]
        $User,
        [string]
        $Mailbox
    )
    
    
    $Access = get-MailboxPermission -Identity $mailbox | Where-Object { $_.user -like $user }
    if ($null -ne $Access) {
        Write-host "User $User has Full Access to $mailbox" -ForegroundColor Black -BackgroundColor Green
    }
    else {
        Write-host "User $User does not have Full Access to $mailbox" -ForegroundColor White -BackgroundColor Red
    }
    $SendAs = get-RecipientPermission -Identity $mailbox | Where-Object { $_.Trustee -like $user }
    if ($null -ne $SendAs) {
        Write-host "User $User has Send As to $mailbox" -ForegroundColor Black -BackgroundColor Green
    }
    else {
        Write-host "User $User does not have Send As to $mailbox" -ForegroundColor White -BackgroundColor Red
    }
}

# function Get-AllUserMailboxAccess {
#     [CmdletBinding()]
#     param (
#         $user
#     )
#     begin {
#         $Mailboxes = get-mailbox -ResultSize Unlimited
#     } 
#     process {      
#     }
#     end { 
#     }
# }

function Repair-MailboxMigration {
    param (
        $user
    )
    set-aduser $user -Clear MsExchMailboxGuid, MsExchRecipientDisplayType, MsExchRecipientTypeDetails, LegacyExchangeDN
}

function Repair-DisabledMailboxMigration {
    param (
        
    )
    $disabled = get-aduser -filter * -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu"

    foreach ($user in $disabled) {
        set-aduser $user -Clear MsExchMailboxGuid, MsExchRecipientDisplayType, MsExchRecipientTypeDetails, LegacyExchangeDN
    }
}

function Add-OutOfOfficeReply {
    [CmdletBinding()]
    param (
        $EmailAddress,
        $Message
    )
    
    begin {
        
    }
    
    process {
        if ($null -eq $Message) {
            Set-MailboxAutoReplyConfiguration -Identity $EmailAddress -InternalMessage "This account is no longer active." -ExternalMessage "This account is no longer active." -AutoReplyState Enabled -ExternalAudience All
        }
        else {
            Set-MailboxAutoReplyConfiguration -Identity $EmailAddress -InternalMessage $Message -ExternalMessage $Message -AutoReplyState Enabled -ExternalAudience All
        }
    }
    
    end {
        
    }
}

function Remove-PotpourriSubscription {
    [CmdletBinding()]
    param (
        $EmailAddress
    )
    $User = Get-AzureADUser -ObjectId $EmailAddress
    $UID = $User.ObjectId
    $PotpourriOptIn = "51917837-ce04-42ef-b2e5-c0427a8dfab0"
    $PotpourriOptOut = "ad9edfc9-337b-4b8f-9a61-e3458d3d38e2"
    $dynPotpourri = "73748c39-bba0-4f92-b555-704a5c140bc5"
    #$dlPotpourri = "69f702dc-6f24-446a-9b28-0a3dc1d7378e"

    $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

    if ($membership.objectid -contains $dynPotpourri) {
        if ($membership.objectid -contains $PotpourriOptOut) {
            Write-Host "User is already opted out of Potpourri" -ForegroundColor Yellow
        }
        elseif ($membership.objectid -contains $PotpourriOptIn) {
            Remove-AzureADGroupMember -ObjectId $PotpourriOptIn -MemberId $UID
            Add-AzureADGroupMember -ObjectId $PotpourriOptOut -RefObjectId $UID
            Write-Host "User removed from Potpourri-OptIn Group" -ForegroundColor Yellow
            Write-Host "User added to Potpourri-OptOut Group" -ForegroundColor Green
        }
        else {
            Add-AzureADGroupMember -ObjectId $PotpourriOptOut -RefObjectId $UID
            Write-Host "User added to Potpourri-OptOut Group" -ForegroundColor Green
        }
    }
    elseif ($membership.objectid -contains $PotpourriOptIn) {
        Remove-AzureADGroupMember -ObjectId $PotpourriOptIn -MemberId $UID
        Write-Host "User removed from Potpourri-OptIn Group" -ForegroundColor Yellow
    }
    else {
        Write-Host "User is not a member of Potpourri" -ForegroundColor Yellow
    }
}

function Add-PotpourriSubscription {
    [CmdletBinding()]
    param (
        $EmailAddress
    )
    $User = Get-AzureADUser -ObjectId $EmailAddress
    $UID = $User.ObjectId
    $PotpourriOptIn = "51917837-ce04-42ef-b2e5-c0427a8dfab0"
    $PotpourriOptOut = "ad9edfc9-337b-4b8f-9a61-e3458d3d38e2"
    $Potpourri = "69f702dc-6f24-446a-9b28-0a3dc1d7378e"
    $dynPotpourri = "73748c39-bba0-4f92-b555-704a5c140bc5"

    $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

    if (!($membership.objectid -contains $Potpourri)) {
        if ($membership.objectid -contains $PotpourriOptOut) {
            Remove-AzureADGroupMember -ObjectId $PotpourriOptOut -MemberId $UID
            if (!($membership.objectid -contains $dynPotpourri)) {
                Add-AzureADGroupMember -ObjectId $PotpourriOptIn -RefObjectId $UID
            }
            Write-Host "User removed from Potpourri-OptOut Group" -ForegroundColor Yellow
            Write-Host "User added to Potpourri-OptIn Group" -ForegroundColor Green
        }
        else {
            Add-AzureADGroupMember -ObjectId $PotpourriOptIn -RefObjectId $UID
            Write-Host "User added to Potpourri-OptIn Group" -ForegroundColor Green
        }
    }
    else {
        Write-Host "User is a member of Potpourri" -ForegroundColor Yellow
    }
}


function Add-MailboxForward {
    [CmdletBinding()]
    param (
        $EmailAddress,
        $ForwardingAddress
    )
    Set-Mailbox -Identity $EmailAddress -ForwardingSmtpAddress $ForwardingAddress
}