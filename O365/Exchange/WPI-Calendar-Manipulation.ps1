<#
    Functions intened to facilitate manipulation and modification of calendars
#>


#How to add people to a calendar
#Set-MailboxFolderPermission -Identity "username:\Calendar" -User "otherusername" -AccessRights "Contributor"

function Add-WPIUsersToCalendar {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Author", "Contributor", "Editor", "None", "NonEditingAuthor", "Owner", "PublishingEditor", "PublishingAuthor", "Reviewer")]
        [string[]]
        $AccessLevel,

        [Parameter(Mandatory = $true)]
        [string[]]
        $Calendar,

        [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
        [string[]]
        $UserList,
        [switch]
        $Delegate
    )

    foreach ($User in $UserList) {
        $UserPermission = (Get-MailboxFolderPermission -Identity "${Calendar}:\Calendar" -User $User -ErrorAction Ignore).AccessRights

        if ($null -eq $UserPermission) {
            if ($Delegate) {
                Write-verbose "Adding $User to Calendar"
                Add-MailboxFolderPermission -Identity "${Calendar}:\Calendar" -User $User -AccessRights $AccessLevel -SharingPermissionFlags Delegate
            }
            else {
                Write-verbose "Adding $User to Calendar"
                Add-MailboxFolderPermission -Identity "${Calendar}:\Calendar" -User $User -AccessRights $AccessLevel
            }
        }
        elseif ($AccessLevel -ne $UserPermission) {
            Write-Verbose "User $User already had rights to the calendar, updating them."
            Set-MailboxFolderPermission -Identity "${Calendar}:\Calendar" -User $User -AccessRights $AccessLevel
        }
        else {
            Write-verbose "User $User already has the requested rights."
        }
        
    }
}

#Modification to calendars that AEptein wants on their boxes
function Set-ThatThingAEpsteinWanted {
    [CmdletBinding()]
    param (
        $Calendar
    )
    
    Set-CalendarProcessing -Identity $Calendar -DeleteSubject $False -AddOrganizerToSubject $False -DeleteComments $True
    
}

function Set-WPIMeetingSubjectRetention {
    [CmdletBinding()]
    param (
        $Calendar
    )
    set-CalendarProcessing $Calendar -AddOrganizerToSubject $false -DeleteSubject $false -AllowConflicts:$true
}

function Remove-WPIUsersFromCalendar {
    param (

        [Parameter(Mandatory = $true)]
        [string[]]
        $Calendar,

        [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
        [string[]]
        $UserList
    )

    foreach ($User in $UserList) {
        Remove-MailboxFolderPermission -Identity "${Calendar}:\Calendar" -User $User -Confirm:$false
    }
}

function Get-WPIUsersFromCalendar {
    param (
        [Parameter(Mandatory = $false, ParameterSetName = "Identity")]
        [Parameter(Position = 0)]
        $Identity,

        [Parameter(Mandatory = $false, ParameterSetName = "Email")]
        [string[]]
        $EmailAddress
    )
    
    if ($EmailAddress -eq $true) {
        $Identity = $EmailAddress
    }

    Get-MailboxFolderPermission -Identity "${Identity}:\Calendar"
}

function Add-WPIUsersToCalendarCSV {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]
        $Calendar,
        [Parameter(Mandatory = $true)]
        [string[]]
        $CSV,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Author", "Contributor", "Editor", "None", "NonEditingAuthor", "Owner", "PublishingEditor", "PublishingAuthor", "Reviewer")]
        [string[]]
        $AccessLevel
    )
    $UserList = Get-Content $CSV
    Add-WPIUsersToCalendar -AccessLevel $AccessLevel -Calendar $Calendar -UserList $UserList
}

function New-RoomMailbox {
    [CmdletBinding()]
    param (
        $RoomName,
        $UserList
    )
    
    begin {
        New-Mailbox -Name $RoomName -Room
        Set-MailboxFolderPermission ${RoomName}:\Calendar -User Default -AccessRights None
    }
    
    process {
        if ($null -ne $userlist) {
            foreach ($User in $UserList) {
                Add-MailboxFolderPermission ${RoomName}:\Calendar -User $User -AccessRights Author
            }
        }
    }
    
    end {
        
    }
}

function Set-BookinPolicy {
    [CmdletBinding()]
    param (
        $Calendar,
        $Users
    )
    
    begin {
        $calendarSettings = Get-CalendarProcessing $Calendar
        if ($calendarSettings.allbookinpolicy) {
            Set-CalendarProcessing $Calendar -AllBookInPolicy $false
        }
        if ($calendarSettings.AutomateProcessing -ne "AutoAccept") {
            Set-CalendarProcessing $Calendar -AutomateProcessing AutoAccept
        }
    }
    
    process {
        foreach ($user in $users) {
            Set-CalendarProcessing $Calendar -BookInPolicy @{add = "$user" }
        }
    }
    
    end {
        
    }
}

function Get-BookinPolicy {
    [CmdletBinding()]
    param (
        $Calendar
    )

    get-calendarprocessing $Calendar | select-object allbookinpolicy, BookInPolicy
}