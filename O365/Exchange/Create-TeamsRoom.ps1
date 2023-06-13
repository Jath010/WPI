#Creates a Teams Room for one of those Logitech Bars


function New-TeamsRoom {
    [CmdletBinding()]
    param (
        $RoomName,
        $RoomEmailAddress,
        $Password
    )
    
    begin {
        Connect-ExchangeOnline

    }
    
    process {
        $RoomEmailAddress = $RoomEmailAddress.split("@")[0]

        $roombox = New-Mailbox -MicrosoftOnlineServicesID "$RoomEmailAddress@wpi.edu" -Name $RoomName -Alias $RoomEmailAddress -Room -EnableRoomMailboxAccount $true  -RoomMailboxPassword (ConvertTo-SecureString -String "$Password" -AsPlainText -Force)

        Set-CalendarProcessing -Identity $roombox.alias -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -ProcessExternalMeetingMessages $true -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "This is a Microsoft Teams Meeting room!"

    }
    
    end {
        
    }
}