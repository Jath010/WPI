function New-GuestTeam {
    [CmdletBinding()]
    param (
        $Team,
        $Faculty,
        $Students
    )
    
    $team = New-team -alias $Team -DisplayName $Team -AddCreatorASMember $false -Template EDU_Class

    foreach($professor in $faculty){
        Add-TeamUser -GroupID $team.GroupID -User "${professor}@wpi.edu" -Role Owner
    }
    foreach($student in $students){
        $AAD = get-azureaduser -searchstring $student
        if($null -ne $AAD.displayname){
                Add-TeamUser -GroupID $Team.GroupID -User $AAD.UserPrincipalName -Role Member            
        }
        else {
            New-AzureADMSInvitation -ErrorAction continue -InvitedUserEmailAddress $student -InvitedUserDisplayName $student
            Add-TeamUser -GroupID $Team.GroupID -User $student -Role Member
        }
    }

}