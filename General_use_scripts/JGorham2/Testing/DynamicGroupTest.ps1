
#Create a new group and disable welcome message
Connect-AzureAD

function New-WPIGroup {
    param (
        $DisplayName,
        $EmailAddress
    )
    New-UnifiedGroup -Alias $DisplayName -AutoSubscribeNewMembers:$false -HiddenGroupMembershipEnabled -EmailAddresses $EmailAddress
}
