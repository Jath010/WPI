#removes and readds the users in a group to cause them to all be resubbed again.
#doesn't disable welcome message

$group= "gr-sfs-program@wpi.edu"

$members = Get-UnifiedGroupLinks -Identity $group -LinkType member
foreach ($member in $members){
    Remove-UnifiedGroupLinks -Identity $group -LinkType member -Links $member.PrimarySmtpAddress -Confirm:$true
    Add-UnifiedGroupLinks -Identity $group -LinkType member -Links $member.PrimarySmtpAddress
}