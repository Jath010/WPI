#$var = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate
#$noNumbers = ($var| where {$_.name -notmatch ".*(\(\d.*\))"})                           # This finds accounts that lack numbers in their name to find non-standard accounts in disabled
#$olderThanThree = ($var| where {$_.lastlogondate -lt (get-date).addyears(-3)}).count    # this finds accounts that haven't been logged into in 3 years



function Move-OldDisabled {
    [CmdletBinding()]
    param (
        
    )
    $NumberOfYearsBack = 3
    $TargetOU = "NoLicense"
    $oldUsers = get-aduser -Filter * -SearchScope OneLevel -SearchBase "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Properties LastLogonDate | where-object {$_.name -match ".*(\(\d.*\))" -and $_.lastlogondate -lt (get-date).addyears("-$($NumberOfYearsBack)")}

    foreach($user in $oldUsers){
        Move-ADObject -TargetPath "OU=$($TargetOU),OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu" -Identity $user.samaccountname
    }
}