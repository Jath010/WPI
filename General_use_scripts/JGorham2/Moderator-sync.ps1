$senders = get-content C:\tmp\Senders.txt
$lists = get-content C:\tmp\MailingLists.txt

<#
Work order
Empty the fields
    ModeratedBy
    BypassModerationFromSendersOrMembers

#>

foreach($list in $lists){
    $DList = Get-DistributionGroup "dl-$list"
    $currentmoderators = $dlist.ModeratedBy
    $currentsenders = $DList.BypassModerationFromSendersOrMembers

    Write-Host "dl-$list"
    foreach($moderator in $currentmoderators){
        Write-host "Removing $moderator from ModeratedBy"
        Set-DistributionGroup "dl-$list" -ModeratedBy @{remove=$moderator} -Verbose
    }

    foreach($sender in $currentsenders){
        Write-host "Removing $sender from BypassModerationFromSendersOrMembers"
        Set-DistributionGroup "dl-$list" -BypassModerationFromSendersOrMembers @{remove=$sender} -Verbose
    }

    foreach($sender in $Senders){
        Write-Host "Adding Moderator ${Sender}"
        Set-DistributionGroup "dl-$list" -ModeratedBy @{add=$sender} -Verbose
        Write-Host "Adding Sender ${Sender}"
        Set-DistributionGroup "dl-$list" -BypassModerationFromSendersOrMembers @{add=$sender} -Verbose
    }
}

foreach($list in $lists){
    $DList = Get-DistributionGroup "dl-$list"
    $currentmoderators = $dlist.ModeratedBy
    $currentsenders = $DList.BypassModerationFromSendersOrMembers
    if($currentmoderators.count -ne $senders.count){
        Write-host $DList.Name Moderator Issue
    }
    if($currentsenders.count -ne $senders.count){
        Write-host $DList.Name Senders Issue
    }
}