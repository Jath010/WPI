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