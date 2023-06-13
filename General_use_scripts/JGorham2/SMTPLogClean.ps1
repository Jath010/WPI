function Remove-WPIMailLog {
    [CmdletBinding()]
    param (
        
    )

    $MailQueuePath = 'C:\inetpub\mailroot\Badmail'

    $List = Get-ChildItem $MailQueuePath -include *.BAD, *.BDP, *.BDR -Recurse
    foreach($item in $List){
        if ($item.CreationTime -lt (get-date).AddDays(-7)) {
            Remove-Item $item
        }
    }

}

Remove-WPIMailLog