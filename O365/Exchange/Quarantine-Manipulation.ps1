#this half exists for the sake of making it obvious what commands to use

function Release-WPIquarantine {
    [CmdletBinding()]
    param (
        $SenderAddress
    )
    
    $messages = Get-QuarantineMessage -SenderAddress $SenderAddress -PageSize 500 -ReleaseStatus NOTRELEASED

    $messages | Release-QuarantineMessage -ReleaseToAll -ReportFalsePositive -AllowSender

}