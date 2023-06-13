
$resource = Get-Mailbox -filter {isResource -eq 'true' -and name -notlike "Calendar -*" -and alias -notlike "cal*"}
$shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter {name -notlike "Calendar -*" -and alias -notlike "cal*"}
$starttime = (get-date).adddays(-10)
$endtime = (get-date)



foreach($rbox in $resource){
    <#
    $var = $rbox|Get-MailboxPermission|Where-Object{$_.IsInherited -ne "*"}
    if($var.count -eq 0){
        $rbox.identity | Export-Csv -NoTypeInformation -path C:\tmp\ResourcePermissions.csv -append
    }
    else{
        $var|Export-Csv -NoTypeInformation -path C:\tmp\ResourcePermissions.csv -append
    }
    #>
    $sent = (GET-MESSAgetrace -SenderAddress "${$rbox.Alias}@wpi.edu" -StartDate $starttime -EndDate $endtime -PageSize 5000|measure).count
    $Received = (GET-MESSAgetrace -RecipientAddress "${$rbox.Alias}@wpi.edu" -StartDate $starttime -EndDate $endtime -PageSize 5000|measure).count
    [PSCustomObject]@{
        Alias                = $rbox.Alias
        Sent                 = $sent
        Received             = $received
    }|Export-Csv -NoTypeInformation -path C:\tmp\ResourceTraffic.csv -append
}

foreach($sbox in $shared){
    <#
    $svar = $sbox|Get-MailboxPermission|Where-Object{$_.IsInherited -ne "*"}
    if($svar.count -eq 0){
        $sbox.identity | Export-Csv -NoTypeInformation -path C:\tmp\SharedPermissions.csv -append
    }
    else{
        $svar|Export-Csv -NoTypeInformation -path C:\tmp\SharedPermissions.csv -append
    }
    #>
    $sent = (GET-MESSAgetrace -SenderAddress "${$sbox.Alias}@wpi.edu" -StartDate $starttime -EndDate $endtime -PageSize 5000|measure).count
    $received = (GET-MESSAgetrace -RecipientAddress "${$sbox.Alias}@wpi.edu" -StartDate $starttime -EndDate $endtime -PageSize 5000|measure).count
    [PSCustomObject]@{
        Alias                = $sbox.Alias
        Sent                 = $sent
        Received             = $received
    }|Export-Csv -NoTypeInformation -path C:\tmp\SharedTraffic.csv -append
}