$boxes = Get-Mailbox -ResultSize unlimited
$sanitizedBoxes = $Boxes|where{$_.MaxSendSize -ne "150 MB (157,286,400 bytes)"}
foreach($box in $sanitizedBoxes){
    Set-Mailbox $box.alias -MaxReceiveSize 150MB -MaxSendSize 150MB -Verbose
}