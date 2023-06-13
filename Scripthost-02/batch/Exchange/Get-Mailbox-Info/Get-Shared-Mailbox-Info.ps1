 $resources = Get-Mailbox -filter {isResource -eq 'true' -and Name -notLike "Calendar*" -and Alias -notLike "calendar*" -and Alias -notLike "cal_*"}
 $shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -filter {Name -notLike "Calendar*" -and Alias -notLike "calendar*" -and Alias -notLike "cal_*"}

$date = (Get-Date).ToString('MM-dd-yyyy_HHmm')
$resourcesPath = $PSScriptRoot + "\Resource-Mailbox-Stats-$date.csv"
$sharedPath = $PSScriptRoot + "\Shared-Mailbox-Stats-$date.csv"

[System.Collections.ArrayList]$resourcesTotals = @()
[System.Collections.ArrayList]$sharedTotals = @()


 foreach ($mailbox in $shared){
     Write-Host "Getting Mailbox Stats for Shared Mailbox: " -noNewLine
     Write-Host -foregroundColor CYAN $mailbox.alias -noNewLine
     Write-Host "..." -noNewLine
     $alias = $mailbox.Alias
     $sent = Get-MessageTrace -SenderAddress $alias@wpi.edu -StartDate (Get-Date).addDays(-10) -endDate (Get-Date) -PageSize 5000
     $received = Get-MessageTrace -RecipientAddress $alias@wpi.edu -StartDate (Get-Date).addDays(-10) -endDate (Get-Date) -PageSize 5000

     if ($null -ne $sent) {
         $sentCount = $sent.count
     }
     else {
         $sentCount = 0
     }

     if ($null -ne $received) {
         $receivedCount = $received.count
     }
     else {
         $receivedCount = 0
     }
     $entry = [PSCustomObject]@{
                    Identity    = $mailbox.Identity
                    Alias       = $mailbox.Alias
                    Sent        = $sentCount
                    Received    = $receivedCount
                }

    $sharedTotals.add($entry) | Out-Null
    Write-Host -foregroundColor GREEN "Done."
 }

 $sharedTotals | Export-CSV -Path $sharedPath