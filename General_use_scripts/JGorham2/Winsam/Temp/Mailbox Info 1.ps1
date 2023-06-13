Clear-Host

$alias = 'calendar_marketing'

#Domain Controller Information
$global:dcs = (Get-ADDomainController -Filter *)
$global:dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
$global:dchostname = $dc.HostName
        
#Get Mailbox Information
$mailbox = Get-Mailbox $alias -DomainController $dchostname -ErrorAction SilentlyContinue
$Inbox = "$($mailbox.Name):\Inbox" 
$Calendar = "$($mailbox.Name):\Calendar" 

$out           = $null    
        
$MailboxUsers  = $null
$InboxUsers    = $null
$CalendarUsers = $null
$Senders       = $null
        
$MailboxUserlist        = @()
$SendOnBehalfUserList   = @()
$SendAsUserList         = @()
$InboxUserlist          = @()
$CalendarUserlist       = @()

$MailboxUsers  = Get-MailboxPermission $alias -DomainController $dchostname | Where {$_.IsInherited -ne $true -and $_.User -notlike "NT AUTHORITY\SELF"}
#$SendOnBehalf  = $mailbox.GrantSendOnBehalfTo
#$SendAs        = Get-Mailbox $alias | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")}
#$InboxUsers    = Get-MailboxFolderPermission -Identity $Inbox -DomainController $dchostname
$CalendarUsers = Get-MailboxFolderPermission -Identity $Calendar -DomainController $dchostname

#$CalendarUsers | ft FolderName,User,Identity,AccessRights,IsValid
   
<#

foreach($user in $InboxUsers) {
    $username = $null
    $ADInfo   = $null
            
    $out = New-Object PSObject
    if ($user.User -eq 'Default' -or $user.User -eq 'Anonymous') {
        $out | add-member noteproperty Name $user.user
        $out | add-member noteproperty Username $null
        $out | add-member noteproperty Department $null
        $out | add-member noteproperty Enabled $true

        }
    else {
        if ($user.User -match "NT User:") {$username = $user.User.Replace("NT User:ADMIN\","")}
        else {$username = (Get-Mailbox $user.User -DomainController $dchostname -ErrorAction SilentlyContinue).alias}
        $ADInfo = Get-ADUser $username -Properties Department -Server $dchostname

        $out | add-member noteproperty Name $ADInfo.Name
        $out | add-member noteproperty Username $ADInfo.samAccountName
        $out | add-member noteproperty Department $ADInfo.Department
        $out | add-member noteproperty Enabled $ADInfo.Enabled
    }
    $out | add-member noteproperty AccessRights $user.AccessRights

    $InboxUserlist += $out        
    }#>

foreach($user in $CalendarUsers) {
    $username   = $null
    $ADInfo     = $null
    $GroupInfo  = $null
    $ObjectType = $null
    
    Write-Host "Processing $($user.User)"
    
    $CalendarUserlist += WinSam-Get-ObjectInfo $user.User
    }


$CalendarUserlist | Where {$_.Enabled -eq $true} | Select Name,Username,Department,AccessRights | Sort AccessRights,Name | FT -AutoSize -Wrap | Out-Default

