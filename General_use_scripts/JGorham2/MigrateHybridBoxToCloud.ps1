function Move-HybridBoxToCloud {
    [CmdletBinding()]
    param (
        $Mailbox
    )
    #Import Sync tools
    
    Import-Module "C:\Users\jmgorham2_prv\WPI\General_use_scripts\JGorham2\AAD-Sync.ps1"

    # Get Existing rules

    $ExistingMailbox = Get-mailbox -identity $Mailbox

    $mailboxRules = Get-InboxRule -mailbox $existingmailbox.ExchangeGuid

    #Get details of the mailbox you want to restore

    $ExistingMailbox = Get-Mailbox -SoftDeletedMailbox -Identity $Mailbox 

    #Create New Cloud Only box

    $NewMailbox = New-Mailbox -shared -name $ExistingMailbox.name -DisplayName $ExistingMailbox.displayName -Alias $ExistingMailbox.alias

    #Create Restore with the details found above

    New-MailboxRestoreRequest -SourceMailbox $ExistingMailbox.ExchangeGuid -TargetMailbox $NewMailbox.ExchangeGuid -AllowLegacyDNMismatch

    #Check status of the restore

    while((Get-MailboxRestoreRequest -TargetMailbox $NewMailbox.ExchangeGuid) -ne "Completed"){
        Write-Host "Waiting for job to complete"
        Start-Sleep -Seconds 10
    }
    Write-Host "Job Complete"

    # Make the email addresses match
    set-mailbox $NewMailbox.ExchangeGuid -EmailAddresses $existingmailbox.EmailAddresses

    #readd rules
    set-inboxrule -
}