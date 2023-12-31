function WinSam-MailResource ($EmailAddress) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    #Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    #. $ScriptPath\WinSam.3.GlobalVariables.ps1
    #. $ScriptPath\WinSam.3.Functions.ps1
    #. $ScriptPath\WinSam.3.Info.ps1
    #. $ScriptPath\WinSam.3.Menu.ps1

    #Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    #Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    while (!$EmailAddress -or $EmailAddress -notmatch '@' -or $EmailAddress -notmatch 'wpi.edu') {
        Write-Host ''
        $EmailAddress = read-host -Prompt 'Please enter an email address (ex: alias@wpi.edu)'
        Clear-Host
        #Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        #Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-MailResourceInfo $EmailAddress
    #WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    #WinSam-Menu-CheckEmail    
    }


function WinSam-Get-MailResourceInfo ($EmailAddress) {
    $RecipientInfo=$null

    if ($EmailAddress -notmatch '@' -or $EmailAddress -notmatch 'wpi.edu') {Write-Host "Please enter a valid WPI email address";WinSam-CheckEmail}

    $RecipientInfo = Get-Recipient $EmailAddress -ErrorAction SilentlyContinue
    if (!$RecipientInfo) {
        Write-Host ''
        Write-Host ''
        Write-Host "     ERROR: [$EmailAddress] There is no such Exchange object." -ForegroundColor Red
        Write-Host "     Please contact a Mail Administrator for more information." -ForegroundColor Red
        Return
        }

    <#
    RecipientType
    -------------
    MailUser     : GuestMailUser, MailUser
    UserMailbox  : UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox, DiscoveryMailbox

    MailContact  : MailContact

    MailUniversalDistributionGroup : GroupMailbox, MailUniversalDistributionGroup, RoomList
    MailUniversalSecurityGroup     : MailUniversalSecurityGroup
    DynamicDistributionGroup       : DynamicDistributionGroup
    #>

    switch ($RecipientInfo.RecipientType) {
        'UserMailbox' {WinSam-Get-MailboxInfo $alias}
        'MailContact' {
            Write-Host ''
            $Contact = Get-MailContact $Alias
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header 'NOTE: This is a mail contact, not a mailbox.' $MenuLength $MenuIndent) -Foregroundcolor Black -BackgroundColor Yellow
            Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor Black -BackgroundColor Yellow
            Write-Host ''
            Write-Host (WinSam-Write-Header 'Mail Contact Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
            Write-Host (WinSam-Write-Header '' $MenuLength -Line)
            Write-Host ''
            Write-Host "Contact Display Name: $($Contact.DisplayName)"
            Write-Host "Contact Alias       : $($Contact.Alias)"
            Write-Host "Contact Target      : $($Contact.PrimarySmtpAddress)"
            Write-Host ''
            Return
            }
        'MailUniversalDistributionGroup' {
            if ($RecipientInfo.RecipientTypeDetails -eq 'GroupMailbox') {WinSam-Get-UnifiedGroupInfo $EmailAddress}
            if ($RecipientInfo.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') {WinSam-Get-DistributionGroupInfo $EmailAddress}
            }
        'MailUniversalSecurityGroup' {WinSam-Get-DistributionGroupInfo $EmailAddress}
        'DynamicDistributionGroup'   {WinSam-Get-DistributionGroupInfo $EmailAddress}
        default {
            ## Make a note of the type of recipient with a note to contact the mail administrator
            Write-Host 'default'
            }
        }
    }

function WinSam-UnifiedGroup ($Alias) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$Alias) {
        Write-Host ''
        Write-Host 'Please enter the email address for the O365 group.'
        Write-Host ''
        $Global:UnifiedGroupAlias = read-host -Prompt '   O365 Group Email Address'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-UnifiedGroupInfo $Global:UnifiedGroupAlias
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-UnifiedGroup
    }

function WinSam-Mailbox ($MBAlias) {
    if ($debug) {$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()}  ## DEBUG
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    if (!$MBAlias) {
        Write-Host ''
        $alias = read-host -Prompt 'Please enter a mailbox alias'
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        }
    WinSam-Get-MailboxInfo $alias
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    WinSam-Menu-Mailbox
    }