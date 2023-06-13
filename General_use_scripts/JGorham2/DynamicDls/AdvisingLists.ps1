<#
So to start
we need to detect if the groups exist for each one we want to create
if they don't exist we create them

so the list will be
    The DL
        The Dynamic Group containing all the users fitting a particular description
        The Opt-in group
        The Opt-out group
    
    so the only important aspect of the opt group creation is making sure they match the naming scheme for later searching

the logic will be to act on a list of descriptions, get the list, check existence, gather the lists of dyn and opt in, subtract opt out then sync the DL

#>

#Create a dynamic list
#needs to take a simple description 

Param (
    [Switch] $testMode,
    [Switch] $sendEmails,
    [Switch] $LoadFunctions
)

# Set path for log files:
$logPath = "$PSScriptRoot\Logs"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_AdvisingMgmt.log" -Force

if ($testMode) {
    "WARNING: Script running in test mode, no actual changes will be made."
}

function New-ADVDynamicGroup {
    # Creates a new Dynamic Group in AAD that has a filter for all users with extensionAttribute6 including the advisor name.
    [CmdletBinding()]
    param (
        $Name,
        $Rule
    )
    process {
        "Creating new list for $name"
        if (-NOT $testMode) {
            New-AzureADMSGroup -DisplayName "Dyn-$name" -Description "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$name" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
        }
    }

}
function New-WPIOptInList {
    # Creates a new AzureADGroup that allows join/leave as needed. Those in this group will be ADDED to the mail list.
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        if (-NOT $testMode) {
            New-AzureADGroup -DisplayName "OptIn-${Name}" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet" 
        }
    }
}
function New-WPIOptOutList {
    # Creates a new AzureADGroup that allows join/leave as needed. Those in this group will be REMOVED to the mail list.
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        if (-NOT $testMode) {
            New-AzureADGroup -DisplayName "OptOut-${Name}" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
        }
    }
}
function New-ADVMailEnabledSyncGroup {
    # Creates a new Exchange Distribution Group that will be what handles the email.
    [CmdletBinding()]
    param (
        $Name
    )
    process {
        if (-NOT $testMode) {
            New-DistributionGroup -Name "ADV-${Name}" -Alias "ADV-${Name}" -Type Distribution -PrimarySmtpAddress "ADV-${Name}@wpi.edu"
            Set-DistributionGroup -Identity "ADV-${Name}" -AcceptMessagesOnlyFrom "${Name}@wpi.edu" -HiddenFromAddressListsEnabled $true 
        }
    }
}
function Sync-WPIDynlist {
    [CmdletBinding()]
    param (
        $name
    )
    # Make sure these are blank each iteration.
    $RemovedFromList = @()
    $AddedToList = @()

    #$OptIn = get-AzureADGroup -SearchString "OptIn-${Name}"
    #$OptOut = get-AzureADGroup -SearchString "OptOut-${Name}"
    
    # Get our DistributionGroup and DynamicGroup based on the advisor name.
    $DynList = get-DistributionGroup -Identity "ADV-${Name}" -ErrorAction SilentlyContinue
    $DynGroup = get-AzureADMSGroup -SearchString "Dyn-$name" -ErrorAction SilentlyContinue | where-object { $_.Displayname -eq "Dyn-$name" }

    # If either of the lists don't exist, something is wrong and we need to know.
    if ($null -eq $DynList) {
        Write-Host -foregroundColor RED "`nERROR: Dynamic AAD Group for $name do not exist, please verify and re-run script.`n"
    } 
    elseif ($null -eq $DynGroup) {
        Write-Host -foregroundColor RED "`nERROR: Distribution Group for $name do not exist, please verify and re-run script.`n"
    }
    else {
        try {
            # Reset these each time. 
            $AddedToList = @()
            $RemovedFromList = @()

            #$OptInMembers = Get-AzureADGroupMember -ObjectId $OptIn.ObjectId | Select-Object UserPrincipalName
            #$OptOutMembers = Get-AzureADGroupMember -ObjectId $OptOut.ObjectId | Select-Object UserPrincipalName

            # Get all the current members of the DynamicGroup
            $DynGroupMembers = Get-AzureADGroupMember -ObjectId $DynGroup.Id | Select-Object UserPrincipalName

            #$regex = '(?i)^(' + (($OptOutMembers | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$'

            # Edit this if we're doing Opt-Ins/Out - add/remove people depending on what opt-group they're in.
            $CorrectMembers = $DynGroupMembers # + $OptInMembers) -notmatch $regex

            # Get the current DistributionGroup members.
            $CurrentMembers = Get-DistributionGroupMember -Identity $DynList.DisplayName | Select-Object @{N = 'UserPrincipalName'; E = { $_.primarysmtpaddress } }

            # If there's no one in the list, just add them all!
            if ($null -eq $CurrentMembers) {
                Write-Host -BackgroundColor GRAY -foregroundColor BLACK "No members currently in list, adding all."

                ForEach ($member in $CorrectMembers) {
                    Write-Host -foregroundColor GREEN "Adding $($member.UserPrincipalName) to ADV-${Name}"
                    $AddedToList += $member.UserPrincipalName

                    if (-NOT $testMode) {
                        Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $member.UserPrincipalName
                    }
                }

                if ($sendEmails) {
                    emailTheAdvisor -Name $name -addedToList $AddedToList -fullList $CorrectMembers
                }
            }
            # If the list has members, sync them.
            else {
                Write-Host -foregroundColor CYAN "ADV-${Name} " -noNewLine
                Write-Host "currently has " -nonewLine
                Write-Host -foregroundColor CYAN $Currentmembers.count -noNewLine 
                Write-Host " members."

                Write-Host -foregroundColor CYAN "ADV-${Name} " -noNewLine
                Write-Host "should have " -nonewLine
                WRite-Host -foregroundColor CYAN $CorrectMembers.count -noNewLine 
                Write-Host " members."

                # #reconcile lists 
                $comparisons = Compare-Object $CurrentMembers $CorrectMembers -Property UserPrincipalName
                
                # Store the users who should and shouldn't be in the lists in variables.
                $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object userprincipalname
                $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object userprincipalname
                    
                # Iterate through the add/remove lists and do what's necessary.
                ForEach ($Removal in $RemoveMembers.userprincipalname) {
                    Write-Host -foregroundCOlor RED "Removing $($Removal) from ADV-${Name}"
                    $RemovedFromList += $Removal

                    if (-NOT $testMode) {
                        Remove-DistributionGroupMember -Identity $DynList.DisplayName -Member $removal -Confirm:$false
                    }
                }
                                
                ForEach ($Addition in $AddMembers.userprincipalname) {
                    Write-Host -foregroundColor GREEN "Adding $($Addition) to ADV-${Name}"
                    $AddedToList += $Addition

                    if (-NOT $testMode) {
                        Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $Addition
                    }               
                }

                # Finally, email the advisor the changes.
                if ($sendEmails) {
                    emailTheAdvisor -name $name -addedToList $AddedToList -RemovedFromList $RemovedFromList -fullList $CorrectMembers 
                }
                
            }
        }
        catch [Exception] {
            $_.Exception.Message
        }
    }
}
function emailTheAdvisor {
    [CmdletBinding()]
    param (
        $name,
        $addedToList,
        $RemovedFromList,
        $fullList
    )

    #Set email information.
    $HeaderFrom = 'its@wpi.edu'
    $SMTPServer = 'smtp.wpi.edu'

    if ($testMode) {
        $Recipients = 'sgemme@wpi.edu'
    }
    else {
        $Recipients = "$($name)@wpi.edu"
    }

    Write-Host "Prepping and sending email to " -NoNewline
    Write-host -ForegroundColor CYAN $Recipients
    #-----------------------------------------------------------------------------------------------
    # Null out our message stuff before filling it.
    $messageParameters = $null
    $Subject = "ADV-$($name) Changes Were Made"
    $Body = $null

    if ($null -ne $RemovedFromList) {
        # Add a new paragraph to the email.
        $Body += "<p>The following accounts were <b>Removed</b> from your advisory list:</p>"
        # Start a new table.
        $Body += "<ul>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $RemovedFromList) {
            $Body += "    <li>$($user)</li>"
        }

        # Close our list.
        $Body += "</ul>"
    }
    if ($null -ne $AddedToList) {
        # Add a new paragraph to the email.
        $Body += "<p>The following accounts were <b>Added</b> to your advisory list:</p>"
        # Start a new table.
        $Body += "<ul>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $AddedToList) {
            $Body += "    <li>$($user)</li>"
        }

        # Close our list.
        $Body += "</ul>"
    }

    # If we have a message to send, send it, but only after adding final details!
    if ($null -ne $Body) {
        # Add a new paragraph to the email.
        $Body += "<p>Your current, updated list:</p>"
        # Start a new table.
        $Body += "<ul>"

        # Iterate through the list of MissingLinuxAccounts and add them all to a table.
        foreach ($user in $fullList.userPrincipalName) {
            $Body += "    <li>$($user)</li>"
        }

        # Close our list.
        $Body += "</ul>"

        Write-Host "Sending email about advising list changes for" -NoNewline
        Write-host -ForegroundColor CYAN $name

        $messageParameters = @{
            Subject    = $Subject
            Body       = $Body
            From       = $HeaderFrom
            To         = $Recipients
            SmtpServer = $SMTPServer
            Priority   = "Low"
        }
        $Body
        $Body | Out-File "$LogPath\$name-Email.html"
        Send-MailMessage @messageParameters -BodyAsHtml
    }
}

function New-WPIAdvisingListComponents {
    [CmdletBinding()]
    param (
        $Name
    )
    #New-WPIOptInList $Name
    #New-WPIOptOutList $Name
    Write-Host "Creating new lists for: " -nonewLine
    Write-Host -foregroundColor GREEN $($name)

    if (-NOT $testMode) {
        try {
            New-ADVMailEnabledSyncGroup -Name $Name
            New-ADVDynamicGroup -Name $Name -Rule "(user.extensionAttribute6 -match "".*${name}.*"") and (user.accountEnabled -eq true)"  
        }
        catch [Exception] {
            $_.Exception.Message
        }
    }
}

function Sync-WPIAdvisingLists {
    param (
        
    )

    # Do some fun regex to get all current advisors by means of parsing everyone who has an advisor.
    "Getting list of all advisors..."
    $advisors = Get-aduser -filter { Enabled -eq $true -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6 | Where-Object { $_.extensionAttribute6 -match "PADV-[^ ](.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";")[0].split("-")[1] } | Sort-Object | Get-Unique
    
    Write-Host "Total advisors: $($advisors.count)"

    # Iterate through everything and make necessary changes.
    foreach ($Advisor in $advisors) {
        Write-Host -backgroundColor GRAY -foregroundColor BLACK "Checking List for Advisor: $($advisor)"

        # If no AAD group exists, start from scratch.
        if ("Dyn-$($advisor)" -ne (get-AzureADMSGroup -filter "startswith(MailNickname,'Dyn-$advisor')" -top 1).DisplayName) {      #It's necessary to check these results in order to makes sure that the search isn't returning a different list dues to it's wildcard nature
            "No list for $($Advisor), creating..."
            New-WPIAdvisingListComponents $Advisor
        }
        else {
            "$($Advisor) already has list, syncing..."
            Sync-WPIDynlist $Advisor
        }
    }
}


#############################
#                           #
#       MAIN PROGRAM        #
#                           #
#############################

if (!$LoadFunctions) {
    $credentials = $null
    $credPath = "$PSScriptRoot\$env:UserName.xml"

    if (Test-path $credPath) {
        $credentials = Import-CliXml -Path $credPath
        
    }
    else {
        Write-Warning "Saved Exchange Credentials not Found. Please enter new credentials."
        $credentials = Get-Credential

        $save = (Read-Host "Would you like to encrypt and save these credentials for future use? [Y|n]").toLower()

        if ($save -eq "y") {
            $credentials | Export-CliXml -Path $credPath

            Write-Host "`nEncrypted credentials saved as: " -nonewLine 
            Write-Host -foregroundColor CYAN $credPath
        }
        else {
            Write-Host -foregroundColor YELLOW "`nCredentials not saved.`n"
        }
    }
    try {
        ## Load Exchange Online
        $ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
        Import-PSSession $ExchangeOnlineSession
        Connect-AzureAD -Credential $credentials
        Write-Host -foregroundColor GREEN "Done."
    }
    catch {
        Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate with saved credentials.`n"
        Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
        exit
    }
        
    Sync-WPIAdvisingLists

}

Stop-Transcript
