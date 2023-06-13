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

Import-Module ExchangeOnlineManagement

Import-module AzureADPreview

# Set path for log files:
$logPath = "D:\wpi\Logs\DYNManagement\Advising"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force
}

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
            $DynGroupMembers = Get-AzureADGroupMember -ObjectId $DynGroup.Id -all $true | Select-Object UserPrincipalName

            #$regex = '(?i)^(' + (($OptOutMembers | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$'

            # Edit this if we're doing Opt-Ins/Out - add/remove people depending on what opt-group they're in.
            $CorrectMembers = $DynGroupMembers # + $OptInMembers) -notmatch $regex

            # Get the current DistributionGroup members.
            $CurrentMembers = Get-DistributionGroupMember -Identity $DynList.DisplayName -ResultSize Unlimited | Select-Object @{N = 'UserPrincipalName'; E = { $_.primarysmtpaddress } }

            # If there's no one in the list, just add them all!
            if ($null -eq $CurrentMembers) {
                Write-Host -BackgroundColor GRAY -foregroundColor BLACK "No members currently in list, adding all."

                ForEach ($member in $CorrectMembers) {
                    Write-Host -foregroundColor GREEN "Added $($member.UserPrincipalName) to ADV-${Name}"
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
                Write-Host "ADV-${Name} currently has $($Currentmembers.count) members."

                Write-Host "ADV-${Name} should have $($CorrectMembers.count) members." 

                if ($null -eq $CorrectMembers) {
                    $RemoveMembers = $CurrentMembers | select-object userprincipalname
                }
                else {
                                   

                    # #reconcile lists 
                    $comparisons = Compare-Object $CurrentMembers $CorrectMembers -Property UserPrincipalName
                
                    # Store the users who should and shouldn't be in the lists in variables.
                    $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object userprincipalname
                    $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object userprincipalname
                }
                  
                # Iterate through the add/remove lists and do what's necessary.
                ForEach ($Removal in $RemoveMembers.userprincipalname) {
                    Write-Host -foregroundCOlor RED "Removed $($Removal) from ADV-${Name}"
                    $RemovedFromList += $Removal

                    if (-NOT $testMode) {
                        Remove-DistributionGroupMember -Identity $DynList.DisplayName -Member $removal -Confirm:$false
                    }
                }
                                
                ForEach ($Addition in $AddMembers.userprincipalname) {
                    Write-Host -foregroundColor GREEN "Added $($Addition) to ADV-${Name}"
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

    Write-Host "Prepping and sending email to $Recipients"
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

        Write-Host "Sending email about advising list changes for $name"

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
    Write-Host "Creating new lists for: $($name)"

    if (-NOT $testMode) {
        try {
            New-ADVMailEnabledSyncGroup -Name $Name                                                                                                                                     #Triggers function to turn an advisor's name into a distribution list
            New-ADVDynamicGroup -Name $Name -Rule "(user.extensionAttribute6 -match ""(.*;)*${name};.*"") and (user.accountEnabled -eq true) and (user.UserType -eq ""Member"") and (user.extensionattribute8 match ""(.*;)*student;.*"")"        #Triggers function to turn an advisor's name into a dynamic group
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
    $advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
    #$advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -match "PADV-(.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.replace("PADV-", "").replace("OADV-", "").split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
    #$advisors = Get-aduser -filter { Enabled -eq $true -and extensionAttribute6 -like 'PADV-*' } -property extensionAttribute6 | Where-Object { $_.extensionAttribute6 -match "PADV-[^ ](.+);*" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";")[0].split("-")[1] } | Sort-Object | Get-Unique
    
    Write-Host "Total advisors: $($advisors.count)"

    # Iterate through everything and make necessary changes.
    foreach ($Advisor in $advisors) {
        $advisor = $advisor.trim()
        Write-Host -backgroundColor GRAY -foregroundColor BLACK "`nChecking List for Advisor: $($advisor)"

        # If no AAD group exists, start from scratch.
        if ("Dyn-$($advisor)" -ne (get-AzureADMSGroup -filter "DisplayName eq 'Dyn-$advisor'").DisplayName) {
            # There was apparently an error with the previous line, so now we're catching only 1 result and using a normal -ne, we need to sanitize due to the wildcard nature of the search
            "No list for $($Advisor), creating..."
            New-WPIAdvisingListComponents $Advisor
        }
        else {
            "$($Advisor) already has list, syncing..."
            Sync-WPIDynlist $Advisor
        }
    }
    Remove-UnusedAdvisingLists $advisors
}

function Remove-UnusedAdvisingLists {
    [CmdletBinding()]
    param (
        $advisors
    )
    
    begin {
        $currentLists = Get-DistributionGroup ADV-* | select-object -ExpandProperty DisplayName | Foreach-Object { ($_.split("-"))[1] } | Sort-Object | Get-Unique
        if ($null -eq $advisors) {
            $advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
        }
        $diff = Compare-Object $currentLists $advisors # "<="" is people who should be deleted
    }
    
    process {
        foreach ($list in $diff) {
            if ($list.SideIndicator -eq "<=") {
                Write-Host "Deleting $($list.InputObject) for not existing in current list of advisors"
                Remove-DistributionGroup -Identity "ADV-$($list.InputObject)"
                Remove-AzureADMSGroup -DisplayName "Dyn-$($list.InputObject)"
            }
        
        }
    }
    
    end {
        
    }
}

function Repair-WPIAdvisingLists {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        Import-Module AzureADPreview -force
        $advisors = Get-aduser -filter { Enabled -eq $true -and extensionattribute7 -like "student" } -property extensionAttribute6, extensionattribute7 | Where-Object { $_.extensionAttribute6 -notmatch "PADV-(.+);*" -and $_.extensionAttribute6 -match ".*;" } | Select-Object -expandproperty extensionAttribute6 | Foreach-Object { $_.split(";") } | where-object { $_ -ne " " -and $_ -ne "" } | Sort-Object | Get-Unique
    }
    
    process {
        $counter = 0
        foreach ($advisor in $advisors) {
            $counter++
            $advisor = $advisor.trim()
            Write-Progress -Activity "Processing Lists" -CurrentOperation $advisor -PercentComplete (($counter / $advisors.count) * 100)
            #"(user.extensionAttribute6 -match ""PADV-(.*; )*${name}; .*"") and (user.accountEnabled -eq true)"
            $ID = (Get-AzureADMSGroup -Filter "DisplayName eq 'Dyn-$advisor'").Id
            Write-Host "Rewriting Membership rule for user $advisor with ID: $ID"
            Set-AzureADMSGroup -Id $ID -MembershipRule "(user.extensionAttribute6 -match ""(.*;)*${advisor};.*"") and (user.accountEnabled -eq true) and (user.UserType -eq ""Member"") and (user.extensionattribute8 match ""(.*;)*Student;.*"")"

        }
    }
    
    end {
        
    }
}

#############################
#                           #
#       MAIN PROGRAM        #
#                           #
#############################

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_AdvisingMgmt.log" -Force

if (!$LoadFunctions) {
    $credentials = $null
    $credPath = "D:\wpi\XML\$env:UserName.xml"

    if (Test-path $credPath) {
        $credentials = Import-CliXml -Path $credPath
        
    }
    else {
        Write-Warning "Saved Exchange Credentials not Found. Please enter new credentials."
        $credentials = Get-Credential

        $save = (Read-Host "Would you like to encrypt and save these credentials for future use? [Y | n]").toLower()

        if ($save -eq "y") {
            $credentials | Export-CliXml -Path $credPath

            Write-Host "`nEncrypted credentials saved as: $credPath"
        }
        else {
            Write-Host -foregroundColor YELLOW "`nCredentials not saved.`n"
        }
    }
    try {
        ## Load Exchange Online
        #$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
        #Import-PSSession $ExchangeOnlineSession
        Connect-ExchangeOnline -credential $credentials -ShowBanner:$false
        Connect-AzureAD -Credential $credentials
        Write-Host -foregroundColor GREEN "Done."
    }
    catch {
        Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate with saved credentials.`n"
        Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
        exit
    }
    # Exchange login blows up logs because of the URL it gives back in the MOTD - Start logging here instead.
    
    Sync-WPIAdvisingLists

}

Stop-Transcript
