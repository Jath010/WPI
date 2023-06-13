<#
.SYNOPSIS 
    DynamicListManagement.ps1 does the automatic syncing and creation of Dynamic emailing lists.

.DESCRIPTION
    If we want to create a new Dynamic DL we can, provided we give the proper args. Otherwise, we will sync current ones.

    A "mailing list" will comprise of:
        The DL (where the mail is sent to)
        The Dynamic Group containing all the users fitting a particular description
        The Opt-in group (if needed)
        The Opt-out group (if needed)
    
    So the only important aspect of the opt group creation is making sure they match the naming scheme for later searching

    The logic will be to act on a list of descriptions, get the list, check existence, gather the lists of dyn and opt in, subtract opt out then sync the DL
    
    Dynamic Rule Example: (user.extensionAttribute6 -match ".*acsabuncu.*") and (user.accountEnabled -eq true)
#>

Param (
    # Reports but doesn't actually make changes.
    [Switch] $testMode,

    # Used to setup brand new lists.
    [Switch] $createNew,
    [Switch] $addOptInList,
    [Switch] $addOptOutList,
    [Switch] $convertDL,

    # Specify based on how we're running this.
    [Switch] $autoSync,
    [Switch] $manualSync,
    [Switch] $syncOptsWithDL,
    [Switch] $dontSyncOpts,

    # Used to run more quickly without needing prompts.
    [String] $dlName,
    [String] $DYNfilter,
    [Switch] $noADCheck,
    [String] $ADfilter,

    # May not end up using these
    [Switch] $removeOptInList,
    [Switch] $removeOptOutList
)

# Set path for log files:
$logPath = "D:\wpi\Logs\DYNManagement\Standing"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Clean out logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | del -Force
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_DYNMgmt.log" -Force

if ($testMode) {
    Write-Warning "Test Mode Enabled, no actual changes will be made."
}


function test-Existing {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        $listName,
        [Parameter(Position = 0, Mandatory = $true)]
        $listType
    )

    # Depending on what kind of list we're checking, we need to do things a little differently.
    if ($listType -eq "AzureAD") {
        $list = (Get-AzureADMSGroup -SearchString $listname -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $listname })
    }
    elseif ($listType -eq "DL") {
        # This is here because dl-employees is dumb.
        if ($listname -eq "DL-employees") {
            $list = get-DistributionGroup "dl-allemployees"
        }
        else {
            $list = get-DistributionGroup $listname
        }        
    }
    elseif ($listType -eq "DynamicDL") {
        $list = Get-DynamicDistributionGroup -Identity $listName -ErrorAction SilentlyContinue
    }


    if ($list) {
        return $true
    }
    else {
        return $false
    }

}

#Create a dynamic list
#needs to take a simple description 
function New-WPIDynamicGroup {
    [CmdletBinding()]
    param (
        $Name,
        $Rule
    )

    # Create a new Dynamic Group that will add members based on a filter. 
    Write-Host "Creating new Dynamic Group: DYN-$($Name)"
    if (-NOT $testMode) {
        # If the list already exists, let the user know and give them a choice. 
        if (test-Existing -listname "Dyn-$($name)" -listType "AzureAD") {
            Write-Warning "Dyn-$($name) already exists!"
            $overwrite = (Read-Host "Overwrite existing? [Y|n]").toLower()

            if ($overwrite -eq "y") {
                Write-Host -foregroundColor RED "Deleting and Re-creating Dyn-$($name)"
                $dyn = Get-AzureADMSGroup -SearchString "Dyn-$($name)"  | Where-Object { $_.DisplayName -eq "Dyn-$($name)" }
                $dyn | Set-AzureADMSGroup "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$name" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
            }
            else {
                Write-Host -foregroundColor CYAN "Dyn-$($name) has not been altered."
            }
        }
        else {
            New-AzureADMSGroup -DisplayName "Dyn-$($name)" -Description "Dynamic group created from PS" -MailEnabled $False -MailNickName "Dyn-$($name)" -SecurityEnabled $True -GroupTypes "DynamicMembership" -MembershipRule $Rule -MembershipRuleProcessingState "On"
        }
    }

}

function New-WPIOptInList {
    [CmdletBinding()]
    param (
        $Name
    )
    # Create a new AzureAD Group that people can be manually added/removed. 
    Write-Host "Creating new AzureADGroup: OptIn-$($Name)"
    if (-NOT $testMode) {
        # If the list already exists, let the user know and give them a choice. 
        if (test-Existing -listname "OptIn-$($name)" -listType "AzureAD") {
            Write-Warning "OptIn-$($name) already exists!"
            $overwrite = (Read-Host "Overwrite existing? [Y|n]").toLower()

            if ($overwrite -eq "y") {
                Write-Host -foregroundColor RED "Deleting and Re-creating OptIn-$($name)"
                $dyn = Get-AzureADMSGroup -SearchString "OptIn-$($name)"  | Where-Object { $_.DisplayName -eq "OPTIn-$($name)" }
                $dyn | Set-AzureADMSGroup -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
            }
            else {
                Write-Host -foregroundColor CYAN "OptIn-$($name) has not been altered."
            }
        }
        else {
            New-AzureADMSGroup -DisplayName "OptIn-$($name)" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
        }
    }
}
function New-WPIOptOutList {
    [CmdletBinding()]
    param (
        $Name
    )

    # Create a new AzureAD Group that people can be manually added/removed. 
    Write-Host "Creating new AzureADGroup: OptOut-$($Name)"
    if (-NOT $testMode) {
        # If the list already exists, let the user know and give them a choice. 
        if (test-Existing -listname "OptOut-$($name)" -listType "AzureAD") {
            Write-Warning "OptOut-$($name) already exists!"
            $overwrite = (Read-Host "Overwrite existing? [Y|n]").toLower()

            if ($overwrite -eq "y") {
                Write-Host -foregroundColor RED "Deleting and Re-creating OptOut-$($name)"
                $dyn = Get-AzureADMSGroup -SearchString "OptOut-$($name)"  | Where-Object { $_.DisplayName -eq "OPTOut-$($name)" } 
                $dyn | Set-AzureADMSGroup -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
            }
            else {
                Write-Host -foregroundColor CYAN "OptOut-$($name) has not been altered."
            }
        }
        else {
            New-AzureADMSGroup -DisplayName "OptOut-$($name)" -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
        }
    }
}
function New-WPIMailEnabledSyncGroup {
    [CmdletBinding()]
    param (
        $Name
    )

    # Create the Distribution List that emails will be sent to.  
    Write-Host "Creating new Distribution List: DL-$($Name)"
    if (-NOT $testMode) {
        # If the list already exists, let the user know and give them a choice. 
        if (test-Existing -listname "DL-$($name)" -listType "DL") {
            Write-Warning "DL-$($name) already exists!"
            $overwrite = (Read-Host "Overwrite existing? [Y|n]").toLower()

            if ($overwrite -eq "y") {
                Write-Host -foregroundColor RED "Deleting and Re-creating DL-$($name)"
                Remove-DistributionGroup -identity "DL-$($name)" -confirm:$false
                New-DistributionGroup -Name "DL-$($Name)" -Alias "DL-$($Name)" -Type Distribution -PrimarySmtpAddress "Dl-$($name)@wpi.edu"
            }
            else {
                Write-Host -foregroundColor CYAN "DL-$($name) has not been altered."
            }
        }
        else {
            New-DistributionGroup -Name "DL-$($Name)" -Alias "DL-$($Name)" -Type Distribution -PrimarySmtpAddress "Dl-$($name)@wpi.edu"
        }
    }
}
function Sync-OldWithNew {
    [CmdletBinding()]
    param (
        $name,
        $rule,
        $filter, 
        $optIn,
        $optOut
    )
    <# 
    .SYNOPSIS
        Okay this is a bit much but need to put it here. 
    .DESCRIPTION
        We need to:
         - Get all users who are currently in the existing DL. 
         - Get the proper Dynamic Filter setup for the new list.
         - Compare/Contrast current DL entries with new filter.
         - Auto add users to OptIn/Out to reflect current DL
         - That way, the new Dynamic list exists but it has already 
           taken into account the old one and anyone who had Rich manually opt in/out.
    #>

    # Get who's currently in the DL
    # This is because dl-allemployees is dumb.
    if ($name -eq "employees") {
        $currentDLMembers = Get-DistributionGroupMember -identity "DL-all$($name)" -ResultSize Unlimited | Select-Object -expandProperty PrimarySMTPAddress
    }
    else {
        $currentDLMembers = Get-DistributionGroupMember -identity "DL-$($name)" -ResultSize Unlimited | Select-Object -expandProperty PrimarySMTPAddress
    }
    Write-Host "`nMembers in Current DL: $($currentDLMembers.count)"

    # Get who should be in the DL based on the Get-ADUser filter.
    $allADMembers = (Get-ADuser -filter "$filter")
    Write-Host "Members Found With Filter: $($allADMembers.count)"

    # Diff the results so we see who is missing/added for each. 
    $diffResults = Compare-Object $allADMembers.UserPrincipalName $currentDLMembers -PassThru
    Write-Host "Total Differences: $($diffResults.count)"
                    
    $AddMembers = $diffResults | Where-Object { $_.SideIndicator -eq '=>' }
    Write-Host "Members to be Opted-IN: $($AddMembers.count)"

    $RemoveMembers = $diffResults | Where-Object { $_.SideIndicator -eq '<=' }
    Write-Host "Members to be Opted-OUT: $($RemoveMembers.count)"

    # Now that we've gathered everything, let's start setting things up!
    if (-NOT $testMode) {
        # I know we do this elsewhere and I should save the code, but I really just wanted everything in here since it was special.
        if ($addOptInList) {
            New-WPIOptInList $Name
            $optInCreated = $true
        }
        else {
            # Make sure before continuing.
            Write-Host -foregroundColor Black -BackgroundColor Gray "`nOpt-In list not specified."
            $youSure = (Read-Host "Create Opt-In List? [Y|n]").toLower()

            if ($youSure -eq "y") {
                New-WPIOptInList $Name
                $optInCreated = $true
            }
        }
        if ($addOptOutList) {
            New-WPIOptOutList $Name
            $optOutCreated = $true
        }
        else {
            # Make sure before continuing.
            Write-Host -foregroundColor Black -BackgroundColor Gray "`nOpt-Out list not specified."
            $youSure = (Read-Host "Create Opt-Out List? [Y|n]").toLower()

            if ($youSure -eq "y") {
                New-WPIOptOutList $Name
                $optOutCreated = $true
            }
        }

        # If we created either, add who we need to add. 
        if ($optInCreated -and (-NOT $dontSyncOpts)) {
            $DynList = Get-AzureADMSGroup -SearchString "OptIn-$($name)"  | Where-Object { $_.DisplayName -eq "OPTIn-$($name)" }

            ForEach ($Addition in $AddMembers) {
                # We have to do this because the compare-object stripped the property away earlier.
                $userToAdd = (Get-AzureADUser -SearchString "$($addition)")

                Write-Host -ForegroundColor GREEN "Adding Member to Opt-In: $($addition)"
                Add-AzureADGroupMember -ObjectID $DynList.ID -RefObjectId $userToAdd.ObjectID                              
            }
        }

        if ($optOutCreated -and (-NOT $dontSyncOpts)) {
            $DynList = Get-AzureADMSGroup -SearchString "OptOut-$($name)"  | Where-Object { $_.DisplayName -eq "OPTOut-$($name)" }

            ForEach ($Removal in $RemoveMembers) {
                # We have to do this because the compare-object stripped the property away earlier.
                $userToRemove = (Get-AzureADUser -SearchString "$($Removal)")

                Write-Host -foregroundColor RED "Adding Member to Opt-Out: $($removal)"                
                Add-AzureADGroupMember -ObjectId $DynList.ID -RefObjectId $userToRemove.ObjectID
            }
        }

        # Now that the opt-in/out lists are set, we need to make the new DYN. 
        New-WPIDynamicGroup -Name $Name -Rule $Rule

        # Note: We don't need to create the DL since it already exists, and we don't need to sync it since we just diff'd everything against it.
    }
}
function Sync-OPTsWithDL {
    [CmdletBinding()]
    param (
        $name
    )
    <# 
    .SYNOPSIS
        Sync an existing DYN and its OPT IN/OUT with the current DL. 
    .DESCRIPTION
        We needed this because we ended up making all the DYN and OPT groups ahead of time. 
        That means we then needed to sync those with the existing DLs before moving forward.
        Otherwise, everyone from the DYN would be added to the DL and we don't want that.

        We need to:
         - Get all users who are currently in the existing DL. 
         - Compare/Contrast current DL entries with who's in the DYN.
         - Auto add users to OptIn/Out to reflect current DL
    #>

    Write-Host -foregroundColor CYAN "`nSyncing List: $($name)"

    # Get who's currently in the DL
    # This is because dl-allemployees is dumb.
    if ($name -eq "employees") {
        $currentDLMembers = Get-DistributionGroupMember -identity "DL-all$($name)" -ResultSize Unlimited | Select-Object -expandProperty PrimarySMTPAddress
    }
    else {
        $currentDLMembers = Get-DistributionGroupMember -identity "DL-$($name)" -ResultSize Unlimited | Select-Object -expandProperty PrimarySMTPAddress
    }
    Write-Host "Members in Current DL: $($currentDLMembers.count)"

    # Get who should be in the DL based on who's in the DYN
    $dyn = (Get-AzureADMSGroup -SearchString "DYN-$($name)" | Where-Object { $_.DisplayName -eq "DYN-$($name)" })
    $dynMembers = (Get-AzureADGroupMember -ObjectId $dyn.ID -all $true)
    Write-Host "Members in Current DYN: $($dynMembers.count)"

    if ($dynMembers.count -le 1) {
        Write-Warning "DYN list has no members, please check rule. No sync will be performed." 
    }
    elseif ($null -eq $currentDLMembers -or $currentDLMembers.count -lt 1) {
        Write-Warning "DL has no members, please double-check DYN accuracy before adding."
    }
    else {
        # Diff the results so we see who is missing/added for each. 
        $diffResults = Compare-Object $dynMembers.UserPrincipalName $currentDLMembers -PassThru
        Write-Host "Total Differences: $($diffResults.count)"
                        
        $AddMembers = $diffResults | Where-Object { $_.SideIndicator -eq '=>' }
        Write-Host "Members to be Opted-IN: $($AddMembers.count)"

        $RemoveMembers = $diffResults | Where-Object { $_.SideIndicator -eq '<=' }
        Write-Host "Members to be Opted-OUT: $($RemoveMembers.count)"

        # Get our OPT lists so we can work on them.
        $inList = Get-AzureADMSGroup -SearchString "OptIn-$($name)" | Where-Object { $_.DisplayName -eq "OPTIn-$($name)" }
        $outList = Get-AzureADMSGroup -SearchString "OptOut-$($name)"  | Where-Object { $_.DisplayName -eq "OPTOut-$($name)" }

        # Now that we've gathered everything, let's start setting things up!
        ForEach ($Addition in $AddMembers) {
            # We have to do this because the compare-object stripped the property away earlier.
            $userToAdd = (Get-AzureADUser -SearchString "$($addition)")

            Write-Host -ForegroundColor GREEN "Adding Member to Opt-In: $($addition)"
            if (-NOT $testMode) {
                try {
                    Add-AzureADGroupMember -ObjectID $inList.ID -RefObjectId $userToAdd.ObjectID
                }
                catch {
                    Write-Host -foregroundColor RED -backgroundColor BLACK "Unable to add user."
                } 
            }                         
        }

        ForEach ($Removal in $RemoveMembers) {
            # We have to do this because the compare-object stripped the property away earlier.
            $userToRemove = (Get-AzureADUser -SearchString "$($Removal)")

            Write-Host -foregroundColor RED "Adding Member to Opt-Out: $($removal)"
            if (-NOT $testMode) {
                try {
                    Add-AzureADGroupMember -ObjectId $outList.ID -RefObjectId $userToRemove.ObjectID
                }
                catch {
                    Write-Host -foregroundColor RED -backgroundColor BLACK "Unable to add user."
                }
            }              
        }
    }
}
function Sync-WPIDynlist {
    [CmdletBinding()]
    param (
        $name
    )
    Write-Host -foregroundColor CYAN "`nChecking List: $($name)"
    # Null these out each time just in case
    $OptIn = $OptOut = $DynGroup = $DynList = $null
    # Get any lists associated with the list we're searching for.
    if (test-Existing -listname "OptIn-$($name)" -listType "AzureAD") {
        $OptIn = get-AzureADMSGroup -SearchString "OptIn-$($name)" -ErrorAction SilentlyContinue  | Where-Object { $_.DisplayName -eq "OPTIn-$($name)" }
    }
    if (test-Existing -listname "OptOut-$($name)" -listType "AzureAD") {
        $OptOut = get-AzureADMSGroup -SearchString "OptOut-$($name)" -ErrorAction SilentlyContinue  | Where-Object { $_.DisplayName -eq "OPTOut-$($name)" }
    }
    if (test-Existing -listname "DYN-$($name)" -listType "AzureAD") {
        $DynGroup = get-AzureADMSGroup -SearchString "Dyn-$name" -ErrorAction SilentlyContinue  | Where-Object { $_.DisplayName -eq "DYN-$($name)" }
    }
    if (test-Existing -listname "DL-$($name)" -listType "DL") {
        # This is because dl-allemployees is dumb.
        if ($name -eq "employees") {
            $DynList = get-DistributionGroup -Identity "DL-all$($name)"
        }
        else {
            $DynList = get-DistributionGroup -Identity "DL-$($name)"
        }
    }

    # Get the members of each group (but only if they exist)
    if ($null -ne $OptIn) {
        $OptInMembers = Get-AzureADGroupMember -ObjectId $OptIn.Id -all $true | Select-Object UserPrincipalName
        Write-Host -foregroundColor YELLOW "OptIn-$($name) members: $($OptInMembers.count)"
    }
    else {
        $OptIn = ""
    }
    if ($null -ne $OptOut) {
        $OptOutMembers = Get-AzureADGroupMember -ObjectId $OptOut.Id -all $true | Select-Object UserPrincipalName
        Write-Host -foregroundColor YELLOW "OptOut-$($name) members: $($OptOutMembers.count)"
    }
    else {
        $OptOut = ""
    }
    if ($null -ne $DynGroup) {
        $DynGroupMembers = Get-AzureADGroupMember -ObjectId $DynGroup.Id -all $true | Select-Object UserPrincipalName
        Write-Host -foregroundColor YELLOW "DYN-$($name) members: $($DynGroupMembers.count)"
    }
    else {
        # If this doesn't exist, we MAY have a problem, so let the user know. 
        Write-Warning "No DYN-$($name) exists, make sure this is what you want (it could be, I dunno)."
    }

    # Make a scary regex that contains the names of all the people who should be opted out.
    $regex = '(?i)^(' + (($OptOutMembers | ForEach-Object { [regex]::escape($_) }) -join "|") + ')$'

    # Use the scary regex to filter out the people who shouldn't be on t he list.
    $CorrectMembers = ([array]$DynGroupMembers + $OptInMembers) -notmatch $regex | Select-Object UserPrincipalName -Unique

    $CurrentMembers = Get-DistributionGroupMember -Identity $DynList.DisplayName -ResultSize Unlimited | Select-Object @{N = 'UserPrincipalName'; E = { $_.primarysmtpaddress } }
    Write-Host -foregroundColor YELLOW "DL-$($name) members: $($CurrentMembers.count)"

    if ($null -eq $CurrentMembers -and $null -eq $CorrectMembers) {
        Write-Host -foregroundColor YELLOW "No members currently in list, but no members to add."
    }
    elseif ($null -eq $CurrentMembers) {    
        Write-Host -foregroundColor YELLOW "No members currently in list, adding all."
        ForEach ($member in $CorrectMembers) {
            if (-NOT $testMode) {
                Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $member.UserPrincipalName
            }
        }
    }
    else {
        # #reconcile lists 
        $comparisons = Compare-Object $CurrentMembers $CorrectMembers -Property UserPrincipalName
                    
        $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object userprincipalname
        $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object userprincipalname

        Write-Host -foregroundColor GREEN "Members to add: $($AddMembers.count)"
        Write-Host -foregroundColor RED "Members to remove: $($RemoveMembers.count)"
            
        ForEach ($Removal in $RemoveMembers.userprincipalname) {
            Write-Host -foregroundColor RED "Removing Member: $removal"
            if (-NOT $testMode) {
                Remove-DistributionGroupMember -Identity $DynList.DisplayName -Member $removal -Confirm:$false
            }
        }
                        
        ForEach ($Addition in $AddMembers.userprincipalname) {
            Write-Host -ForegroundColor GREEN "Adding Member: $addition"
            if (-NOT $testMode) {
                try {
                    Add-DistributionGroupMember -Identity $DynList.DisplayName -Member $Addition
                }
                catch {
                    Write-host -ForegroundColor Red "User $addition already in group"
                }
            }               
        }
    }
}

function Get-ExchangeConnection($credentials) {
    Write-Host "Trying to login to Exchange with credentials..."
 
    try {
        ## Load Exchange Online
        Connect-ExchangeOnline -Credential $credentials -ShowBanner:$false
        Connect-AzureAD -Credential $credentials
        Write-Host -foregroundColor GREEN "Done."
    }
    catch [Exception] {
        $_.Exception.Message
             
        Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate with saved credentials.`n"
        Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
        exit
    }
}

function checkWhatWeAreDoing {
    # If we want to make a new list, go for it!
    iF ($createNew) {

        Write-Host -foregroundColor YELLOW "`nNOTE: Please enter ONLY the name - naming schemes will be automatically applied."
        Write-Host "AKA: If you enter 'testDL' the actual DL will be named 'DL-testDL' and Dynamic will be 'DYN-testDL'"
        $name = Read-Host "Desired List Name:"

        # We want to define our rule before we continue, so they are 100% certain they know what they're getting into.
        Write-Host -foregroundColor CYAN "Please Specify the Dynamic List Rule you would like to use."
        $rule = Read-Host ":"

        # If we need to make an optin/out, do it.
        if ($addOptInList) {
            New-WPIOptInList $Name  
        }
        else {
            # Make sure they didn't forget to specify this.
            Write-Host -foregroundColor Black -BackgroundColor Gray "Opt-In list not specified."
            $youSure = (Read-Host "Create Opt-In List? [Y|n]").toLower()

            if ($youSure -eq "y") {
                New-WPIOptInList $Name  
            }
        }
        if ($addOptOutList) {
            New-WPIOptOutList $Name 
        }
        else {
            # Make sure they didn't forget to specify this.
            Write-Host -foregroundColor Black -BackgroundColor Gray "Opt-Out list not specified."
            $youSure = (Read-Host "Create Opt-Out List? [Y|n]").toLower()

            if ($youSure -eq "y") {
                New-WPIOptOutList $Name 
            }
        }

        # Do this no matter what.
        New-WPIMailEnabledSyncGroup $Name
        New-WPIDynamicGroup -Name $Name -Rule $Rule

        
    }
    elseif ($convertDL) {
        if ($null -eq $dlName) {
            # Converting from an old one is a bit more work, but get what we need first.
            Write-Host -foregroundColor YELLOW "`nNOTE: Please enter ONLY the name - naming schemes will be automatically applied."
            Write-Host "AKA: If you enter 'testDL' the actual DL will be named 'DL-testDL' and Dynamic will be 'DYN-testDL'`n"
            $dlName = (Read-Host "Desired List Name")
        }

        if ($null -eq $DYNfilter) {
            # We want to define our rule before we continue, so they are 100% certain they know what they're getting into.
            Write-Host -foregroundColor CYAN "`nPlease Specify the Dynamic List Rule you would like to use."
            $DYNfilter = (Read-Host " ")
        }

        if ($null -eq $ADFilter -and (-NOT $noADCheck)) {
            # We want to define our AD filter before we continue, so we can compare/contrast our results.
            Write-Host -foregroundColor CYAN "`nPlease Specify the Get-ADUser Filter you would like to use."
            $ADfilter = ((Read-Host "(Enabled Check added Automatically)") + " -and Enabled -eq 'true'")
        }

        # Pass all our parameters so we set things up right.
        if ($noADCheck) {
            Sync-OldWithNew -name $dlName -rule $DYNfilter -optIn $addOptInList -optOut $addOptOutList
        }
        else {
            Sync-OldWithNew -name $dlName -rule $DYNfilter -filter $ADfilter -optIn $addOptInList -optOut $addOptOutList
        }
    }
    elseif ($manualSync) {
        # Manually run the sync on a single entity.
        Write-Host -ForegroundColor YELLOW "`nPut listname only, not DL- or DYN-"
        $name = Read-Host "Listname to Sync"
        Sync-WPIDynlist -name $name
    }
    elseif ($autoSync) {
        # Every list SHOULD have an OptIn list, so search for those to get our targets. 
        $allLists = (Get-AzureADMSGroup -SearchString "OptIn-" -all $true | Select-Object -expandProperty DisplayName)

        Write-Host "Found $($allLists.count) lists to sync."
        foreach ($list in $allLists) {
            # Pull our listname out to make it easier.
            $listName = $list.split("-", 2)[1]
            # Send it!
            Sync-WPIDynlist -name $listName
        }
    }
    elseif ($syncOptsWithDL) {
        # Every list SHOULD have an OptIn list, so search for those to get our targets. 
        $allLists = (Get-AzureADMSGroup -SearchString "OptIn-" -all $true | Select-Object -expandProperty DisplayName)

        Write-Host "Found $($allLists.count) lists to sync."
        foreach ($list in $allLists) {
            # Pull our listname out to make it easier.
            $listName = $list.split("-", 2)[1]
            # Send it!
            Sync-OPTsWithDL -name $listName
        }
    }
}

#############################
#                           #
#       MAIN PROGRAM        #
#                           #
#############################

             
$credential = $null
#$credPath = "$PSScriptRoot\$env:UserName.xml"
$credpath = "D:\wpi\XML\exch_automation\exch_automation@wpi.edu.xml"
 
if (Test-path $credPath) {
    $credential = Import-CliXml -Path $credPath
 
    Get-ExchangeConnection($credential)
       
    checkWhatWeAreDoing
}
else {
    Write-Warning "Saved Exchange Credentials not Found. Please enter new credentials."
    $credential = Get-Credential
 
    $save = (Read-Host "Would you like to encrypt and save these credentials for future use? [Y|n]").toLower()
 
    if ($save -eq "y") {
        $credential | Export-CliXml -Path $credPath
 
        Write-Host "`nEncrypted credentials saved as: " -nonewLine 
        Write-Host -foregroundColor CYAN $credPath
    }
    else {
        Write-Host -foregroundColor YELLOW "`nCredentials not saved.`n"
    }
 
    Get-ExchangeConnection($credential)

    checkWhatWeAreDoing
 
}
 


Stop-Transcript