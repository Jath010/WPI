<#
.SYNOPSIS
   AdvisingListManagement.ps1 - Created to automagically manage the new dynamic maling lists.

.DESCRIPTION
    This script is meant to run regularly so lists are kept up to date. It checks AD for all qualified users
who have primary advisors. It creates/updates dynamic lists/filters appropriately so that those who should
be getting emails for/from these advisors will be. 

.NOTES
    Created By: Stephen Gemme
    Created On: 01/05/2021

    *All modifications should be recorded in Git.

    Check AD Extension mapping with: Get-AzureADUserExtension

#>

# Set path for log files:
$logPath   = "D:\wpi\batch\Exchange\DYNManagement\Logs"

# Get date for logging and file naming:
$date      = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Get our arrays ready.
[System.Collections.ArrayList]$Global:enabledPADV = @()
[System.Collections.ArrayList]$Global:disabledPADV = @()

# We make this global so it can be used in multiple functions without having to pass it around.
$Global:usersWithPADV

# Start recording what we're doing.
Start-Transcript -Append -Path "$($logPath)\$($datestamp)_Transcript.log" -Force

function loginToEchange($credentials){
   Write-Host "Trying to login to Exchange with credentials..."
   $ExchangeOnlineSession=$null

   try {
      ## Load Exchange Online
      #$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
      #Import-PSSession $ExchangeOnlineSession
      Connect-ExchangeOnline -Credential $credentials
      Connect-AzureAD -Credential $credentials
      Write-Host -foregroundColor GREEN "Done."
      
      getPADVInfo
   }
   catch [Exception] {
      $_.Exception.Message
            
      Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate with saved credentials.`n"
      Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
      exit
   }
}

# Simply check if the list exists or not.
function test-Existing {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true)]
        $listName
    )

    $list =  Get-DynamicDistributionGroup -Identity $listName -ErrorAction SilentlyContinue

    if ($list){
        return $true
    }
    else {
        return $false
    }

}

function getPADVInfo(){
    <# 
    .SYNOPSIS 
        Let's start by doing the gross thing to get a list of all currently listed advisors we see in AD. 
     
    .NOTES 
        This may be changed later for a more updated version provided by the Advising Office. 

        This can all be done in a single line, but I broke it up so someone 5 years from now doens't hate me (including myself).
        Here's the one-liner: 
        Get-aduser -filter {Enabled -eq $true -and extensionAttribute6 -like 'PADV-*'} -property extensionAttribute6 | Where-Object {$_.extensionAttribute6 -match "PADV-[^ ](.+);*"} | Select-Object -expandproperty extensionAttribute6 | Foreach-Object {$_.split(";")[0].split("-")[1]} | Sort | Get-Unique | foreach-object {Get-ADUser $_ | Where-Object {$_.Enabled -eq $false}}
    #>

    # Get our list of users who have a primary advisor. 
    # For the Regex-noobs, "PADV-[^ ](.+);*" translates to: Get everything that starts with "PADV-", doesn't have a space after that [^ ], has at least 1 character (.+), then a semi colon ;, then is followed by anything *
    Write-Host "Getting all users who are enabled with Primary Advisors..." -noNewLine
    $Global:usersWithPADV = (Get-aduser -filter {Enabled -eq $true -and extensionAttribute6 -like 'PADV-*'} -property extensionAttribute6 | Where-Object {$_.extensionAttribute6 -match "PADV-[^ ](.+);*"} )
    $found = $Global:usersWithPADV.count
    Write-Host -foregroundColor GREEN "Found $found active users with a Primary Advisor."

    # Extract the names with some fancy doodads.
    Write-Host "Getting all unique advisors and checking their status..." -noNewLine
    $AllPADVs = ($Global:usersWithPADV | Select-Object -expandproperty extensionAttribute6 | Foreach-Object {$_.split(";")[0].split("-")[1]})

    # Get just the unique ones so we know how many there are.
    $UniquePADVs = ($AllPADVs | Sort | Get-Unique)

    # Get relevant data for each advisor and store it all in an array.
    # We do the Try-Catch because Get-ADUser doesn't handle -ErrorAction and we need to add users who don't exist to the disabled report.
    foreach ($name in $UniquePADVs){
        try {
            # Get our AD info on hte advisor.
            $info = Get-ADUser $name

            $entry = [PSCustomObject]@{
                        UserName    = $name
                        First       = $info.GivenName
                        Last        = $info.Surname
                        Email       = $info.UserPrincipalName
                        Advisees    = ($AllPADVs | Group | Where-Object {$_.Name -eq $name} | Select-Object -expandProperty Count)
                    }

            if ($info.enabled){
                $Global:enabledPADV.Add($entry) | Out-Null
            }
            else {
                $Global:disabledPADV.Add($entry) | Out-Null
            }
            
        }
        catch {
            # They don't exist in AD.
            $entry = [PSCustomObject]@{
                        UserName    = $name
                        First       = "Not Found In AD"
                        Last        = "Not Found In AD"
                        Email       = "Not Found In AD"
                        Advisees    = ($AllPADVs | Group | Where-Object {$_.Name -eq $name} | Select-Object -expandProperty Count)
                    }

            $Global:disabledPADV.Add($entry) | Out-Null
        }
    }

    Write-Host -foregroundColor GREEN "Done.`n"

    Write-Host -foregroundColor CYAN "Enabled Primary Advisors ($($enabledPADV.count))"
    $Global:enabledPADV | Format-Table

    Write-Host -foregroundColor CYAN "Disabled Primary Advisors ($($disabledPADV.count))"
    $Global:disabledPADV | Format-Table

    # Now that we're done gathering/displaying data, time to check the dynamic lists.
    make-changes -enable

    # Call this again to deal with the disbaled list.
    make-changes
}

function make-changes(){
    [CmdletBinding()]
    Param(
        [Switch] $enable
    )

    <#
    .SYNOPSIS
        Create or Delete lists based on our list of PADVs.

    .NOTES 
        Use Stephen's DistributionList-Manager.ps1 to update lists, not this.
    #>

    # Iterate through our PADVs and check which ones need to be created or deleted (existing lists *should* always remain the same)
    if ($enable){
        foreach ($entry in $Global:enabledPADV){
            # Extract our username so it's easier to reference.
            $username = $entry.Username

            # Derive our list name from the username.
            $listname = "$username-advising@wpi.edu"

            if (-NOT (test-Existing -listName $listName)){
                # List doesn't exist but it needs to.
                Write-Host -foregroundColor YELLOW "DYN Not Found for Enabled Advisor: $username"
                # Create the new list with an appropriate name based on the username.
                new-List -username $username -listname $listname
            }
            else {
                # List exists and ought to
                Write-Host -foregroundColor GREEN "DYN Found for Enabled Advisor: $username"
            }
        }
    }
    else {
        foreach ($entry in $Global:disabledPADV){
            
            if (test-Existing -listName $listName){
                # List exists but shouldn't
                Write-Host -foregroundColor YELLOW "DYN Found for Disabled Advisor: $username"
                delete-List -listname $listname
            }
            else {
                # List shouldn't exist and doesn't.
                Write-Host -foregroundColor GREEN "DYN Not Found for Disabled Advisor: $username"
            }
        }
    }
    
}

function update-list {
    [CmdletBinding()]
    Param(
        [Parameter(Position=1, Mandatory=$true)]
        $listname
    )

    # Get a list of who was in the DL before we edit it.
    $preDLmembers = Get-DistributionGroupMember -Identity $listname -ResultSize unlimited | Select-Object -ExpandProperty PrimarySMTPAddress
    # Before we do anything, let's diff the input and current lists so we don't bother if there's no changes to be made.
    # Diff returns $null if they are equal, but IF sees null as $false, so only act if it's not null. Seems backwards but it works.
    # If there are currently no members in the list, that means we add them all.
    $diff = Compare-Object -ReferenceObject @($preDLMembers | Select-Object) -DifferenceObject @($memberEmailList | Select-Object)         
        
    if ($diff) {
        # Legacy, not sure why we're logging this.
        $jsonCount = ($InputObject  | select -ExpandProperty members | Measure-Object).count

        # Do a little logging for logging's sake.
        $preDLcount = $preDLmembers.count
        Write-Output "Member total: $preDLcount","JSON file total: $jsonCount", $dlAlias, $preDLmembers | `
            Out-File "$($logPath)\$($datestamp)_Sync-DLmembership_$dlAlias-before ADD.txt"

        # If we have more than 30 changes to make, it's faster if we just replace the entier thing instead of going 1 at a time.
        if ($diff.count -gt 30){
        
            # New, faster, DL member sync - just replace them instead of adding and removing each individual one.
            Update-DistributionGroupMember -Identity $dlAlias -Members $memberEmailList -confirm:$false

            # Display what differences were found/changed.
            foreach ($change in $diff){
                if ($change.SideIndicator -eq "<="){
                    Write-Host -foregroundColor RED "Removed $($change.InputObject)"
                }
                elseif ($change.SideIndicator -eq "=>") {
                    Write-Host -foregroundColor GREEN "Added $($change.InputObject)"
                }
            }
        }
        else {
            # We have under 30 changes to make, do them 1 at a time since it's quicker on large lists.
            foreach ($change in $diff){
                # As of 09/01/2020 we've needed to start adding users by full email to avoid errors.
                if ($change.SideIndicator -eq "<="){
                    Remove-DistributionGroupMember -Identity $dlAlias -Member "$($change.InputObject)" -confirm:$false
                    Write-Host -foregroundColor RED "Removed $($change.InputObject)"
                }
                elseif ($change.SideIndicator -eq "=>") {
                    Add-DistributionGroupMember -Identity $dlAlias -Member "$($change.InputObject)" -confirm:$false
                    Write-Host -foregroundColor GREEN "Added $($change.InputObject)"
                }
            }
        }

        # Log the things!
        $postDLcount = $memberEmailList.count
        Write-Output "Member total: $postDLcount","JSON file total: $jsonCount", $dlAlias, $memberEmailList | `
            Out-File "$($logPath)\$($datestamp)_Sync-DLmembership_$dlAlias-after REMOVAL.txt"
        $changesMade | Out-File -append "$($logPath)\$($datestamp)_Sync-DLmembership_$DL-CHANGES.txt"
    }
    else {
        Write-Host "No changes to be made for dl-$($InputObject.displayname)"
    }
}

function new-List {

    <#
        .NOTES Potentially useful parameters to consider.
            -AcceptMessagesOnlyFrom
            -AcceptMessagesOnlyFromDLMembers
            -AcceptMessagesOnlyFromSendersOrMembers
            -GrantSendOnBehalfTo
            -MailTip
            -MaxSendSize
            -RejectMessagesFrom
            -RejectMessagesFromDLMembers
            -RejectMessagesFromSendersOrMembers 
    #>

    [CmdletBinding()]
    Param(
        # We take the username here because we need it for the filter also.
        [Parameter(Position=0, Mandatory=$true)]
        $username,

        [Parameter(Position=1, Mandatory=$true)]
        $listname
    )

    # We already checked if it exists, so just do the thing.
    try {
        Write-Host "Creating new DL: " -nonewLine
        Write-Host -foregroundColor CYAN "$listName"

        # Make sure it's hidden from the address book if we're testing.
        if ($testMode){
            #Set-DynamicDistributionGroup -Identity $listName -HiddenFromAddressListsEnabled:$true
        }

        # If we get this far, check to make sure it worked.
        if (test-Existing -listName $listName){
            Write-Host -foregroundColor GREEN "`nList Successfully Created!"
        }
        else {
            Write-Host -foregroundColor YELLOW "Something unexpected happened, please try again."
        }

    }
    catch [Exception] {
        $_.Exception.Message
    }
}

function delete-List{
    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true)]
        $listName
    )

    try {
        Remove-DynamicDistributionGroup -identity $listName -confirm:$false
        Write-Host -foregroundColor GREEN "Done."
    }
    catch [Exception] {
        $_.Exception.Message
    }
}

#############################
#                           #
#       MAIN PROGRAM        #
#                           #
#############################

# If we are alreayd logged onto Exchange, no need to try again
try{ 
   # If we can get a random, known list, we can get them all.
   Write-Host "Testing connection to Exchange..." -nonewLine
   if(Get-DistributionGroup "dl-fuller"){
      Write-Host -foregroundColor GREEN "Done."
      getPADVInfo
   }
}
catch [Exception] {
   $_.Exception.Message
            
   $credential = $null
   $credPath = "$PSScriptRoot\$env:UserName.xml"

   if (Test-path $credPath){
      $credential = Import-CliXml -Path $credPath

      loginToEchange($credential)
      
   }
   else {
      Write-Warning "Saved Exchange Credentials not Found. Please enter new credentials."
      $credential = Get-Credential

      $save = (Read-Host "Would you like to encrypt and save these credentials for future use? [Y|n]").toLower()

      if ($save -eq "y"){
         $credential | Export-CliXml -Path $credPath

         Write-Host "`nEncrypted credentials saved as: " -nonewLine 
         Write-Host -foregroundColor CYAN $credPath
      }
      else {
         Write-Host -foregroundColor YELLOW "`nCredentials not saved.`n"
      }

      loginToEchange($credential)

   }
}

# Stop recording what we're doing.
Stop-Transcript