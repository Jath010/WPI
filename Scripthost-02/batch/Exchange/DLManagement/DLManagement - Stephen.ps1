<#
.SYNOPSIS
   DLmgmt.ps1 - PowerShell functions to manage Office 365 distribution groups

.DESCRIPTION
This script adds eleven functions that can be used individually and outside of
this PS1 file, and one function (process-DL) that is only relevant to the other 
functions within this script.

.EXAMPLE
. .\DLmgmt.ps1

To import modules into the current session:
import-module .\DLmgmt.ps1

.NOTES
The following variables should be defined and updated accordingly prior to 
running this script:

    $jsonFiles, $logPath

    The JSON files live in \\storage.wpi.edu\dept\Information Technology\Strategic_Projects\it_proj_Unix_Mail_Transition\json-files\json-[date]

See 'DEFINE VARIABLES' section below.

Automate credentials by following instructions to use Paul Cunningham's Get-StoredCredential module:
https://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/

If you choose to use Paul Cunningham's Functions-PSStoredCredentials.ps1 script for
automating your connection to Office 365, you will also want to update the location of
this script in the 'CONNECT TO OFFICE 365' section below.

Alternatively, you can manually create a new PS Exchange session by capturing credentials and passing them as an argument, then opening a new session, like so:

$credentials = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking

#>

Param (
   [Switch] $testMode
)


#################################
#                               #
#     CONNECT TO OFFICE 365     #
#                               #
#################################

Clear-Host
#

##################################################################################################################
## Powershell Load Credentials
##################################################################################################################

Add-Type -assembly "system.io.compression.filesystem"

write-output `r`n

################################
#                              #
#       DEFINE VARIABLES       #
#                              #
################################

Write-Host -BackgroundColor Black -ForegroundColor Magenta `
"$((get-date).ToShortTimeString()): Defining variables" `r`n

# Set path for log files:
$logPath        = "D:\wpi\batch\Exchange\DLManagement\Logs"
$logArchivePath = "D:\wpi\batch\Exchange\DLManagement\Logs-Archive"

# Get date for logging and file naming:
$date      = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")
$timestamp = $date.ToString("yyyy-MM-dd - HH:mm")

# Set the path to search for today's files.
$jsonData = @()
if ($testMode){
   $jsonPath = "$PSScriptRoot\JSONData\Test-Data"
}
else {
   $jsonPath = "\\storage-02.wpi.edu\dept\Information Technology\Strategic_Projects\it_proj_Unix_Mail_Transition\json-files\json-" + (Get-Date).toString("yyyy-MM-dd")
}
   
# Now that we're ready, start logging, but only if we're not in test mode - Actual Program starts below the functions.
if (-NOT $testMode) {

   Start-Transcript -Append -Path "$($logPath)\$($datestamp)_TRANSCRIPT-DLmgmt.log" -Force
}

Write-Host "Begin file"

#############################
#                           #
#    PRIMARY FUNCTIONS      #
#                           #
#############################

Write-Host -BackgroundColor Black -ForegroundColor Magenta `
"$((get-date).ToShortTimeString()): Loading functions" `r`n

# Can probably do away with this in favor of checking as I go rather than making it its own function.
function Test-ExistingDL {
    <#
    .DESCRIPTION
      Test-ExistingDL - Checks input for an existing DL.  If one exists, it skips it.
      If not, it feeds the input into the New-DLfromJSON function.

    .EXAMPLE
      $jsonData | Test-ExistingDL

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
      Write-Host -BackgroundColor DarkGray -ForegroundColor White `
      "$((get-date).ToShortTimeString()): Checking to see if dl-$($InputObject.DisplayName) exists"

      $TestGroup = Get-DistributionGroup "dl-$($inputobject.DisplayName)" -ErrorAction silentlycontinue

      # CORE TASK: UPDATES GROUP IF GROUP EXISTS
      if($TestGroup) {
         Write-Host -BackgroundColor DarkGreen -ForegroundColor White `
         "$((get-date).ToShortTimeString()): $TestGroup already exists, updating..."

         $InputObject | Sync-DLmembership

         Write-Host `r`n
      }
      
      # CORE TASK: FEEDS INPUT INTO 'New-DLfromJSON' FUNCTION
      if(!($TestGroup)) {
         ForEach-Object{$TestGroup} {
            $InputObject | New-DLfromJSON
         }
      }
   }
}

function New-DLfromJSON {
    <#
    .DESCRIPTION
    New-DLfromJSON - Create a new Office 365 distribution group from JSON source data

    .EXAMPLE
    $jsonData | new-DLfromJSON

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
      Write-Host -BackgroundColor black -ForegroundColor green `
      "$((get-date).ToShortTimeString()): Creating Office 365 DL: dl-$($InputObject.DisplayName)"

      # Get just the data we want.
      $members = @()
      $members = $InputObject | select -ExpandProperty members 

      $ErrorAction = "SilentlyContinue"

      # Create our splatted info for the new list.
      $newDLInfo = @{
         Name                                = "dl-$($InputObject.Name)"
         DisplayName                         = "dl-$($InputObject.DisplayName)"
         Alias                               = "dl-$($InputObject.Alias)"
         Type                                = $InputObject.Type
         PrimarySmtpAddress                  = "dl-$($InputObject.DisplayName)@wpi.edu"
         ManagedBy                           = $InputObject.ManagedBy
         ModerationEnabled                   = $($InputObject.ModerationEnabled)
         ModeratedBy                         = $InputObject.ModeratedBy
         Members                             = $jMembers
         Notes                               = $InputObject.Notes
         RequireSenderAuthenticationEnabled  = $($InputObject.RequireSenderAuthenticationEnabled)
         IgnoreNamingPolicy                  = $true
         MemberJoinRestriction               = $InputObject.MemberJoinRestriction
         MemberDepartRestriction             = $InputObject.MemberDepartRestriction
         ErrorAction                         = $ErrorAction
      }

      $createDL = New-DistributionGroup @newDLInfo

         # CORE TASK
         if($createDL) {
            $InputObject | Set-AcceptMessagesOnlyFromSendersOrMembers -ErrorAction $ErrorAction
            $InputObject | Set-HiddenFromAddressListsEnabled -ErrorAction $ErrorAction
            $InputObject | Add-DLproxyAddress -ErrorAction $ErrorAction
            $InputObject | Set-BypassModerationFromSendersOrMembers -ErrorAction $ErrorAction
            Write-Host `r`n
                  
            # DATA GATHERING / LOGGING
            $newDLemail = Get-DistributionGroup "dl-$($InputObject.DisplayName)" | select -ExpandProperty primarysmtpaddress
            $newDLmembers = Get-DistributionGroupMember "dl-$($InputObject.DisplayName)" -ResultSize Unlimited | select -ExpandProperty alias

            Write-Output "$($timestamp): `r`nNew DL: dl-$($inputobject.DisplayName)`r`nPrimarySMTPAddress: $newDLemail`r`nMembers: $newDLmembers" `r`n | `
            Out-File -Append "$($logPath)\$($datestamp)_new-DLfromJSON.log"
            # END DATA GATHERING
            }
         else {
            Write-Host -BackgroundColor red -ForegroundColor white `
            "$((get-date).ToShortTimeString()): $($InputObject.DisplayName) is missing mandatory values and was not created" `r`n
            }
   }
}

function Sync-DLmembership {
   <#
    .DESCRIPTION
    Sync-DLmembership - Synchronizes the membership of a DL from JSON source data.

    .EXAMPLE
    $jsonData | Sync-DLmembership

    .NOTES
    In an effort to save time with so many lists, I've updated this function to use the 
    'Update-DistributionList' cmdlet for lists with over 30 changes because, instead of 
    iterating through the list 1 user at a time and adding/removing, we simply replace 
    the entire list with the new one in 1 fell swoop.

    Testing has shown that the time-to-complete for the entier script has been reduced by about 50%.

     - Stephen

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
      Write-Host -BackgroundColor Blue -ForegroundColor White `
      "$((get-date).ToShortTimeString()): Synchronizing dl-$($InputObject.displayname) membership"

      # Validate that the members field exists in the JSON
      if(-NOT $InputObject.Members) {
         Write-Host -BackgroundColor yellow -ForegroundColor black `
         "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'Members' field or it is empty"
               
         $ErrorActionPreference = 'silentlycontinue'
      }
      else {
         # We have data, let's act on it. First, null everything out each time we're here in case it cached (it shouldn't).
         $diff = $null
         $additions = $null
         $removals = $null

         # As of 09/01/2020 we've needed to start adding users by full email to avoid errors, especially with aliases.
         $memberEmailList = @()
         Foreach ($member in ($InputObject | Select-Object -ExpandProperty members)){
            $memberEmailList += "$member@wpi.edu"
         }
         # Get an alias of the group so we don't have to recalculate it's name each time.
         $dlAlias = Get-DistributionGroup "dl-$($InputObject.displayname)" -ResultSize unlimited | Select-Object -ExpandProperty alias
         # Get a list of who was in the DL before we edit it.
         $preDLmembers = Get-DistributionGroupMember -Identity $dlAlias -ResultSize unlimited | Select-Object -ExpandProperty PrimarySMTPAddress
         # Before we do anything, let's diff the input and current lists so we don't bother if there's no changes to be made.
         # Diff returns $null if they are equal, but IF sees null as $false, so only act if it's not null. Seems backwards but it works.
         # If there are currently no members in the list, that means we add them all.
         $diff = Compare-Object -ReferenceObject @($preDLMembers | Select-Object) -DifferenceObject @($memberEmailList | Select-Object)

            #$diff = Compare-Object $preDLmembers $memberEmailList
         
        
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
         # If there's less than 30 but still stuff in it, push them 1 at a time. 
         elseif ($diff) {

         }
         else {
            Write-Host "No changes to be made for dl-$($InputObject.displayname)"
         }
      }
   }
}


function Set-AcceptMessagesOnlyFromSendersOrMembers {
   <#
    .DESCRIPTION
    Set-AcceptMessagesOnlyFromSendersOrMembers - Updates the 'AcceptMessagesOnlyFromSendersOrMembers'
    setting on a DL from JSON source data.

    .EXAMPLE
    $jsonData | Set-AcceptMessagesOnlyFromSendersOrMembers

    .NOTES
    This function will override this setting for each DL from JSON source data.

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               # CORE TASK
               if($InputObject.accept_messages_only_from_senders_or_members -eq $true) {
                  Write-Host -BackgroundColor black -ForegroundColor green `
                  "$((get-date).ToShortTimeString()): Setting AcceptMessagesOnlyFromSendersOrMembers on dl-$($InputObject.DisplayName) to true"


                  Set-DistributionGroup -identity "dl-$($InputObject.displayname)" `
                  -AcceptMessagesOnlyFromSendersOrMembers "dl-$($InputObject.displayname)" 
                        
                  Set-DistributionGroup -identity "dl-$($InputObject.displayname)" `
                  -ErrorAction silentlycontinue `
                  -AcceptMessagesOnlyFromSendersOrMembers @{add=$($InputObject.Senders)}

                  if(!($InputObject.senders)) {
                     Write-Host -BackgroundColor red -ForegroundColor white `
                     "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'Senders' field"
    
                  $ErrorActionPreference = 'silentlycontinue'
                  }
                  
                  elseif($InputObject.accept_messages_only_from_senders_or_members -eq $false) {
                     Write-Host -BackgroundColor black -ForegroundColor green `
                     "$((get-date).ToShortTimeString()): Setting AcceptMessagesOnlyFromSendersOrMembers on dl-$($InputObject.DisplayName) to false"
                     }
                        
                     elseif($InputObject.accept_messages_only_from_senders_or_members -eq $null) {
                        Write-Host -BackgroundColor Black  -ForegroundColor Yellow `
                        "$((get-date).ToShortTimeString()): AcceptMessagesOnlyFromSendersOrMembers for dl-$($InputObject.DisplayName) does not exist"

                        Write-Output "$($timestamp):  Displayname: $($_.displayname):  Attribute false"  `r`n | `
                        Out-File -Append "$($logPath)\$($datestamp)_Set-AcceptMessagesOnlyFromSendersOrMembers.log"
                        }
               }
            }
}


function Set-HiddenFromAddressListsEnabled {
   <#
    .DESCRIPTION
    Set-HiddenFromAddressListsEnabled - Updates the 'HiddenFromAddressListsEnabled'
    setting on a DL from JSON source data.

    .EXAMPLE
    $jsonData | Set-HiddenFromAddressListsEnabled

    .NOTES
    This function will override this setting for each DL from JSON source data.

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
        )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting HiddenFromAddressListsEnabled on dl-$($InputObject.DisplayName) to $($InputObject.HiddenFromAddressListsEnabled)"

               # CORE TASK
               "dl-$($InputObject.displayname)" | `
               Set-DistributionGroup -WarningAction silentlycontinue `
               -ErrorAction silentlycontinue `
               -HiddenFromAddressListsEnabled:$($InputObject.HiddenFromAddressListsEnabled)
                
               if($InputObject.HiddenFromAddressListsEnabled -eq $null) {
                  Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                  "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'HiddenFromAddressListsEnabled' field"
    
               $ErrorActionPreference = 'silentlycontinue'
               }
            }
}


function Add-DLproxyAddress {
   <#
    .DESCRIPTION
    Add-DLproxyAddress - Adds a secondary legacy email address for backwards compatibility.
    The primary SMTP address will still be prefixed with dl-.

    .EXAMPLE
    $jsonData | Add-DLproxyAddress

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Adding additional SMTP addresses: $($InputObject.displayname)@wpi.edu"
                
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -WarningAction silentlycontinue `
               -emailaddresses: @{add="$($InputObject.DisplayName)@wpi.edu"}

               # DATA GATHERING / LOGGING
               $DLSMTP = @(Get-DistributionGroup "dl-$($InputObject.displayname)" | `
               select emailaddresses -ExpandProperty emailaddresses | out-string)

               write-output "$($timestamp):`r`nDisplayname: dl-$($InputObject.displayname);`r`nEmail/Proxy Addresses:`r`n$($DLSMTP.split(","))" `r`n | `
               Out-File -Append "$($logPath)\$($datestamp)_Add-DLProxyAddress.log"
               # END DATA GATHERING
            }
}


function Set-BypassModerationFromSendersOrMembers {
   <#
    .DESCRIPTION
    Set-BypassModerationFromSendersOrMembers - Sets the BypassModeration setting as per the JSON source data.
    This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-BypassModerationFromSendersOrMembers

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'BypassModeration' on DL: dl-$($InputObject.displayname)"
                       
               # CORE TASK
               if($InputObject.bypass_moderation_from_senders -eq $true) {
                  Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
                  -WarningAction silentlycontinue `
                  -ErrorAction silentlycontinue `
                  -BypassModerationFromSendersOrMembers $($InputObject.Senders)

                  if(!($InputObject.Senders)) {
                     Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                     "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'Senders' field"
    
                     $ErrorActionPreference = 'silentlycontinue'
                     }
               }
            }
}


function Set-ManagedBy {
   <#
    .DESCRIPTION
    Set-ManagedBy - Sets the ManagedBy setting as per the JSON source data.
    This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-ManagedBy

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'ManagedBy' to $($InputObject.ManagedBy) on dl-$($InputObject.displayname)"
                       
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -ErrorAction silentlycontinue `
               -ManagedBy $($InputObject.ManagedBy)

               if(!($InputObject.ManagedBy)) {
                  Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                  "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'ManagedBy' field"
    
                  $ErrorActionPreference = 'silentlycontinue'
                  }
            }
}


function Set-ModerationEnabled {
   <#
    .DESCRIPTION
    Set-ModerationEnabled - Sets the ModerationEnabled setting as per the JSON source data.
    This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-ModerationEnabled

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'ModerationEnabled' on dl-$($InputObject.displayname) to $($InputObject.ModerationEnabled)"
               
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -WarningAction silentlycontinue `
               -ErrorAction silentlycontinue `
               -ModerationEnabled:$($InputObject.ModerationEnabled)

               if(!($InputObject.ModerationEnabled)) {
                  Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                  "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'ModerationEnabled' field"
    
                  $ErrorActionPreference = 'silentlycontinue'
                  }
            }
}


function Set-ModeratedBy {
   <#
    .DESCRIPTION
    Set-ModeratedBy - Sets the ModeratedBy setting as per the JSON source data.
    This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-ModeratedBy

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Adding $($InputObject.ModeratedBy) to 'ModeratedBy'on dl-$($InputObject.displayname)"

               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -ErrorAction silentlycontinue `
               -ModeratedBy $($InputObject.ModeratedBy)

               if(!($InputObject.ModeratedBy)) {
                  Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                  "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'ModeratedBy' field"
    
                  $ErrorActionPreference = 'silentlycontinue'
                  }
            }
}


function Set-RequireSenderAuthenticationEnabled {
   <#
    .DESCRIPTION
    Set-RequireSenderAuthenticationEnabled - Sets the RequireSenderAuthenticationEnabled setting 
    as per the JSON source data.  This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-RequireSenderAuthenticationEnabled

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'RequireSenderAuthenticationEnabled' on dl-$($InputObject.displayname) to $($InputObject.RequireSenderAuthenticationEnabled)"
            
               # CORE TASK
               if((($InputObject.RequireSenderAuthenticationEnabled) -eq $true) -or (($InputObject.RequireSenderAuthenticationEnabled) -eq $false)) {
                  Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
                  -WarningAction silentlycontinue `
                  -ErrorAction silentlycontinue `
                  -RequireSenderAuthenticationEnabled:$($InputObject.RequireSenderAuthenticationEnabled)
                  }
                       
                  elseif(!($InputObject.RequireSenderAuthenticationEnabled)) {
                     Write-Host -BackgroundColor Black -ForegroundColor Yellow `
                     "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'RequireSenderAuthenticationEnabled' field"
    
                     $ErrorActionPreference = 'silentlycontinue'
                     }
            }
}


function Set-Notes {
   <#
    .DESCRIPTION
    Set-Notes - Sets the Notes field of a DL as per the JSON source data.  
    This will override the existing setting.

    .EXAMPLE
    $jsonData | Set-Notes

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'Notes' on dl-$($InputObject.displayname)"
               
               # NOT A TYPO, THE COMMAND IS 'SET-GROUP'
               # CORE TASK
               Set-Group -Identity "dl-$($InputObject.DisplayName)" `
               -WarningAction silentlycontinue `
               -Notes $($InputObject.Notes)
            }
}


function Delete-DL {
   <#
    .DESCRIPTION
    Delete-DL - Doesn't actually delete the group, but clears the membership, hides the DL
    from the Global Address List, if not already hidden, and renames the group prefixed with 
    "_TBD-".

    .EXAMPLE
    $jsonData | Delete-DL

    .NOTES

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               Write-Host -BackgroundColor red -ForegroundColor white `
               "$((get-date).ToShortTimeString()): PREPARING dl-$($InputObject.DisplayName) FOR REMOVAL"
               
               Get-DistributionGroup "dl-$($InputObject.DisplayName)" | fl | `
               Out-File -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

               Get-DistributionGroupMember "dl-$($InputObject.DisplayName)" -ResultSize Unlimited  | `
               Out-File -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

               Get-Group "dl-$($InputObject.DisplayName)" | fl | `
               Out-File -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

               Update-DistributionGroupMember -Identity "dl-$($InputObject.DisplayName)" `
               -Members $null -Confirm:$false

               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -HiddenFromAddressListsEnabled:$true `
               -ManagedBy Listmaster `
               -ModerationEnabled:$false `
               -ModeratedBy $null `
               -AcceptMessagesOnlyFromSendersOrMembers Listmaster `
               -MemberJoinRestriction closed `
               -MemberDepartRestriction closed `
               -RequireSenderAuthenticationEnabled:$true `
               -IgnoreNamingPolicy `
               -DisplayName "_TBD-dl-$($InputObject.DisplayName)"

               Set-Group -Identity "dl-$($InputObject.DisplayName)" `
               -Notes "THIS GROUP IS SCHEDULED TO BE DELETED"
            }
}

Function Archive-Logs ($DataPath,$ArchiveRoot,$Date) {
    ## Requires the assembly type below.  Typically called once in the main script.
    #Add-Type -assembly "system.io.compression.filesystem"

    ## Verify Archive Directory Path
    $ArchivePath = "$ArchiveRoot\$($Date.Year)\$($Date.ToString("MM"))"
    if (!(Test-Path $ArchivePath)) {[VOID](New-Item $ArchivePath -Type Directory)}

    $datestamp    = get-date ($Date) -Format yyyy_MM_dd_HHmmss
    $Archive_Zip   = "$ArchivePath\$($datestamp)_Logs.zip"
    
    ## Zip and Copy all files to archive
    [io.compression.zipfile]::CreateFromDirectory($DataPath, $Archive_Zip) 
    
    ## Remove files
    $files = Get-ChildItem $DataPath -File
    foreach ($file in $files) {Remove-Item -Path $file.FullName -Confirm:$false}
}

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
      
      beginParsingJSON
   }
   catch [Exception] {
      $_.Exception.Message
            
      Write-Host -ForegroundColor YELLOW "`n`nUnable to authenticate with saved credentials.`n"
      Write-Host -foregroundColor YELLOW "This may be due to already having loading the Exchange plugins, please try closing this shell and restarting."
      exit
   }
}

function beginParsingJSON(){
   Write-Host "Checking for JSON folder at $jsonPath"
   if (Test-Path $jsonPath){
      Write-Host "Path Found, getting JSON data from sub-folders."
      # Grab our nested files, wherever they may be.
      Foreach ($path in Get-ChildItem -path $jsonPath | Select-Object -ExpandProperty Name){
         $jsonFileList = Get-ChildItem -path "$jsonPath\$path" -filter *.json | Sort-Object
         Write-Host -ForegroundColor CYAN "`nFound Files:"
         $jsonFiles = $jsonFileList | Sort Length | Get-Content -raw
         $jsonData = $jsonFiles | ConvertFrom-Json

         Write-Host -BackgroundColor Black -ForegroundColor Magenta `
         "$((get-date).ToShortTimeString()): Begin processing of distribution lists in $path" `r`n

         $jsonData | Test-ExistingDL
      }
   }
   else {
      Write-Host -ForegroundColor RED "`nPath not found: $jsonPath`n" 
      Stop-Transcript
      exit
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
      beginParsingJSON
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



<#
# IF THE DL IS TO BE DELETED:
$jsonData | Delete-DL
#>


## Clean Up JSON Files
#$jsonFileList | Remove-Item -Confirm:$false

<#
## Clean up Sessions
Remove-PSSession $ExchangeOnlineSession
Disconnect-AzureAD
#>

# Only log stuff if we're not testing.
if (-NOT $testMode){
   Stop-Transcript

   ## Create Error Log
   Get-Content "$($logPath)\$($datestamp)_TRANSCRIPT-DLmgmt.log" | `
      Select-String -Pattern '(does not exist)|(does not contain)|(was not created)|(has missing mandatory)' | `
      Get-Unique | `
      Out-File -Encoding default -append "$($logPath)\$($datestamp)_DLmgmt-ERROR.log"

   ## Clean up Logs
   Archive-Logs $logPath $logArchivePath $date

}