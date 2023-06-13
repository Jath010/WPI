################################
#                              #
#     GENERAL SCRIPT NOTES     #
#                              #
################################

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

See 'DEFINE VARIABLES' section below.

Automate credentials by following instructions to use Paul Cunningham's Get-StoredCredential module:
https://practical365.com/blog/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks/

If you choose to use Paul Cunningham's Functions-PSStoredCredentials.ps1 script for
automating your connection to Office 365, you will also want to update the location of
this script in the 'CONNECT TO OFFICE 365' section below.

#>



#################################
#                               #
#     CONNECT TO OFFICE 365     #
#                               #
#################################

Clear-Host
#Get list of adusers that have been modified in the last 24 hours and verify that there are no termination mail restrictions
$datestamp = Get-Date -Format "yyyyMMdd-HHmm"
$timestamp = Get-Date -Format "yyyy-MM-dd - HH:mm"
#<#
##################################################################################################################
## Powershell Load Credentials
##################################################################################################################
#$credentials=$null;$credPath=$null
## The following credentials use the UPN "s_tcollins@wpitest.org" for the username.  This is used to connect to remote PSSessions for Exchange Online
#$credPath = 'D:\wpi\batch\Exchange\CheckMailRestrictions\exch_automation@wpitest.org.xml'
#$credentials = Import-CliXml -Path $credPath

##################################################################################################################
## Powershell Load Modules
##################################################################################################################
#$ExchangeOnlineSession=$null

## Load Exchange Online
#$ExchangeOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $credentials -Authentication Basic -AllowRedirection
#Import-PSSession $ExchangeOnlineSession
#Connect-AzureAD -Credential $credentials
#>

write-output `r`n

################################
#                              #
#       DEFINE VARIABLES       #
#                              #
################################

write-host -BackgroundColor Black -ForegroundColor Magenta `
"$((get-date).ToShortTimeString()): Defining variables" `r`n

# Variable that cycles through each .json file and converts it to a Powershell object
$jsonData = @()
$jsonFileList = Get-ChildItem -path "C:\Users\s_tcollins\Documents\WindowsPowerShell\Scripts\DLManagement\JSONData" -filter *.json | Sort-Object
$jsonFiles = $jsonFileList | Sort Length | Get-Content -raw
$jsonData = $jsonFiles | ConvertFrom-Json

# Set path for log files:
$logPath = "C:\Users\s_tcollins\Documents\WindowsPowerShell\Scripts\DLManagement\Logs"

# Get date for logging and file naming:
$datestamp = Get-Date -Format "yyyyMMdd-HHmm"
$timeFull = $((get-date).toString("yyyy-MM-dd--HH:mm:ss"))

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_TRANSCRIPT-DLmgmt.log" -Force
Write-Host "Begin file"

## Get a full list of available mailboxes.  This will reduce the time it takes to validate if a member exists and reduce the number of calls to Exchange Online
$global:mailboxes = Get-Mailbox -ResultSize unlimited | Sort Alias
$global:mailboxAliasList = $mailboxes | Sort Alias | Select -ExpandProperty Alias

#############################
#                           #
#    PRIMARY FUNCTIONS      #
#                           #
#############################

write-host -BackgroundColor Black -ForegroundColor Magenta `
"$((get-date).ToShortTimeString()): Loading functions" `r`n

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
               write-host -BackgroundColor DarkGray -ForegroundColor White `
               "$((get-date).ToShortTimeString()): Checking to see if dl-$($InputObject.DisplayName) exists"

               $TestGroup = Get-DistributionGroup "dl-$($inputobject.DisplayName)" -ErrorAction silentlycontinue

                  # CORE TASK: UPDATES GROUP IF GROUP EXISTS
                  if($TestGroup) {
                     write-host -BackgroundColor DarkGreen -ForegroundColor White `
                     "$((get-date).ToShortTimeString()): $TestGroup already exists, updating..."

                     # YOU WILL TYPICALLY ONLY DO ONE OF THE FOLLOWING TWO AT RUNTIME:
                     # $InputObject | Update-DLmembership
                     $InputObject | Sync-DLmembership
                     #

                     $InputObject | Set-AcceptMessagesOnlyFromSendersOrMembers
                     $InputObject | Set-HiddenFromAddressListsEnabled
                     $InputObject | Add-DLproxyAddress
                     $InputObject | Set-BypassModerationFromSendersOrMembers
                     $InputObject | Set-ManagedBy
                     $InputObject | Set-ModerationEnabled
                     $InputObject | Set-ModeratedBy
                     $InputObject | Set-RequireSenderAuthenticationEnabled
                     $InputObject | Set-Notes
                     write-host `r`n
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
    that has been validated by the Test-ExistingDL function.

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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Creating Office 365 DL: dl-$($InputObject.DisplayName)"

               # FILTER OUT NON-EXISTENT MEMBERS FROM SOURCE DATA AND ADD TO NEW-ARRAY
               $jMembers = @()
               $jMembers = $InputObject | select -ExpandProperty members
               $populateNewList = @()
               foreach ($member in $jMembers) {
                    if($global:mailboxAliasList | Where {$_ -eq $member}) {$populateNewList += $member}
                    else {
                        write-host -BackgroundColor red -ForegroundColor white `
                        "$((get-date).ToShortTimeString()): $member in $($InputObject.displayname) does not exist - skipping member"
                                  
                        write-output "$member in $($InputObject.DisplayName) does not exist" | `
                        Out-File -Append "$($logPath)\$($datestamp)_New-DLfromJSON_$($InputObject.DisplayName)_Non-Existing Members.txt"
                        }  
                    }

               # TEST MANDATORY SETTINGS FOR CORE TASK
               $ErrorActionPreference = 'silentlycontinue'

               $jPrimarySMTPAddress = @()
               $jPrimarySMTPAddress = $InputObject | select -ExpandProperty PrimarySMTPAddress 
               
               foreach($group in $InputObject.displayname) {
                  if($jPrimarySMTPAddress -eq $null) {
                     write-host -BackgroundColor red -ForegroundColor white `
                     "$((get-date).ToShortTimeString()): 'PrimarySMTPAddress' in $($InputObject.displayname) does not exist or is empty"
                     }
               }

               $jManagedBy = @()
               $jManagedBy = $InputObject | select -ExpandProperty ManagedBy 
               
               foreach($group in $InputObject.displayname) {
                  if($jManagedBy -eq $null) {
                     write-host -BackgroundColor red -ForegroundColor white `
                     "$((get-date).ToShortTimeString()): 'ManagedBy' in $($InputObject.displayname) does not exist or is empty"
                     }
               }

               $ErrorActionPreference = 'continue'

               $createDL = New-DistributionGroup -Name "dl-$($InputObject.DisplayName)" `
                  -DisplayName "dl-$($InputObject.DisplayName)" `
                  -Alias "dl-$($InputObject.DisplayName)" `
                  -Type "Distribution" `
                  -PrimarySmtpAddress "dl-$($InputObject.DisplayName)@wpitest.org" `
                  -ManagedBy $InputObject.ManagedBy `
                  -ModerationEnabled:$($InputObject.ModerationEnabled) `
                  -ModeratedBy $InputObject.ModeratedBy `
                  -Members $populateNewList `
                  -Notes $InputObject.Notes `
                  -RequireSenderAuthenticationEnabled:$($InputObject.RequireSenderAuthenticationEnabled) `
                  -IgnoreNamingPolicy `
                  -MemberJoinRestriction Closed `
                  -MemberDepartRestriction Closed `
                  -ErrorAction silentlycontinue


                # CORE TASK
                if($createDL) {
                    $InputObject | Set-AcceptMessagesOnlyFromSendersOrMembers -ErrorAction SilentlyContinue
                    $InputObject | Set-HiddenFromAddressListsEnabled -ErrorAction SilentlyContinue
                    $InputObject | Add-DLproxyAddress -ErrorAction SilentlyContinue
                    $InputObject | Set-BypassModerationFromSendersOrMembers -ErrorAction SilentlyContinue
                    write-host `r`n
                         
                    # DATA GATHERING / LOGGING
                    $newDLemail = Get-DistributionGroup "dl-$($InputObject.DisplayName)" | select -ExpandProperty primarysmtpaddress
                    $newDLmembers = Get-DistributionGroupMember "dl-$($InputObject.DisplayName)" -ResultSize Unlimited | select -ExpandProperty alias

                    Write-Output "$($timeFull): `r`nNew DL: dl-$($inputobject.DisplayName)`r`nPrimarySMTPAddress: $newDLemail`r`nMembers: $newDLmembers" `r`n | `
                    out-file -Append "$($logPath)\$($datestamp)_new-DLfromJSON.log"
                    # END DATA GATHERING
                    }
                else {
                    write-host -BackgroundColor red -ForegroundColor white `
                    "$((get-date).ToShortTimeString()): $($InputObject.DisplayName) is missing mandatory values and was not created" `r`n
                    }
            }
}


function Update-DLmembership {
    <#
    .DESCRIPTION
    Update-DLmembership - Updates the membership of a DL from JSON source data.

    .EXAMPLE
    $jsonData | Update-DLmembership

    .NOTES
    This function will remove all existing DL members and add members from source data.
    However, this has proven to be faster than synchronizing membership one by one when a 
    large number of changes is involved.

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               write-host -BackgroundColor Blue -ForegroundColor White `
               "$((get-date).ToShortTimeString()): Updating dl-$($InputObject.displayname) membership"

               # FILTER OUT NON-EXISTENT MEMBERS FROM SOURCE DATA AND ADD TO NEW-ARRAY
               $jMembers = @()
               $jMembers = $InputObject | select -ExpandProperty members
               $updateMembers = @()

               foreach ($member in $jMembers) {
                    if($global:mailboxAliasList | Where {$_ -eq $member}) {$updateMembers += $member}
                    else {
                        write-host -BackgroundColor red -ForegroundColor white `
                        "$((get-date).ToShortTimeString()): $member in $($InputObject.displayname) does not exist - skipping member"
                                  
                        write-output "$member in $($InputObject.DisplayName) does not exist" | `
                        Out-File -Append "$($logPath)\$($datestamp)_Update-DLmembership_$($InputObject.DisplayName)_Non-Existing Members.txt"
                        }  
                    }

               # DATA GATHERING / LOGGING BEFORE CORE TASK
               $DL = Get-DistributionGroup "dl-$($InputObject.displayname)" -ResultSize unlimited | `
                select -ExpandProperty alias
               $jsonCount = ($InputObject  | select -ExpandProperty members | Measure-Object).count
               $preDLmembers = $DL | Get-DistributionGroupMember -ResultSize unlimited | `
                select -ExpandProperty alias
               $preDLcount = ($preDLmembers | measure-object).count
               Write-Output "Member total: $preDLcount","JSON file total: $jsonCount",$DL,$preDLmembers | `
               out-file "$($logPath)\$($datestamp)_Update-DLmembership_$DL-BEFORE.txt"
               # END DATA GATHERING

               # CORE FUNCTION TASK
               Update-DistributionGroupMember -Identity "dl-$($InputObject.DisplayName)" `
               -ErrorAction silentlycontinue `
               -Members $updateMembers `
               -Confirm:$false

               # DATA GATHERING / LOGGING POST CORE TASK
               $postDLmembers = $DL | Get-DistributionGroupMember -ResultSize unlimited | `
                select -ExpandProperty alias
               $postDLcount = ($postDLmembers | measure-object).count
               Write-Output "Member total: $postDLcount","JSON file total: $jsonCount",$DL,$postDLmembers | `
               out-file "$($logPath)\$($datestamp)_Update-DLmembership_$DL-AFTER.txt"
               
               ForEach($postDLmember in $postDLmembers) {
                    if($postDLmember -notin $preDLmembers){
                        write-output  "$postDLmember added" | `
                        out-file -append "$($logPath)\$($datestamp)_Update-DLmembership_$DL-CHANGES.txt"
                        }
               }

               ForEach($preDLmember in $preDLmembers) {
                    if($preDLmember -notin $postDLmembers){
                        write-output  "$preDLmember removed" | `
                        out-file -append "$($logPath)\$($datestamp)_Update-DLmembership_$DL-CHANGES.txt"
                        }
               }
               # END DATA GATHERING
            }
}


function Sync-DLmembership {
   <#
    .DESCRIPTION
    Sync-DLmembership - Synchronizes the membership of a DL from JSON source data.

    .EXAMPLE
    $jsonData | Sync-DLmembership

    .NOTES
    This function will synchronize only the changes for each DL from source data.
    However, this function is better for smaller one-off changes.  For a larger number 
    of changes, the Update-DLmembership function has proven to be faster and more 
    efficient.

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        $InputObject
    )

    Process {
               write-host -BackgroundColor Blue -ForegroundColor White `
               "$((get-date).ToShortTimeString()): Synchronizing dl-$($InputObject.displayname) membership"

               # VALIDATES THAT THE MEMBERS FIELD EXISTS IN THE SOURCE DATA
               $membersField = $InputObject.Members
                  if($membersField -eq $null) {
                     write-host -BackgroundColor yellow -ForegroundColor black `
                     "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'Members' field or it is empty"
                         
                     $ErrorActionPreference = 'silentlycontinue'
                     }
                     # END ADD

               $jMembers = @()
               $jMembers = $InputObject | select -ExpandProperty members
               $syncMembers = @()

               # VALIDATES EACH MEMBER IN THE SOURCE DATA - ADDS ONLY VALID USERS TO NEW ARRAY
               foreach ($member in $jMembers) {
                    if($global:mailboxAliasList | Where {$_ -eq $member}) {$syncMembers += $member}
                    else {
                        write-host -BackgroundColor red -ForegroundColor white `
                        "$((get-date).ToShortTimeString()): $member in $($InputObject.displayname) does not exist - skipping member"
                                  
                        write-output "$member in $($InputObject.DisplayName) does not exist" | `
                        Out-File -Append "$($logPath)\$($datestamp)_Sync-DLmembership_$($InputObject.DisplayName)_Non-Existing Members.txt"
                        }  
                    }

               $dList = Get-DistributionGroupMember "dl-$($InputObject.displayname)" -ResultSize unlimited | `
                select -ExpandProperty alias

               # DATA GATHERING / LOGGING PRE CORE TASK
               $DL = Get-DistributionGroup "dl-$($InputObject.displayname)" -ResultSize unlimited | `
                select -ExpandProperty alias
               $jsonCount = ($InputObject  | select -ExpandProperty members | Measure-Object).count
               $preDLmembers = $DL | Get-DistributionGroupMember -ResultSize unlimited | `
                select -ExpandProperty alias
               $preDLcount = ($preDLmembers | measure-object).count
               Write-Output "Member total: $preDLcount","JSON file total: $jsonCount",$DL,$preDLmembers | `
               out-file "$($logPath)\$($datestamp)_Sync-DLmembership_$DL-before ADD.txt"
               # END DATA GATHERING

               # CORE FUNCTION TASK - ADD
               foreach ($member in $syncMembers) {
                        
                  if($dlist -notcontains $member){
                     Add-DistributionGroupMember -Identity "dl-$($InputObject.DisplayName)" -Member $member

                     write-host -BackgroundColor black -ForegroundColor green `
                     "$((get-date).ToShortTimeString()): Adding $member to dl-$($InputObject.displayname)"
                     }
                  }

               # CORE FUNCTION TASK - REMOVE
               foreach ($dlMembers in $dList) {

                  if($InputObject.members -notcontains $dlMembers) {
                     write-host -BackgroundColor black -ForegroundColor green `
                     "$((get-date).ToShortTimeString()): Removing $dlMembers from dl-$($InputObject.displayname)"

                     Remove-DistributionGroupMember -Identity "dl-$($InputObject.DisplayName)" -Member $dlMembers `
                     -confirm:$false
                     }
               }

               # DATA GATHERING / LOGGING POST CORE TASKS
               $postDLmembers = $DL | Get-DistributionGroupMember -ResultSize unlimited | `
                select -ExpandProperty alias
               $postDLcount = ($postDLmembers | measure-object).count
               Write-Output "Member total: $postDLcount","JSON file total: $jsonCount",$DL,$postDLmembers | `
               out-file "$($logPath)\$($datestamp)_Sync-DLmembership_$DL-after REMOVAL.txt"
                     
               ForEach($postDLmember in $postDLmembers) {
                  if($postDLmember -notin $preDLmembers){
                     write-output  "$postDLmember added" | `
                     out-file -append "$($logPath)\$($datestamp)_Sync-DLmembership_$DL-CHANGES.txt"
                     }
               }

               ForEach($preDLmember in $preDLmembers) {
                  if($preDLmember -notin $postDLmembers){
                     write-output  "$preDLmember removed" | `
                     out-file -append "$($logPath)\$($datestamp)_Sync-DLmembership_$DL-CHANGES.txt"
                     }
               }
               # END DATA GATHERING
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
                  write-host -BackgroundColor black -ForegroundColor green `
                  "$((get-date).ToShortTimeString()): Setting AcceptMessagesOnlyFromSendersOrMembers on dl-$($InputObject.DisplayName) to true"


                  Set-DistributionGroup -identity "dl-$($InputObject.displayname)" `
                  -AcceptMessagesOnlyFromSendersOrMembers "dl-$($InputObject.displayname)" 
                        
                  Set-DistributionGroup -identity "dl-$($InputObject.displayname)" `
                  -ErrorAction silentlycontinue `
                  -AcceptMessagesOnlyFromSendersOrMembers @{add=$($InputObject.Senders)}

                  if(!($InputObject.senders)) {
                     write-host -BackgroundColor red -ForegroundColor white `
                     "$((get-date).ToShortTimeString()): $($InputObject.displayname) does not contain a 'Senders' field"
    
                  $ErrorActionPreference = 'silentlycontinue'
                  }
                  
                  elseif($InputObject.accept_messages_only_from_senders_or_members -eq $false) {
                     write-host -BackgroundColor black -ForegroundColor green `
                     "$((get-date).ToShortTimeString()): Setting AcceptMessagesOnlyFromSendersOrMembers on dl-$($InputObject.DisplayName) to false"
                     }
                        
                     elseif($InputObject.accept_messages_only_from_senders_or_members -eq $null) {
                        write-host -BackgroundColor Black  -ForegroundColor Yellow `
                        "$((get-date).ToShortTimeString()): AcceptMessagesOnlyFromSendersOrMembers for dl-$($InputObject.DisplayName) does not exist"

                        Write-Output "$($timeFull):  Displayname: $($_.displayname):  Attribute false"  `r`n | `
                        out-file -Append "$($logPath)\$($datestamp)_Set-AcceptMessagesOnlyFromSendersOrMembers.log"
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting HiddenFromAddressListsEnabled on dl-$($InputObject.DisplayName) to $($InputObject.HiddenFromAddressListsEnabled)"

               # CORE TASK
               "dl-$($InputObject.displayname)" | `
               Set-DistributionGroup -WarningAction silentlycontinue `
               -ErrorAction silentlycontinue `
               -HiddenFromAddressListsEnabled:$($InputObject.HiddenFromAddressListsEnabled)
                
               if($InputObject.HiddenFromAddressListsEnabled -eq $null) {
                  write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Adding additional SMTP addresses: $($InputObject.displayname)@wpitest.org"
                
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -WarningAction silentlycontinue `
               -emailaddresses: @{add="$($InputObject.DisplayName)@wpitest.org"}

               # DATA GATHERING / LOGGING
               $DLSMTP = @(Get-DistributionGroup "dl-$($InputObject.displayname)" | `
               select emailaddresses -ExpandProperty emailaddresses | out-string)

               write-output "$($timeFull):`r`nDisplayname: dl-$($InputObject.displayname);`r`nEmail/Proxy Addresses:`r`n$($DLSMTP.split(","))" `r`n | `
               out-file -Append "$($logPath)\$($datestamp)_Add-DLProxyAddress.log"
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'BypassModeration' on DL: dl-$($InputObject.displayname)"
                       
               # CORE TASK
               if($InputObject.bypass_moderation_from_senders -eq $true) {
                  Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
                  -WarningAction silentlycontinue `
                  -ErrorAction silentlycontinue `
                  -BypassModerationFromSendersOrMembers $($InputObject.Senders)

                  if(!($InputObject.Senders)) {
                     write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'ManagedBy' to $($InputObject.ManagedBy) on dl-$($InputObject.displayname)"
                       
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -ErrorAction silentlycontinue `
               -ManagedBy $($InputObject.ManagedBy)

               if(!($InputObject.ManagedBy)) {
                  write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'ModerationEnabled' on dl-$($InputObject.displayname) to $($InputObject.ModerationEnabled)"
               
               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -WarningAction silentlycontinue `
               -ErrorAction silentlycontinue `
               -ModerationEnabled:$($InputObject.ModerationEnabled)

               if(!($InputObject.ModerationEnabled)) {
                  write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Adding $($InputObject.ModeratedBy) to 'ModeratedBy'on dl-$($InputObject.displayname)"

               # CORE TASK
               Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
               -ErrorAction silentlycontinue `
               -ModeratedBy $($InputObject.ModeratedBy)

               if(!($InputObject.ModeratedBy)) {
                  write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
               "$((get-date).ToShortTimeString()): Setting 'RequireSenderAuthenticationEnabled' on dl-$($InputObject.displayname) to $($InputObject.RequireSenderAuthenticationEnabled)"
            
               # CORE TASK
               if((($InputObject.RequireSenderAuthenticationEnabled) -eq $true) -or (($InputObject.RequireSenderAuthenticationEnabled) -eq $false)) {
                  Set-DistributionGroup -Identity "dl-$($InputObject.DisplayName)" `
                  -WarningAction silentlycontinue `
                  -ErrorAction silentlycontinue `
                  -RequireSenderAuthenticationEnabled:$($InputObject.RequireSenderAuthenticationEnabled)
                  }
                       
                  elseif(!($InputObject.RequireSenderAuthenticationEnabled)) {
                     write-host -BackgroundColor Black -ForegroundColor Yellow `
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
               write-host -BackgroundColor black -ForegroundColor green `
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
               write-host -BackgroundColor red -ForegroundColor white `
               "$((get-date).ToShortTimeString()): PREPARING dl-$($InputObject.DisplayName) FOR REMOVAL"
               
               Get-DistributionGroup "dl-$($InputObject.DisplayName)" | fl | `
               out-file -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

               Get-DistributionGroupMember "dl-$($InputObject.DisplayName)" -ResultSize Unlimited  | `
               out-file -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

               Get-Group "dl-$($InputObject.DisplayName)" | fl | `
               out-file -append "$($logPath)\$($datestamp)_dl-$($InputObject.DisplayName)_PRE-REMOVAL-DETAILS.txt"

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



#############################
#                           #
#       MAIN PROGRAM        #
#                           #
#############################

write-host -BackgroundColor Black -ForegroundColor Magenta `
"$((get-date).ToShortTimeString()): Begin processing of distribution lists" `r`n

$jsonData | Test-ExistingDL

<#
# IF THE DL IS TO BE DELETED:
$jsonData | Delete-DL
#>


## Clean Up JSON Files
### $jsonFileList | Remove-Item -Confirm:$false

### Remove-PSSession $ExchangeOnlineSession
### Disconnect-AzureAD
Stop-Transcript

Get-Content "$($logPath)\$($datestamp)_TRANSCRIPT-DLmgmt.log" | `
    Select-String -Pattern '(does not exist)|(does not contain)|(was not created)|(has missing mandatory)' | `
    Get-Unique | `
    out-file -Encoding default -append "$($logPath)\$($datestamp)_DLmgmt-ERROR.log"


