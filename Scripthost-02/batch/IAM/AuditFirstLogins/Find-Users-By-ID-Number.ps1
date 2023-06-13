<#
.SYNOPSIS
    This script is intended to be used to "convert" a list of ID numbers to SamAccountName
    and output them to CSV so they can be used in the Get-New-Account-First-Logons.ps1 script.

.DESCRIPTION

.NOTES
    Created By: Stephen Gemme
    Created On: 06/29/2020
    
    Modified By:
    Modified On:
    Modifications:
#>

Param (
   [Parameter(Mandatory = $true)]
   [String]
   $csvPath, 

   [Parameter(Mandatory = $false)] 
   [Switch] 
   $testMode
)

# Where we're going to store all the necessary info.
[System.Collections.ArrayList]$Global:userInfo = @()

function convertIDs(){
    try {
        $entries = Import-CSV $csvPath
    }
    catch {
        Write-Host -foregroundColor YELLOW "Error when trying to import CSV."
        exit
    }

    if (-NOT $null -eq $entries){
        foreach ($entry in $entries){
            if ($null -eq $entry."Student ID"){
                Write-Host -foregroundColor RED "'Student ID' field not found, please re-check CSV that this field exists."
                exit
            }
            else {
                $info = $null
                $idNumber = $entry."Student ID"
                Write-Host "Attempting to get info on user with ID: " -noNewLine
                Write-Host -foregroundColor CYAN $idNumber -noNewLine
                Write-Host "..." -noNewLine

                $search = @{
                    Filter      = "Name -like '*$idNumber*'"
                    SearchBase  = "OU=Accounts,DC=admin,DC=wpi,DC=edu"
                    Property    = "whenCreated"
                }
                $info = Get-AdUser @search

                if (-NOT $null -eq $info){
                    $entry = [PSCustomObject]@{
                        ID                  = $idNumber
                        First               = $info.GivenName
                        Last                = $info.Surname
                        SamAccountName      = $info.SamAccountName
                        UserPrincipalName   = $info.UserPrincipalName
                        whenCreated         = $info.whenCreated
                    }

                    Write-Host -foregroundColor GREEN "Done"

                    $Global:userInfo.Add($entry) | Out-Null
                }
                else {
                    Write-Host -foregroundColor RED "User not found."
                }
            }
        }

        # Derive our new path based on the given path.
        $newPath = ($csvPath.split("\")[1].split(".")[0] + "-Results.csv")
        # Export the data we want to the new file. 
        $Global:userInfo | Export-CSV -path $newPath
    }
}

if (Test-Path $csvPath){
    convertIDs
}
else {
    Write-Host -foregroundColor YELLOW "`nNo CSV found at given path " -noNewLine
    Write-Host -foregroundColor CYAN $csvpath -noNewLine
    Write-Host -foregroundColor YELLOW ", please check path and try again."
}