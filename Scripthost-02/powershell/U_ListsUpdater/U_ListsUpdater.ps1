Param (
    # Reports but doesn't actually make changes.
    [Switch] $Sync
)

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")
$logPath = "D:\wpi\powershell\U_ListsUpdater\Logfiles"
Start-Transcript -Append -Path "$($logPath)\$($datestamp)_U_Sync.log" -Force

function Sync-LegacyGroups {
    [CmdletBinding()]
        
    #exists to search everything
    $LDAProot = 'LDAP://DC=admin,DC=wpi,DC=edu'
    #sets paged query default - groups of 1000
    $pg = '1000'

    Write-Host "Getting Azure Group Populations"
    $Students = Get-AzureADGroup -ObjectId 01a5ccb6-128f-4abe-84ce-a1e8328de0e1 | Get-AzureADGroupMember -All $true | Select-Object -expandproperty UserPrincipalName
    Write-Host "Got Students"
    $Staff = Get-AzureADGroup -ObjectId 9a4ad58f-eaea-4165-b5bc-e64a9c2ecc35 | Get-AzureADGroupMember -All $true | Select-Object -expandproperty UserPrincipalName
    Write-Host "Got Staff"
    $Faculty = Get-AzureADGroup -ObjectId 1c1303b4-7a4f-4cd3-98b2-e8dbecf27428 | Get-AzureADGroupMember -All $true | Select-Object -expandproperty UserPrincipalName
    Write-Host "Got Faculty"
    $allStaffFac = Get-AzureADGroup -ObjectId 8bf7e4f6-d861-4ab9-acc3-f4be9b376bba | Get-AzureADGroupMember -All $true | Select-Object -expandproperty UserPrincipalName
    Write-Host "Got Staff/Fac"

    $GroupMappings = @(
         
        @{
            AzureFilter = $allStaffFac
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=U_Employees,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "U_Employees"    
        },
        @{
            AzureFilter = $Students
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=U_Students,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "U_Students"       
        },
        @{
            AzureFilter = $Faculty
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=U_Faculty,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "U_Faculty"       
        },
        @{
            AzureFilter = $Staff
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=U_Staff,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "U_Staff"    
        }
        ,
        @{
            AzureFilter = $Students
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=CLA_Student_Eligibility,OU=WPI Global,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "CLA_Student_Eligibility"       
        },
        @{
            AzureFilter = $Faculty
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=CLA_Faculty_Eligibility,OU=WPI Global,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "CLA_Faculty_Eligibility"       
        },
        @{
            AzureFilter = $Staff
            LDAPFilter  = '(&(objectCategory=person)(memberOf=CN=CLA_Staff_Eligibility,OU=WPI Global,OU=Groups,DC=admin,DC=wpi,DC=edu))'
            ADGroupName = "CLA_Staff_Eligibility"    
        }
    
    
    )
    
    ForEach ($group in $GroupMappings) {
            
        #GET AZURE AD GROUP
        $AzurePeeps = $group.AzureFilter
        $ADGroup = $group.ADGroupName
    
        #GET AD GROUP
        $Searcher = New-Object DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = $LDAProot
        $Searcher.Filter = $group.LDAPFilter
        $Searcher.PageSize = $pg
        $CurrentMembersList = $Searcher.FindAll() | Sort-Object path
        $ADPeeps = $CurrentMembersList.properties.userprincipalname

        # COMPARE OBJECTS
        Write-Host "Comparing the query results for $($group.ADGroupName)" -ForegroundColor Yellow -BackgroundColor Green
        #           Write-Output "Comparing the query results" | Out-File "D:\wpi\powershell\U_ListsUpdater\Logfiles\ADGroupPopulate.log" -Append

        $comparisons = Compare-Object -ReferenceObject $ADPeeps -DifferenceObject $AzurePeeps
        $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object InputObject
        $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object InputObject

        Write-Host "No. of Additions: $($AddMembers.count)" -ForegroundColor Yellow
        Write-Host "No. of Removals: $($RemoveMembers.count)" -ForegroundColor Yellow
    
        #        Write-Host "============" -ForegroundColor DarkMagenta
        ForEach ($Removal in $RemoveMembers) {
            Write-Host "Removing $($Removal.InputObject)" -ForegroundColor Gray
            #           Write-Host "Removing "$Removal -ForegroundColor green 
            #           Write-Output "Removing $($Removal.InputObject)" | Out-File "D:\wpi\powershell\U_ListsUpdater\Logfiles\ADGroupPopulate.log" -Append
            try{
                Remove-ADPrincipalGroupMembership $Removal.InputObject.Split('@')[0] $ADGroup -confirm:$false #-WhatIf
            }
            catch{
                Write-Host "$($($Removal.InputObject)) Couldn't Be Found"
            }
            
        }
        ForEach ($Addition in $AddMembers) {
            Write-Host "Adding $($Addition.InputObject)" -ForegroundColor Gray
            #           Write-Host "Adding "$Addition -ForegroundColor green
            #           Write-Output "Adding $($Addition.InputObject)" | Out-File "D:\wpi\powershell\AdobeUserSync\Logfiles\ADGroupPopulate.log" -Append
            try{
                Add-ADPrincipalGroupMembership $Addition.InputObject.Split('@')[0] $ADGroup -confirm:$false #-WhatIf
            }
            catch{
                Write-Host "$($($Removal.InputObject)) Couldn't Be Found"
            } 
        }
    }
}

$credpath = "D:\wpi\powershell\U_ListsUpdater\exch_automation"
$credential = Import-CliXml -Path $credPath
Connect-AzureAD -Credential $credential



# C:\Users\jmgorham2_prv\WPI\General_use_scripts\JGorham2\U_ListsUpdater\U_ListsUpdater.ps1 -Sync

if($Sync){
    Sync-LegacyGroups
}
Stop-Transcript