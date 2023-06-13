<#
    So the question to answer is how to recognize that a user has logged into a new account

    https://docs.microsoft.com/en-us/powershell/module/azuread/get-azureadauditsigninlogs?view=azureadps-2.0-preview
    get-azureadauditsigninlogs


    All students who are freshmen created since january 1st until today
    output UPN, student ID#, and date of login if available
#>

Import-module AzureADPreview -force

$Date = "2022-01-01" #((Get-Date).AddDays(-15)).tostring("yyyy-MM-dd")

$FreshmanMembers = 0
$RecentMembers = Get-AzureADUser -Filter "UserType eq 'Member'" -All:$true | Where-Object { $_.ExtensionProperty.createdDateTime -lt $date -and $_.extensionproperty.onPremisesDistinguishedName -like "*,OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu" }


[System.Collections.ArrayList]$ActiveMembers = @() #New-Object System.Array
[System.Collections.ArrayList]$inactiveMembers = @() #New-Object System.Array
$counter = 0
foreach ($Member in $RecentMembers) {
    $counter++
    Write-Progress -Activity "Processing Account Stats" -Status "Checking $($Member.UserPrincipalName)" -PercentComplete (($counter/$RecentMembers.Count) * 100)
    Try {
        #Write-Host "testing $($Member.UserPrincipalName.split("@")[0])"
        $adData = get-aduser $Member.UserPrincipalName.split("@")[0] -property extensionattribute3, pwdlastset
        #$adData.extensionattribute3
    }
    catch {
    }
    if ($adData.extensionattribute3 -eq "freshman") {
        $FreshmanMembers++
        $searchVar = "createdDateTime gt $($Date) and UserPrincipalName eq '$($Member.UserPrincipalName)'"
        try {
            $LoginEvent = Get-AzureADAuditSignInLogs -Filter "$($searchVar)" -Top 1
        }
        catch {
            Start-Sleep -Seconds 15
            $LoginEvent = Get-AzureADAuditSignInLogs -Filter "$($searchVar)" -Top 1
        }
        if ($null -ne $loginevent -or $adData.pwdlastset -ne 0) {
            $Hash = @{
                UserPrincipalName = $Member.UserPrincipalName
                StudentID         = $Member.ExtensionProperty.employeeId
                Login             = $LoginEvent.CreatedDateTime
                PasswordSet       = ([datetime]::fromfiletime($var.pwdlastset)).tostring("MM-dd-yyy")
            }
            $ActiveMembers += $Hash
        }
        else {
            $Hash = @{
                UserPrincipalName = $Member.UserPrincipalName
                StudentID         = $Member.ExtensionProperty.employeeId
                Login             = "None"
                PasswordSet       = ([datetime]::fromfiletime($var.pwdlastset)).tostring("MM-dd-yyy")
            }
            $inactiveMembers += $Hash
        }
    }
    
}

if ($FreshmanMembers -eq $inactiveMembers.count + $ActiveMembers.count) {
    Write-host "All Users Accounted for"
}
else {
    Write-Host "Mismatch in total count to sorted" -BackgroundColor Red -ForegroundColor Black
}

$OutputTarget = "D:\tmp\StudentMelt"

$rundate = Get-Date
$datestamp = $rundate.ToString("yyyyMMdd-HHmm")

$inactiveMembers | Select-Object -Property @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName } }, @{Name = 'StudentID'; Expression = { $_.StudentID } }, @{Name = 'Login'; Expression = { $_.Login }}, @{Name = 'PasswordSet'; Expression = { $_.PasswordSet } } | Export-Csv -Path "$($OutputTarget)\StudentMelt_Inactive_$($datestamp).csv" -NoTypeInformation
$ActiveMembers | Select-Object -Property @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName } }, @{Name = 'StudentID'; Expression = { $_.StudentID } }, @{Name = 'Login'; Expression = { $_.Login }}, @{Name = 'PasswordSet'; Expression = { $_.PasswordSet } } | Export-Csv -Path "$($OutputTarget)\StudentMelt_Active_$($datestamp).csv" -NoTypeInformation