# Example string
#  New-DynamicDistributionGroup -Name "Full Time Employees" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute10 -eq 'FullTimeEmployee')}

<#

((((RecipientTypeDetails -eq 'UserMailbox') -and (Office -like 'Fuller Labs*'))) -and (-not(Name -like 'SystemMailbox{*')) -and (-not(Name -like 'CAS_{*')) -and (-not(RecipientTypeDetailsValue -eq 'MailboxPlan')) -and (-not(RecipientTypeDetailsValue -eq 'DiscoveryMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'PublicFolderMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'ArbitrationMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuxAuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'SupervisoryReviewPolicyMailbox')))

#>
<#
New-DynamicDistributionGroup -Name "$($Building) DynList" -PrimarySmtpAddress "DLDyn-$($Building)@wpi.edu" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute5 -like "DB-$($building);*")} -whatif

#Gemme's concat hack
$array = @('FN','FA','F1','F2','DH','EP','NR')
$filterstring = ""extensionAttribute1 -eq '$($array -join ""' -or extensionAttribute1 -eq '"")'""
#####

$building = "DH"
$BuildingString = "DB-" + $Building + ";*"
New-DynamicDistributionGroup -Name "$($Building) DynList" -PrimarySmtpAddress "DLDyn-$($Building)@wpi.edu" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute5 -like $BuildingString)}



$filter = {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute3 -eq "junior" -or CustomAttribute3 -eq "sophomore" -or CustomAttribute3 -eq "freshman" -or CustomAttribute3 -eq "senior") -and ((CustomAttribute5 -like "DB-25T*") -or (CustomAttribute5 -like "DB-22S*") -or (CustomAttribute5 -like "DB-26H*") -or (CustomAttribute5 -like "DB-16E*") -or (CustomAttribute5 -like "DB-DH*") -or (CustomAttribute5 -like "DB-EH*") -or (CustomAttribute5 -like "DB-FD*") -or (CustomAttribute5 -like "DB-FH*") -or (CustomAttribute5 -like "DB-ME*") -or (CustomAttribute5 -like "DB-MH*") -or (CustomAttribute5 -like "DB-RH*") -or (CustomAttribute5 -like "DB-IH*") -or (CustomAttribute5 -like "DB-SA*") -or (CustomAttribute5 -like "DB-SB*") -or (CustomAttribute5 -like "DB-SC*") -or (CustomAttribute5 -like "DB-EA*") -or (CustomAttribute5 -like "DB-FA*"))}
New-DynamicDistributionGroup -Name "$($Building) DynList" -PrimarySmtpAddress "DLDyn-$($Building)@wpi.edu" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute3 -eq "junior" -or CustomAttribute3 -eq "sophomore" -or CustomAttribute3 -eq "freshman" -or CustomAttribute3 -eq "senior") -and ((CustomAttribute5 -like "DB-25T*") -or (CustomAttribute5 -like "DB-22S*") -or (CustomAttribute5 -like "DB-26H*") -or (CustomAttribute5 -like "DB-16E*") -or (CustomAttribute5 -like "DB-DH*") -or (CustomAttribute5 -like "DB-EH*") -or (CustomAttribute5 -like "DB-FD*") -or (CustomAttribute5 -like "DB-FH*") -or (CustomAttribute5 -like "DB-ME*") -or (CustomAttribute5 -like "DB-MH*") -or (CustomAttribute5 -like "DB-RH*") -or (CustomAttribute5 -like "DB-IH*") -or (CustomAttribute5 -like "DB-SA*") -or (CustomAttribute5 -like "DB-SB*") -or (CustomAttribute5 -like "DB-SC*") -or (CustomAttribute5 -like "DB-EA*") -or (CustomAttribute5 -like "DB-FA*"))}




New-DynamicDistributionGroup -Name "$($Building) DynList" -PrimarySmtpAddress "DLDyn-$($Building)@wpi.edu" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute3 -eq "junior" -or "sophomore" -or "freshman" -or "senior") -and ((CustomAttribute5 -like "DB-25T" -or "DB-22S" -or "DB-26H" -or "DB-16E" -or "DB-DH" -or "DB-EH" -or "DB-FD" -or "DB-FH" -or "DB-ME" -or "DB-MH" -or "DB-RH" -or "DB-IH" -or "DB-SA" -or "DB-SB" -or "DB-SC" -or "DB-EA" -or "DB-FA"))}
Set-DynamicDistributionGroup -Identity "$($Building) DynList" -HiddenFromAddressListsEnabled:$true
New-DistributionGroup -Name "$($Building) List" -PrimarySmtpAddress "DL-$($Building)@wpi.edu" -Members "DLDyn-$($Building)@wpi.edu"
Set-DistributionGroup -Identity "$($Building) List" -HiddenFromAddressListsEnabled:$true

#this bit will just resolve the conditional at run time, ex. I ran this in january, so it just feeds a 01 into the filter
New-DynamicDistributionGroup -Name "DLDyn-jmgorham-january" -PrimarySmtpAddress "DLDyn-jmgorham-january@wpi.edu" -IncludedRecipients "MailboxUsers" -ConditionalCustomAttribute14 "$((get-date).Tostring('MM'))" -whatif

"(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute14 -like ""$((get-date).Tostring('MM'))"")"
#That line resolves to (RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute14 -like "01") so if it's accepted by a filter then it means we can programmaically fill

{(RecipientTypeDetails -eq 'UserMailbox') -and (CustomAttribute15 -like "?*CS/")}


#Get Group Membership
#$FTE = Get-DynamicDistributionGroup "Full Time Employees"
#Get-Recipient -RecipientPreviewFilter $FTE.RecipientFilter -OrganizationalUnit $FTE.RecipientContainer

function get-DynList {
    [CmdletBinding()]
    param (
        $DynamicList
    )
    $List = Get-DynamicDistributionGroup -Identity $DynamicList
    Get-Recipient -RecipientPreviewFilter $List.RecipientFilter -OrganizationalUnit $List.RecipientContainer
}

#original attempt ((((((CustomAttribute6 -like 'PADV-xkong*') -or (CustomAttribute6 -like 'PADV-.*xkong.*'))) -and (RecipientType -eq 'UserMailbox'))) -and (-not(Name -like 'SystemMailbox{*')) -and (-not(Name -like 'CAS_{*')) -and (-not(RecipientTypeDetailsValue -eq 'MailboxPlan')) -and (-not(RecipientTypeDetailsValue -eq 'DiscoveryMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'PublicFolderMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'ArbitrationMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuxAuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'SupervisoryReviewPolicyMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'GuestMailUser')))
#fails ((((((CustomAttribute6 -like 'PADV-*xkong*'))) -and (RecipientType -eq 'UserMailbox'))) -and (-not(Name -like 'SystemMailbox{*')) -and (-not(Name -like 'CAS_{*')) -and (-not(RecipientTypeDetailsValue -eq 'MailboxPlan')) -and (-not(RecipientTypeDetailsValue -eq 'DiscoveryMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'PublicFolderMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'ArbitrationMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuxAuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'SupervisoryReviewPolicyMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'GuestMailUser')))
#this works, so we can't use foo*bar* only foobar* 
((((CustomAttribute6 -like 'PADV-xkong*') -and (RecipientType -eq 'UserMailbox'))) -and (-not(Name -like 'SystemMailbox{*')) -and (-not(Name -like 'CAS_{*')) -and (-not(RecipientTypeDetailsValue -eq 'MailboxPlan')) -and (-not(RecipientTypeDetailsValue -eq 'DiscoveryMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'PublicFolderMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'ArbitrationMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuxAuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'SupervisoryReviewPolicyMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'GuestMailUser')))


((((((Alias -eq 'jmgorham2') -or (CustomAttribute6 -like 'PADV-xkong*'))) -and (RecipientType -eq 'UserMailbox'))) -and (-not(Name -like 'SystemMailbox{*')) -and (-not(Name -like 'CAS_{*')) -and (-not(RecipientTypeDetailsValue -eq 'MailboxPlan')) -and (-not(RecipientTypeDetailsValue -eq 'DiscoveryMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'PublicFolderMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'ArbitrationMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'AuxAuditLogMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'SupervisoryReviewPolicyMailbox')) -and (-not(RecipientTypeDetailsValue -eq 'GuestMailUser')))


function FunctionName {
    [CmdletBinding()]
    param (
        OptionalParameters
    )
    
}

Get-aduser -filter {Enabled -eq $true -and extensionAttribute6 -like 'PADV-*'} -property extensionAttribute6 | Where-Object {$_.extensionAttribute6 -match "PADV-[^ ](.+);*"} | Select-Object -expandproperty extensionAttribute6 | Foreach-Object {$_.split(";")[0].split("-")[1]} | Sort | Get-Unique


#>