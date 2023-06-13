#Oh Boy We're Counting!
# This is the script for getting the total number of licensed and affiliated users in the environment

$progress = 0
$steps = 9
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Service Licenses" -PercentComplete (($progress/$steps)*100)
$Service = (Get-AzureADGroupMember -ObjectId "1209bcb3-9986-4d1e-9050-07269637ac59" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Faculty Licenses" -PercentComplete (($progress++/$steps)*100)
$faculty = (Get-AzureADGroupMember -ObjectId "18bcd282-89e8-4851-a282-10ced64a149d" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Affiliate Licenses" -PercentComplete (($progress++/$steps)*100)
$affiliates = (Get-AzureADGroupMember -ObjectId "194e5567-072e-435f-82f8-1108b8dd1304" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Retiree Licenses" -PercentComplete (($progress++/$steps)*100)
$retirees = (Get-AzureADGroupMember -ObjectId "27e1890b-ebcf-4765-b0f8-5cebc04af6c4" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Priv Licenses" -PercentComplete (($progress++/$steps)*100)
$priv = (Get-AzureADGroupMember -ObjectId "3961f6c6-7d15-49dc-817b-d7bed7fb0335" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Staff Licenses" -PercentComplete (($progress++/$steps)*100)
$staff = (Get-AzureADGroupMember -ObjectId "5d0b8e03-556d-4355-81de-4e362efd4961" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Student Licenses" -PercentComplete (($progress++/$steps)*100)
$students = (Get-AzureADGroupMember -ObjectId "7cf8b200-27c8-4ded-aba8-303145d7ca89" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Sysops Licenses" -PercentComplete (($progress++/$steps)*100)
$sysops = (Get-AzureADGroupMember -ObjectId "99661794-f302-4c9e-9a47-18cb0965d57c" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Counting Alumni Licenses" -PercentComplete (($progress++/$steps)*100)
$alumni = (Get-AzureADGroupMember -ObjectId "ac9142c1-f2a6-4aa6-aeba-46a1395ba5a1" -all:$true).count
Write-Progress -Activity "Gathering Azure Counts" -Status "Azure Count Complete" -PercentComplete (($progress++/$steps)*100)

$progress = 0
$steps = 9
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Affiliates" -PercentComplete (($progress/$steps)*100)
$ADAffiliate = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Affiliate'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Alumni" -PercentComplete (($progress++/$steps)*100)
$ADAlum = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Alum'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Faculty" -PercentComplete (($progress++/$steps)*100)
$ADFaculty = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Faculty'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Leave of Absence" -PercentComplete (($progress++/$steps)*100)
$ADLA = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'LA'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting `"None`"" -PercentComplete (($progress++/$steps)*100)
$ADNone = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'None'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Resource Mailboxes" -PercentComplete (($progress++/$steps)*100)
$ADResourceMailbox = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Resource Mailbox'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Service Accounts" -PercentComplete (($progress++/$steps)*100)
$ADService = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Service'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Staff" -PercentComplete (($progress++/$steps)*100)
$ADStaff = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Staff'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "Counting Students" -PercentComplete (($progress++/$steps)*100)
$ADStudent = (Get-ADUser -filter "Enabled -eq 'True' -and extensionattribute7 -like 'Student'" -Properties extensionattribute7).count
Write-Progress -Activity "Gathering AD Counts" -Status "AD Count Complete" -PercentComplete (($progress++/$steps)*100)

$ADTotal = $ADAffiliate + $ADAlum + $ADFaculty + $ADLA + $ADNone + $ADResourceMailbox + $ADService + $ADStaff + $ADStudent

$guest = (Get-AzureADUser -Filter "UserType eq 'Guest'" -all:$true).count

$total = $Service + $faculty + $affiliates + $retirees + $priv + $staff + $students + $sysops +$alumni

$enabled = (Get-ADUser -filter "Enabled -eq 'True'").count
$disabled = (Get-ADUser -filter "Enabled -eq 'False'").count

Write-Host "Azure License Counts"
Write-Host "Total number of Azure Service Licensed accounts: $Service"
Write-Host "Total number of Azure Faculty Licensed accounts: $Faculty"
Write-Host "Total number of Azure Affiliate Licensed accounts: $Affiliates"
Write-Host "Total number of Azure Retiree Licensed accounts: $retirees"
Write-Host "Total number of Azure Privileged  Licensed accounts: $priv"
Write-Host "Total number of Azure Staff Licensed accounts: $staff"
Write-Host "Total number of Azure Student Licensed accounts: $students"
Write-Host "Total number of Azure Sysops Licensed accounts: $sysops"
Write-Host "Total number of Azure Alumni Licensed accounts: $alumni"
Write-Host "Total number of Azure Licensed accounts: $Total"
Write-host "Total number of Azure Guest accounts: $guest"
Write-Host "Active Directory ExtensionAttribute7"
Write-host "Total number of Affiliate accounts: $ADAffiliate"
Write-host "Total number of Alum accounts: $ADAlum"
Write-host "Total number of Faculty accounts: $ADFaculty"
Write-host "Total number of LA accounts: $ADLA"
Write-host "Total number of None accounts: $ADNone"
Write-host "Total number of ResourceMailbox accounts: $ADResourceMailbox"
Write-host "Total number of Service accounts: $ADService"
Write-host "Total number of Staff accounts: $ADStaff"
Write-host "Total number of Student accounts: $ADStudent"
Write-host "Total number of Extension Attribute accounts: $ADTotal"
Write-host "Total number of Enabled accounts: $enabled"
Write-host "Total number of Disabled accounts: $disabled"