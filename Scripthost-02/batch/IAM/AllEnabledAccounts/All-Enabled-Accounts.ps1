<# All-Enabled-accounts.ps1

This script simply finds all enabled user accounts who aren't in the All Enabled Accounts security group in AD and adds them to it. 
The reason this exists is because the symptom tracker requires users be in a security group to have access.

Created By: Stephen Gemme
Created On: 08/18/20

All modification notes are in git.

#>

$dateString = Get-Date -format "MM-dd-yyyy_HHmm"
$logpath = "D:\wpi\Logs\IAM\AllEnabledAccounts\"
Start-Transcript -Path "$logpath\Changes-$dateString.txt"

# Delete files older than 1 week
Get-ChildItem $logpath -Recurse -Force -ea 0 | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-7)} |
ForEach-Object {
   $_ | del -Force
}

Get-ADUser -Filter 'Enabled -eq $true -and MemberOf -ne "CN=All Enabled Accounts,OU=Groups,DC=admin,DC=wpi,DC=edu" ' | Foreach-Object { 
    Write-Host "Adding user to group: " -noNewLine
    Write-Host -foregroundColor CYAN $_.SamAccountName
    Add-ADGroupMember -Identity "All Enabled Accounts" -Members $_ 
}

Stop-Transcript