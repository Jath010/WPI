$BlockedWords = ""
$PrefixSuffix = ""

Connect-AzureAD

try
    {
        $template = Get-AzureADDirectorySettingTemplate | Where-Object {$_.displayname -eq "group.unified"}
        $settingsCopy = $template.CreateDirectorySetting()
        New-AzureADDirectorySetting -DirectorySetting $settingsCopy
        $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
    }
catch
    {
        $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id       
    }

$settingsCopy = Get-AzureADDirectorySetting -Id $settingsObjectID

$SettingsCopy["PrefixSuffixNamingRequirement"] = $PrefixSuffix

$SettingsCopy["EnableMSStandardBlockedWords"] = $true

$SettingsCopy["CustomBlockedWordsList"] = $BlockedWords

Set-AzureADDirectorySetting -Id $settingsObjectID -DirectorySetting $settingsCopy

(Get-AzureADDirectorySetting -Id $settingsObjectID).Values