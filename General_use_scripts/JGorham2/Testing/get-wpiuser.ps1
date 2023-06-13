function Get-WPIADUser {
    [CmdletBinding()]
    param (
        $User
    )
    
    begin {
        $User = $user.split("@")[0]
    }
    
    process {
        $UserData = get-aduser $user -Properties *
        $DataOut = @{
            DisplayName = $UserData.DisplayName
            Email = $UserData.mail
            Year = $UserData.extensionattribute3
            Affiliation = $UserData.extensionAttribute7
            Major = ($UserData.extensionAttribute4 -match "MJ-(.*);MN-") | ForEach-Object {$Matches[1]}
            Minor = ($UserData.extensionAttribute4 -match "MJ-.*;MN-(.*)") | ForEach-Object {$Matches[1]}
            "Primary Advisor" = ($UserData.extensionAttribute6 -match "PADV-(.*);OADV-") | ForEach-Object {$Matches[1]}
            "Off Advisor" = ($UserData.extensionAttribute6 -match "PADV-.*;OADV-(.*)") | ForEach-Object {$Matches[1]}
            "Student Code" = $UserData.extensionattribute2
            "Employee Code" = $UserData.extensionattribute1
        }
    }
    
    end {
        $DataOut
    }
}