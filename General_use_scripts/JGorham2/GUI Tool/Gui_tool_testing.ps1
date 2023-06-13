#Check to see if RsatTools are installed

if ($null -eq (Get-Module -Name ActiveDirectory)) {
    Install-WindowsFeature -Name "RSAT-AD-PowerShell" -IncludeAllSubFeature
    if ($null -eq (Get-Module -Name ActiveDirectory)) {
        Import-Module -Name ActiveDirectory
    }else{
        Write-Host "Couldn't install RSAT/AD Module" -BackgroundColor Red -ForegroundColor Black
        Return
    }
}else{
    Import-Module -Name ActiveDirectory
}

# Check Exchange

if ($null -eq (get-module -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -Confirm:$false
    if ($null -eq (get-module -Name ExchangeOnlineManagement)) {
        Import-Module -Name ExchangeOnlineManagement
    }else{
        Write-Host "Couldn't install Exchange Module" -BackgroundColor Red -ForegroundColor Black
        Return
    }
}else{
    Import-Module -Name ExchangeOnlineManagement
}