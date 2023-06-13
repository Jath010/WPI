#Remove Licenses from a WPI Account

#Sku for P1: 9XacWBXK-UGIS1XsFaBnKgQrjQe98RFBu9S0sbNUzvQ
<#
ObjectId                                                                  SkuPartNumber                     PrepaidUnits                                                                 ConsumedUnits
--------                                                                  -------------                     ------------                                                                 -------------
589c76f5-ca15-41f9-884b-55ec15a0672a_e433b246-63e7-4d0b-9efa-7940fa3264d6 PROJECTESSENTIALS_FACULTY         class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_42ad914d-58fc-495b-9d65-45a9f9cbdb14 PROJECTESSENTIALS_STUDENT         class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_dbb6dc54-c03c-4c6a-89aa-dac1bdc81653 ECAL_SERVICES_FACULTY             class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_26ad4b5c-b686-462e-84b9-d7c22b46837f ATP_ENTERPRISE_FACULTY            class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_60023c66-283d-4785-9334-1d4ca7fd3a18 RIGHTSMANAGEMENT_STANDARD_FACULTY class LicenseUnitsDetail {...                                                2496
589c76f5-ca15-41f9-884b-55ec15a0672a_6470687e-a428-4b7a-bef2-8a291ad947c9 WINDOWS_STORE                     class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_314c4481-f395-4525-be8b-2ec4bb1e9d91 STANDARDWOFFPACK_STUDENT          class LicenseUnitsDetail {...                                                16920
589c76f5-ca15-41f9-884b-55ec15a0672a_e82ae690-a2d5-4d76-8d30-7c6e01e6022e STANDARDWOFFPACK_IW_STUDENT       class LicenseUnitsDetail {...                                                8350
589c76f5-ca15-41f9-884b-55ec15a0672a_84a661c4-e949-4bd2-a560-ed7766fcaf2b AAD_PREMIUM_P2                    class LicenseUnitsDetail {...                                                2258
589c76f5-ca15-41f9-884b-55ec15a0672a_28db6bcc-8442-405b-9ebb-e2f4da7355ed EXCHANGE_STANDARD_ALUMNI          class LicenseUnitsDetail {...                                                13029
589c76f5-ca15-41f9-884b-55ec15a0672a_a403ebcc-fae0-4ca2-8c8c-7a907fd6c235 POWER_BI_STANDARD                 class LicenseUnitsDetail {...                                                2
589c76f5-ca15-41f9-884b-55ec15a0672a_efccb6f7-5641-4e0e-bd10-b4976e1bf68e EMS                               class LicenseUnitsDetail {...                                                2265
589c76f5-ca15-41f9-884b-55ec15a0672a_99fc2803-fa72-42d3-ae78-b055e177d275 INTUNE_A_VL                       class LicenseUnitsDetail {...                                                0
589c76f5-ca15-41f9-884b-55ec15a0672a_078d2b04-f1bd-4111-bbd4-b4b1b354cef4 AAD_PREMIUM                       class LicenseUnitsDetail {...                                                8348
589c76f5-ca15-41f9-884b-55ec15a0672a_12b8c807-2e20-48fc-b453-542b6ee9d171 OFFICESUBSCRIPTION_FACULTY        class LicenseUnitsDetail {...                                                2267
589c76f5-ca15-41f9-884b-55ec15a0672a_c32f9321-a627-406d-a114-1f9c81aaafac OFFICESUBSCRIPTION_STUDENT        class LicenseUnitsDetail {...                                                8348
589c76f5-ca15-41f9-884b-55ec15a0672a_94763226-9b3c-4e75-a931-5c89701abe66 STANDARDWOFFPACK_FACULTY          class LicenseUnitsDetail {...                                                2729
589c76f5-ca15-41f9-884b-55ec15a0672a_ff14db38-7582-4a15-aa7d-a856f1e5c23c RIGHTSMANAGEMENT_STANDARD_STUDENT class LicenseUnitsDetail {...                                                8348
#>

function Remove-WPIAzureLicenses {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )
    
    begin {
    }
    
    process {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $AzureADUser = Get-AzureADUser -ObjectId $User"@wpi.edu"
        foreach ($License in $AzureADUser.AssignedLicenses) {
            $Licenses.RemoveLicenses = $License.SkuID

            Write-Verbose "$((Get-Date).ToString('HH:mm'))  [REMOVE] $License for $username ($displayName)"
            try {
                Set-AzureADUserLicense -ObjectId $User"@wpi.edu" -AssignedLicenses $Licenses -ErrorAction:Ignore
            }
            catch { }
        }
        
    }
    
    end {
    }
}

function Remove-AllDirectAzureLicenses {
    [CmdletBinding()]
    param ()
    Remove-WPIAzureLicenses
}


function Remove-DirectP1License {
    [CmdletBinding()]
    param (
        
    )
    
    $userlist = Get-AzureADUser -all:$true

    $planName = "AAD_PREMIUM"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID

    foreach ($user in $userlist) {
        $userUPN = $user.UserPrincipalName
        
        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $license.SkuId = $planSkuID
        $licenses.AddLicenses = $license
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
        $Licenses.AddLicenses = @()
        $Licenses.RemoveLicenses = $planSkuID
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    }
    

}

function Remove-DirectAlumLicense {
    [CmdletBinding()]
    param (
        
    )
    
    $userlist = Get-AzureADUser -all:$true

    $planName = "EXCHANGE_STANDARD_ALUMNI"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID

    foreach ($user in $userlist) {
        $userUPN = $user.UserPrincipalName
        
        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $license.SkuId = $planSkuID
        #$licenses.AddLicenses = $license
        #Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
        $Licenses.AddLicenses = @()
        $Licenses.RemoveLicenses = $planSkuID
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    }
    

}

function Remove-StudentPPlusSerial {
    param (
        
    )

    $count = 0
    $userlist = Get-AzureADUser -all:$true

    $planName = "OFFICESUBSCRIPTION_STUDENT"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $license.SkuId = $planSkuID
    $Licenses.AddLicenses = @()
    $Licenses.RemoveLicenses = $planSkuID

    Write-Host "Starting Parallel Process"
    #Remove-DirectP1LicenseParallel -userlist $userlist -licenses $licenses
    foreach ($user in $userlist) {
        $userUPN = $user.UserPrincipalName
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses -ErrorAction Ignore
        $count++
        Write-Host $count
    }
}

function Remove-P1Serial {
    param (
        
    )
    $userlist = Get-AzureADUser -all:$true

    $planName = "AAD_PREMIUM"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $license.SkuId = $planSkuID
    $Licenses.AddLicenses = @()
    $Licenses.RemoveLicenses = $planSkuID

    Write-Host "Starting Parallel Process"
    #Remove-DirectP1LicenseParallel -userlist $userlist -licenses $licenses
    foreach ($user in $userlist) {
        $userUPN = $user.UserPrincipalName
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    }
}

workflow Remove-DirectP1LicenseParallel {
    [CmdletBinding()]
    param (
        $userlist,
        $licenses
    )

    foreach -parallel ($user in $userlist) {
        $userUPN = $user.UserPrincipalName
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    }
    

}

function Get-StudentPPlusSerial {
    param (
        
    )
    $userlist = Get-AzureADUser -all:$true

    $planName = "OFFICESUBSCRIPTION_STUDENT"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID

    # $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    # $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    # $license.SkuId = $planSkuID
    # $Licenses.AddLicenses = @()
    # $Licenses.RemoveLicenses = $planSkuID

    Write-Host "Starting Parallel Process"
    #Remove-DirectP1LicenseParallel -userlist $userlist -licenses $licenses
    foreach ($user in $userlist) {
        Write-Progress -Activity "Search in Progress" -CurrentOperation $user.UserPrincipalName
        $userUPN = $user.UserPrincipalName
        $userlicenses = Get-AzureADUserLicenseDetail -ObjectId $userUPN
        foreach ($skuid in $userlicenses.SkuID) {
            if ($skuid -eq $planSkuID) {
                $userUPN
            }
        }
    }
}

function Get-StudentPPlusParallel {
    param (
        $Userlist
    )
    if($null -eq $userlist){
        $userlist = Get-AzureADUser -all:$true
    }
    

    $planName = "OFFICESUBSCRIPTION_STUDENT"
    $planSkuID = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
    
    get-ParallelLicense -userlist $userlist -planSkuID $planSkuID
}


workflow get-ParallelLicense {
    param (
        $userlist,
        $planSkuID
    )
    foreach -parallel ($user in $userlist) {
        Write-Progress -Activity "Search in Progress" -CurrentOperation $user.UserPrincipalName
        $userUPN = $user.UserPrincipalName
        $userlicenses = Get-AzureADUserLicenseDetail -ObjectId $userUPN
        foreach ($skuid in $userlicenses.SkuID) {
            if ($skuid -eq $planSkuID) {
                $userUPN
            }
        }
    }
    
}
