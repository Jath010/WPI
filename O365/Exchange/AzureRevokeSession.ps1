##TODO: Add password scramble

<#
$password = ConvertTo-SecureString 'Please enter the new password' -AsPlainText -Force

Set-AzureADUserPassword -ObjectId  "a8d5e982-6c3d-406e-a533-a21b275e3d37" -Password $password
#>


function Revoke-WPIAzureSessions {
    [cmdletbinding()]
    param (
        $User
    )
    Connect-AzureAD
    $ObjectID = (Get-AzureADUser -SearchString $user).ObjectID
    Revoke-AzureADUserAllRefreshToken -ObjectID $ObjectID
}

function Clear-WPIAzureLogin {
    [CmdletBinding()]
    param (
        $User
    )
    
    begin {
        Connect-AzureAD
        $minLength = 8 ## characters
        $maxLength = 10 ## characters
        $length = Get-Random -Minimum $minLength -Maximum $maxLength
        $nonAlphaChars = 5
        $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
        $secPw = ConvertTo-SecureString -String $password -AsPlainText -Force
    }
    
    process {
        $ObjectID = (Get-AzureADUser -SearchString $user).ObjectID
        Revoke-AzureADUserAllRefreshToken -ObjectID $ObjectID
        Set-AzureADUserPassword -ObjectId $ObjectID -Password $secPw
    }
    
    end {
        Disconnect-AzureAD
    }
}