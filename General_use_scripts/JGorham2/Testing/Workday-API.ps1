

function Get-WorkdayCreds {
    [CmdletBinding()]
    param (
        $Username,
        [SecureString] $Password
    )
    
    $WorkdayURI = "wd5-impl-services1.workday.com/ccx/service/wpi_preview/Student_core/v39.0"
    $requestFile = 'request.xml'
    $responseFile = 'response.xml'

    $authorizationHeaderValue = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($Username):$($Password)"))

    $Headers = @{
        Authorization = $authorizationHeaderValue
    }

    Invoke-WebRequest -Uri $WorkdayURI -Headers $Headers -Method Post -ContentType "text/xml" -InFile $requestFile -OutFile $responseFile

}