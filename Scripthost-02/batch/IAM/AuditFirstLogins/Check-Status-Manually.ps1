Param (
   [Parameter(Mandatory = $true)]
   [String]
   $csvPath
)

[System.Collections.ArrayList]$Global:allNotAccessed = @()
[System.Collections.ArrayList]$Global:allAccessed = @()


if (Test-Path $csvPath){
    $list = Import-CSV -Path $csvPath

    Foreach ($user in $list) {
        Write-Host "Checking user: " -noNewLine
        Write-Host -foregroundColor CYAN $user.SamAccountName -noNewLine

        $whenCreated = Get-ADUser $user.SamAccountName -Property whenCreated | Select-Object -expandProperty whenCreated
        $MFAStatus = Get-MSOLUser -userprincipalname $user.userPrincipalName | Select-Object -expandproperty StrongAuthenticationMethods

        $entry = [PSCustomObject]@{
                'Name'              = $user.Name
                'Banner ID'         = $user.ID
                'Username'          = $user.SamAccountName
                'WPI Email'         = $user.userprincipalname
            }

        if ($null -eq $MFAStatus){
            Write-Host -foregroundColor RED " No MFA enabled."

            # We use LastUserActionTime because it's the most accurate. See: https://o365reports.com/2019/06/18/office-365-users-last-logon-time-incorrect/
            $lastUserActionTime = Get-MailboxStatistics -Identity $user.userPrincipalName | Select-Object -expandProperty lastUserActionTime

            # If they also have no action on their Exchange, they definitely haven't logged in to email.
            if ($null -eq $lastUserActionTime){
               $Global:allNotAccessed.Add($entry) | Out-Null
            }
            # Check if they were created more than 2 weeks ago (they have 14 days to defer MFA before it's forced)
            elseif ((Get-Date $user.whenCreated) -lt (Get-Date).AddDays(-14) ){
               # They definitely haven't interacted with their account yet. 
               $Global:allNotAccessed.Add($entry) | Out-Null
            }
            else {
               # If they've been created more than a few days ago, we check if they've been interacting without MFA.
               # Otherwise, we give them more time to login and whatnot, so we don't report them.
                if ((Get-Date $user.whenCreated) -gt (Get-Date).AddDays(-3)) {
                    # If they've interacted with their email within 48 hours, they've likely been suppressing MFA.
                    if (-NOT (Get-Date $lastUserActionTime) -gt (Get-Date).AddDays(-2)){
                        $Global:allNotAccessed.Add($entry) | Out-Null
                    }
                    else {
                        $Global:allAccessed.Add($entry) | Out-Null
                    }
                }
            }
        }
        else {
            Write-Host -foregroundColor GREEN " MFA Enabled"
            $Global:allAccessed.Add($entry) | Out-Null
        }
    }

    # Derive our new path based on the given path.
    $newPath = "$PSScriptRoot\Manual-Accessed-Results.csv"
    Write-Host "Saving results to $newPath"

    $Global:allAccessed | Export-CSV -path $newPath -noTypeInformation
}
else {
    Write-Host -foregroundColor RED "CSV Not Found"
}
