function Set-WPIExpirationDate {
    [cmdletbinding()]
    param (
        # Samaccountname
        [Parameter(ParameterSetName="Username")]
        [String]
        $Username,
        # Id Number
        [Parameter(ParameterSetName="IDNumber")]
        [int]
        $WPIIDNumber,
        [Parameter(ValueFromRemainingArguments = $true)]
        [string]
        $ExpirationDate
    )
    try{
        $date = [datetime]$ExpirationDate
    }
    catch{
        Write-host "Expiration date is not a valid date" -ForegroundColor Red
        Break
    }
    Switch ($pscmdlet.ParameterSetName){
        "IDNumber" {
            $User = Get-ADUser -filter {EmployeeID -eq $WPIIDNumber} -Properties EmployeeID -searchbase "OU=Accounts,DC=admin,DC=wpi,DC=edu"
            Write-Verbose "Setting Expiration Date for $($User.SamAccountName)"
            Set-ADUser $User.SamAccountName -AccountExpirationDate $date.ToString("MM/dd/yyy")
            $status = Get-ADUser $User.SamAccountName -Properties AccountExpirationDate
        }
        "Username" {
            Write-Verbose "Setting Expiration Date for $($Username)"
            Set-ADUser $Username -AccountExpirationDate $date.ToString("MM/dd/yyy")
            $status = get-aduser $Username -Properties AccountExpirationDate
        }
    }
    Write-Verbose "Expiration date for $($status.samaccountname) is now $($status.AccountExpirationDate)"
}

#Scrub the input file with a replace regex like "(\d+).*$" or "(\d{9})*$"

function Set-WPIExpirationDateFromFile {
    [cmdletbinding()]
    param(
        $FilePath,
        [switch]
        $NoScrub,
        $ExpirationDate
    )
    
    if(!$NoScrub){
        $file = get-content $FilePath
        Write-Verbose "File Reads: $($file)"
        $grep = $file -match [regex]::new('\d{9}')
        Write-Verbose "After initial scrubbing: $($grep)"
        $FilePath = $grep -replace "(\d{9})\s.*", '$1'
        Write-verbose "After final scrub: $($FilePath)"
    }
    if($FilePath -match "^C:"){$FilePath = get-content $FilePath}
    foreach($ID in $FilePath){
        Set-WPIExpirationDate -WPIIDNumber $ID -ExpirationDate $ExpirationDate
    }
}