function New-PhishingSearch {
    [CmdletBinding()]
    param (
        $Address
    )
    
    begin {
        if (test-path C:\tmp\jmgorham2_prvCredential.xml) { $Credentials = Import-Clixml C:\tmp\jmgorham2_prvCredential.xml }
        if (!(Get-PSSession | Where-Object { $_.ComputerName -match 'protection.outlook.com' })) {
            if (!$Credentials) { $Credentials = Get-Credential -UserName "$($env:username)@wpi.edu" -Message 'Please enter your email address and password' }
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Credentials -Authentication Basic -AllowRedirection 
            Import-PSSession $Session -AllowClobber -DisableNameChecking 
            #$Host.UI.RawUI.WindowTitle = $UserCredential.UserName + " (Office 365 Security & Compliance Center)" 
        } 
    }
    
    process {
        New-ComplianceSearch -name "Phishing - $Address" -ExchangeLocation All -ContentMatchQuery "(c:c)(senderauthor=$Address)"
        Start-ComplianceSearch -Identity "Phishing - $Address"
    }
    
    end {
    }
}