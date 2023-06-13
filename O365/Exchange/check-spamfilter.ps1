##Function to check if a name is in the spam filter

if (!$Credentials) { $Credentials = Get-Credential -UserName "$($env:username)@wpi.edu" -Message 'Please enter your email address and password' }
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $Credentials -Authentication Basic -AllowRedirection 
Import-PSSession $Session -AllowClobber -DisableNameChecking 

$targetAddress = Read-Host "Please enter target address"

$policy = Get-HostedContentFilterPolicy -Identity 'Default'
$spamfilter = $policy.BlockedSenders|Select-Object Sender
foreach($address in $spamfilter){
    if($address.Sender -like $targetAddress){Write-Host "Target already blocked."}
}

Get-PSSession | Remove-PSSession