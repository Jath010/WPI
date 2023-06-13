#C:\Users\jmgorham2_prv\WPI\General_use_scripts\JGorham2\Export-Office365-Users-MFA-Status\AdditionalMFAData.ps1

function Get-MFAExtraData {
    [CmdletBinding()]
    param (
        $Path
    )
    
    begin {
        $CSV = Import-Csv $Path
        $ExportCSVReport=".\MFAOUUserReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
    }
    
    process {
        foreach($Entry in $CSV){
            $DisplayName = $Entry.DisplayName
            $UPN = $Entry.UserPrincipalName
            $MFAStatus = $entry.MFAStatus
            try{$samaccountname = ($Entry.UserPrincipalName.substring(0,$Entry.UserPrincipalName.IndexOf('@')))}
            catch{continue}
            try{$ADObject = Get-ADUser $samaccountname -Properties DistinguishedName}
            Catch{continue}
            $OU = $ADObject.DistinguishedName.substring($ADObject.DistinguishedName.IndexOf('OU'))

            $Result=@{'DisplayName'=$DisplayName;'UserPrincipalName'=$upn;'MFAStatus'=$MFAStatus;'OU'=$OU} 
            $Results= New-Object PSObject -Property $Result 
            $Results | Select-Object DisplayName,UserPrincipalName,MFAStatus,OU | Export-Csv -Path $ExportCSVReport -Notype -Append
        }
    }
    
    end {
    }
}