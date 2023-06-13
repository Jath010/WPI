function New-WPIgMSA {
    [CmdletBinding()]
    param (
        $Name,
        $TargetDC,
        $AllowedMachinesGroup
    )
    
    begin {
        if ($null -eq $targetDC) {
            $TargetDC = "AD-DC-P-W01"
        }
        $DNSHostName = $Name + ".wpi.edu"
    }
    
    process {
        Invoke-Command -computerName $TargetDC -scriptblock {New-ADServiceAccount -Name $Using:Name -DNSHostName $Using:DNSHostName -PrincipalsAllowedToRetrieveManagedPassword $Using:AllowedMachinesGroup -KerberosEncryptionType AES128, AES256}
    }
    
    end {
        
    }
}