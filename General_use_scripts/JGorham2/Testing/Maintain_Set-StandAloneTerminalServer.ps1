function Set-StandAloneTerminalServer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, Position = 0)]
        [string] $ServerAdminFQDN,
        $LocalServer,
        $KMS
    )
    ## Import Remote Desktop module
    Import-Module RemoteDesktop
    ## Set server variable
    $Server = $ServerAdminFQDN
    ## Set KMS variable
    if ($null -eq $KMS) {
        $KMS = 'kms-p-w01.wpi.edu'
    }
    ## Set Local Server variable
    if ($null -eq $LocalServer) {
        $LocalServer = $env:computername + '.admin.wpi.edu'
    }

    Write-Verbose "This WILL cause the remote server to reboot"

    ## Local Server Check
    If ( $Server -eq $LocalServer ) {

        Write-Verbose "This can not be run locally.  The RDS-Connection-Broker role won't install when running the setup locally. Run this from a remote server."

    }
    else {
        #Run a get and check if the collection already exists
        $CollectionState = get-RDSessionCollection -CollectionName $Server -ConnectionBroker $Server -ErrorAction silentlycontinue
            
        if ($null -eq $CollectionState) {
            ## Command below adds all the necessary roles for a single stand alone Terminal Server.
            ## This needs to be run from a remote server, can not be run locally, it errors when it tries to add the Connection Broker role
            New-RDSessionDeployment -ConnectionBroker $Server -WebAccessServer $Server -SessionHost $Server

            ## If this command fails for any reason remotely, run it locally
            # New-RDSessionCollection –CollectionName $Server –SessionHost $Server –CollectionDescription $Server -ConnectionBroker $Server
            try {
                New-RDSessionCollection -CollectionName $Server -SessionHost $Server -CollectionDescription $Server -ConnectionBroker $Server
                Write-verbose "Successfully installed"
            }
            catch {
                Write-verbose "Remote creation of RDSessionCollection failed, attempting locally"
                invoke-command -computername $server -scriptblock {New-RDSessionCollection -CollectionName $Using:Server -SessionHost $Using:Server -CollectionDescription $Using:Server -ConnectionBroker $Using:Server}
            }
        }else {
            Write-Verbose "$server Is already Configured as a Remote Desktop Host."
        }

        #Get the existence on the licence configuration
        $LicenseState = get-RDLicenseConfiguration -ConnectionBroker $Server -erroraction silentlycontinue

        if ($null -eq $LicenseState) {
            Set-RDLicenseConfiguration -LicenseServer $KMS -Mode PerUser -ConnectionBroker $Server -Force
            Write-Verbose "Successfully configured license server."
        }elseif ( $kms -ne $LicenseState.LicenseServer) {
            Set-RDLicenseConfiguration -LicenseServer $KMS -Mode PerUser -ConnectionBroker $Server -Force
        }else {
            Write-verbose "$Server is already pointed at the correct KMS"
        }

        #Add-RDServer -Server $KMS -Role RDS-LICENSING -ConnectionBroker $Server
    
    }
    
}

function Get-IntendedStandAloneTerminalServers {
    [CmdletBinding()]
    param (
        
    )
    ## Import Remote Desktop module
    Import-Module RemoteDesktop

    #Set Local Server and KMS vars
    $LocalServer = $env:computername + '.admin.wpi.edu'
    $KMS = 'kms-p-w01.wpi.edu'

    #Example Searchbase is where I found arc-research-01
    $Searchbase = "OU=Research,OU=Terminal Servers,DC=admin,DC=wpi,DC=edu"
    #Now we get all the computers in our target OU and make a list of their FQDNs
    $TerminalServerList = get-adcomputer -SearchBase $Searchbase -filter "*"|select-object DNSHostName

    foreach ($Server in $TerminalServerList) {
        Set-StandAloneTerminalServer -ServerAdminFQDN $Server.DNSHostName -localserver $LocalServer -KMS $KMS
    }
}