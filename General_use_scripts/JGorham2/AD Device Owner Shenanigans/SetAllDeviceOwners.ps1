#
# Script to get all machines that do not currently have ADMIN\download as their owner and set their owner to ADMIN\download
#


function Get-HereticDevices {
    param (
        
    )
    Get-ADComputer -Filter * -Properties ntSecurityDescriptor | where-object { $_.ntSecurityDescriptor.Owner -ne "ADMIN\download" }
}

function Set-DevicesToDownloadUser {
    param (
        
    )

    $DesiredOwner = "download"
    #Get our list of machines
    $Computers = Get-ADComputer -Filter * -Properties ntSecurityDescriptor | where-object { $_.ntSecurityDescriptor.Owner -ne "ADMIN\$DesiredOwner" }

    #For each machine in our list, get the acl and then change it to the desired owner
    foreach ($Computer in $Computers) {
        #get our target
        try {
            $oAceObj = Get-Acl -Path ("ActiveDirectory:://RootDSE/" + $Computer.DistinguishedName)
        }
        catch {
            Write-Error "Failed to find the source object."
            return
        }

        try {
            $oNewOwnAce = New-Object System.Security.Principal.NTAccount($DesiredOwner)
        }
        catch {
            Write-Error "Failed to find the new owner object."
            return
        }

        try {
            $oAceObj.SetOwner($oNewOwnAce)
            Set-Acl -Path ("ActiveDirectory:://RootDSE/" + $oADObject.DistinguishedName) -AclObject $oAceObj
        }
        catch {
            $errMsg = "Failed to set the new new ACE on " + $Computer.Name
            Write-Error $errMsg
        }
    }
}