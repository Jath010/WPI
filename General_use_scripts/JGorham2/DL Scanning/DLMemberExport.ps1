function Export-WPIDLMembers {
    [CmdletBinding()]
    param (
        
    )

    begin {
        $i = 0 
        
        $output = @()

        $CSVfile = "D:\wpi\Logs\DynamicGroupMembersExport.csv" #Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 
    
        #$Dgname = Read-Host "Enter the DG name or Range (Eg. DGname , DG*,*DG)" # overload? -filter "Name -like 'dl-*' -or Name -like 'adv-*'"
    
        $AllDG = Get-DistributionGroup -filter "Name -like 'dl-*' -or Name -like 'adv-*'" -resultsize unlimited
    } 

    process {
        Foreach ($dg in $allDg) {
    
            $Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
    
            if ($members.count -eq 0) {
                #$managers = $Dg | Select-Object @{Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
                #$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited
    
                $userObj = New-Object PSObject
    
                $userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
                $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
                $userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
                #$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
                #$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled
    
                $output += $UserObj  
    
            }
            else {
                Foreach ($Member in $members) {
    
                    #$managers = $Dg | Select-Object @{Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
                    #$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited
    
                    $userObj = New-Object PSObject
    
                    $userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
                    #$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
                    #$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
                    #$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
                    $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
                    $userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
                    #$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled
    
                    $output += $UserObj  
    
                }
            }
            # update counters and write progress
            $i++
            Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
            $output | Export-csv -Path $CSVfile -NoTypeInformation -Encoding UTF8
        }

    }
}
function Export-WPIADVMembers {
    [CmdletBinding()]
    param (
        
    )

    begin {
        $i = 0 
        
        $output = @()

        $CSVfile = "D:\wpi\Logs\AdvisingGroupMembersExport.csv" #Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 
    
        #$Dgname = Read-Host "Enter the DG name or Range (Eg. DGname , DG*,*DG)" # overload? -filter "Name -like 'dl-*' -or Name -like 'adv-*'"
    
        $AllDG = Get-DistributionGroup -filter "Name -like 'adv-*'" -resultsize unlimited
    } 

    process {
        Foreach ($dg in $allDg) {
    
            $Members = Get-DistributionGroupMember $Dg.name -resultsize unlimited
    
            if ($members.count -eq 0) {
                #$managers = $Dg | Select-Object @{Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
                #$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited
    
                $userObj = New-Object PSObject
    
                $userObj | Add-Member NoteProperty -Name "DisplayName" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Alias" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "RecipientType" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value EmptyGroup
                $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value EmptyGroup
                #$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
                $userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
                $userObj | Add-Member NoteProperty -Name "Advisor" -Value $DG.PrimarySmtpAddress.split("-")[1].split("@")[0]
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
                #$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
                #$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
                #$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled
    
                $output += $UserObj  
    
            }
            else {
                Foreach ($Member in $members) {
    
                    #$managers = $Dg | Select-Object @{Name = 'DistributionGroupManagers'; Expression = { [string]::join(";", ($_.Managedby)) } }
                    #$manageremail = Get-Mailbox $managers.DistributionGroupManagers -ErrorAction SilentlyContinue -resultsize unlimited
    
                    $userObj = New-Object PSObject
    
                    $userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Member.Name
                    #$userObj | Add-Member NoteProperty -Name "Alias" -Value $Member.Alias
                    #$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Member.RecipientType
                    #$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Member.OrganizationalUnit
                    $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Member.PrimarySmtpAddress
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group" -Value $DG.Name
                    $userObj | Add-Member NoteProperty -Name "Distribution Group Primary SMTP address" -Value $DG.PrimarySmtpAddress
                    $userObj | Add-Member NoteProperty -Name "Advisor" -Value $DG.PrimarySmtpAddress.split("-")[1].split("@")[0]
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers" -Value $managers.DistributionGroupManagers
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Managers Primary SMTP address" -Value $manageremail.primarysmtpaddress
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group OU" -Value $DG.OrganizationalUnit
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Type" -Value $DG.GroupType
                    #$userObj | Add-Member NoteProperty -Name "Distribution Group Recipient Type" -Value $DG.RecipientType
                    #$userObj | Add-Member NoteProperty -Name "Not Allowed from Internet" -Value $DG.RequireSenderAuthenticationEnabled
    
                    $output += $UserObj  
    
                }
            }
            # update counters and write progress
            $i++
            Write-Progress -activity "Scanning Groups . . ." -status "Scanned: $i of $($allDg.Count)" -percentComplete (($i / $allDg.Count) * 100)
            $output | Export-csv -Path $CSVfile -NoTypeInformation -Encoding UTF8
        }

    }
}

