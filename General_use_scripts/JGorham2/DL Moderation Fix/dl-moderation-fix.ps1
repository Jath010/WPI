function Switch-DLModeratorToSublist {
    [CmdletBinding()]
    param (
        $HostList
    )
    
    begin {
        $SMTP = "MB-" + $hostList + "@wpi0.onmicrosoft.com"
        $MBList = "MB-" + $HostList
        if ($null -eq (Get-DistributionGroup -Identity $MBList -ErrorAction SilentlyContinue)) {
            new-distributiongroup -Name $MBList -PrimarySmtpAddress $SMTP -HiddenGroupMembershipEnabled:$true
        }
        $BypassList = (get-distributiongroup $HostList).BypassModerationFromSendersOrMembers
    }
    
    process {
        foreach ($user in $BypassList) {
            try {
                Add-DistributionGroupMember -Identity $MBList -Member $user
            }
            catch {
                Write-Host "User $User is already in group"
            }
        }
    }
    
    end {
        Set-DistributionGroup -Identity $HostList -BypassModerationFromSendersOrMembers $MBList
    }
}

function Switch-DLSendersToSublist {
    [CmdletBinding()]
    param (
        $HostList
    )
    
    begin {
        $SMTP = "AS-" + $hostList + "@wpi0.onmicrosoft.com"
        $ASList = "AS-" + $HostList
        if ($null -eq (Get-DistributionGroup -Identity $ASList -ErrorAction SilentlyContinue)) {
            new-distributiongroup -Name $ASList -PrimarySmtpAddress $SMTP -HiddenGroupMembershipEnabled:$true
        }
        $SenderList = (get-distributiongroup $HostList).AcceptMessagesOnlyFromSendersOrMembers
    }
    
    process {
        foreach ($user in $SenderList) {
            try {
                Add-DistributionGroupMember -Identity $ASList -Member $user
            }
            catch {
                Write-Host "User $User is already in group"
            }
        }
    }
    
    end {
        Set-DistributionGroup -Identity $HostList -AcceptMessagesOnlyFromSendersOrMembers $ASList
    }
}

function get-DLStudentStatus {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $Student = Get-DistributionGroup dl-students
        $undergrad = Get-DistributionGroup dl-undergraduates
        $other = Get-DistributionGroup dl-other-students
        $part = Get-DistributionGroup dl-grads-parttime
        $full = Get-DistributionGroup dl-grads-fulltime
        $lists = $Student, $undergrad, $other, $part, $full
    }
    
    process {
        foreach ($List in $lists) {
            Write-host $list.Identity
            if($list.ModerationEnabled){
                Write-Host "Moderation Enabled" -BackgroundColor Green -ForegroundColor Black
            }else {
                Write-Host "Moderation Disabled" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -ModerationEnabled:$true
            }
            if ("$($list.Name)@wpi.edu" -eq $List.PrimarySmtpAddress) {
                Write-Host "Primary SMTP Correct" -BackgroundColor Green -ForegroundColor Black
            }else {
                Write-Host "Primary SMTP Wrong: $($list.PrimarySmtpAddress)" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -PrimarySmtpAddress "$($list.Name)@wpi.edu"
            }
            if($list.BypassNestedModerationEnabled){
                Write-Host "Nested Moderation Enabled" -BackgroundColor Green -ForegroundColor Black
            }else {
                Write-Host "Nested Moderation Disabled" -BackgroundColor Red
                Write-host "Attempting Repair"
                Set-DistributionGroup $list -BypassNestedModerationEnabled:$true
            }
        }
    }
    
    end {
        
    }
}

function get-DLEmployeeStatus {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $List = Get-DistributionGroup dl-allemployees
    }
    
    process {
        Write-host $list.Identity
            if(!$list.ModerationEnabled){
                Write-Host "Moderation Disabled" -BackgroundColor Green -ForegroundColor Black
            }else {
                Write-Host "Moderation Enabled" -BackgroundColor Red
            }
            if ("$($list.Name)@wpi.edu" -eq $List.PrimarySmtpAddress) {
                Write-Host "Primary SMTP Correct" -BackgroundColor Green -ForegroundColor Black
            }else {
                Write-Host "Primary SMTP Wrong: $($list.PrimarySmtpAddress)" -BackgroundColor Red
            }
    }
    
    end {
        
    }
}

function Get-DLModerationBypassMember {
    [CmdletBinding()]
    param (
        $DL,
        [parameter(ValueFromPipeline)]
        $User
    )
    
    begin {
        $BypassList = Get-DistributionGroupMember -identity mb-$DL
        $addresses = $BypassList.primarysmtpaddress
    }
    
    process {
        if ($addresses.Contains("$user")) {
            Write-Host "User Bypassing Moderation" -BackgroundColor Green -ForegroundColor Black
        }elseif ($BypassList -like "$User*") {
            Write-Host "User ($($BypassList -like "$User*")) Bypassing Moderation" -BackgroundColor Green -ForegroundColor Black
        }else {
            Write-Host "User Not Found in Bypass List" -BackgroundColor Red
        }
    }
    
    end {
        
    }
}

function Get-DistributionGroupAuditLog {
    [CmdletBinding()]
    param (
        $DL
    )
    
    begin {
        $log = Search-AdminAuditLog -Cmdlets set-distributiongroup -EndDate (get-date) -StartDate (get-date).AddDays(-7) -IsSuccess $true
    }
    
    process {
        $StudentsLog = $log | Where-Object {$_.ObjectModified -eq $DL}

        foreach ($Log in $StudentsLog) {
            Write-Host "Rundate: " $log.rundate -BackgroundColor Green -ForegroundColor Black
            write-host "Caller: " $log.Caller
            Write-Host "Cmdlet Parameters:"
            $log.cmdletparameters|format-table
        }
    }
    
    end {
        
    }
}