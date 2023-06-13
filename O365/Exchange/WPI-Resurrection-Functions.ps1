# Functions to aid in the restoration of WPI accounts
# Functions to remove and see the reciept filters of mailboxes for the purposes of restoring accounts

function Get-WPIRecieptFilter {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )

    process {
        $mailbox = Get-mailbox $user
        if ("" -ne $mailbox.AcceptMessagesOnlyFrom -and "" -ne $mailbox.AcceptMessagesOnlyFromSendersOrMembers) {
            $answer = Read-Host "User $($User) has their reciept filter set. Would you like to remove? Y/N"
            if ($answer -eq "Y") {
                Remove-WPIRecieptFilter $User
            }
        }
    }
}

function Remove-WPIRecieptFilter {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )
    
    begin {
    }
    
    process {
        Write-Verbose "Clearing Reciept Filter from User $($User)"
        Set-Mailbox $User -AcceptMessagesOnlyFrom $null
        Set-Mailbox $User -AcceptMessagesOnlyFromSendersOrMembers $null
    }
    
    end {
    }
}

function Get-WPIExpirationStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )
    process {
        $ADEntry = Get-ADuser $user -Properties AccountExpirationDate
        if ($null -ne $ADEntry.AccountExpirationDate) {
            $answer = Read-Host "User $($User) has an expiration date. Would you like to remove? Y/N"
            if ($answer -eq "Y") {
                Remove-WPIExpirationDate $User
            }
        }
    }
}

function Remove-WPIExpirationDate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )
    
    begin {
    }
    
    process {
        Write-Verbose "Clearing expiration date from User $($User)"
        Set-ADUser $User -AccountExpirationDate $null
    }
    
    end {
    }
}

function Repair-WPIResurrection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [String]
        $User
    )
    
    begin {
    }
    
    process {
        Get-WPIExpirationStatus $User
        Get-WPIRecieptFilter $User
        Write-Host "Removed filter and expiration date for User $($User)"
    }
    
    end {
    }
}

function Get-WPIFuckedAccount {
    [CmdletBinding()]
    param (
        
    )
    
    $FilteredUsers = Get-Mailbox -filter { AcceptMessagesOnlyFrom -ne $null } -ResultSize Unlimited | select-object Alias
    foreach ($User in $FilteredUsers) {
        $ADUser = (get-aduser $User.Alias -Properties DistinguishedName, Enabled)
        $DistinguishedName = $ADUser.DistinguishedName
        if ($DistinguishedName -like "*,OU=Students*" -or $DistinguishedName -like "*,OU=Employees*" -and $ADUser.Enabled -eq $True) {
            $User.Alias
        }
    }
}

function Get-WPILoopedUsers {
    [CmdletBinding()]
    param (
        
    )
    
    $FilteredUsers = Get-Mailbox -filter { AcceptMessagesOnlyFrom -ne $null } -ResultSize Unlimited | select-object Alias
    foreach ($User in $FilteredUsers) {
        $ADUser = (get-aduser $User.Alias -Properties DistinguishedName, Enabled)
        $DistinguishedName = $ADUser.DistinguishedName
        if ($DistinguishedName -like "*,OU=Students*" -or $DistinguishedName -like "*,OU=Employees*" -and $ADUser.Enabled -eq $True) {
            $User.Alias
        }
    }
}

function Repair-WPILoopedUsers {
    [CmdletBinding()]
    param (
        
    )
    
    $FilteredUsers = Get-Mailbox -filter { AcceptMessagesOnlyFrom -ne $null } -ResultSize Unlimited | select-object Alias
    foreach ($User in $FilteredUsers) {
        $ADUser = (get-aduser $User.Alias -Properties DistinguishedName, Enabled)
        $DistinguishedName = $ADUser.DistinguishedName
        if ($DistinguishedName -like "*,OU=Students*" -or $DistinguishedName -like "*,OU=Employees*" -and $ADUser.Enabled -eq $True) {
            Remove-WPIRecieptFilter $User.Alias
        }
    }
}

function Repair-WPIFuckedAccount {
    [CmdletBinding()]
    param (
        
    )
    
    $FilteredUsers = Get-Mailbox -filter { AcceptMessagesOnlyFrom -ne $null } -ResultSize Unlimited | select-object Alias
    foreach ($User in $FilteredUsers) {
        $ADUser = (get-aduser $User.Alias -Properties DistinguishedName, Enabled)
        $DistinguishedName = $ADUser.DistinguishedName
        if (($DistinguishedName -like "*,OU=Students*" -or $DistinguishedName -like "*,OU=Employees*") -and $ADUser.Enabled -eq $True) {
            Write-Verbose "Clearing Reciept Filter from User $($User)"
            Set-Mailbox $User -AcceptMessagesOnlyFrom $null
            Set-Mailbox $User -AcceptMessagesOnlyFromSendersOrMembers $null
            Write-Verbose "Clearing expiration date from User $($User)"
            Set-ADUser $User -AccountExpirationDate $null
        }
    }
}

function Get-WPIExpiredAccount {
    [CmdletBinding()]
    param (
        
    )
    $date = Get-date
    $ExpiredStudents = Get-ADUser -filter { Enabled -eq $True -and AccountExpirationDate -lt $date } -Properties Enabled, AccountExpirationDate -SearchBase "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    $ExpiredEmployees = Get-ADUser -filter { Enabled -eq $True -and AccountExpirationDate -lt $date } -Properties Enabled, AccountExpirationDate -SearchBase "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    forEach ($Student in $ExpiredStudents) {
        if (Get-Mailbox $student.SamAccountName -filter { AcceptMessagesOnlyFrom -ne $null }) {
            $student.SamAccountName
        }
    }
    forEach ($Employees in $ExpiredEmployees) {
        if (Get-Mailbox $Employees.SamAccountName -filter { AcceptMessagesOnlyFrom -ne $null }) {
            $Employees.SamAccountName
        }
    }
}

function Repair-WPIExpiredAccount {
    [CmdletBinding()]
    param (
        
    )

    $date = Get-date
    $ExpiredAbsence = Get-ADUser -filter * -Properties Enabled, AccountExpirationDate -SearchBase "OU=Leave Of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    $ExpiredStudents = Get-ADUser -filter { Enabled -eq $True -and AccountExpirationDate -lt $date } -Properties Enabled, AccountExpirationDate -SearchBase "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    $ExpiredEmployees = Get-ADUser -filter { Enabled -eq $True -and AccountExpirationDate -lt $date } -Properties Enabled, AccountExpirationDate -SearchBase "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    forEach ($Student in $ExpiredStudents) {
        if ("" -ne (Get-Mailbox $student.SamAccountName).AcceptMessagesOnlyFrom) {
            Write-Verbose "Clearing Reciept Filter from User $($student.SamAccountName)"
            Set-Mailbox $student.SamAccountName -AcceptMessagesOnlyFrom $null
            Set-Mailbox $student.SamAccountName -AcceptMessagesOnlyFromSendersOrMembers $null
            Write-Verbose "Clearing expiration date from User $($student.SamAccountName)"
            Set-ADUser $student.SamAccountName -AccountExpirationDate $null
        }
    }
    forEach ($Employees in $ExpiredEmployees) {
        if ("" -ne (Get-Mailbox $Employees.SamAccountName).AcceptMessagesOnlyFrom) {
            Write-Verbose "Clearing Reciept Filter from User $($Employees.SamAccountName)"
            Set-Mailbox $Employees.SamAccountName -AcceptMessagesOnlyFrom $null
            Set-Mailbox $Employees.SamAccountName -AcceptMessagesOnlyFromSendersOrMembers $null
            Write-Verbose "Clearing expiration date from User $($Employees.SamAccountName)"
            Set-ADUser $Employees.SamAccountName -AccountExpirationDate $null
        }
    }
    forEach ($Absent in $ExpiredAbsence) {
        if ("" -ne (Get-Mailbox $Absent.SamAccountName).AcceptMessagesOnlyFrom) {
            Write-Verbose "Clearing Reciept Filter from User $($Absent.SamAccountName)"
            Set-Mailbox $Absent.SamAccountName -AcceptMessagesOnlyFrom $null
            Set-Mailbox $Absent.SamAccountName -AcceptMessagesOnlyFromSendersOrMembers $null
            #Write-Verbose "Clearing expiration date from User $($Absent.SamAccountName)"
            #Set-ADUser $Absent.SamAccountName -AccountExpirationDate $null
        }
    }
}

function Get-WPILOAAcceptOnly {
    param (
        
    )
    $LOAUser = Get-ADUser -filter * -Properties Enabled, AccountExpirationDate -SearchBase "OU=Leave Of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu"
    foreach ($User in $LOAUser) {
        $Box = Get-Mailbox $User.samaccountname
        Write-Host $box.Alias $box.AcceptMessagesOnlyFrom
    }
}

function Remove-WPIMailLoop {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        $users = get-aduser -filter * -ResultSetSize $null 
    }
    
    process {
        foreach ($user in $users) {
            if ("" -ne (Get-Mailbox $user.SamAccountName -ErrorAction Ignore).AcceptMessagesOnlyFrom -and [bool](Get-Mailbox $user.SamAccountName -ErrorAction Ignore)) {
                Write-Verbose "Clearing Reciept Filter from User $($user.SamAccountName)"
                Set-Mailbox $user.SamAccountName -AcceptMessagesOnlyFrom $null
                Set-Mailbox $user.SamAccountName -AcceptMessagesOnlyFromSendersOrMembers $null
            }
        }
    }
    
    end {
    }
}

function Repair-MailboxAdminRights {
    [CmdletBinding()]
    param (
        $Mailbox
    )
    
    begin {
    }
    
    process {
        Add-MailboxPermission -Identity $Mailbox -User "Organization Management" -AccessRights FullAccess -InheritanceType All
    }
    
    end {
    }
}