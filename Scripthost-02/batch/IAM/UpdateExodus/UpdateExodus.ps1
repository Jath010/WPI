Clear-Host
## Update SQL Database for Account Self-Service
Import-Module ActiveDirectory

## Declare array to hold all users
$Users = @()

## Declare the OUs of the account locations
$OU_Employee = "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$OU_Student  = "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$OU_Alumni   = "OU=Alumni,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$OU_LOA      = "OU=Leave Of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu"

## Use Get-ADUser to pull all accounts and needed properties
$Users  = Get-ADUser -Filter * -SearchBase $OU_Employee -Properties DisplayName,DistinguishedName,UserPrincipalName,Title,Department,Description,EmployeeID,EmployeeNumber,WhenCreated
$Users += Get-ADUser -Filter * -SearchBase $OU_Student  -Properties DisplayName,DistinguishedName,UserPrincipalName,Title,Department,Description,EmployeeID,EmployeeNumber,WhenCreated
$Users += Get-ADUser -Filter * -SearchBase $OU_Alumni   -Properties DisplayName,DistinguishedName,UserPrincipalName,Title,Department,Description,EmployeeID,EmployeeNumber,WhenCreated
$Users += Get-ADUser -Filter * -SearchBase $OU_LOA      -Properties DisplayName,DistinguishedName,UserPrincipalName,Title,Department,Description,EmployeeID,EmployeeNumber,WhenCreated

## Specify the information needed for writing logs to Exodus DB.
$date = Get-Date -f M/d/yyyy
$localhost = $env:COMPUTERNAME
$operator = $env:USERNAME

foreach ($user in $users) {
    ## Clear all generated information within the foreach loop
    $username=$null;$id=$null;$pidm=$null;$time=$null;$record=$null

    ## Set the information needed to validate EXODUS records
    $username = $user.SamAccountName
    $id = $user.EmployeeID
    $pidm = $user.EmployeeNumber
    $time = Get-Date -f HH:mm:ss

    if (!$pidm) {continue}  ## Ignore if there is no PIDM specified on the account.

    ## Check the database to see if the user record exists.  If exists, continue to next record.
    $record =     Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
        "SELECT * FROM Logins WHERE Login = '$username'"
    if ($record) {continue}

    ## If no recrod exists, restore to database
    Write-Host "INSERT INTO Logins VALUES ('$PIDM','$ID','$Date','$username')"

    Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
        "INSERT INTO Logins VALUES ('$PIDM','$ID','$Date','$username')"

    Invoke-Sqlcmd -Database 'Exodus' -ServerInstance 'MSDB-01' -Query `
        "INSERT INTO Log VALUES ('$date','$time','$localhost','Powershell','Powershell','','$ID','$PIDM','$username','User $username successfully added to EXODUS by $operator')"
    }

c:
