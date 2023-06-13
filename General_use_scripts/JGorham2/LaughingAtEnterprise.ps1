function Get-Competent {
    [CmdletBinding()]
    param (
        $FirstName,
        $MiddleName,
        $LastName
    )
    
    begin {
        if ($null -eq $MiddleName) {
            $possibleUsername = $FirstName.substring(0, 1).ToLower() + $LastName.ToLower()
        }
        else {
            $possibleUsername = $FirstName.substring(0, 1).ToLower() + $MiddleName.Substring(0, 1).ToLower() + $LastName.ToLower()
        }
    }
    
    process {
        $count = 2
        try{
            $test = get-aduser $possibleUsername -ErrorAction SilentlyContinue
            $test = $possibleUsername
            $taken = 1
        }catch{
            $test = $possibleUsername
            $taken = 0
        }

        if($taken -eq 1){
            do {
                $test = $possibleUsername + $count
                try {
                    $test = get-aduser $test -ErrorAction SilentlyContinue
                    $count++
                }
                catch {
                    $possibleUsername = $test
                    $taken = 0
                }
            } until ($taken -eq 0)
        }
    }
    
    end {
        $possibleUsername
    }
}

[int]$currentNumber = (Get-ADUser -Filter "EmployeeID -like '*'" -Properties EmployeeID | Sort-Object [int]EmployeeID -Descending | Select-Object -First 1).EmployeeID

$freeNumber = $currentNumber + 1

$freeNumber

#757279089

$Filter = "(&(objectCategory=person)(employeeid=*))"
$RootOU = "OU=Accounts,DC=admin,DC=wpi,DC=edu"
$scope = "Subtree"

$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Searcher.Filter = $Filter
$Searcher.PageSize = "20000"
$Searcher.SearchScope = $Scope # Either: "Base", "OneLevel" or "Subtree"
$Users = $Searcher.FindAll()

#[int]$currentNumber = ($Users|Sort-Object [int]EmployeeID|Select-Object -First 1).properties.employeeid

[int]$CurrentNumber = 999000999
foreach ($user in $Users) {
    if ([int]($user.properties.employeeid)[0] -lt $currentNumber -and [int]($user.properties.employeeid)[0] -ne 1234567 -and [int]($user.properties.employeeid)[0] -ne 0) {
        [int]$currentNumber = [int]($user.properties.employeeid)[0]
        $currentNumber
    }
}

#Top number = 999000999 || 899914804
#Bottom number = 100001942 || 100016468

($UserList|Sort-Object [int]EmployeeID|Select-Object -First 1).properties.employeeid