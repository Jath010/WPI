#Function to pull data from banner
#Joshua Gorham 5/16/2019
#Install-Module InvokeQuery

Import-Module InvokeQuery

<# This is the old scripting I started working from
########Banner Connection info######################
#encrypted oracle connection string
#$BannerConnectionString = Get-Content $connfile 
#TNSNames.ora Prod line : PROD=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=bannerprod.wpi.edu)(PORT=1527)))(CONNECT_DATA=(SERVICE_NAME=prod.admin.wpi.edu)(server=dedicated)))
$BannerConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=bannerprod.wpi.edu)(PORT=1527)))(CONNECT_DATA=(SERVICE_NAME=prod.admin.wpi.edu)(server=dedicated)));User ID=locksys;Password=**********"

####################################################


#get list of group members from banner
#$command.CommandText = "Select distinct replace(wpi_email(pidm,'UNIX'),'@WPI.EDU','') as username, IND as is_instructor from courses where subj = '$subjectcode' AND crse = '$coursenumber' and sect = '$sectionnumber' and etrm = '$bnrterm' and ptrm = '$bnrpartofterm'"

#For InvokeQuery cmdlet
#$command.CommandText = New-SqlQuery "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username from pebempl, person where pebempl_pidm = person.pidm and PEBEMPL_ECLS_CODE = 'NE' and PEBEMPL_EMPL_STATUS = 'A' and dead is null and UNIX_EMAIL is not null and not exists (select * from swvlpos where swvlpos_pidm = person.pidm) and pebempl_orgn_code_home not in ('28200','245200')"
$QueryText = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from pebempl, person 
                        where pebempl_pidm = person.pidm
                        and PEBEMPL_ECLS_CODE = 'NE'
                        and PEBEMPL_EMPL_STATUS = 'A'
                        and dead is null
                        and UNIX_EMAIL is not null
                        and not exists (select * from swvlpos where swvlpos_pidm = person.pidm)
                        and pebempl_orgn_code_home not in ('28200','245200')"


#$command.CommandText = $usernamequery
$data = Invoke-OracleQuery -Sql $QueryText -ConnectionString $BannerConnectionString
#>

###############################################################################################
#   Current Saviynt View
$SaviyntTable = "GWVSVNT_EXPANDED"
#   I Tossed this in because GWVSVNT_EXPANDED was a temporary table to be replaced with GWVSVNT
###############################################################################################


function Invoke-WPIBannerQuery {
    param (
        $Query
    )
    $BannerConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=bannerprod.wpi.edu)(PORT=1527)))(CONNECT_DATA=(SERVICE_NAME=prod.admin.wpi.edu)(server=dedicated)));User ID=locksys;Password=**********"
    Invoke-OracleQuery -Sql $Query -ConnectionString $BannerConnectionString
}

<# 
$nonemployeeQuery = "select *
                    from GWVSVNT_EXPANDED
                    FETCH NEXT 1 ROWS ONLY"

$peopleList = Invoke-WPIBannerQuery -Query $nonemployeeQuery
 #>

function Get-WPIPrimaryAffiliation {
    param (
        $Username
    )
    $AffiliationQuery = "Select PRIMARY_AFFILIATION from $SaviyntTable Where USERNAME = '$Username' FETCH NEXT 1 ROWS ONLY"
    
    $Affiliation = Invoke-WPIBannerQuery -Query $AffiliationQuery

    $Affiliation.PRIMARY_AFFILIATION
}

function Get-WPIUserFromBanner {
    param (
        $Username
    )
    $UserQuery = "Select * from $SaviyntTable Where USERNAME = '$Username' FETCH NEXT 1 ROWS ONLY"
    
    Invoke-WPIBannerQuery -Query $UserQuery
}

function Get-WPIStudents {
    param (
        
    )
    $StudentQuery = "SELECT distinct USERNAME FROM $SaviyntTable WHERE PRIMARY_AFFILIATION = 'student'"

    Invoke-WPIBannerQuery -Query $StudentQuery
}

function Get-WPIActiveStatus {
    param (
        $username
    )

    $UnixAddress = $Username +"@WPI.EDU"

    $ActivityQuery = "SELECT IS_ACTIVE, PRIMARY_AFFILIATION FROM $SaviyntTable WHERE USERNAME = '$UnixAddress' FETCH NEXT 1 ROWS ONLY"

    Invoke-WPIBannerQuery -Query $ActivityQuery
}

function Get-WPIUserFromID{
    param (
        $IDNumber
    )
    $Search = Invoke-WPIBannerQuery "Select USERNAME as username FROM $SaviyntTable WHERE WPI_ID = '$IDNumber' FETCH NEXT 1 ROWS ONLY"
    $Search.USERNAME
}

function Get-WPIEmployeeUsernames {
    param (
        
    )
    $Search = Invoke-WPIBannerQuery "Select distinct CONCAT(USERNAME,'@WPI.EDU') as UserPrincipalName FROM $SaviyntTable WHERE PRIMARY_AFFILIATION = 'staff' or PRIMARY_AFFILIATION = 'affiliate' and USERNAME is not null and DEAD_IND = 'N' and IS_ACTIVE = '1' and EMP_CODE != 'RE'"
    $Search
}

function Get-WPIStudentUsernames {
    param (
        
    )
    $Search = Invoke-WPIBannerQuery "Select distinct CONCAT(USERNAME,'@WPI.EDU') as UserPrincipalName FROM $SaviyntTable WHERE PRIMARY_AFFILIATION = 'student' and USERNAME is not null and DEAD_IND = 'N' and IS_ACTIVE = '1'"
    $search
}

function Get-WPIFacultyUsernames {
    param (
        
    )
    $Search = Invoke-WPIBannerQuery "Select distinct CONCAT(USERNAME,'@WPI.EDU') as UserPrincipalName FROM $SaviyntTable WHERE PRIMARY_AFFILIATION = 'faculty' and USERNAME is not null and DEAD_IND = 'N' and IS_ACTIVE = '1'"
    $search
}

function Get-WPIAlumniUsernames {
    param (
        
    )
    #Alum don't have Unix Addresses in their entries
    $Search = Invoke-WPIBannerQuery "Select distinct concat(USERNAME,'@wpi.edu') as UserPrincipalName FROM $SaviyntTable WHERE PRIMARY_AFFILIATION = 'alum' and DEAD_IND = 'N' and IS_ACTIVE = '1'"
    $search
}
#Putting it here until I have a better location

function Set-WPIExpirationDate {
    [cmdletbinding()]
    param (
        # Samaccountname
        [Parameter(ParameterSetName="Username")]
        [String]
        $Username,
        # Id Number
        [Parameter(ParameterSetName="IDNumber")]
        [int]
        $WPIIDNumber,
        [Parameter(ValueFromRemainingArguments = $true)]
        [string]
        $ExpirationDate
    )
    try{
        $date = [datetime]$ExpirationDate
    }
    catch{
        Write-host "Expiration date is not a valid date" -ForegroundColor Red
        Break
    }
    Switch ($pscmdlet.ParameterSetName){
        "IDNumber" {
            $User = Get-ADUser -filter {EmployeeID -eq $WPIIDNumber} -Properties EmployeeID -searchbase "OU=Accounts,DC=admin,DC=wpi,DC=edu"
            Set-ADUser $User.SamAccountName -AccountExpirationDate $date.ToString("MM/dd/yyy")
        }
        "Username" {
            Set-ADUser $Username -AccountExpirationDate $date.ToString("MM/dd/yyy")
        }
    }
}

function Set-WPIExpirationDateFromFile {
    [cmdletbinding()]
    param(
        $FilePath
    )
    Get-Content $FilePath | foreach-object {Set-WPIExpirationDate -WPIIDNumber $_ -ExpirationDate 8/9/2019}
}
