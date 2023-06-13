
##################################################################################

#Install-Module InvokeQuery
Import-Module ActiveDirectory
Import-Module InvokeQuery

function Invoke-WPIBannerQuery {
    param (
        $Query
    )
    $BannerConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=bannerprod.wpi.edu)(PORT=1527)))(CONNECT_DATA=(SERVICE_NAME=prod.admin.wpi.edu)(server=dedicated)));User ID=locksys;Password=ihatecbord"
    Invoke-OracleQuery -Sql $Query -ConnectionString $BannerConnectionString
}

function Sync-Group($usernamequery, $groupname) {
    try {
        <# ########Banner Connection info######################
        $connfile = "C:\wpi\batch\ADGroups\locksys.txt"
        #encrypted oracle connection string
        $BannerConnectionString = Get-Content $connfile 

        $connection = New-Object System.Data.OracleClient.OracleConnection($BannerConnectionString)
        [System.Data.OracleClient.OracleCommand] $command = new-Object System.Data.OracleClient.OracleCommand
        $command.Connection = $connection
        $connection.Open()
        #################################################### #>

        # Banner Connection

        #Verify AD group exists
        Try {Get-ADGroup $groupname}
        Catch {
            "Active Directory group $groupname wasn't found."
            [GC]::Collect()
            return
        }

        #get list of group members from banner
        $data = Invoke-WPIBannerQuery $usernamequery

        $users = $data
        #A list of the users from banner for the comparison process
        $userarray = New-Object System.Collections.ArrayList
        
        #ADDING USERS TO AD Group
        
        foreach($user in $users) {
            #for each member in banner and not in group, add them
            try {
                Add-ADGroupMember -Identity $groupname -Member $user.username
            } 
            catch {
                #username is already in the group
            }
            #$count = $userarray.Add($_.username)
        }
        
        #REMOVING USERS
        #for each member in the AD group but not in banner, remove them
        #first get the list of usernames in the group
        if ($count -gt 0) {
            $group = get-ADGroup $groupname -properties *
            $group.Members | ForEach-Object { 
                $username = (Get-ADUser $_ -properties SamAccountName).SamAccountName
                if (-not ($userarray.Contains($username))) {
                    "Removing $username"
                    Remove-ADGroupMember -Identity $groupname -Member $username -Confirm:$false
                }
            }
        }
    }
    Catch {

    }

    [GC]::Collect()
}

#Sync Non Employee Group
$nonemployeeQuery = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                    from pebempl, person 
                    where pebempl_pidm = person.pidm
                    and PEBEMPL_ECLS_CODE = 'NE'
                    and PEBEMPL_EMPL_STATUS = 'A'
                    and dead is null
                    and UNIX_EMAIL is not null
                    and not exists (select * from swvlpos where swvlpos_pidm = person.pidm)
                    and pebempl_orgn_code_home not in ('28200','245200')"

Sync-Group $nonemployeeQuery "Nonemployees"


#Sync Retirees 
$retireesQuery = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                    from pebempl, person 
                    where pebempl_pidm = person.pidm
                    and PEBEMPL_ECLS_CODE = 'RE'
                    and PEBEMPL_EMPL_STATUS = 'A'
                    and dead is null
                    and UNIX_EMAIL is not null
                    and not exists (select * from swvlpos where swvlpos_pidm = person.pidm)"
                    
Sync-Group $retireesQuery "Retirees"


#Sync Banner Job Submission Users
$JobSubQuery = "select jobsub_user as username
                    from gwvjsus"
        
Sync-Group $JobSubQuery "L_Banner_Job_Submission_Users"


$CLA_Student_Eligibility = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpop, person
                        where swvlpop_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and (hwwkmedia.f_is_student(swvlpop_pidm) = 'Y')
                        minus
                        select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpop, person
                        where swvlpop_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and (hwwkmedia.f_is_employee(swvlpop_pidm) = 'Y')"

Sync-Group $CLA_Student_Eligibility "CLA_Student_Eligibility"                        
                        
$CLA_Staff_Eligibility = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpoe, person
                        where swvlpoe_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and (hwwkmedia.f_is_employee(swvlpoe_pidm) = 'Y')
                        and (swvlpoe_type <> 'Faculty')"
                        
Sync-Group $CLA_Staff_Eligibility "CLA_Staff_Eligibility"    

$CLA_Faculty_Eligibility = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpoe, person
                        where swvlpoe_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and (hwwkmedia.f_is_employee(swvlpoe_pidm) = 'Y')
                        and swvlpoe_type = 'Faculty'"
                        
Sync-Group $CLA_Faculty_Eligibility "CLA_Faculty_Eligibility"    


###Groups for Students, Employees (Staff and Faculty), Staff, and Faculty.  

$Employees = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpoe, person
                        where swvlpoe_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null"

Sync-Group $Employees "U_Employees"                        

$Staff = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpoe, person
                        where swvlpoe_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and (swvlpoe_type <> 'Faculty')"
                     
Sync-Group $Staff "U_Staff"    

$Faculty = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpoe, person
                        where swvlpoe_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null
                        and swvlpoe_type = 'Faculty'"
                        
Sync-Group $Faculty "U_Faculty"    

$Students = "select distinct replace(UNIX_EMAIL,'@WPI.EDU','') as username
                        from swvlpos, person
                        where swvlpos_pidm = person.pidm
                        and dead is null
                        and UNIX_EMAIL is not null"
                        
Sync-Group $Students "U_Students"    

