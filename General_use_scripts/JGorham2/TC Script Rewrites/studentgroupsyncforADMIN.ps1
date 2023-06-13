#Install-Module InvokeQuery
Import-Module ActiveDirectory
Import-Module InvokeQuery

#Note:  This does 1 course section
#Need to be in the courses OU first

function Invoke-WPIBannerQuery {
    param (
        $Query
    )
    $BannerConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=bannerprod.wpi.edu)(PORT=1527)))(CONNECT_DATA=(SERVICE_NAME=prod.admin.wpi.edu)(server=dedicated)));User ID=locksys;Password=ihatecbord"
    Invoke-OracleQuery -Sql $Query -ConnectionString $BannerConnectionString
}

function Sync-CourseGroup($bnrterm, $bnrpartofterm, $subjectcode, $coursenumber, $sectionnumber) {
    ########Banner Connection info######################

    <# #encrypted oracle connection string
    $BannerConnectionString = Get-Content $connfile 

    $connection = New-Object System.Data.OracleClient.OracleConnection($BannerConnectionString)
    [System.Data.OracleClient.OracleCommand] $command = new-Object System.Data.OracleClient.OracleCommand
    $command.Connection = $connection
    $connection.Open()
    #################################################### #>

    #create term OU if it doesn't exist
    Set-Location $ADCourseOUPath
    $BannerOU = "Banner_Term_" + $bnrterm + "_" + $bnrpartofterm 
    if (-not(Test-Path "OU=$BannerOU")) {New-ADOrganizationalUnit $BannerOU}
    Set-Location OU=$BannerOU

    #get list of group members from banner
    $BannerQuery = "Select distinct replace(wpi_email(pidm,'UNIX'),'@WPI.EDU','') as username, IND as is_instructor from courses where subj = '$subjectcode' AND crse = '$coursenumber' and sect = '$sectionnumber' and etrm = '$bnrterm' and ptrm = '$bnrpartofterm'"
    $data = Invoke-WPIBannerQuery $BannerQuery

    $groupname = "" + $subjectcode + "_" + $coursenumber + "_" + $sectionnumber + "_" + $bnrterm + "_" + $bnrpartofterm
    $instructorgroupname = "" + $subjectcode + "_" + $coursenumber + "_" + "Faculty"
    $coursegroup = "" + $subjectcode + "_" + $coursenumber
 
    $students = $data
  
    #if someone is taking the course do { 
    if ($students.Rows.Count -gt 0) {
 
        #if course section group doesn't exist, create it
        Try {Get-ADGroup $groupname}
        Catch {New-ADGroup -Name $groupname -GroupScope Universal -GroupCategory Security }
    
        Set-Location ..
    
        #if course instructor group doesn't exist, create it
        Try {Get-ADGroup $instructorgroupname}
        Catch {New-ADGroup -Name $instructorgroupname -GroupScope Universal -GroupCategory Security}
    
        #if non-term based course group doesn't exist, create it
        Try {Get-ADGroup $coursegroup}
        Catch {
            New-ADGroup -Name $coursegroup -GroupScope Universal -GroupCategory Security
            Add-ADGroupMember -Identity $coursegroup -Member $instructorgroupname
        }
        try {
            Add-ADGroupMember -Identity $coursegroup -Member $groupname
        }
        catch {
            #group is already a member 
        }
    
    
        Set-Location OU=$BannerOU
       
        $Bannerstudentarray = New-Object System.Collections.ArrayList
        #ADDING STUDENTS
        $students | ForEach-Object {
            #for each member in banner and not in group, add them
            try {
                Add-ADGroupMember -Identity $groupname -Member $_.username
            } 
            catch {
                #username is already in the group
            }
            $Bannerstudentarray.Add($_.username)
        }
        
        #REMOVING STUDENTS
        #for each member in group but not in banner, remove them
        #first get the list of usernames in the group
        $group = get-ADGroup $groupname -properties *
        $group.Members | ForEach-Object { 
            $username = (Get-ADUser $_ -properties SamAccountName).SamAccountName
            if (-not ($Bannerstudentarray.Contains($username))) {
                Remove-ADGroupMember -Identity $groupname -Member $username -Confirm:$false
            }
        }
        
    }# end if someone is taking the course.

    [GC]::Collect()
}
#Note:  This feeds all the course sections for the given term to Sync-CourseGroup
function Sync-TermGroups($bannerterm, $bannerpartofterm) {
    <# ########Banner Connection info######################

    #encrypted oracle connection string
    $BannerConnectionString = Get-Content $connfile 

    $connection = New-Object System.Data.OracleClient.OracleConnection($BannerConnectionString)
    [System.Data.OracleClient.OracleCommand] $command = new-Object System.Data.OracleClient.OracleCommand
    $command.Connection = $connection
    $connection.Open()

    ################################################# #>
    #get course section list
    $BannerQuery = "Select distinct subj, crse, sect from section where term = '$bannerterm' and ptrm = '$bannerpartofterm' and schd <> 'P'"
  
    $data = Invoke-WPIBannerQuery $BannerQuery

    $courses = $data

    ###for each course### 
    foreach($course in $courses){
        Sync-CourseGroup $bannerterm $bannerpartofterm $course.SUBJ $course.CRSE $course.SECT
    }

    [GC]::Collect()
}

function Start-AutoSyncTerms() {
    <# ########Banner Connection info######################

    #encrypted oracle connection string
    $BannerConnectionString = Get-Content $connfile 

    $connection = New-Object System.Data.OracleClient.OracleConnection($BannerConnectionString)
    [System.Data.OracleClient.OracleCommand] $command = new-Object System.Data.OracleClient.OracleCommand
    $command.Connection = $connection
    $connection.Open()
    #################################################### #>

    #get course section list
    $BannerQuery = "Select distinct term, ptrm from section where term >= date_to_term(sysdate) order by term, ptrm"
  
    $data = Invoke-WPIBannerQuery $BannerQuery

    $terms = $data

    ###for each course### 
    foreach($term in $terms) { 
        #make sure the groups exist
        Sync-TermGroups $term.term $term.ptrm
    
    }

    [GC]::Collect()
}

##################################################
#       Main Process                             #
##################################################

#$connfile = "C:\wpi\batch\ADGroups\locksys.txt"

If (test-Path Student) {
    Remove-PSDrive Student
}
else {}
New-PSDrive -Name Student -PSProvider ActiveDirectory -Root "" -server admin.wpi.edu:389

#Student:\OU=Courses,OU=Banner Groups,DC=student   

##############CONTROL VARIABLES####################
$ADCourseOUPath = "Student:\OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"
#$bannerterm = "201101" #Read-Host "Banner Term (ie 201001):"
#$bannerpartofterm = "1" #Read-Host "Banner Part of Term (ie A):"

Start-AutoSyncTerms

Remove-PSDrive Student