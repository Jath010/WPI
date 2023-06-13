##Get the Contents of the relevant flat files
##Then we want to run down the contents of the courses file and create the course
##After creating the course, run down the contents of the enrollment file and get everyone who is supposed to be in the course
##diff the current course with the intended population

######
# Functions to get data from Workday
######
function Get-WorkdayTerms {
    [CmdletBinding()]
    param (
        [PSCredential]
        [Parameter(Mandatory = $true)]
        $Cred
    )

    $Call = @{
        Uri = "https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_Changes_Terms?format=json"
        #wtth data from course
        # Uri = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212-_AD_Instructors__Course_Section_Definitions_?Effective_Date=2021-07-12-07:00&Academic_Period!Academic_Period_ID=' + $codes + '&format=json'
        # with data from section
        #URI = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_group_changes_Student_Enrollments__student_course_registration_records_?Starting_Academic_Period!Academic_Period_ID=' + $codes + '&Course_Section!Section_Listing_ID=' + $section +'&format=json'
        #Authentication = "Basic"
        Credential = $cred
    }
    (Invoke-RestMethod @Call).Report_Entry #| fl

}

function Get-WorkdayTermContents {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("A", "B", "C", "D", "Summer Session I", "Summer Session II")]
        $term,
        [Parameter(Mandatory = $true)]
        [String]
        $Year,
        [Parameter(Mandatory = $true)]
        $WorkdayTermData,
        [PSCredential]
        [Parameter(Mandatory = $true)]
        $Cred
    )

    $termReferenceID = ($WorkdayTermData | where-object { $_.Academic_Period -match "$year.* $Term .*" } | select-object referenceID).referenceID

    $Call2 = @{
        Uri = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212-_AD_Instructors__Course_Section_Definitions_?Effective_Date=2021-07-12-07:00&Academic_Period!Academic_Period_ID=' + $termReferenceID + '&format=json'
        Credential = $cred
    }
    (Invoke-RestMethod @Call2).Report_Entry
}

##
# Import Credentials
##
$Directory = "D:\tmp\Bifrost copy\CourseGroups"
$credentials = $null
$credPath = "$Directory\ISU_Active_Directory.xml"

if (Test-path $credPath) {
    $credentials = Import-CliXml -Path $credPath
        
}

##
# Get the List of Available Terms
##

$WorkdayTermData = Get-WorkdayTerms $credentials

$year = (get-date).year
















$ExportDirectory = "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi2\files"
$EnrollmentData = Import-Csv $ExportDirectory"\StudentEnrollments.csv"
$CourseData = Import-Csv $ExportDirectory"\Section.csv"
$FacultyEnrollment = Import-Csv $ExportDirectory"\TeacherRoster.csv"

$TopOU = "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"

##So, act on each course by creating the group if it doesn't exist, attaching it to the top level, then collecting the users
foreach ($Course in $CourseData) {
    $CourseStudentUIDs = $EnrollmentData | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
    $CourseTeacherUIDs = $FacultyEnrollment | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
    $CourseStudent = @()
    $CourseTeachers = @()
    $ClassCode = $Course.'SIS ID'.split("-")[0]
    $TermYear = $Course.'SIS ID'.split("-")[1]
    $Term = $TermYear.substring(0, $TermYear.Length - 2)
    $year = "20" + $TermYear.substring($TermYear.Length - 2)
    $Semester = "01"
    $section = $Course.'SIS ID'.split("-")[2]
    $DeptCode = $ClassCode -replace '\d+.*', ''
    $ClassNumber = $ClassCode.TrimStart($DeptCode)
    switch ($Term) {
        A { $Semester = "01" }
        B { $Semester = "01" }
        C { $Semester = "02" }
        D { $Semester = "02" }
        E { $Semester = "03" }
        Default { $Semester = "01" }
    }

    $OU = "OU=" + $year + "_" + $Term + "," + $TopOU
    
    #Verify the OU already exists for the term, else create
    try {
        Get-ADOrganizationalUnit $OU
    }
    catch {
        New-ADOrganizationalUnit $OU
    }
    #Compose Names based on the existing naming scheme
    $GroupName = $DeptCode + "_" + $ClassNumber + "_" + $section + "_" + $year + $Semester + "_" + $Term
    $TopGroupName = $DeptCode + "_" + $ClassNumber
    $FacultyGroupName = $DeptCode + "_" + $ClassNumber + "_" + $section + "_" + $year + $Semester + "_" + $Term

    #Verify Course Group Exists
    try {
        Get-ADGroup $TopGroupName
    }#Else Create
    catch {
        New-ADGroup $TopGroupName -Path $TopOU
    }

    #Verify Faculty Group Exists
    try {
        Get-ADGroup $FacultyGroupName
    }#Else Create and Nest
    catch {
        New-ADGroup $FacultyGroupName -Path $TopOU
        Add-ADPrincipalGroupMembership -Identity $FacultyGroupName -MemberOf $TopGroupName
    }

    #Verify SEction Group Exists
    try {
        Get-ADGroup $GroupName
    }#Else Create and Nest
    catch {
        New-ADGroup $GroupName -Path $OU
        Add-ADPrincipalGroupMembership -Identity $GroupName -MemberOf $TopGroupName
    }

    foreach ($UID in $CourseStudentUIDs) {
        $CourseStudent += [PSCustomObject]@{'samaccountname' = (get-aduser -Filter "uidNumber -like $($UID.'SIS ID')").samaccountname }
    }
    foreach ($UID in $CourseTeacherUIDs) {
        $CourseTeachers += [PSCustomObject]@{'samaccountname' = (get-aduser -Filter "uidNumber -like $($UID.'SIS ID')").samaccountname }
    }

    #sync population

    #Faculty Group Sync
    $currentFaculty = Get-ADGroupMember $FacultyGroupName | Select-Object samaccountname
    $comparisons = Compare-Object $currentFaculty $CourseTeachers -Property samaccountname
    $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object samaccountname
    $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object samaccountname
    foreach ($removal in $RemoveMembers.samaccountname) {
        Remove-ADGroupMember -Identity $FacultyGroupName -Members $removal
    }
    foreach ($addition in $AddMembers.samaccountname) {
        Add-ADGroupMember -Identity $FacultyGroupName -Members $addition
    }

    #Section Group Sync
    $currentSection = Get-ADGroupMember $GroupName | Select-Object samaccountname
    $comparisons = Compare-Object $currentSection $CourseStudent -Property samaccountname
    $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object samaccountname
    $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object samaccountname
    foreach ($removal in $RemoveMembers.samaccountname) {
        Remove-ADGroupMember -Identity $GroupName -Members $removal
    }
    foreach ($addition in $AddMembers.samaccountname) {
        Add-ADGroupMember -Identity $GroupName -Members $addition
    }

}