workflow Sync-WPICourseGroups {
    param (
    )
    ##Get the Contents of the relevant flat files
    ##Then we want to run down the contents of the courses file and create the course
    ##After creating the course, run down the contents of the enrollment file and get everyone who is supposed to be in the course
    ##diff the current course with the intended population

    #$ExportDirectory = "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi-prod\files"
    $EnrollmentData = Import-Csv "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi-prod\files\StudentEnrollments.csv"
    $CourseData = Import-Csv "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi-prod\files\Section.csv"
    $FacultyEnrollment = Import-Csv "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi-prod\files\TeacherRoster.csv"

    $TopOU = "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"

    ##So, act on each course by creating the group if it doesn't exist, attaching it to the top level, then collecting the users
    foreach -parallel -throttlelimit 25 ($Course in $CourseData) {

        $ClassCode = $Course.'SIS ID'.split("-")[0]
        $ClassCode = $ClassCode.replace("/", "-")
        if ($ClassCode -match "MQP$|IQP$|ISG$|DR$|PHD|THES$") {
            return
        }
    
        $CourseStudentUIDs = $EnrollmentData | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
        $CourseTeacherUIDs = $FacultyEnrollment | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
        $CourseStudent = @()
        $CourseTeachers = @()
    

        if ($Course.'SIS ID'.split("-")[1] -eq "Lab") {
            $TermYear = $Course.'SIS ID'.split("-")[2]
            $Term = $TermYear.substring(0, $TermYear.Length - 2)
            $year = "20" + $TermYear.substring($TermYear.Length - 2)
            $Semester = "01"
            $section = ($Course.'SIS ID'.split("-")[3]).split(" ")[0]
            $DeptCode = $ClassCode -replace '\d+.*', ''
            $ClassNumber = ($ClassCode.TrimStart($DeptCode)).split(" ")[0]
            if ($Term -eq "A") {
                $Semester = "01"
            }
            elseif ($term -eq "B") {
                $Semester = "01"
            }
            elseif ($term -eq "C") {
                $Semester = "02"
            }
            elseif ($term -eq "D") {
                $Semester = "02"
            }
            elseif ($term -eq "E") {
                $Semester = "03"
            }
            elseif ($term -eq "F") {
                $Semester = "03"
            }
            else {
                $Semester = "01"
            }
        }
        else {
            $TermYear = $Course.'SIS ID'.split("-")[1]
            $Term = $TermYear.substring(0, $TermYear.Length - 2)
            $year = "20" + $TermYear.substring($TermYear.Length - 2)
            $Semester = "01"
            $section = ($Course.'SIS ID'.split("-")[2]).split(" ")[0]
            $DeptCode = $ClassCode -replace '\d+.*', ''
            $ClassNumber = ($ClassCode.TrimStart($DeptCode)).split(" ")[0]
            if ($Term -eq "A") {
                $Semester = "01"
            }
            elseif ($term -eq "B") {
                $Semester = "01"
            }
            elseif ($term -eq "C") {
                $Semester = "02"
            }
            elseif ($term -eq "D") {
                $Semester = "02"
            }
            elseif ($term -eq "E") {
                $Semester = "03"
            }
            elseif ($term -eq "F") {
                $Semester = "03"
            }
            else {
                $Semester = "01"
            }
        }

        $OU = "OU=Term_" + $year + "_" + $Term + "," + $TopOU
        $OU_Name = "Term_" + $year + "_" + $Term
    
        #Verify the OU already exists for the term, else create

        if ((Get-ADOrganizationalUnit -filter 'Name -like $OU_Name').Name -ne "$OU_Name") {
            New-ADOrganizationalUnit -path $TopOU -name $OU_Name
        }

        #Compose Names based on the existing naming scheme
        [string]$GroupName = $DeptCode + "_" + $ClassNumber + "_" + $section + "_" + $year + $Semester + "_" + $Term
        $TopGroupName = $DeptCode + "_" + $ClassNumber
        [string]$FacultyGroupName = $DeptCode + "_" + $ClassNumber + "_Faculty"

        #Verify Course Group Exists
        try {
            Get-ADGroup $TopGroupName
        }#Else Create
        catch {
            New-ADGroup $TopGroupName -Path $TopOU -GroupScope Universal -GroupCategory Security
        }

        #Verify Faculty Group Exists
        try {
            Get-ADGroup $FacultyGroupName
        }#Else Create and Nest
        catch {
            New-ADGroup $FacultyGroupName -Path $TopOU -GroupScope Universal -GroupCategory Security
            $timer = [diagnostics.stopwatch]::startnew()
            while ($timer.elapsed.totalseconds -lt 20) {
                if (Get-ADGroup -filter 'samaccountname -eq $GroupName') {
                    Add-ADPrincipalGroupMembership -Identity $FacultyGroupName -MemberOf $TopGroupName
                }
                else {
                    start-sleep -seconds 5
                }
            }
        }

        #Verify Section Group Exists
        try {
            Get-ADGroup $GroupName
        }#Else Create and Nest
        catch {
            New-ADGroup -Name $GroupName -Path $OU -GroupScope Universal -GroupCategory Security
            $timer = [diagnostics.stopwatch]::startnew()
            while ($timer.elapsed.totalseconds -lt 20) {
                if (Get-ADGroup -filter 'samaccountname -eq $GroupName') {
                    Add-ADPrincipalGroupMembership -Identity $GroupName -MemberOf $TopGroupName
                }
                else {
                    start-sleep -seconds 5
                }
            }
        }

        foreach ($UID in $CourseStudentUIDs) {
            $CourseStudent += [PSCustomObject]@{'samaccountname' = (get-aduser -Filter "uidNumber -like $($UID.'SIS ID')").samaccountname }
        }
        foreach ($UID in $CourseTeacherUIDs) {
            $CourseTeachers += [PSCustomObject]@{'samaccountname' = (get-aduser -Filter "uidNumber -like $($UID.'SIS ID')").samaccountname }
        }

        #sync population

        $comparisons = $null
        $AddMembers = $null
        $RemoveMembers = $null
        #Faculty Group Sync
        ##$currentFaculty = Get-ADGroupMember $FacultyGroupName | Select-Object samaccountname
        ##if ($null -eq $currentFaculty) {
        $AddMembers = $CourseTeachers
        ##}
        ##else {
        ##    $comparisons = Compare-Object $currentFaculty $CourseTeachers -Property samaccountname
        ##    $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object samaccountname
        ##    $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object samaccountname
        ##}
        ##foreach ($removal in $RemoveMembers.samaccountname) {
        ##    Remove-ADGroupMember -Identity $FacultyGroupName -Members $removal -Confirm:$false
        ##}
        foreach ($addition in $AddMembers.samaccountname) {
            Add-ADGroupMember -Identity $FacultyGroupName -Members $addition -Confirm:$false
        }

        $comparisons = $null
        $AddMembers = $null
        $RemoveMembers = $null
        #Section Group Sync
        $currentSection = Get-ADGroupMember $GroupName | Select-Object samaccountname
        if ($null -eq $currentSection) {
            $AddMembers = $CourseStudent
        }
        else {
            $comparisons = Compare-Object $currentSection $CourseStudent -Property samaccountname
            $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object samaccountname
            $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object samaccountname
        }
        foreach ($removal in $RemoveMembers.samaccountname) {
            Remove-ADGroupMember -Identity $GroupName -Members $removal -Confirm:$false
        }
        foreach ($addition in $AddMembers.samaccountname) {
            Add-ADGroupMember -Identity $GroupName -Members $addition -Confirm:$false
        }

    }
    
}

Measure-Command {
    Sync-WPICourseGroups
}