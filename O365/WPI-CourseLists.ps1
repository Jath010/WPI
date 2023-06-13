#Get List of students in course based on year and term

#Input Scheme
#Get-WPICourseStudentList IMGD_2101 A19
function Get-WPICourseStudentList {
    param (
        $CourseID,
        $Term
    )

    # build student list, ensuring the proper term is selected
    $CourseSectionGroup = $CourseID+"_"+($Term.SubString(0,1))
    $year = $Term.SubString(1,2)

    # So the banner information gets tossed into groups with the name scheme like IMGD_2101_A01_201701_A

    $StudentGroupList = Get-ADGroupMember $CourseID  | Where-Object {($_.Name -Like "$CourseSectionGroup*") -and ($_.Name.Contains("20" + $year))} | Select-Object Name
    foreach ($group in $StudentGroupList) {
        $ContactList += Get-ADGroupMember -Recursive $group.Name
    }
    foreach($user in $ContactList) {
        $user.SamAccountName
    }
}

function Get-WPICourseFacultyList {
    param (
        $CourseID,

        [switch]
        $csv
    )
    $FacultyGroup = $CourseID+"_Faculty"
    $FacultyList = Get-ADGroupMember $FacultyGroup

    foreach($user in $FacultyList){
        $user.SamaccountName
    }
}