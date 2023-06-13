Import-Module "D:\wpi\powershell\CourseShareGeneration\build-course-shares.ps1"
#I need to pull from OU OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu

function Get-CurrentYearCourses {
    [CmdletBinding()]
    param (
        
    )
    
    $TopOU = "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"

    # We need to be able to determine what year this is

    $Date = get-date
    $currentYear = $date.Year

    $termOUs = Get-ADOrganizationalUnit -filter "name -like 'Term_$($currentyear)_*'" -SearchBase $TopOU
    #$termOU[0].Name = CS_3013_C02_202202_C

    $allCourses = $null

    foreach ($term in $termOUs) {
        #Filter on numbers in section ($AtermSections|where {$_.name.split("_")[2] -match ".*\d+"}).count
        # Alternatively add | where-object {$null -ne (get-adgroupmember $_.name)} , but this destroys runtime
        $TermCourseGroups = Get-ADGroup -filter * -SearchBase $term.distinguishedname | Select-Object @{ Name = 'Course'; Expression = { "$($_.Name.split("_")[0])_$($_.Name.split("_")[1])_$($_.Name.split("_")[4])$($_.Name.split("_")[3].substring(2,2))" } } | Sort-Object -Property Course -Unique
        $allCourses += $TermCourseGroups
    }
    return $allCourses
}

function Get-CurrentTerms {
    [CmdletBinding()]
    param (
        $date
    )

    $terms = New-Object System.Collections.Generic.List[string]

    # A term is between date:
    $AtermStart = "August 15"
    $AtermEnd = "October 20"
    if ($date -lt $AtermEnd -and $date -gt $AtermStart) {
        $terms.Add("A")
    }
    # B term is between date:
    $BtermStart = "October 15"
    $BtermEnd = "December 20"
    if ($date -lt $BtermEnd -and $date -gt $BtermStart) {
        $terms.Add("B")
    }
    # C term is between date:
    $CtermStart = "January 5"
    $CtermEnd = "March 10"
    if ($date -lt $CtermEnd -and $date -gt $CtermStart) {
        $terms.Add("C")
    }
    # D term is between date:
    $DtermStart = "March 5"
    $DtermEnd = "May 10"
    if ($date -lt $DtermEnd -and $date -gt $DtermStart) {
        $terms.Add("D")
    }
    # F term is between date:
    $EtermStart = "May 15"
    $EtermEnd = "August 25"
    if ($date -lt $EtermEnd -and $date -gt $EtermStart) {
        $terms.Add("F")
    }
    # S term is between date:
    $StermStart = "May 15"
    $StermEnd = "August 25"
    if ($date -lt $StermEnd -and $date -gt $StermStart) {
        $terms.Add("S")
    }
        
    return $terms

}

function Get-TermCourses {
    [CmdletBinding()]
    param (
        $terms,
        $year
    )
    
    $TopOU = "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"

    # We need to be able to determine what year this is

    $Date = get-date
    if ($null -eq $year) {
        $currentYear = $date.Year
    }
    if ($null -eq $terms) {
        $terms = get-currentTerms $Date
    }

    $termOUs = New-Object System.Collections.Generic.List[Microsoft.ActiveDirectory.Management.ADObject]

    foreach ($term in $terms) {
        $termOUs += Get-ADOrganizationalUnit -filter "name -like 'Term_$($currentyear)_$($term)'" -SearchBase $TopOU
        #$termOU[0].Name = CS_3013_C02_202202_C
    }

    $allCourses = $null

    foreach ($term in $termOUs) {
        #Filter on numbers in section ($AtermSections|where {$_.name.split("_")[2] -match ".*\d+"}).count
        # Alternatively add | where-object {$null -ne (get-adgroupmember $_.name)} , but this destroys runtime
        $TermCourseGroups = Get-ADGroup -filter * -SearchBase $term.distinguishedname | Select-Object @{ Name = 'Course'; Expression = { "$($_.Name.split("_")[0])_$($_.Name.split("_")[1])_$($_.Name.split("_")[4])$($_.Name.split("_")[3].substring(2,2))" } } | Sort-Object -Property Course -Unique
        $allCourses += $TermCourseGroups
    }
    return $allCourses
}


function Set-CourseGroups {
    [CmdletBinding()]
    param (
        $courselist
    )
    
    foreach ($course in $courselist.course) {
        $ClassCode = "$($course.split("_")[0])_$($course.split("_")[1])"
        $Term = $course.split("_")[2]
        $Folder = $course.split("_")[0]

        Try {
            set-wpiacLs-NoClassVision $term $ClassCode $Folder
        }
        Catch {
            Write-host "Something went wrong with $course"
        }
    }
}

function Remove-OldCourseGroups {
    [CmdletBinding()]
    param (
        [switch]$whatif
    )

    $Path = "\\storage-02\academics\courses"

    $DeptList = Get-ChildItem $path | Select-Object Name

    foreach ($Dept in $Deptlist) {
        $CourseList = get-childitem "$($path)/$($dept.Name)"
        foreach ($Course in $CourseList) {
            $OldTerms = get-childitem "$($path)\$($dept.Name)\$($course.Name)" | where-object { (Get-Date).addyears("-2") -gt $_.LastWriteTime }
            foreach ($Folder in $OldTerms) {
                if (!($whatif)) {
                    try {
                        Remove-Item "$($path)\$($dept.Name)\$($course.Name)\$($folder.name)" -Recurse
                    }
                    catch {
                        Write-Host "Error Attempting to Delete: $($path)\$($dept.Name)\$($course.Name)\$($folder.name)"
                    }
                }
                else {
                    Write-Host "Would Delete: $($path)\$($dept.Name)\$($course.Name)\$($folder.name)"
                }
            }
        }
    } 
}