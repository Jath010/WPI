#Needs PnP.Powershell Module

function New-StudentFolder {
    [CmdletBinding()]
    param (
        $Student,
        $teacherList,
        $Connection
    )
    if ($null -eq (Get-PnPFolder -Url "Shared Documents/Course Storage/submissions/${student}" -Connection $Connection -ErrorAction Ignore)) {
        Add-PnPFolder -Name $Student -Folder "Shared Documents/Course Storage/submissions" -Connection $Connection
        Set-PnPFolderPermission -List "Shared Documents" -Identity "Shared Documents/Course Storage/submissions/${student}" -User "$Student@wpi.edu" -AddRole "Contribute" -Connection $Connection -ClearExisting
        foreach ($Teacher in $teacherList) {
            Set-PnPFolderPermission -List "Shared Documents" -Identity "Shared Documents/Course Storage/submissions/${student}" -User "$Teacher@wpi.edu" -AddRole "Full Control" -Connection $Connection
        }
    }
}


function Set-WPITeamClassStorage {
    [CmdletBinding()]
    param (
        $TeamName,
        $Term,
        $CourseID,
        $Dept
    )
    
    ###Connection
    ##$Team = Get-Team -DisplayName $TeamName
    ##$SharepointURL = (Get-UnifiedGroup -Identity (Get-Team -DisplayName $TeamName).GroupId).SharePointSiteUrl
    $connection = Connect-PnPOnline -Url ((Get-UnifiedGroup -Identity (Get-Team -DisplayName $TeamName).GroupId).SharePointSiteUrl) -Interactive
    ###

    ###
    # Gather Class Data
    ###

    #Getting the shifted course name might be pointless
    #$course = $CourseID.Replace("_", "") #Convert CourseID from XXXX_#### to XXXX#### for the folder path
    #Build Teacher List
    $TeacherList = Get-ADGroupMember "${CourseID}_faculty" | Select-Object SamAccountName
    Write-Host "$CourseID $Term has the following faculty: $($TeacherList.SamAccountName)"

    #Build Student List
    # build student list, ensuring the proper term is selected
    $CourseSectionGroup = $CourseID + "_" + ($Term.SubString(0, 1))
    $year = $Term.SubString(1, 2)

    $StudentGroupList = Get-ADGroupMember $CourseID  | Where-Object { $_.Name -Like "$CourseSectionGroup*_20$Year*" } | Select-Object Name
    foreach ($group in $StudentGroupList) {
        #write-host "checking $($group.Name)"
        if ($null -ne (get-adgroup -LDAPFilter "(Name=$($group.Name))" -SearchBase "OU=Term_20$($year)_$($Term.SubString(0, 1)),OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu")) {
            #Write-Host "group $($group.name) found"
            [array]$StudentList += Get-ADGroupMember -Recursive $group.Name | Select-Object SamAccountName
        }
        #[array]$StudentList += Get-ADGroupMember -Recursive $group.Name
    }
    Write-Host "$CourseID $Term has the following students: $($StudentList.SamAccountName)"

    ###
    #  Create the folders
    ###

    Add-PnPFolder -Name "Course Storage" -Folder "Shared Documents"
    Add-PnPFolder -Name "course-resources" -Folder "Shared Documents/Course Storage"
    Add-PnPFolder -Name "submissions" -Folder "Shared Documents/Course Storage"

    foreach ($Student in $StudentList) {
        New-StudentFolder $Student -teacherList $TeacherList -Connection $connection
    }
}