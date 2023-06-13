#####################################################################################
# Created by:  Tom Collins, Jack O'Brien 10/16/2012                                 #
# Modified by:  Jack O'Brien 10/25/2012 (added Set-WPIACLs-SingleUser)              #
# Modified by:  Jack O'Brien 12/20/2012 (added get-course-groups and Set-DeptACLs)  #
# Modified by:  Jack O'Brien 3/13/2015 (final version includes submissions and      #
#               course resource folders at the request of IMGD faculty, updated     #
#               readme)                                                             #
# Modified by: Andrew Stone 9/22/2016 (added buildinng folder structures for teams  #
#               from a csv file                                                     #
# Modified By: Joshua Gorham 10/25/2018 (Added Section for NoClassVision and cleaned#
#               up some of the warnings for thigslike improper names)               #
# Modified By: Joshua Gorham 03/06/2019 (Fixed the bit where it would have matched  #
#               every section for some classes in 2020                              #
#####################################################################################

# README:  By default, the folder will be created under \\storage.wpi.edu\academics\courses\DEPT\COURSENAME\TERM
# 
# Functions # 
#
# Set-WPIACLs ($Term <String>, $CourseID <AD Group>, $dept <department short name>)
# ex. Set-WPIACLs B12 IMGD_1000 IMGD
#
# The Set-WPIACLs-Single-User function doesn't seem that useful anymore, but I'll leave it in.  Typically I'll just run the main function (Set-WPIACLs) again and it'll do what it needs to
# 
# Run this after runninng Set-WPIACLs for a particular course
# New-Team-folderstructure "\\storage.wpi.edu\academics\courses\RBE\RBE2001\A17" '.\RBE 2001 A16 Team Assignments.csv' "RBE_2001_Faculty"
#
# End of Functions #
# 
#
# For testing:  You can change the $rootPath variable at the top of the code to point locally if you want to do a test run before hitting the Isilon.  Probably not a bad idea.
#
#
# KNOWN ISSUE:  When running this in A/B term, go one year ahead (e.g. for a course in A term of 2015, specify "A16" as the term in the function call). Then go into the share
# in your GUI of choice and manually change the name of the folder to the appropriate term ("A15" in our example).  This is due to the logic I use to determine the correct
# AD course section group.  I am too lazy to correct this, renaming the folder is sufficient.  You may not be so lazy, but trust me it's annoying to fix and not worth the
# effort.
# EDIT: Messed with this, you still enter the FY instead of the current for A and B Terms, but no longer need to rename the folder
# EDIT: Strike that, since the swap from banner to workday they don't change the year, so it's just w/e we're in.
#
# NOTE ABOUT HU 3900 COURSES:  Replace "$CourseSectionGroup*" in line 161 with "AD_Group_For_The_Section_You_Need*" to pick the right group:
# 
# This is the line:  $StudentGroupList = Get-ADGroupMember $CourseID  | Where {($_.Name -Like "$CourseSectionGroup*") -and ($_.Name.Contains($year))} | Select Name
# 
# Don't forget the asterisk, and don't forget to change it back to $CourseSectionGroup* when you're done.  If there is more than one HU3900
# seminar course in a particular term, you may need to edit the $CoursePath = "$rootPath\$course\$term\submissions" on line 97 to look like:
# $CoursePath = "$rootPath\$course\$term\SECTION\submissions".  Of course you should remember to change this back too.
#
#
# NOTE ABOUT CROSS-LISTED COURSES:  Pick which one will be the "primary" (e.g. for AR3201 and IMGD3201, the path is \\storage\academics\courses\IMGD\AR3201).
# Which is the primary is completely up to you.  Once you've selected the primary, pick the AD group for thecorresponding SECTION (not the course, just the section) 
# and insert it into the primary AD groups SECTION.  For example, for AR3201, I would place the group IMGD_3201_D01_201502_D inside of AR_3201_D01_201502_D.  The 
# code is recurisve and will find all the correct members and you'll be good to go.  This can be confusing if the same cross listed course is being taught by two 
# different professors.  Use your judgement and assign one professor to the AR3201 course, and one to the IMGD3201 course.
#
#

############################
# declare variables

$course = $null
$StudentGrouplist = $null
$CourseSectionGroup = $null
$FacultyList = $null
$StudentList = $null
$group = $null
$student = $null
    
#Set path information
# $rootPath = "\\wheeljack.wpi.edu\C$\testing"
$rootPath = "\\storage-02.wpi.edu\academics\courses"

#end of variable declaration
############################


Import-Module ActiveDirectory


# this will find all banner course groups for a particular academic department
function get-course-groups($dept) {
    $trimmedList = @()
    #build LDAPFilter
    $deptUnderScore = $dept + "_"
    #get list of course groups in academic dept
    $fullList = Get-ADGroup -SearchBase "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu" -LDAPFilter "(name=$deptUnderScore*)" | Select-Object -Expand Name
    
    #trim course information from each group, leaving only the top level course groups
    foreach ($group in $fullList) {
        $groupSplit = $group.Split("_")
        $groupName = $groupSplit.Get(0) + "_" + $groupSplit.Get(1)
        $trimmedList += $groupName
    }
    
    #remove duplicates from the list
    $trimmedList = $trimmedList | Select-Object -Unique

    return $trimmedList
}

# this will configure a dropbox style setup in a single folder for all sections of a course in a particular term
function Set-WPIACLs($Term, $CourseID, $dept) {
    $rootPath = "$rootPath\$dept"
    
    #Change Term number for A and B Terms
    if ($Term.StartsWith("B") -or $Term.StartsWith("A")) {
        $TermFolder = $Term[0] + ($Term.Substring(1, 2) - 1)
    }
    else {
        $TermFolder = $Term[0] + $Term.Substring(1, 2)
    }

    #build course path
    $course = $CourseID.Replace("_", "") #Convert CourseID from XXXX_#### to XXXX#### for the folder path
    $CoursePath = "$rootPath\$course\$TermFolder\submissions"

    # build faculty list
    $FacultyGroup = $CourseID + "_Faculty"
    $FacultyList = Get-ADGroupMember $FacultyGroup
    Write-Host "$CourseID $Term has the following faculty: $($FacultyList.SamAccountName)" #| Out-File c:\$CourseID.log
    "$CourseID $Term has the following students: $($facultyList.SamAccountName)" #| Out-File c:\$CourseID.log

    # create acls for submissions folder
    if (!(Test-Path $CoursePath)) {
        ###############
        #Create Folder#
        ###############
        [VOID](New-Item $CoursePath -type directory)

        ############
        #create acl#
        ############
        
        #new blank acl
        $acl = New-Object System.Security.AccessControl.DirectorySecurity

        #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
        $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
        $objType = [System.Security.AccessControl.AccessControlType]::Allow 

        #hosting team permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #faculty permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #course permissions on the parent folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #remove inheritance
        $acl.SetAccessRuleProtection($true, $true)

        ###########
        #apply acl#
        ###########

        Write-Host "Applying ACL for $CoursePath..."
        "Applying ACL for $CoursePath..." #| Out-File c:\$CourseID.log -Append
        
        try {$acl | Set-Acl $CoursePath}
        catch {
            $output = "" + $username + "- " + $error[0]
            Write-Host $output
            $output #| Out-File c:\$CourseID.log -Append
        }
        Write-Host "ACL applied."
        "ACL applied." #| Out-File c:\$CourseID.log -Append
    }

    
    
    # build student list, ensuring the proper term is selected
    $CourseSectionGroup = $CourseID + "_" + ($Term.SubString(0, 1))
    $year = $Term.SubString(1, 2)

    $StudentGroupList = Get-ADGroupMember $CourseID  | Where-Object {$_.Name -Like "$CourseSectionGroup*_20$Year*"} | Select-Object Name
    foreach ($group in $StudentGroupList) {
        $StudentList += Get-ADGroupMember -Recursive $group.Name
    }
    Write-Host "$CourseID $Term has the following students: $($StudentList.SamAccountName)"
    "$CourseID $Term has the following students: $($StudentList.SamAccountName)" #| Out-File c:\$CourseID.log

    # create student folders and permissions
    foreach ($student in $StudentList) {
        $username = $student.SamAccountName

        #create folder if not exist
        if (!(Test-Path $CoursePath\$username)) {
            ###############
            #Create Folder#
            ###############
            [VOID](New-Item $CoursePath\$username -type directory)

            ############
            #create acl#
            ############
        
            #new blank acl
            $acl = New-Object System.Security.AccessControl.DirectorySecurity

            #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
            $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
            $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
            $objType = [System.Security.AccessControl.AccessControlType]::Allow 

            #hosting team permissions on the folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            $acl.AddAccessRule($objACE) 

            #student permissions on the student folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE)

            #faculty permissions on the student folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE) 

            #course permissions on the student folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE)

            #remove inheritance
            $acl.SetAccessRuleProtection($true, $true)

            ###########
            #apply acl#
            ###########

            Write-Host "Applying ACL for $CoursePath\$username..."
            "Applying ACL for $CoursePath\$username..." #| Out-File c:\$CourseID.log -Append
        
            try {$acl | Set-Acl $CoursePath\$username}
            catch {
                $output = "" + $username + "- " + $error[0]
                Write-Host $output
                $output #| Out-File c:\$CourseID.log -Append
            }
            Write-Host "ACL applied."
            "ACL applied." #| Out-File c:\$CourseID.log -Append
        }
    
        # create faculty folders and permissions
        foreach ($faculty in $FacultyList) {
            $username = $faculty.SamAccountName

            #create folder if not exist
            if (!(Test-Path $CoursePath\$username)) {
                ###############
                #Create Folder#
                ###############
                [VOID](New-Item $CoursePath\$username -type directory)

                ############
                #create acl#
                ############
        
                #new blank acl
                $acl = New-Object System.Security.AccessControl.DirectorySecurity

                #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
                $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
                $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
                $objType = [System.Security.AccessControl.AccessControlType]::Allow 

                #hsoting team permissions on the folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                $acl.AddAccessRule($objACE) 

                #individual faculty permissions on the faculty folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                #add to acl
                $acl.AddAccessRule($objACE)

                #all faculty permissions on the faculty folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                #add to acl
                $acl.AddAccessRule($objACE) 

                #remove inheritance
                $acl.SetAccessRuleProtection($true, $true)

                ###########
                #apply acl#
                ###########

                Write-Host "Applying ACL for $CoursePath\$username..."
                "Applying ACL for $CoursePath\$username..." #| Out-File c:\$CourseID.log -Append
        
                try {$acl | Set-Acl $CoursePath\$username}
                catch {
                    $output = "" + $username + "- " + $error[0]
                    Write-Host $output
                    $output | Out-File c:\$CourseID.log -Append
                }
                Write-Host "ACL applied."
                "ACL applied." | Out-File c:\$CourseID.log -Append
            }
        }
    }


    # create additional folder for resources
    if (!(Test-Path "$CoursePath\..\course-resources")) {
        ###############
        #Create Folder#
        ###############
        [VOID](New-Item "$CoursePath\..\course-resources" -type directory)

        ############
        #create acl#
        ############
        
        #new blank acl
        $acl = New-Object System.Security.AccessControl.DirectorySecurity

        #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
        $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
        $objType = [System.Security.AccessControl.AccessControlType]::Allow 

        #hosting team permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #course permissions on the resources folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #faculty permissions on the resources folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE) 

        #remove inheritance
        $acl.SetAccessRuleProtection($true, $true)

        ###########
        #apply acl#
        ###########

        Write-Host "Applying ACL for $CoursePath\..\course-resources..."
        "Applying ACL for $CoursePath\..\course-resources..." #| Out-File c:\$CourseID.log -Append
        
        try {$acl | Set-Acl "$CoursePath\..\course-resources"}
        catch {
            $output = "" + $username + "- " + $error[0]
            Write-Host $output
            $output #| Out-File c:\$CourseID.log -Append
        }
        Write-Host "ACL applied."
        "ACL applied." #| Out-File c:\$CourseID.log -Append
    }

}

# This will act as Set-WPIACLs except it will not grant read access to the entire class for each student folder
function Set-WPIACLs-NoClassVision([string]$Term, $CourseID, $dept) {
    [CmdletBinding()]
    $rootPath = "$rootPath\$dept"
    
    # Change Term number for A and B Terms
    # if ($Term.StartsWith("B") -or $Term.StartsWith("A")) {
    #     $TermFolder = $Term[0] + ($Term.Substring(1, 2) - 1)
    # }
    # else {
    $TermFolder = $Term[0] + $Term.Substring(1, 2)
    # }
    

    #build course path
    $course = $CourseID.Replace("_", "") #Convert CourseID from XXXX_#### to XXXX#### for the folder path
    $CoursePath = "$rootPath\$course\$TermFolder\submissions"

    # build faculty list
    $FacultyGroup = $CourseID + "_Faculty"
    $FacultyList = Get-ADGroupMember $FacultyGroup
    Write-Host "$CourseID $Term has the following faculty: $($FacultyList.SamAccountName)" #| Out-File c:\$CourseID.log
    #"$CourseID $Term has the following students: $($facultyList.SamAccountName)" #| Out-File c:\$CourseID.log

    # create acls for submissions folder
    if (!(Test-Path $CoursePath)) {
        ###############
        #Create Folder#
        ###############
        [VOID](New-Item $CoursePath -type directory)

        ############
        #create acl#
        ############
        
        #new blank acl
        $acl = New-Object System.Security.AccessControl.DirectorySecurity

        #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
        $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
        $objType = [System.Security.AccessControl.AccessControlType]::Allow 

        #hosting team permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #faculty permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 
        
        #course permissions on the parent folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #remove inheritance
        $acl.SetAccessRuleProtection($true, $true)

        ###########
        #apply acl#
        ###########

        Write-Host "Applying ACL for $CoursePath..."
        "Applying ACL for $CoursePath..." #| Out-File c:\$CourseID.log -Append
        
        try {$acl | Set-Acl $CoursePath}
        catch {
            $output = "" + $username + "- " + $error[0]
            Write-Host $output
            $output #| Out-File c:\$CourseID.log -Append
        }
        Write-Host "ACL applied."
        "ACL applied." #| Out-File c:\$CourseID.log -Append
    }

    
    
    # build student list, ensuring the proper term is selected
    $CourseSectionGroup = $CourseID + "_" + ($Term.SubString(0, 1))
    $year = $Term.SubString(1, 2)

    $StudentGroupList = Get-ADGroupMember $CourseID  | Where-Object {$_.Name -Like "$CourseSectionGroup*_20$Year*"} | Select-Object Name
    foreach ($group in $StudentGroupList) {
        #write-host "checking $($group.Name)"
        if ($null -ne (get-adgroup -LDAPFilter "(Name=$($group.Name))" -SearchBase "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu")) { #I removed OU=Term_20$($year)_$($Term.SubString(0, 1)), from the start, caused accuracy issues
            #Write-Host "group $($group.name) found"
            [array]$StudentList += Get-ADGroupMember -Recursive $group.Name
        }
        #[array]$StudentList += Get-ADGroupMember -Recursive $group.Name
    }
    Write-Host "$CourseID $Term has the following students: $($StudentList.SamAccountName)"
    #"$CourseID $Term has the following students: $($StudentList.SamAccountName)" #S| Out-File c:\$CourseID.log

    # create student folders and permissions
    foreach ($student in $StudentList) {
        $username = $student.SamAccountName

        #create folder if not exist
        if (!(Test-Path $CoursePath\$username)) {
            ###############
            #Create Folder#
            ###############
            [VOID](New-Item $CoursePath\$username -type directory)

            ############
            #create acl#
            ############
        
            #new blank acl
            $acl = New-Object System.Security.AccessControl.DirectorySecurity

            #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
            $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
            $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
            $objType = [System.Security.AccessControl.AccessControlType]::Allow 

            #hosting team permissions on the folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            $acl.AddAccessRule($objACE) 

            #student permissions on the student folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE)

            #faculty permissions on the student folder
            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE) 

            #remove inheritance
            $acl.SetAccessRuleProtection($true, $true)

            ###########
            #apply acl#
            ###########

            Write-Host "Applying ACL for $CoursePath\$username..."
            "Applying ACL for $CoursePath\$username..." #| Out-File c:\$CourseID.log -Append
        
            try {$acl | Set-Acl $CoursePath\$username}
            catch {
                $output = "" + $username + "- " + $error[0]
                Write-Host $output
                $output #| Out-File c:\$CourseID.log -Append
            }
            Write-Host "ACL applied."
            "ACL applied." #| Out-File c:\$CourseID.log -Append
        }
    
        # create faculty folders and permissions
        foreach ($faculty in $FacultyList) {
            $username = $faculty.SamAccountName

            #create folder if not exist
            if (!(Test-Path $CoursePath\$username)) {
                ###############
                #Create Folder#
                ###############
                [VOID](New-Item $CoursePath\$username -type directory)

                ############
                #create acl#
                ############
        
                #new blank acl
                $acl = New-Object System.Security.AccessControl.DirectorySecurity

                #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
                $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
                $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
                $objType = [System.Security.AccessControl.AccessControlType]::Allow 

                #hsoting team permissions on the folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                $acl.AddAccessRule($objACE) 

                #individual faculty permissions on the faculty folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                #add to acl
                $acl.AddAccessRule($objACE)

                #all faculty permissions on the faculty folder
                $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
                $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
                #add to acl
                $acl.AddAccessRule($objACE) 

                #remove inheritance
                $acl.SetAccessRuleProtection($true, $true)

                ###########
                #apply acl#
                ###########

                Write-Host "Applying ACL for $CoursePath\$username..."
                "Applying ACL for $CoursePath\$username..." #| Out-File c:\$CourseID.log -Append
        
                try {$acl | Set-Acl $CoursePath\$username}
                catch {
                    $output = "" + $username + "- " + $error[0]
                    Write-Host $output
                    $output #| Out-File c:\$CourseID.log -Append
                }
                Write-Host "ACL applied."
                "ACL applied." #| Out-File c:\$CourseID.log -Append
            }
        }
    }


    # create additional folder for resources
    if (!(Test-Path "$CoursePath\..\course-resources")) {
        ###############
        #Create Folder#
        ###############
        [VOID](New-Item "$CoursePath\..\course-resources" -type directory)

        ############
        #create acl#
        ############
        
        #new blank acl
        $acl = New-Object System.Security.AccessControl.DirectorySecurity

        #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
        $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
        $objType = [System.Security.AccessControl.AccessControlType]::Allow 

        #hosting team permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #course permissions on the resources folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #faculty permissions on the resources folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE) 

        #remove inheritance
        $acl.SetAccessRuleProtection($true, $true)

        ###########
        #apply acl#
        ###########

        Write-Host "Applying ACL for $CoursePath\..\course-resources..."
        "Applying ACL for $CoursePath\..\course-resources..." #| Out-File c:\$CourseID.log -Append
        
        try {$acl | Set-Acl "$CoursePath\..\course-resources"}
        catch {
            $output = "" + $username + "- " + $error[0]
            Write-Host $output
            $output #| Out-File c:\$CourseID.log -Append
        }
        Write-Host "ACL applied."
        "ACL applied." #| Out-File c:\$CourseID.log -Append
    }

}

function Set-DeptACLs ($dept, $term) {
    
    #get courses for the deptartment
    $courses = (get-course-groups $dept)

    foreach ($course in $courses) {
        Set-WPIACLs $term $course $dept
    }
}

# this function is to be used if a student joins the course late
function Set-WPIACLs-SingleUser($username, $CourseID, $path) {
    $CoursePath = $path
    $FacultyGroup = $CourseID + "_Faculty"
        
    #create folder if not exist
    if (!(Test-Path $CoursePath\$username)) {
        ###############
        #Create Folder#
        ###############
        [VOID](New-Item $CoursePath\$username -type directory)

        ############
        #create acl#
        ############
        
        #new blank acl
        $acl = New-Object System.Security.AccessControl.DirectorySecurity

        #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
        $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
        $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
        $objType = [System.Security.AccessControl.AccessControlType]::Allow 

        #hosting team permissions on the folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"FullControl" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\Access_Hosting_Services")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        $acl.AddAccessRule($objACE) 

        #student permissions on the student folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #faculty permissions on the student folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$FacultyGroup")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE) 

        #course permissions on the student folder
        $colRights = [System.Security.AccessControl.FileSystemRights]"ReadAndExecute"
        $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$CourseID")
        $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
        #add to acl
        $acl.AddAccessRule($objACE)

        #remove inheritance
        $acl.SetAccessRuleProtection($true, $true)

        ###########
        #apply acl#
        ###########

        Write-Host "Applying ACL for $CoursePath\$username..."
        "Applying ACL for $CoursePath\$username..." #| Out-File c:\$CourseID.log -Append
        
        try {$acl | Set-Acl $CoursePath\$username}
        catch {
            $output = "" + $username + "- " + $error[0]
            Write-Host $output
            $output #| Out-File c:\$CourseID.log -Append
        }
        Write-Host "ACL applied."
        "ACL applied." #S| Out-File c:\$CourseID.log -Append
    }
}


function New-Team-folderstructure ($folder, $file, $faculty) {
    #for each team create folder if it doesn't exist
    $teammembers = Import-Csv $file
    $log = "$file.log" 
    
    #Set Standard Flags/Object Types (These can be set manually/overridden for a specific ACL if needed)
    $InheritanceFlag = [System.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::None 
    $objType = [System.Security.AccessControl.AccessControlType]::Allow 
    $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
      
    Write-Host "Building directories for teams in file $file"
    "Building directorties for teams in file $file" #| Out-File $log -Append
    
    Write-Host "There are $teammembers.length people who need access."
        
    foreach ( $member in $teammembers) {
   
        $teamnumber = $member.TeamNumber
        $username = $member.username
        $foldername = "$folder\Team$teamnumber" 
   
        Write-Host "User $username, Team $teamnumber"
        "User $username, Team $teamnumber" #| Out-File $log -Append
   
   
        if (Test-Path  $foldername) {
            Write-Host "$foldername exists"
            "$foldername exists" #| Out-File $log -Append
      
            Write-Host "Granting permissions for $username on $foldername"
            "Granting permissions for $username on $foldername" | Out-File $log -Append
   
            #teamfolder exists, grant permissions for this team member
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl = get-acl $foldername
            $acl.AddAccessRule($objACE) 

            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$Faculty")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE)

            $acl | Set-Acl $foldername
        
        }
        else {
            #create folder and grant permission
            Write-Host "Creating directory $foldername"
            "Creating directory $foldername" | Out-File $log -Append
                
            [VOID](New-Item $foldername -type directory)

            Write-Host "Granting permissions for $username on $foldername"
            "Granting permissions for $username on $foldername" | Out-File $log -Append
   
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$username")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl = get-acl $foldername
            $acl.AddAccessRule($objACE) 

            $colRights = [System.Security.AccessControl.FileSystemRights]"Modify" 
            $objUser = New-Object System.Security.Principal.NTAccount("ADMIN\$Faculty")
            $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule ($objUser, $colRights, $InheritanceFlag, $PropagationFlag, $objType) 
            #add to acl
            $acl.AddAccessRule($objACE)


            $acl | Set-Acl $foldername
            
        }
    }

}


# Here is a really cool example that I left here on 011217
# Set-WPIACLs C17 IMGD_2048 IMGD