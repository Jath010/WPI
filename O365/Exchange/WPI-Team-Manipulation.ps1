# Collection of functions for manipulating MS Teams

Import-Module MicrosoftTeams

#The MicrosoftTeams module requires you to connect to Teams to run the pieces, I'm looking into allowing it to do so without an interactive prompt
Connect-MicrosoftTeams

#Accepts Input as New-WPIClassTeam IMGD_2101 A19
#This will create a group with an alias of Team_IMGD_2101_A19 and you can find that with Get-UnifiedGroup -Filter {alias -like "Team_IMGD_2101_A19"}
#the team groupID is the ExternalDirectoryObjectId value in the group object.
function New-WPIClassTeam {
    [CmdletBinding()]
    param (
        $CourseID,
        $Term
    )
    $SplitName = $CourseID.Split("_")

    #Create the group
    #Format of the email address should be something like IMGD_2101_A19@wpi.edu
    #Format of the DisplayName should be "IMGD 2101 A19"
    $group = New-Team -alias "Team_${CourseID}_${Term}" -displayname "$SplitName $Term" -AddCreatorASMember $false -Template EDU_Class

    #Pull the lists from AD
    $faculty = Get-WPICourseFacultyList -CourseID $CourseID
    $students = Get-WPICourseStudentList -CourseID $CourseID -Term $Term

    #Make each member of the faculty group into an Owner for the Group
    foreach($professor in $faculty){
        Add-TeamUser -GroupID $group.GroupID -User "${professor}@wpi.edu" -Role Owner
    }

    #Make everyone in the student groups into regular Members
    foreach($student in $students){
        Add-TeamUser -GroupID $group.GroupID -User "${student}@wpi.edu" -Role Member
    }
}

function Get-WPIClassTeamMembers {
    param (
        $ClassAlias
    )
    $team = Get-UnifiedGroup -Identity $ClassAlias
    $UserList = Get-TeamUser -GroupId $team.ExternalDirectoryObjectId -Role Member
    
    $UserList.User
}

function Get-WPIClassTeamOwners {
    param (
        $ClassAlias
    )
    $team = Get-UnifiedGroup -Identity $ClassAlias
    $UserList = Get-TeamUser -GroupId $team.ExternalDirectoryObjectId -Role Owner
    
    $UserList.User
}

function Sync-WPIClassTeam {
    [CmdletBinding()]
    param (
        $ClassAlias
    )
    ###################################################################
    #Data Section
    #Split Up the Alias into 
    $SplitAlias = $ClassAlias.Split("_")
    $Dept = $SplitAlias[1]
    $CourseNumber = $SplitAlias[2]
    $CourseID = "${Dept}_${CourseNumber}"
    $Term = $SplitAlias[3]

    #Load the team 
    $group = Get-UnifiedGroup -Identity $ClassAlias
    $team = $group.ExternalDirectoryObjectId

    #Get Team Lists
    $TeamMemberList = Get-TeamUser -GroupId $team -Role Member
    $TeamOwnerList = Get-TeamUser -GroupId $team -Role Owner

    #Pull the lists from AD
    $ADFaculty = Get-WPICourseFacultyList -CourseID $CourseID
    $ADStudents = Get-WPICourseStudentList -CourseID $CourseID -Term $Term

    #Pull SysOps from AD
    $ADSysOps = Get-ADGroupMember SysOps
    ####################################################################

    #Sync Section
    #Faculty Section
    #Check if a Faculty Member has been removed, if so, remove them
    foreach($Owner in $TeamOwnerList) {
        Write-Verbose "Checking if ${Owner} should be an Owner."
        $exists = $false
        #Compare Owner with each member of the Faculty group
        foreach($professor in $ADFaculty) {
            $ProfessorEmail = "${professor}@wpi.edu"
            if($ProfessorEmail -eq $owner) {
                Write-Verbose "${Owner} is a Professor."
                $exists = $true
            }
        }
        #Also make sure not to remove any SysOps User that's in the group
        foreach($Op in $ADSysOps) {
            $OpEmail = "${Op}@wpi.edu"
            if($OpEmail -eq $Owner) {
                Write-Verbose "${Owner} is an Administrator."
                $exists = $true
            }
        }

        if($exists -eq $false) {
            Write-Verbose "${Owner} should not be an Owner. They will now be removed."
            Remove-TeamUser -GroupId $team -user $owner
        }
    }
    #Check if a Faculty menber has been added, if so, add them
    foreach($professor in $ADFaculty) {
        Write-Verbose "Checking to see if ${professor} should be added."
        $exists = $false
        $email = "${professor}@wpi.edu"
        #Compare Faculty member with each Owner
        foreach($Owner in $TeamOwnerList) {
            if($email -eq $Owner) {
                Write-Verbose "${professor} is already an Owner."
                $exists = $true
            }
        }
        if($exists -eq $false) {
            Write-Verbose "${professor} needs to be added as an Owner. Doing so."
            Add-TeamUser -GroupId $team -User $email -Role Owner
        }
    }

    #Student Section
    #Check if a Student has been removed, if so, remove them
    foreach($Member in $TeamMemberList) {
        Write-Verbose "Checking if ${Member} should be a Member."
        $exists = $false
        #Compare Member with each Student in the group
        foreach($student in $ADStudents) {
            $email = "${student}@wpi.edu"
            if($email -eq $Member) {
                Write-Verbose "${Owner} is a Student."
                $exists = $true
            }
        }
        if($exists -eq $false) {
            Write-Verbose "${Owner} is not a Student. Removing from Team."
            Remove-TeamUser -GroupId $team -user $Member
        }
    }
    #Check if a Student has been added, if so, add them
    foreach($student in $ADStudents) {
        Write-Verbose "Checking if ${Member} needs to be added as a Member."
        $exists = $false
        $email = "${student}@wpi.edu"
        #Compare Student with each member of the team
        foreach($Member in $TeamMemberList) {
            if($email -eq $Member) {
                Write-Verbose "${Member} is already a Member."
                $exists = $true
            }
        }
        if($exists -eq $false) {
            Write-Verbose "${Member} needs to be added as a Member. Doing so."
            Add-TeamUser -GroupId $team -User $email -Role Member
        }
    }
}

#From WPI-Courselists
#Get List of students in course based on year and term

#Input Scheme
#Get-WPICourseStudentList IMGD_2101 A19
#TODO:Add handing for IMGD/AR crossposting
function Get-WPICourseStudentList {
    param (
        $CourseID,
        $Term
    )
    $SplitName = $CourseID.Split("_")

    # build student list, ensuring the proper term is selected
    $CourseSectionGroup = $CourseID+"_"+($Term.SubString(0,1))
    $year = $Term.SubString(1,2)

    # So the banner information gets tossed into groups with the name scheme like IMGD_2101_A01_201701_A

    $StudentGroupList = Get-ADGroupMember $CourseID | Where-Object {$_.Name -Like "$CourseSectionGroup*_20$Year*"} | Select-Object Name
    foreach ($group in $StudentGroupList) {
        $ContactList += Get-ADGroupMember -Recursive $group.Name
    }
    if($Splitname[0] -eq "IMGD"){
        $ARClass = "AR_" + $SplitName[1]
        $StudentGroupList = Get-ADGroupMember $ARClass | Where-Object {$_.Name -Like "$CourseSectionGroup*_20$Year*"} | Select-Object Name
        foreach ($group in $StudentGroupList) {
            $ContactList += Get-ADGroupMember -Recursive $group.Name
        }
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
    $SplitName = $CourseID.Split("_")

    $FacultyGroup = $CourseID+"_Faculty"
    $FacultyList += Get-ADGroupMember $FacultyGroup
    if($SplitName[0] -eq "IMGD"){
        $FacultyGroup = "AR_"+ $SplitName[1] + "_Faculty"
        $FacultyList += Get-ADGroupMember $FacultyGroup
    }

    foreach($user in $FacultyList){
        $user.SamaccountName
    }
}

function Set-TeamBoxUnhide {
    [CmdletBinding()]
    param (
        $team
    )
    
    begin {
        $group = Get-UnifiedGroup $team
    }
    
    process {
        Set-UnifiedGroup $group.PrimarySmtpAddress -HiddenFromExchangeClientsEnabled:$false
    }
    
    end {
        
    }
}