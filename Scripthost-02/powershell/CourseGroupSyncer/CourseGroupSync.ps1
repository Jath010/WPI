##Get the Contents of the relevant flat files
##Then we want to run down the contents of the courses file and create the course
##After creating the course, run down the contents of the enrollment file and get everyone who is supposed to be in the course
##diff the current course with the intended population


# Set path for log files:
$logPath = "D:\wpi\Logs\ADCourseSyncer"

# Get date for logging and file naming:
$date = Get-Date
$datestamp = $date.ToString("yyyyMMdd-HHmm")

# Delete any logs older than 30 days.
Get-ChildItem $logPath -Recurse -Force -ea 0 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
ForEach-Object {
    $_ | Remove-Item -Force -Recurse
}

Start-Transcript -Append -Path "$($logPath)\$($datestamp)_ADCourseSyncer.log" -Force


$ExportDirectory = "\\storage.wpi.edu\dept\Workday Integrations\Information Technology\WPI_INT1211 - Students and Course info to School Data Sync\wpi-prod\files"
$EnrollmentData = Import-Csv $ExportDirectory"\StudentEnrollments.csv"
$CourseData = Import-Csv $ExportDirectory"\Section.csv"
$FacultyEnrollment = Import-Csv $ExportDirectory"\TeacherRoster.csv"

$TopOU = "OU=Courses,OU=Banner Groups,OU=Groups,DC=admin,DC=wpi,DC=edu"

##So, act on each course by creating the group if it doesn't exist, attaching it to the top level, then collecting the users
foreach ($Course in $CourseData) {

    Write-host "Beginning Course $($Course.'SIS ID')"

    $ClassCode = $Course.'SIS ID'.split("-")[0]
    if ($ClassCode.contains("/")) {
        $ClassCode = $ClassCode.replace("/", "-")
    }
    if($ClassCode -match "MQP$|IQP$|ISG$|DR$|PHD|THES$"){
        continue
    }
    
    $CourseStudentUIDs = $EnrollmentData | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
    $CourseTeacherUIDs = $FacultyEnrollment | Where-Object { $_.'Section SIS ID' -eq $Course.'SIS ID' } | Select-Object 'SIS ID'
    $CourseStudent = @()
    $CourseTeachers = @()
    

    if ($Course.'SIS ID'.split("-")[1] -eq "Lab") {
        $TermYear = $Course.'SIS ID'.split("-")[1]
        $Term = $TermYear.substring(0, $TermYear.Length - 2)
        $year = "20" + $TermYear.substring($TermYear.Length - 2)
        $Semester = "01"
        $section = ($Course.'SIS ID'.split("-")[3]).split(" ")[0]
        $DeptCode = $ClassCode -replace '\d+.*', ''
        $ClassNumber = ($ClassCode.TrimStart($DeptCode)).split(" ")[0]
        switch ($Term) {
            A { $Semester = "01" }
            B { $Semester = "01" }
            C { $Semester = "02"<# ; [string]$year = [int]$year+1 #> }
            D { $Semester = "02"<# ; [string]$year = [int]$year+1 #> }
            S { $Semester = "03"<# ; [string]$year = [int]$year+1 #> }
            F { $Semester = "03"<# ; [string]$year = [int]$year+1 #> }
            Default { $Semester = "01" }
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
        switch ($Term) {
            A { $Semester = "01" }
            B { $Semester = "01" }
            C { $Semester = "02"<# ; [string]$year = [int]$year+1 #> }
            D { $Semester = "02"<# ; [string]$year = [int]$year+1 #> }
            S { $Semester = "03"<# ; [string]$year = [int]$year+1 #> }
            F { $Semester = "03"<# ; [string]$year = [int]$year+1 #> }
            Default { $Semester = "01" }
        }
    }

    $OU = "OU=Term_" + $year + "_" + $Term + "," + $TopOU
    $OU_Name = "Term_" + $year + "_" + $Term
    
    #Verify the OU already exists for the term, else create

    if ((Get-ADOrganizationalUnit -filter 'Name -like $OU_Name').Name -ne "$OU_Name") {
        Write-Host "Term/Year OU Not Found. Creating"
        New-ADOrganizationalUnit -path $TopOU -name $OU_Name
    }
    else{
        Write-Host "Term/Year OU Found: $($OU_Name)"
    }

    #Compose Names based on the existing naming scheme
    [string]$GroupName = $DeptCode + "_" + $ClassNumber + "_" + $section + "_" + $year + $Semester + "_" + $Term

    Write-Host "$groupname"

    $TopGroupName = $DeptCode + "_" + $ClassNumber
    [string]$FacultyGroupName = $DeptCode + "_" + $ClassNumber + "_Faculty"

    #Verify Course Group Exists
    try {
        Get-ADGroup $TopGroupName | Out-Null
        Write-Host "Top Group Found"
    }#Else Create
    catch {
        Write-Host "Creating Top Group"
        New-ADGroup $TopGroupName -Path $TopOU -GroupScope Universal -GroupCategory Security
    }

    #Verify Faculty Group Exists
    try {
        Get-ADGroup $FacultyGroupName | Out-Null
        Write-Host "Faculty Group Found"
    }#Else Create and Nest
    catch {
        Write-Host "Creating Faculty Group"
        New-ADGroup $FacultyGroupName -Path $TopOU -GroupScope Universal -GroupCategory Security
        $timer = [diagnostics.stopwatch]::startnew()
        while($timer.elapsed.totalseconds -lt 20){
            if(Get-ADGroup -filter 'samaccountname -eq $GroupName'){
                Add-ADPrincipalGroupMembership -Identity $FacultyGroupName -MemberOf $TopGroupName
            }
            else{
                start-sleep -seconds 5
            }
        }
    }

    #Verify Section Group Exists
    try {
        Get-ADGroup $GroupName | Out-Null
        Write-Host "Group Found"
    }#Else Create and Nest
    catch {
        Write-Host "Creating Group"
        New-ADGroup -Name $GroupName -Path $OU -GroupScope Universal -GroupCategory Security
        $timer = [diagnostics.stopwatch]::startnew()
        while($timer.elapsed.totalseconds -lt 20){
            if(Get-ADGroup -filter 'samaccountname -eq $GroupName'){
                Add-ADPrincipalGroupMembership -Identity $GroupName -MemberOf $TopGroupName
            }
            else{
                start-sleep -seconds 5
            }
        }
    }

    foreach ($UID in $CourseStudentUIDs) {
        $CourseStudent += [PSCustomObject]@{'samaccountname' = ((get-aduser -Filter "uidNumber -like $($UID.'SIS ID') -and enabled -eq 'true'" -Properties extensionattribute8) | Where-Object {$_.extensionattribute8 -match "(.*;)*Student;.*"} ).samaccountname } | Where-Object {$null -ne $_.samaccountname}
    }
    foreach ($UID in $CourseTeacherUIDs) {
        $CourseTeachers += [PSCustomObject]@{'samaccountname' = (get-aduser -Filter "uidNumber -like $($UID.'SIS ID') -and enabled -eq 'true'").samaccountname } | Where-Object {$null -ne $_.samaccountname}
    }

    #cleanup output
    $CourseStudent = $CourseStudent | Sort-Object -Property samaccountname -Unique
    $CourseTeachers = $CourseTeachers | Sort-Object -Property samaccountname -Unique

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
    Write-Host "Checking $GroupName"
    $currentSection = Get-ADGroupMember $GroupName | Select-Object samaccountname
    if ($null -eq $currentSection) {
        Write-Host "Group is empty. Initializing."
        $AddMembers = $CourseStudent
    }
    elseif ($null -eq $courseStudent) {
        Write-Host "Group should be is empty. Clearing."
        $RemoveMembers = $currentSection
    }
    else {
        Write-Host "Running Comparison"
        $comparisons = Compare-Object $currentSection $CourseStudent -Property samaccountname
        Write-Host "$($comparisons.count) differences."
        $AddMembers = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object samaccountname
        Write-Host "$($AddMembers.count) Additions"
        $RemoveMembers = $comparisons | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object samaccountname
        Write-Host "$($RemoveMembers.count) Removals"
    }
    foreach ($removal in $RemoveMembers.samaccountname) {
        Write-Host "Removing $removal"
        Remove-ADGroupMember -Identity $GroupName -Members $removal -Confirm:$false
    }
    foreach ($addition in $AddMembers.samaccountname) {
        Write-Host "Adding $addition"
        Add-ADGroupMember -Identity $GroupName -Members $addition -Confirm:$false
    }

}

Stop-Transcript