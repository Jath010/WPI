#Jmgorham2 10/28/2020
#Declare Paths

#To run just run the script followed by term number
#EX. .\SDSBannerImport.ps1 C21

#real target will be '\\bannerprod.wpi.edu\canvas$'
$DataPath = '\\bannerprod.wpi.edu\canvas$'
$ExportPath = 'D:\tmp\Bifrost copy\School Data Sync\testing'

#Grab Args
if (!$args[0]) {
    Write-Host "Please enter a Term number such as A21"
    Exit
}
$terms = "A","B","C","D","E","F","S"
if($args[0].substring(0,1) -in $terms){
    $Term = $args[0]
}
else{
    write-host "Valid Terms begin with A,B,C,D,E,F, or S"
    exit
}

$Courses = Import-Csv -path (Get-ChildItem -path $DataPath -filter "Courses-$Term.csv").FullName
$enrollment = Import-Csv -path (Get-ChildItem -path $DataPath -filter "Enrollment-$Term.csv").FullName
$users = Import-Csv -path (Get-ChildItem -path $DataPath -filter 'Users*.csv').FullName

#De-dupe Enrollment
$UniqueEnrollment = $enrollment | Sort-Object user_id -Unique

#create Section.csv
$section = $courses | Select-Object @{expression = { $_.course_id }; label = 'SIS ID' }, @{expression = { "1" }; label = 'School SIS ID' }, @{expression = { $_.long_name }; label = 'Section Name' }, @{expression = { $_.term_id }; label = 'Term Name' }

$section | Export-Csv -Path "$ExportPath\Section.csv" -NoTypeInformation

#Create Student.csv

#Compare-Object -ReferenceObject $users -DifferenceObject $studentids -Property SamAccountName -ExcludeDifferent -IncludeEqual -PassThru

$studentIds = $UniqueEnrollment | Where-Object { $_.role -eq "Student" -and $_.status -eq "Active" } | Select-Object user_ID

$students = Compare-Object -ReferenceObject $users -DifferenceObject $studentids -Property user_id -ExcludeDifferent -IncludeEqual -PassThru | Select-Object @{expression = { $_.user_id }; label = 'SIS ID' }, @{expression = { "1" }; label = 'School SIS ID' }, @{expression = { $_.login_id }; label = 'Username' }

$students | Export-Csv -Path "$ExportPath\Student.csv" -NoTypeInformation
#$students = $users|Where-Object {$_.user_ID -in $studentIds} |Select-Object @{expression={$_.user_id}; label='SIS ID'},@{expression={"1"}; label='School SIS ID'},@{expression={$_.login_id}; label='Username'}



#Create Teacher.csv
$TeacherIds = $UniqueEnrollment | Where-Object { $_.role -eq "Teacher" -and $_.status -eq "Active" } | Select-Object user_id

$Teachers = Compare-Object -ReferenceObject $users -DifferenceObject $Teacherids -Property user_id -ExcludeDifferent -IncludeEqual -PassThru | Select-Object @{expression = { $_.user_id }; label = 'SIS ID' }, @{expression = { "1" }; label = 'School SIS ID' }, @{expression = { $_.login_id }; label = 'Username' }

$Teachers | Export-Csv -Path "$ExportPath\Teacher.csv" -NoTypeInformation

#Create StudentEnrollment.csv

$StudentEnrollment = $enrollment | Where-Object { $_.role -eq "Student" } | Select-Object @{expression = { $_.course_id }; label = 'Section SIS ID' }, @{expression = { $_.user_id }; label = 'SIS ID' }

$StudentEnrollment | Export-Csv -Path "$ExportPath\StudentEnrollment.csv" -NoTypeInformation

#Create Teacher Enrollmentcsv

$TeacherRoster = $enrollment | Where-Object { $_.role -eq "Teacher" } | Select-Object @{expression = { $_.course_id }; label = 'Section SIS ID' }, @{expression = { $_.user_id }; label = 'SIS ID' }

$TeacherRoster | Export-Csv -Path "$ExportPath\TeacherRoster.csv" -NoTypeInformation

#Create School
$school = New-Object -TypeName psobject
$School | Add-Member -MemberType NoteProperty -Name "SIS ID" -Value 1
$school | Add-Member -MemberType NoteProperty -Name "Name" -Value "Worcester Polytechnic Institute"

$school | Export-Csv -Path "$ExportPath\School.csv" -NoTypeInformation