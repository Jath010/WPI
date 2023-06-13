$studentlist = Get-ADUser -Filter {extensionattribute7 -eq "Student" -and enabled -eq "true"} -Properties extensionattribute7
$facultylist = Get-ADUser -Filter {extensionattribute7 -eq "faculty" -and enabled -eq "true"} -Properties extensionattribute7
$affiliatelist = Get-ADUser -Filter {extensionattribute7 -eq "affiliate" -and enabled -eq "true"} -Properties extensionattribute7
$alumlist = Get-ADUser -Filter {extensionattribute7 -eq "alum" -and enabled -eq "true"} -Properties extensionattribute7
$stafflist = Get-ADUser -Filter {extensionattribute7 -eq "staff" -and enabled -eq "true"} -Properties extensionattribute7


$studentcount = $studentlist.count
$facultycount = $facultylist.count
$affiliatecount = $affiliatelist.count
$alumcount = $alumlist.count
$staffcount = $stafflist.count

$studentMissingCount = 0
$facultyMissingCount = 0
$affiliateMissingCount = 0
$alumMissingCount = 0
$staffMissingCount = 0
Connect-MsolService | Out-Null
$c1=0
foreach($student in $studentlist){
    $c1++
    Write-Progress -Activity "Processing Students" -CurrentOperation $Student.samaccountname -PercentComplete (($c1/$studentcount)*100)
    if($null -like (Get-MsolUser -UserPrincipalName $student.UserPrincipalName).StrongAuthenticationMethods){
        $studentMissingCount++
        $student | Export-Csv -Path .\MissingMFAStudent.csv -NoTypeInformation -Append
    }
}
$c1=0
foreach($faculty in $facultylist){
    $c1++
    Write-Progress -Activity "Processing Faculty" -CurrentOperation $faculty.samaccountname -PercentComplete (($c1/$facultycount)*100)
    if($null -like (Get-MsolUser -UserPrincipalName $faculty.UserPrincipalName).StrongAuthenticationMethods){
        $facultyMissingCount++
        $faculty | Export-Csv -Path .\MissingMFAFaculty.csv -NoTypeInformation -Append
    }
}
$c1=0
foreach($affiliate in $affiliatelist){
    $c1++
    Write-Progress -Activity "Processing Affiliates" -CurrentOperation $affiliate.samaccountname -PercentComplete (($c1/$affiliatecount)*100)
    if($null -like (Get-MsolUser -UserPrincipalName $affiliate.UserPrincipalName).StrongAuthenticationMethods){
        $affiliateMissingCount++
        $affiliate | Export-Csv -Path .\MissingMFAAffiliate.csv -NoTypeInformation -Append
    }
}
$c1=0
foreach($alum in $alumlist){
    $c1++
    Write-Progress -Activity "Processing Alumni" -CurrentOperation $alum.samaccountname -PercentComplete (($c1/$alumcount)*100)
    if($null -like (Get-MsolUser -UserPrincipalName $alum.UserPrincipalName).StrongAuthenticationMethods){
        $alumMissingCount++
        $alum | Export-Csv -Path .\MissingMFAAlum.csv -NoTypeInformation -Append
    }
}
$c1=0
foreach($staff in $stafflist){
    $c1++
    Write-Progress -Activity "Processing Staff" -CurrentOperation $alum.samaccountname -PercentComplete (($c1/$staffcount)*100)
    if($null -like (Get-MsolUser -UserPrincipalName $staff.UserPrincipalName).StrongAuthenticationMethods){
        $staffMissingCount++
        $staff | Export-Csv -Path .\MissingMFAStaff.csv -NoTypeInformation -Append
    }
}
Write-Host "Students Missing "$studentMissingCount" out of "$studentcount
Write-Host "Faculty Missing "$facultyMissingCount" out of "$facultycount
Write-host "Affiliates Missing "$affiliateMissingCount" out of "$affiliatecount
Write-Host "Alumni Missing "$alumMissingCount" out of "$alumcount
Write-Host "Staff Missing "$staffMissingCount" out of "$staffcount