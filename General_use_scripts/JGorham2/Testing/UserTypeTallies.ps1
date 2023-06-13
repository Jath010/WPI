#Get the number of each user group that fits each extensionattribute7 total

write-host "Number of Active Staff "(get-aduser -filter {extensionattribute7 -eq "Staff" -and Enabled -eq "True"}).count
write-host "Number of Active Student "(get-aduser -filter {extensionattribute7 -eq "Student" -and Enabled -eq "True"}).count
write-host "Number of Active Faculty "(get-aduser -filter {extensionattribute7 -eq "Faculty" -and Enabled -eq "True"}).count
write-host "Number of Active Affiliate "(get-aduser -filter {extensionattribute7 -eq "Affiliate" -and Enabled -eq "True"}).count
write-host "Number of Active Alum "(get-aduser -filter {extensionattribute7 -eq "Alum" -and Enabled -eq "True"}).count