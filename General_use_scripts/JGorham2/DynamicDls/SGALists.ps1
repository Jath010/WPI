$undergraduates = get-aduser -Filter {extensionattribute3 -eq "freshman" -or extensionattribute3 -eq "sophomore"-or extensionattribute3 -eq "junior"-or extensionattribute3 -eq "senior"} -Properties employeeID
$UndergradIDs = $undergraduates.EmployeeID
for($i=0;$i -lt 100;$i++){
    $num = "{0:D2}" -f $i
    Write-Host $num ($UndergradIDs|Where-Object {$_.endswith($num)}).count
}