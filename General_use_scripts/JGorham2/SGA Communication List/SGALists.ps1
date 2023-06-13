## Get all the students
## Split them into 100 chunks
## make each of them a list

$SGASortedGroupUID = "db4e1daf-530f-4008-bdf2-4f253d2e84cb"
$AllStudentsGUID = "01a5ccb6-128f-4abe-84ce-a1e8328de0e1"

$studentList = Get-AzureADGroupMember -ObjectId $AllStudentsGUID -All:$true                 #get all our active students
$previouslySorted = Get-AzureADGroupMember -ObjectId $SGASortedGroupUID -All:$true          #and then get the list of everyone we don't need to sort
$studentNumber = $studentList.count                                                         #this is going to get used later to generate the max size we should be having for each list

#gets the number of users that should be in each group, rounded down
$populationNumber = [math]::Truncate($studentNumber / 100)                                  #later

$groupFormat = "dl-SGA-"            #Just Hardcode in the start of the string
$groupCounter = 0                   #and make the tail end programmatic

$GroupPopulation = @()

$comparisons = Compare-Object -ReferenceObject $studentList -DifferenceObject $previouslySorted -Property UserPrincipalName     #standard syncer code minus the need for removals
$NewStudentList = $comparisons | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object UserPrincipalName

for ($i = 0; $i -le 99; $i++) {                                                 #this populates an array with group names and their current populations, this was so that
    $group = "$($groupFormat)$($i)"                                             #I didn't need to keep recalling the get line. Do it once then just track it in memory
    $currentPop = (Get-DistributionGroupMember -Identity $group).count
    $GroupPopulation += @{$group = $currentPop }
}

foreach ($student in $NewStudentList) {                                                 # For each user who hasn't been already sorted
    $group = "$($groupformat)$($groupcounter)"                                          # compose name of the group we're working on currently
    if ($GroupPopulation.group -le $populationNumber) {                                 # Add them to the lowest group till it matches our calculated pop
        Add-DistributionGroupMember -Identity $group -Member $student
        #adds user to the SGASorted Group
        Add-AzureADGroupMember -ObjectId $SGASortedGroupUID -RefObjectId $student
        $GroupPopulation.group++
    }
    elseif ($groupCounter -eq 99) {                                                     # if there are already the maximum number of people in the group "Are we at the end of the list?"
        $GroupPopulation.group++                                                        # If so, then we increase the max allowance of the groups
        $groupCounter = 0                                                               # and start from the top
    }
    else {
        $groupCounter++                                                                 #otherwise, if we've just capped out this one group we move on to the next
    }    
}

<# So this version just adds people to each group until they meet the calculated number, then dumps the remainder into 99
    This verion _relies_ on the timely removal of students from these groups, because it only adds, it does this in order to keep people in the 
    same groups for their lifetime at wpi.
#>