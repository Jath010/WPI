function Get-WorkdayTerms {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Cred
    )

    $Params = @{
        Uri        = "https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_Changes_Terms?format=json"
        #wtth data from course
        # Uri = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212-_AD_Instructors__Course_Section_Definitions_?Effective_Date=2021-07-12-07:00&Academic_Period!Academic_Period_ID=' + $codes + '&format=json'
        # with data from section
        #URI = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_group_changes_Student_Enrollments__student_course_registration_records_?Starting_Academic_Period!Academic_Period_ID=' + $codes + '&Course_Section!Section_Listing_ID=' + $section +'&format=json'
        #Authentication = "Basic"
        Credential = $cred
    }
    (Invoke-RestMethod @params).Report_Entry #| fl

}

function Get-WorkdayTermContents {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet("A","B","C","D","Summer Session I","Summer Session II")]
        $term,
        [Parameter(Mandatory=$true)]
        [String]
        $Year,
        [Parameter(Mandatory=$true)]
        $WorkdayTermData
    )

    $termReferenceID = ($WorkdayTermData | where-object { $_.Academic_Period -match "$year.* $Term .*" } | select-object referenceID).referenceID

    $Call2 = @{
        #  Uri = "https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_Changes_Terms?format=json"
        #wtth data from course
        Uri        = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212-_AD_Instructors__Course_Section_Definitions_?Effective_Date=2021-07-12-07:00&Academic_Period!Academic_Period_ID=' + $termReferenceID + '&format=json'
        # with data from section
        #URI = 'https://wd5-impl-services1.workday.com/ccx/service/customreport2/wpi2/jplunkett/WPI_INT1212_-_AD_group_changes_Student_Enrollments__student_course_registration_records_?Starting_Academic_Period!Academic_Period_ID=' + $codes + '&Course_Section!Section_Listing_ID=' + $section +'&format=json'
        #Authentication = "Basic"
        Credential = $cred
    }
    (Invoke-RestMethod @Call2).Report_Entry
}



#Jobs bullshit

$saturdayTrigger = New-JobTrigger -At "01/29/2022 1:00:00"
$sundayTrigger = New-JobTrigger -At "01/30/2022 1:00:00"

Register-ScheduledJob -Name "Remove Card Access Addresses" -Trigger $saturdayTrigger -ScriptBlock {
    $group = Get-UnifiedGroup gr-campuscardaccess
    Set-UnifiedGroup -Identity $group.name -EmailAddresses @{remove="smtp:gr-campuscardaccess@wpi.edu"}
    Set-UnifiedGroup -Identity $group.name -EmailAddresses @{remove="SMTP:campuscardaccess@wpi.edu"}
}

Register-ScheduledJob -Name "Create CardAccess Mailbox" -Trigger $sundayTrigger -ScriptBlock {
    $Box = New-Mailbox -Shared -Name "Campus Card Access" -DisplayName "Campus Card Access" 
    $box | Set-Mailbox -ForwardingAddress (Get-Mailbox registrar@wpi.edu).ForwardingAddress

}

$users = get-aduser -Filter "extensionAttribute7 -eq 'staff' -or extensionAttribute7 -eq 'faculty'" -Properties extensionAttribute7, extensionAttribute5
$users | where-object {$_.extensionAttribute5 -match ".*;F"}