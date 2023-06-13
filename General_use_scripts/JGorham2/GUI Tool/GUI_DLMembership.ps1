param([switch] $Run)
function New-DLManager {
    #################################
    # Connect to Azure
    #################################
    #Import-Module AzureAD
	try{
		Connect-AzureAD
	}
	catch {
		Return "Could not connect to Azure AD"
	}


    function Remove-DLSubscription {
        [CmdletBinding()]
        param (
            $EmailAddress,
            $List
        )

        if (-not $emailaddress.Contains("@") -or $emailaddress -eq "User Email Address") {
            Return "Please Enter a Valid Email Address"
        }
        if ($list -eq "Distribution List") {
            Return "Please Enter a Valid Distribution List"
        }

        try {
            $User = Get-AzureADUser -ObjectId $EmailAddress
        }
        catch {
            Return "User $emailaddress Not Found"
        }
        $UID = $User.ObjectId

        if ($list.startswith("dl-") ) {
            $OFS = '-'
            [string]$trash,[string]$firstSplit = $list -split $OFS
            $ListCore = $firstSplit.split("@")[0]
            $OFS = ' '
        }
        else {
            $ListCore = $list.Split("@")[0]
        }

        if ($ListCore -eq "allemployees") {
            $ListCore = "employees"        
        }

        #$DL_ID =(Get-DistributionGroup dl-$ListCore).ExternalDirectoryObjectId
        try {
            $OptIn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Optin-$ListCore'").ObjectId
            $OptOut_ID = (Get-AzureADGroup -filter "DisplayName eq 'OptOut-$ListCore'").ObjectId
            $Dyn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Dyn-$ListCore'").ObjectId
        }
        catch {
            Return "$List Does not Appear to be a Valid List"
        }

        $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

        if ($membership.objectid -contains $Dyn_ID) {
            if ($membership.objectid -contains $OptOut_ID) {
                Return "User $($user.UserPrincipalName) is already opted out of $ListCore"
            }
            elseif ($membership.objectid -contains $OptIn_ID) {
                Remove-AzureADGroupMember -ObjectId $OptIn_ID -MemberId $UID
                Add-AzureADGroupMember -ObjectId $OptOut_ID -RefObjectId $UID
                Return "User $($user.UserPrincipalName) removed from $ListCore-OptIn Group and added to $ListCore-OptOut Group"
            }
            else {
                Add-AzureADGroupMember -ObjectId $OptOut_ID -RefObjectId $UID
                Return "User $($user.UserPrincipalName) added to $ListCore-OptOut Group"
            }
        }
        elseif ($membership.objectid -contains $OptIn_ID) {
            Remove-AzureADGroupMember -ObjectId $OptIn_ID -MemberId $UID
            Return "User $($user.UserPrincipalName) removed from $ListCore-OptIn Group"
        }
        else {
            Return "User $($user.UserPrincipalName) is not a member of $ListCore"
        }
    }

    function Add-DLSubscription {
        [CmdletBinding()]
        param (
            $EmailAddress,
            $List
        )

        if (-not $emailaddress.Contains("@") -or $emailaddress -eq "User Email Address") {
            Return "Please Enter a Valid Email Address"
        }
        if ($list -eq "Distribution List") {
            Return "Please Enter a Valid Distribution List"
        }

        try {
            $User = Get-AzureADUser -ObjectId $EmailAddress
        }
        catch {
            Return "User $emailaddress Not Found"
        }
        $UID = $User.ObjectId

        if ($list.startswith("dl-") ) {
            $ListCore = $list.split("-")[1].split("@")[0]
        }
        else {
            $ListCore = $list.Split("@")[0]
        }

        if ($ListCore -eq "allemployees") {
            $ListCore = "employees"        
        }

        try {
            $OptIn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Optin-$ListCore'").ObjectId
            $OptOut_ID = (Get-AzureADGroup -filter "DisplayName eq 'OptOut-$ListCore'").ObjectId
            $Dyn_ID = (Get-AzureADGroup -filter "DisplayName eq 'Dyn-$ListCore'").ObjectId
        }
        catch {
            Return "$List Does not Appear to be a Valid List"
        }

        $membership = Get-AzureADUserMembership -ObjectId $uid -all:$true

        if ($membership.objectid -contains $OptOut_ID) {
            Remove-AzureADGroupMember -ObjectId $OptOut_ID -MemberId $UID
            if ($membership.objectid -contains $Dyn_ID) {
                Return "User $($user.UserPrincipalName) removed from $ListCore-OptOut Group"
            }
        }

        if (!($membership.objectid -contains $OptIn_ID -or $membership.objectid -contains $Dyn_ID)) {
                Add-AzureADGroupMember -ObjectId $OptIn_ID -RefObjectId $UID
                Return "User $($user.UserPrincipalName) added to $ListCore-OptIn Group"
        }
        else {
            Return "User $($user.UserPrincipalName) is a member of $ListCore"
        }
    }

    function Get-TextTest {
        $ListCore = "Content"
        Return "User $($user.UserPrincipalName) removed from $Listcore-OptOut Group and added to $ListCore-OptIn Group"
    }

    #########################################
    #   Main
    #########################################

    Add-Type -assembly System.Windows.Forms
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.Text = 'DL Subscription Management'
    $main_form.Width = 500
    $main_form.Height = 150
    $main_form.AutoSize = $true

    $ListLabel = New-Object System.Windows.Forms.Label
    $ListLabel.Text = "Distribution List:"
    $ListLabel.Location = New-Object System.Drawing.Point(10, 5)
    $ListLabel.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($ListLabel)

    $ListField = New-Object System.Windows.Forms.TextBox
    $ListField.Text = ""
    $ListField.Location = New-Object System.Drawing.Point(10, 24)
    $ListField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($ListField)

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Text = "User Email Address:"
    $UserLabel.Location = New-Object System.Drawing.Point(10, 49)
    $UserLabel.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserLabel)

    $UserField = New-Object System.Windows.Forms.TextBox
    $UserField.Text = ""
    $UserField.Location = New-Object System.Drawing.Point(10, 69)
    $UserField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserField)

    $ReturnLabel = New-Object System.Windows.Forms.Label
    $ReturnLabel.Text = ""
    $ReturnLabel.Location = New-Object System.Drawing.Point(170, 5)
    $ReturnLabel.Size = New-Object System.Drawing.Size(350, 60)
    $ReturnLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $main_form.Controls.Add($ReturnLabel)

    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Point(10, 94)
    $AddButton.Size = New-Object System.Drawing.Size(150, 20)
    $AddButton.Text = "Subscribe User to List"

    $main_form.Controls.Add($AddButton)
    $AddButton.Add_Click(
        {
            $ReturnLabel.Text = Add-DLSubscription -List $ListField.Text -EmailAddress $UserField.Text
            $ListField.Text = ""
            $UserField.Text = ""
        })

    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Point(10, 119)
    $RemoveButton.Size = New-Object System.Drawing.Size(150, 20)
    $RemoveButton.Text = "Unsubscribe User to List"

    $main_form.Controls.Add($RemoveButton)
    $RemoveButton.Add_Click(
        {
            $ReturnLabel.Text = Remove-DLSubscription -List $ListField.Text -EmailAddress $UserField.Text
            $ListField.Text = ""
            $UserField.Text = ""
        })

    $main_form.ShowDialog()
    
}

if ($Run) {
    New-DLManager
}