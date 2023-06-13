param([switch] $Run)
function New-DLOwnerManager {
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


    function Remove-OtpOwnership {
        [CmdletBinding()]
        param (
            $EmailAddress
        )

        if (-not $emailaddress.Contains("@") -or $emailaddress -eq "User Email Address") {
            Return "Please Enter a Valid Email Address"
        }

        try {
            $User = Get-AzureADUser -ObjectId $EmailAddress
        }
        catch {
            Return "User $emailaddress Not Found"
        }
        $UID = $User.ObjectId

        $opts = Get-AzureADGroup -SearchString "Opt" -All:$true|Where-Object { $_.displayname -like "OptIn-*" -or $_.displayname -like "OptOut-*"}

        foreach ($List in $opts) {
            try{
                Remove-AzureADGroupOwner -ObjectId $list.objectid -RefObjectId $UID
            }
            Catch{

            }
        }
        return "User $emailaddress successfully removed as owner to Opt lists."
    }

    function Add-OtpOwnership {
        [CmdletBinding()]
        param (
            $EmailAddress
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

        $opts = Get-AzureADGroup -SearchString "Opt" -All:$true|Where-Object { $_.displayname -like "OptIn-*" -or $_.displayname -like "OptOut-*"}

        foreach ($List in $opts) {
            try{
                Add-AzureADGroupOwner -ObjectId $list.objectid -RefObjectId $UID
            }
            Catch{

            }
        }
        return "User $emailaddress successfully added as owner to Opt lists."
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
    $main_form.Text = 'DL Ownership Management'
    $main_form.Width = 500
    $main_form.Height = 150
    $main_form.AutoSize = $true

    # $ListLabel = New-Object System.Windows.Forms.Label
    # $ListLabel.Text = "Distribution List:"
    # $ListLabel.Location = New-Object System.Drawing.Point(10, 5)
    # $ListLabel.Size = New-Object System.Drawing.Size(150, 20)
    # $main_form.Controls.Add($ListLabel)

    # $ListField = New-Object System.Windows.Forms.TextBox
    # $ListField.Text = ""
    # $ListField.Location = New-Object System.Drawing.Point(10, 24)
    # $ListField.Size = New-Object System.Drawing.Size(150, 20)
    # $main_form.Controls.Add($ListField)

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Text = "User Email Address:"
    $UserLabel.Location = New-Object System.Drawing.Point(10, 5)
    $UserLabel.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserLabel)

    $UserField = New-Object System.Windows.Forms.TextBox
    $UserField.Text = ""
    $UserField.Location = New-Object System.Drawing.Point(10, 24)
    $UserField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserField)

    $ReturnLabel = New-Object System.Windows.Forms.Label
    $ReturnLabel.Text = ""
    $ReturnLabel.Location = New-Object System.Drawing.Point(170, 5)
    $ReturnLabel.Size = New-Object System.Drawing.Size(350, 60)
    $ReturnLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $main_form.Controls.Add($ReturnLabel)

    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Point(10, 49)
    $AddButton.Size = New-Object System.Drawing.Size(150, 20)
    $AddButton.Text = "Add User as Owner"

    $main_form.Controls.Add($AddButton)
    $AddButton.Add_Click(
        {
            $ReturnLabel.Text = Add-OtpOwnership -EmailAddress $UserField.Text
            $ListField.Text = ""
            $UserField.Text = ""
        })

    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Point(10, 69)
    $RemoveButton.Size = New-Object System.Drawing.Size(150, 20)
    $RemoveButton.Text = "Remove User Ownership"

    $main_form.Controls.Add($RemoveButton)
    $RemoveButton.Add_Click(
        {
            $ReturnLabel.Text = Remove-OtpOwnership -EmailAddress $UserField.Text
            $ListField.Text = ""
            $UserField.Text = ""
        })

    $main_form.ShowDialog()
    
}

if ($Run) {
    New-DLOwnerManager
}