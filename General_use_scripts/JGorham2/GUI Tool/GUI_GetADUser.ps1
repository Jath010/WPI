param([switch] $Run)
function New-ADLookup {

    function Get-WPIADUser_sub {
        [CmdletBinding()]
        param (
            $User
        )
        
        begin {
            $User = $user.split("@")[0]
        }
        
        process {
            try {
                $UserData = get-aduser $user -Properties * -ErrorAction SilentlyContinue
            }
            catch {
                $DataOut = @{DisplayName    = "Invalid User"}
                Return $DataOut
            }
            
            $DataOut = @{
                DisplayName                 = $UserData.DisplayName
                Email                       = $UserData.mail
                Year                        = $UserData.extensionattribute3
                "Primary Affiliation"       = $UserData.extensionAttribute7
                "All Affiliations"          = $UserData.extensionAttribute8
                Major                       = ($UserData.extensionAttribute4 -match "MJ-(.*);MN-") | ForEach-Object { $Matches[1] }
                Minor                       = ($UserData.extensionAttribute4 -match "MJ-.*;MN-(.*)") | ForEach-Object { $Matches[1] }
                "Advisors"                  = $UserData.extensionAttribute6
                "Student Code"              = $UserData.extensionattribute2
                "Employee Code"             = $UserData.extensionattribute1
            }
            Return $DataOut
        }
        
        end {
            
        }
    }

    function Add-ADFormContents {
        param (
        )
        $UserView.Items.Clear()
        $UserData = Get-WPIADUser_sub -user $UserField.Text
        $UserData | clip.exe
        foreach ($Key in $UserData.Keys) {
            $value = $UserData.$Key
            $ListViewItem = New-Object System.Windows.Forms.ListViewItem([System.String[]](@($key, $Value)), -1)
            $ListViewItem.StateImageIndex = 0
            $UserView.Items.AddRange([System.Windows.Forms.ListViewItem[]](@($ListViewItem)))
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
    $main_form.Text = 'User Lookup'
    $main_form.Size = New-Object System.Drawing.Size(500, 250)
    $main_form.MinimumSize = New-Object System.Drawing.Size(500, 250)
    $main_form.MaximumSize = New-Object System.Drawing.Size(500, 250)
    $main_form.AutoSize = $true

    $UserField = New-Object System.Windows.Forms.TextBox
    $UserField.Text = ""
    $UserField.Location = New-Object System.Drawing.Point(160, 5)
    $UserField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserField)

    $UserField.Add_KeyDown({
            if ($_.KeyCode -eq "Enter") {
                Add-ADFormContents
            }
        })

    $UserLabel = New-Object System.Windows.Forms.Label
    $UserLabel.Text = "Enter Username:"
    $UserLabel.Location = New-Object System.Drawing.Point(10, 5)
    $UserLabel.Size = New-Object System.Drawing.Size(160, 20)
    $UserLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $main_form.Controls.Add($UserLabel)

    $LookupButton = New-Object System.Windows.Forms.Button
    $LookupButton.Location = New-Object System.Drawing.Point(320, 5)
    $LookupButton.Size = New-Object System.Drawing.Size(150, 20)
    $LookupButton.Text = "Lookup"

    $main_form.Controls.Add($LookupButton)
    $LookupButton.Add_Click({
            Add-ADFormContents
        })

    $UserView = New-Object System.Windows.Forms.ListView
    $UserView.Location = New-Object System.Drawing.Point(10, 30)
    $UserView.Size = New-Object System.Drawing.Size(460, 163)
    $UserView.View = [System.Windows.Forms.View]::Details
    $UserView.FullRowSelect = $true
    $Userview.MultiSelect = $true

    $LVcol1 = New-Object System.Windows.Forms.ColumnHeader
    $LVcol1.Text = "Property"
    $LVcol1.Width = 90
    $LVcol2 = New-Object System.Windows.Forms.ColumnHeader
    $LVcol2.Text = "Value"
    $LVcol2.Width = -2

    $UserView.Columns.AddRange([System.Windows.Forms.ColumnHeader[]](@($LVcol1, $LVcol2)))

    # $UserView.Add_KeyDown({
    #     if($_.KeyCode -eq "Enter"){
    #         $LookupButton.Text = $UserView.SelectedItems.Text
    #         #Set-Clipboard -Value $UserView.SelectedItems.Value.Text
    #     }
    # })

    $UserView.AutoArrange = $true
    $UserView.GridLines = $true
    $main_form.Controls.Add($UserView)

    $main_form.ShowDialog()
    
}

if ($Run) {
    New-ADLookup
}