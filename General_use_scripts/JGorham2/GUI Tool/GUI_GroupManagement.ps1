#########################################
#   Main
#########################################

Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = 'Group Management'
$main_form.Width = 600
$main_form.Height = 250
$main_form.AutoSize = $true

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = 'Please make a selection from the list below:'
$main_form.Controls.Add($label)

$GroupField = New-Object System.Windows.Forms.TextBox
$GroupField.Text = "Target Group"
$GroupField.Location = New-Object System.Drawing.Point(290, 20)
$GroupField.Size = New-Object System.Drawing.Size(200, 20)
$main_form.Controls.Add($GroupField)


$MemberBox = New-Object System.Windows.Forms.ListBox
$MemberBox.SelectionMode = 'MultiExtended'
$MemberBox.MultiColumn = $true
$MemberBox.Location = New-Object System.Drawing.Point(10, 40)
$MemberBox.Size = New-Object System.Drawing.Size(400, 200)



$MemberBox.Items.Clear()
$main_form.Controls.Add($MemberBox)

$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Point(410, 40)
$AddButton.Size = New-Object System.Drawing.Size(90, 20)
$AddButton.Text = "Search Azure"

$main_form.Controls.Add($AddButton)
$AddButton.Add_Click(
    { 
        $group = get-AzureADMSGroup -SearchString $GroupField.Text -ErrorAction SilentlyContinue
        Write-host $group.objectid
        $members = Get-AzureADGroupMember -ObjectId $group.id -ErrorAction SilentlyContinue
        $GroupField.Text = $group.DisplayName
        $MemberBox.BeginUpdate()
        foreach ($Member in $members) {
            $MemberBox.Items.Add($member.UserPrincipalName)
        }
        $MemberBox.EndUpdate()
    })







$main_form.ShowDialog()