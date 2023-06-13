#################################
# part 2
#################################
function Open-Notepad {
    [CmdletBinding()]
    param (
        
    )
Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='NEW WINDOW'
$main_form.Width = 600
$main_form.Height = 400
$main_form.AutoSize = $true




$Button = New-Object System.Windows.Forms.Button

$Button.Text = "notepad"

$main_form.Controls.Add($Button)



$Button.Add_Click(

{
notepad.exe

}

)



$main_form.ShowDialog()
}


##################################
# Main
##################################

Add-Type -assembly System.Windows.Forms
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='DL Subscription Management'
$main_form.Width = 600
$main_form.Height = 400
$main_form.AutoSize = $true




$Button = New-Object System.Windows.Forms.Button
$Button.Text = "DO THING"
$main_form.Controls.Add($Button)

$Button.Add_Click(
{

    Open-Notepad

}

)



$main_form.ShowDialog()




