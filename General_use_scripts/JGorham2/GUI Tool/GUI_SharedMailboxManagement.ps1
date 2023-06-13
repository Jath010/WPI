param([switch] $Run)
function New-ShareMailboxManager {
    #################################
    # Connect to Azure
    #################################

    Connect-ExchangeOnline -ShowBanner:$false

    function Add-MailboxAccess_sub {
        [CmdletBinding()]
        param (
            $User,
            $Mailbox
        )
        Add-MailboxPermission -Identity $mailbox -User $User -AccessRights fullaccess -InheritanceType all -AutoMapping $False
        Add-RecipientPermission -Identity $mailbox -Trustee $User -AccessRights SendAs -Confirm:$false
        Return (Get-UserMailboxAccess_sub -User $User -Mailbox $mailbox)
            
    }
    function Remove-MailboxAccess_sub {
        [CmdletBinding()]
        param (
            $User,
            $Mailbox
        )
        Remove-MailboxPermission -Identity $mailbox -User $User -AccessRights fullaccess -InheritanceType all -Confirm:$False
        Remove-RecipientPermission -Identity $mailbox -Trustee $User -AccessRights SendAs -Confirm:$false
        Return (Get-UserMailboxAccess_sub -User $User -Mailbox $mailbox)
            
    }

    function Get-UserMailboxAccess_sub {
        [CmdletBinding()]
        param (
            [string]
            $User,
            [string]
            $Mailbox
        )
        
        
        $Access = get-MailboxPermission -Identity $mailbox | Where-Object { $_.user -like $user }
                
        $SendAs = get-RecipientPermission -Identity $mailbox | Where-Object { $_.Trustee -like $user }
        if ($null -ne $Access -and $null -ne $SendAs) {
            Return "User $User has Full Access and Send As to $mailbox"
        }
        elseif ($null -eq $Access -and $null -ne $SendAs) {
            Write-host "User $User does NOT have Full Access to $mailbox BUT HAS Send Ad"
        }
        elseif ($null -eq $Access -and $null -eq $SendAs) {
            Write-host "User $User does NOT have Send As to $mailbox BUT HAS Full Access"
        }
        else {
            Write-host "User $User does NOT have Send As OR Full Access to $mailbox"
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
    $main_form.Text = 'Shared Mailbox Management'
    $main_form.Width = 500
    $main_form.Height = 150
    $main_form.AutoSize = $true

    $ListField = New-Object System.Windows.Forms.TextBox
    $ListField.Text = "Mailbox Address"
    $ListField.Location = New-Object System.Drawing.Point(10, 5)
    $ListField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($ListField)

    $UserField = New-Object System.Windows.Forms.TextBox
    $UserField.Text = "User Email Address"
    $UserField.Location = New-Object System.Drawing.Point(10, 30)
    $UserField.Size = New-Object System.Drawing.Size(150, 20)
    $main_form.Controls.Add($UserField)

    $ReturnLabel = New-Object System.Windows.Forms.Label
    $ReturnLabel.Text = ""
    $ReturnLabel.Location = New-Object System.Drawing.Point(170, 5)
    $ReturnLabel.Size = New-Object System.Drawing.Size(350, 60)
    $ReturnLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $main_form.Controls.Add($ReturnLabel)

    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Point(10, 55)
    $AddButton.Size = New-Object System.Drawing.Size(150, 20)
    $AddButton.Text = "Add User"

    $main_form.Controls.Add($AddButton)
    $AddButton.Add_Click(
        {
            $ReturnLabel.Text = Add-MailboxAccess_sub -Mailbox $ListField.Text -user $UserField.Text
            $ListField.Text = "Mailbox Address"
            $UserField.Text = "User Email Address"
        })

    $RemoveButton = New-Object System.Windows.Forms.Button
    $RemoveButton.Location = New-Object System.Drawing.Point(10, 80)
    $RemoveButton.Size = New-Object System.Drawing.Size(150, 20)
    $RemoveButton.Text = "Remove User"

    $main_form.Controls.Add($RemoveButton)
    $RemoveButton.Add_Click(
        {
            $ReturnLabel.Text = Remove-MailboxAccess_sub -Mailbox $ListField.Text -user $UserField.Text
            $ListField.Text = "Mailbox Address"
            $UserField.Text = "User Email Address"
        })

    $main_form.ShowDialog()
    
}

if ($Run) {
    New-ShareMailboxManager
}