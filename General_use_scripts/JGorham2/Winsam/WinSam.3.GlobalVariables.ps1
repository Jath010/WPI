##################################################################################################################
#Powershell Command Shell Customizations
##################################################################################################################
$Global:MenuLength = 90
$Global:MenuIndent = 5
$Global:WinSamVersion = '3.5.0 Beta 2 - Dev1'
$Global:HeaderTitle = "Windows Samaritan $WinSamVersion"

If ($Host.Name -ne "Windows PowerShell ISE Host"){
    $h = (Get-Host).UI.RawUI
    $h.WindowTitle = $HeaderTitle
    $h.BackgroundColor = "Black"
    $h.ForegroundColor = "White"
    $buffer = $h.BufferSize
    $buffer.Height = 9999
    $buffer.Width = $MenuLength+1
    $win = $h.WindowSize
    if ($win.Height -le 50) {$win.Height = 50}
    $win.Width = $MenuLength+1

    if ($MenuLength -lt $h.BufferSize.Width) {
        $h.Set_WindowSize($win)
        $h.Set_BufferSize($buffer)
        }
    Else {
        $h.Set_BufferSize($buffer)
        $h.Set_WindowSize($win)
        }
        
    $MaxHistoryCount = 1000
    }    

##################################################################################################################
# Global Variables
##################################################################################################################
$global:ADInfo = $null
$global:GroupInfo = $null
$Global:GroupMembers = $null
$Global:GroupMemberOf = $null
$global:mailbox = $null
$Global:BetaAccess = $false
$Global:BetaAccessList = $null
$Global:Beta = $null

#Domain Controller Information
$global:dcs = (Get-ADDomainController -Filter *)
$global:dc = $dcs | Where {$_.OperationMasterRoles -like "*RIDMaster*"}
$global:dchostname = $dc.HostName

#Environment Variables
$global:today = Get-Date
$global:currentuser = $env:username
$global:localhost = $env:COMPUTERNAME

#Console Color Management
If ($Host.Name -eq "Windows PowerShell ISE Host"){
    $global:fcolor = "White"
    $global:bcolor = "DarkBlue"
    }
Else {
    $global:fcolor = $host.UI.RawUI.ForegroundColor
    $global:bcolor = $host.UI.RawUI.BackgroundColor
    }

#Testing Code
$Global:Debug = $false          #Set value to $true to enable debugging code
$Global:Beta  = $false          #Set value to $true to enable beta version testing
if ($ScriptPath -match "berlin") {$Global:BetaPath = "\\berlin.wpi.edu\c$\wpi\Scripts\WinSam3B"}
else {$Global:BetaPath = "c:\wpi\Scripts\WinSam3B"}

$Global:BetaAccessList  = 'tcollins','s_tcollins','tcollins2.test','winsam.student','winsam.staff'
$Global:BetaAccessList += 'kwheeler','rcriley'
#Get-ADGroupMember hdstaff -Recursive | %{$Global:BetaAccessList += $_.SamAccountName}
#Get-ADGroupMember U_Helpdesk_Student_Staff -Recursive | %{$Global:BetaAccessList += $_.SamAccountName}

if ($Global:Beta -and ($Global:BetaAccessList | Where {$_ -eq $currentuser})) {$Global:BetaAccess = $true}

<## DEBUG ##############
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
#######################>
