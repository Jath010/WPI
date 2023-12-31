# Reference: http://mspowershell.blogspot.com/2009/02/cli-menu-in-powershell.html

function DrawMenu {
    ## supportfunction to the Menu function below
    param ($menuItems, $menuPosition, $menuTitel)
    If ($Host.Name -eq "Windows PowerShell ISE Host"){
        $fcolor = "White"
        $bcolor = "DarkBlue"
        }
    Else {
        $fcolor = $host.UI.RawUI.ForegroundColor
        $bcolor = $host.UI.RawUI.BackgroundColor
        }
    $l = $menuItems.length + 1
    cls
    $menuwidth = $menuTitel.length + 4
    Write-Host "`t" -NoNewLine
    Write-Host ("*" * $menuwidth) -ForegroundColor $fcolor -BackgroundColor $bcolor
    Write-Host "`t" -NoNewLine
    Write-Host "* $menuTitel *" -ForegroundColor $fcolor -BackgroundColor $bcolor
    Write-Host "`t" -NoNewLine
    Write-Host ("*" * $menuwidth) -ForegroundColor $fcolor -BackgroundColor $bcolor
    Write-Host ""
    Write-debug "L: $l MenuItems: $menuItems MenuPosition: $menuposition"
    for ($i = 0; $i -le $l;$i++) {
        Write-Host "`t" -NoNewLine
        if ($i -eq $menuPosition) {
            Write-Host "$($menuItems[$i])" -ForegroundColor $bcolor -back $fcolor
        } else {
            Write-Host "$($menuItems[$i])" -ForegroundColor $fcolor -BackgroundColor $bcolor
        }
    }
}

function Menu {
    ## Generate a small "DOS-like" menu.
    ## Choose a menuitem using up and down arrows, select by pressing ENTER
    param ([array]$menuItems, $menuTitel = "MENU")
    $vkeycode = 0
    $pos = 0
    DrawMenu $menuItems $pos $menuTitel
    While ($vkeycode -ne 13) {
        $press = $host.ui.rawui.readkey("NoEcho,IncludeKeyDown")
        $vkeycode = $press.virtualkeycode
        Write-host "$($press.character)" -NoNewLine
        If ($vkeycode -eq 38) {$pos--}
        If ($vkeycode -eq 40) {$pos++}
        if ($pos -lt 0) {$pos = $menuItems.length -1}
        if ($pos -ge $menuItems.length) {$pos = 0}
        DrawMenu $menuItems $pos $menuTitel
    }
    Write-Output $($menuItems[$pos])
}


Clear-Host
$bad = "Format c:","Send spam to boss","Truncate database *","Randomize user password","Download dilbert","Hack local AD"
#$selection = Menu $bad "WHAT DO YOU WANNA DO?"
#Write-Host "YOU SELECTED : $selection ... DONE!`n" 
