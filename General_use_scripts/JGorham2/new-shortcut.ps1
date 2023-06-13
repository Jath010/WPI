function New-Shortcut {
    [CmdletBinding()]
    param (
        $targetFile, # The path we want the shortcut to run
        $shortcutFile           # The location and name of our shortcut
    )
    
    $WScriptShell = New-Object -ComObject WScript.shell
    $Shortcut = $WScriptShell.CreateShortcut($shortcutFile)
    $Shortcut.TargetPath = $targetFile
    $Shortcut.Save()
}

function New-InternetShortcut {
    [CmdletBinding()]
    param (
        $targetPath,
        $IconPath,
        $shortcutFile
    )

    $shell = New-Object -ComObject WScript.Shell
    #$destination = $shell.SpecialFolders.Item("AllUsersDesktop")
    $shortcutPath = $shortcutFile #Join-Path -Path $destination -ChildPath 'Test Intranet.url'
    # create the shortcut
    $shortcut = $shell.CreateShortcut($shortcutPath)
    # for a .url shortcut only set the TargetPath
    $shortcut.TargetPath = $targetPath #'https://sharepoint.com/'
    $shortcut.Save()

    if ($IconPath) {
        # next update the shortcut with a path to the icon file and the index of that icon
        # you can do that because a .url file is just a text file in INI format
        Add-Content -Path $shortcutPath -Value "IconFile=$($IconPath)" # example C:\Users\Public\Pictures\ShortcutIcon.ico
        Add-Content -Path $shortcutPath -Value "IconIndex=0"
    }

    # clean up the COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function New-Echo360Shortcut {
    [CmdletBinding()]
    param (
        
    )
    
    $shortcutTarget = $(HOSTNAME.EXE) + "-echo.echo360.wpi.edu"
    $WScriptShell = New-Object -ComObject WScript.shell
    $destination = $shell.SpecialFolders.Item("AllUsersDesktop")
    $shortcutPath = Join-Path -Path $destination -ChildPath 'Echo360.url'
    $shortcut = $WScriptShell.createShortcut($shortcutPath)
    $shortcut.TargetPath = "https://$($shortcutTarget)/"
    $shortcut.save()

    Add-Content -Path $shortcutPath -Value "IconFile=C:\Users\Public\Pictures\ShortcutIcon.ico"
    Add-Content -Path $shortcutPath -Value "IconIndex=0"

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WScriptShell) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}