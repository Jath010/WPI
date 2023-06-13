function WinSam-Menu-Main {
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1


    $option = $null
    while ($option -eq $null) {
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Main Menu' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        Write-Host '     (1) User Information'
        Write-Host '     (2) Group Information'
        Write-Host '     (3) Mailbox Information'
        Write-Host '     (4) Computer Information'
        if ($BetaAccess) {Write-Host '';Write-Host '     (B) Run the Beta Version'}
        Write-Host ''
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        Write-Host ''
        WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
        Write-Host ''

        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {WinSam-User} 
            2 {WinSam-Group} 
            3 {WinSam-Mailbox} 
            4 {WinSam-PC}
            'b' {if ($BetaAccess) {. $BetaPath\WinSam.3.ps1}}
            'Q' {exit} 
            'exit' {exit}
            default {
                $option = $null
                Write-Host 'Please specify one of the available options' -foregroundcolor Red
                }
            }
        }
    }

function WinSam-Menu-User {
    $optionaction = $null
		Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray

    if ($UserAccessLevel -eq 'PasswordReset' -or $UserAccessLevel -eq 'SysAdmin') {
        while ($optionaction -eq $null) {
            Write-Host '     (1) Look up another account'
            Write-Host '     (2) Recheck the current user:' $username
            Write-Host '     (3) Show file share group memberships'
            Write-Host '     (4) Reset PIN'
            Write-Host '     (5) Reset Password'
            Write-Host ''
            Write-Host '     (M) Main Menu' -ForegroundColor Yellow
            Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
            $option = read-Host 'Choose an option'
            Write-Host ''
            switch ($option) { 
                1 {$optionaction='AnotherAccount'} 
                2 {$optionaction='RepeatSearch'}
                3 {$optionaction='FileshareMemberships'} 
                4 {$optionaction='ResetPIN'}
                5 {$optionaction='ResetPassword'} 
                'Q' {$optionaction='Exit'}
                'M' {$optionaction='MainMenu'}
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    elseif ($UserAccessLevel -eq 'Unlock' ) {
        while ($optionaction -eq $null) {
            Write-Host '     (1) Look up another account'
            Write-Host '     (2) Recheck the current user:' $username
            Write-Host '     (3) Show file share group memberships'
            Write-Host '     (4) Reset PIN'
            Write-Host ''
            Write-Host '     (M) Main Menu' -ForegroundColor Yellow
            Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
            $option = read-Host 'Choose an option'
            Write-Host ''
            switch ($option) { 
                1 {$optionaction='AnotherAccount'} 
                2 {$optionaction='RepeatSearch'}
                3 {$optionaction='FileshareMemberships'} 
                4 {$optionaction='ResetPIN'}
                'M' {$optionaction='MainMenu'} 
                'Q' {$optionaction='Exit'}
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    elseif ($UserAccessLevel -eq 'ReadOnly') {
        while ($optionaction -eq $null) {
            Write-Host '     (1) Look up another account'
            Write-Host '     (2) Recheck the current user:' $username
            Write-Host '     (3) Show file share group memberships'
            Write-Host ''
            Write-Host '     (M) Main Menu' -ForegroundColor Yellow
            Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
            $option = read-Host 'Choose an option'
            Write-Host ''
            switch ($option) { 
                1 {$optionaction='AnotherAccount'} 
                2 {$optionaction='RepeatSearch'}
                3 {$optionaction='FileshareMemberships'} 
                'M' {$optionaction='MainMenu'} 
                'Q' {$optionaction='Exit'}
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }    
    else {
        Write-Host "Your access level of '$UserAccessLevel' has caused an error.  Please contact an administrator" -ForegroundColor Red
        Write-Host "Current User: $currentuser" -ForegroundColor Red
        Write-Host "Press any key to continue ..."
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
        }

    switch ($optionaction) {
        'AnotherAccount' {WinSam-User}
        'RepeatSearch' {WinSam-User $username}
        'FileshareMemberships' {WinSam-Get-ADGroups}
        'ResetPassword' {WinSam-Reset-Password}
        'ResetPIN' {WinSam-Reset-PIN}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }

function WinSam-Menu-PC {
    $optionaction = $null
    while ($optionaction -eq $null) {
        Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Look up another computer'
        Write-Host '     (2) Recheck the current computer:' $hostname
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {$optionaction='AnotherComputer'} 
            2 {$optionaction='RepeatSearch'}
            'M' {$optionaction='MainMenu'} 
            'Q' {$optionaction='Exit'}
            'exit' {$optionaction='Exit'}
            default {
                $optionaction=$null
                Write-Host 'Please specify one of the available options' -foregroundcolor Red
                }
            }
        }
    switch ($optionaction) {
        'AnotherComputer' {WinSam-PC}
        'RepeatSearch' {WinSam-PC $hostname}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }


function WinSam-Menu-Mailbox {
    $optionaction = $null
    while ($optionaction -eq $null) {
        Write-Host ''
	    Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Look up another mailbox'
        Write-Host '     (2) Recheck the current mailbox:' $alias
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {$optionaction='AnotherMailbox'} 
            2 {$optionaction='RepeatSearch'}
            'M' {$optionaction='MainMenu'} 
            'Q' {$optionaction='Exit'}
            'exit' {$optionaction='Exit'}
            default {
                $optionaction=$null
                Write-Host 'Please specify one of the available options' -foregroundcolor Red
                }
            }
        }
    switch ($optionaction) {
        'AnotherMailbox' {WinSam-Mailbox}
        'RepeatSearch' {WinSam-Mailbox $alias}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }

function WinSam-Menu-Group {
    $optionaction = $null
    while ($optionaction -eq $null) {
        Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Look up another group'
        Write-Host '     (2) Recheck the current group:' $GroupName
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {$optionaction='AnotherGroup'} 
            2 {$optionaction='RepeatSearch'}
            'M' {$optionaction='MainMenu'} 
            'Q' {$optionaction='Exit'}
            'exit' {$optionaction='Exit'}
            default {
                $optionaction=$null
                Write-Host 'Please specify one of the available options' -foregroundcolor Red
                }
            }
        }
    switch ($optionaction) {
        'AnotherGroup' {WinSam-Group}
        'RepeatSearch' {WinSam-Group $GroupName}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }