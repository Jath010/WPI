function WinSam-Menu-Main {
    Clear-Host
    ## Refresh sub-functions to update with any changes since they last ran
    . $ScriptPath\WinSam.3.GlobalVariables.ps1
    . $ScriptPath\WinSam.3.Functions.ps1
    . $ScriptPath\WinSam.3.Info.ps1
    . $ScriptPath\WinSam.3.Menu.ps1
    $username=$null
    Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host ''
    Write-Host (WinSam-Write-Header 'Main Menu' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
    Write-Host (WinSam-Write-Header '' $MenuLength -Line)
    Write-Host '     (1) User Information'
    Write-Host '     (2) Mail Resource (Look-up email address)'
    Write-Host '     (3) Groups: AD Security Groups'
    Write-Host '     (4) Computer Information'
    Write-Host '     (5) Reset PIN'
    Write-Host '     (6) Search User by ID Number'
    if ($BetaAccess) {Write-Host '';Write-Host '     (B) Run the Beta Version' -ForegroundColor Cyan}
    Write-Host ''
    Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
    Write-Host ''
    WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
    Write-Host ''

    $option = $null
    while ($option -eq $null) {
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {WinSam-User} 
            2 {WinSam-MailResource}
            3 {WinSam-ADGroup} 
            4 {WinSam-PC}
            5 {WinSam-PIN}
            6 {WinSam-SearchByID}
            'm' {WinSam-Menu-Main}
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

    if ($UserAccessLevel -eq 'PasswordReset_lvl2' -or $UserAccessLevel -eq 'PasswordReset' -or $UserAccessLevel -eq 'SysAdmin') {
        while ($optionaction -eq $null) {
            Write-Host '     (1) Look up another account'
            Write-Host '     (2) Recheck the current user:' $username
            Write-Host '     (3) Show file share group memberships'
            Write-Host "     (4) Reset PIN for $username"
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
            Write-Host "     (4) Reset PIN for $username"
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
        'ResetPassword' {WinSam-Menu-User-PasswordReset}
        'ResetPIN' {WinSam-Reset-PIN $username}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }


function WinSam-Menu-User-PasswordReset {
    $option = $null
    while ($option -eq $null) {
        Clear-Host
        Write-Host (WinSam-Write-Header $HeaderTitle $MenuLength -Center) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host ''
        Write-Host (WinSam-Write-Header 'General User Information' $MenuLength $MenuIndent) -ForegroundColor $Global:bcolor -BackgroundColor $Global:fcolor
        Write-Host (WinSam-Write-Header '' $MenuLength -Line)
        WinSam-Get-InfoBanner
        Write-Host 'Name                :' $Global:ADInfo.DisplayName
        Write-Host 'Email               :' $Global:ADInfo.UserPrincipalName
        Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Set New Temporary Password (requires change at login)'
        if ($UserAccessLevel -eq 'PasswordReset_lvl2' -or $UserAccessLevel -eq 'SysAdmin') {
            Write-Host '     (2) Set New Permanent Password (will not require a change)'
            Write-Host '     (3) Set Random Temporary Password (requires change at at login)'
            }
        Write-Host ''
        Write-Host '     (U) User Information'
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        
        $option = read-Host 'Choose an option'
        Write-Host ''
        if ($UserAccessLevel -eq 'PasswordReset_lvl2' -or $UserAccessLevel -eq 'SysAdmin') {
            switch ($option) { 
                1 {WinSam-Reset-Password 'ResetManualTemp'} 
                2 {WinSam-Reset-Password 'ResetManualPerm'}
                3 {WinSam-Reset-Password 'ResetRandomTemp'}
                'U' {WinSam-Menu-User} 
                'M' {WinSam-Menu-Main} 
                'Q' {exit}
                'exit' {exit}
                default {
                    $option=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        else {
            switch ($option) { 
                1 {WinSam-Reset-Password 'ResetManualTemp'} 
                'U' {WinSam-Menu-User} 
                'M' {WinSam-Menu-Main} 
                'Q' {exit}
                'exit' {exit}
                default {
                    $option=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    }


function WinSam-Menu-PIN {
    $optionaction = $null
    while ($optionaction -eq $null) {
        Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Reset another PIN'
        if ($username) {Write-Host "     (2) Look up User Information for $username"}
        else {Write-Host "     (2) Look up User Information"}
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {$optionaction='PINMenu'} 
            2 {
                if ($username) {$optionaction='UserSearchUsername'}
                else {$optionaction='UserSearch'}
                }
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
        'PINMenu' {
            $username=$null
            WinSam-PIN
            }
        'UserSearchUsername' {WinSam-User $username}
        'UserSearch' {WinSam-User}
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


function WinSam-Menu-MailResource {
    $optionaction = $null
    while ($optionaction -eq $null) {
        Write-Host ''
	    Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Look up another mail resource'
        Write-Host '     (2) Recheck the current resource:' $EmailAddress
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Green
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {$optionaction='AnotherResource'} 
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
        'AnotherResource' {WinSam-MailResource}
        'RepeatSearch' {WinSam-MailResource $EmailAddress}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }


function WinSam-Menu-ADGroup {
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
        'AnotherGroup' {WinSam-ADGroup}
        'RepeatSearch' {WinSam-ADGroup $GroupName}
        'MainMenu' {WinSam-Menu-Main}
        'Exit' {exit}
        }
    }


function WinSam-Menu-SearchByID ($username) {
    $option = $null
    while ($option -eq $null) {
        Write-Host ''
        Write-Host ''
        Write-Host (WinSam-Write-Header 'Menu Options' $MenuLength $MenuIndent) -ForegroundColor Black -BackgroundColor Gray
        Write-Host '     (1) Start New Search by ID'
        if ($username) {Write-Host "     (2) Lookup User Information for [$username]"}
        Write-Host ''
        Write-Host '     (M) Main Menu' -ForegroundColor Yellow
        Write-Host '     (Q) Quit WinSamaritan' -ForegroundColor Red
        Write-Host ''
        WinSam-Print-DebugInfo ##Prints DEBUG information when $debug is $true
        Write-Host ''

        $option = read-Host 'Choose an option'
        Write-Host ''
        switch ($option) { 
            1 {WinSam-SearchByID} 
            2 {
                if ($username) {WinSam-User $username}
                else {$option = $null;Write-Host 'Please specify one of the available options' -foregroundcolor Red}
                } 
            3 {WinSam-SearchByID 'Alumni'} 
            4 {WinSam-SearchByID 'All'} 
            'M' {WinSam-Menu-Main} 
            'Q' {exit} 
            'exit' {exit}
            default {
                $option = $null
                Write-Host 'Please specify one of the available options' -foregroundcolor Red
                }
            }
        }
    }
