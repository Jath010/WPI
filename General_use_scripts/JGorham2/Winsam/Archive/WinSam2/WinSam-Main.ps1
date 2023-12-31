##################################################################################################################
#MAIN
##################################################################################################################

#While loop for entire process to check if you want to look up another account
$anotheraccount = 'y'
$RepeatSearch = $false

while ($anotheraccount -eq 'y') {
    cls
    Write-Host '========================= WinSamaritan 3.0.0 ========================' -foregroundcolor Cyan

    $admindomain = 'admin.wpi.edu'
    If ($RepeatSearch -ne $true) {
        $UserName = ''
        }
    Else {$RepeatSearch = $false}
    $option = $null 
    $optionaction=$null                                 
    $permissionsstatus = $null
    $permissionsstatus = checkpermissions
    
    if ($permissionsstatus -ne 'None'){
        while (!$username) {
            $username = Read-Host -Prompt 'Please enter a username'
            $adminresult = isindomain $admindomain $username
            if ($adminresult -eq $null) {
                write-host ''
                Write-Host 'The account ' $username ' does not exist.' -foregroundcolor Red
                write-host ''
                $UserName = '' #Username failed so clear field and restart while loop
                }
            }
        cls
        Write-Host '========================= WinSamaritan 3.0.0 ========================' -foregroundcolor Cyan
        accountlookup $username
        write-host ''                  
        Write-Host '===================================================================='
        }
    else {write-host 'Your account ' $env:username ' does not have permissions to run WinSamaritan' -ForegroundColor Red}

    if ($permissionsstatus -eq 'Reset') {
        while ($optionaction -eq $null) {
			Write-Host ''
            Write-Host '     Menu options:'
            Write-Host '     (1) Show file share group memberships' -ForegroundColor Cyan
            Write-Host '     (2) Recheck the current user:' $username -ForegroundColor Cyan
            Write-Host '     (3) Reset Password' -ForegroundColor Yellow
            Write-Host ''
            Write-Host '     (9) Look up another account' -ForegroundColor Green
            Write-Host '     (0) Exit WinSamaritan' -ForegroundColor Red
            $option = read-Host 'Choose an option'
            Write-Host ''
            switch ($option) { 
                1 {$optionaction='FileshareMemberships'} 
                2 {$optionaction='RepeatSearch'}
                3 {$optionaction='ResetPassword'} 
                9 {$optionaction='AnotherAccount'} 
                0 {$optionaction='Exit'} 
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    elseif ($permissionsstatus -eq 'ReadOnly' -or $permissionsstatus -eq 'Unlock' ) {
        while ($optionaction -eq $null) {
			Write-Host ''
            Write-Host '     Menu options:'
            Write-Host '     (1) Show file share group memberships' -ForegroundColor Green
            Write-Host '     (2) Recheck the current user:' $username -ForegroundColor Green
            Write-Host ''
            Write-Host '     (9) Look up another account'
            Write-Host '     (0) Exit WinSamaritan' -ForegroundColor Red
            $option = read-Host 'Choose an option'
            Write-Host ''
            switch ($option) { 
                1 {$optionaction='FileshareMemberships'} 
                2 {$optionaction='RepeatSearch'}
                9 {$optionaction='AnotherAccount'} 
                0 {$optionaction='Exit'} 
                'exit' {$optionaction='Exit'}
                default {
                    $optionaction=$null
                    Write-Host 'Please specify one of the available options' -foregroundcolor Red
                    }
                }
            }
        }
    else {exit}

    $anotheraccount = ''

    switch ($optionaction) {
        'AnotherAccount' {$anotheraccount='y'}
        'RepeatSearch' {
            $anotheraccount='y' 
            $RepeatSearch=$true
            }
        'FileshareMemberships' {GetFileshareMemberships $username}
        'ResetPassword' {resetpassword $username}
        'Exit' {exit}
        }

    while ($anotheraccount -ne 'y' -and $anotheraccount -ne 'n') {
        write-host ''
        $anotheraccount = read-host 'Would you like to look up another account? (y/n)'
        write-host ''
        if ($anotheraccount -ne 'y' -and $anotheraccount -ne 'n') {write-host 'Please specify one of the available options' -ForegroundColor Red}
        }        
}