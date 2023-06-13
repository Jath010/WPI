function Get-ProgramRunConfirm {
    [CmdletBinding()]
    param (
        $Runpath,
        $StartupDelay,
        $ShutdownDelay
    )
    
    $app = start-process $Runpath -PassThru
    if($null -ne $StartupDelay){
        Start-Sleep -Seconds $StartupDelay
    }
    if(get-process -Id $app.id){
        Write-Verbose "Startup successful"
        Stop-Process -Id $app.Id
    }
    else{
        Write-Verbose "Startup Failed"
        return 1
    }
    if($null -ne $ShutdownDelay){
        Start-Sleep -Seconds $ShutdownDelay
    }
    if(!(Get-process -id $app.id -ErrorAction SilentlyContinue)){
        Write-Verbose "Shutdown Successful"
    }
    else{
        Write-Verbose "Shutdown Failed"
        return 2
    }
    return 0
}