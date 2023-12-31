Clear-Host
function WinSam-Get-MainMenu {
    <#
    .SYNOPSIS
    .DESCRIPTION
    .EXAMPLE
    WinSam-Get-MainMenu AccessLevel
    .PARAMETER AccessLevel
    #>
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [ValidateLength(1,64)]
        [string]$AccessLevel
        )
    #$ErrorActionPreference = "SilentlyContinue"
$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()

sleep 5
Write-Host "Index update complete.  Total time: "$ElapsedTime.Elapsed
    # End of WinSam-Get-MainMenu
    }