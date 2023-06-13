#This takes in a list of Machine names such as the kind exported from the remediation page in security.microsoft.com
#It then spits out a file that can be input to a azad group bulk import

$inputFile = "C:\tmp\SecurityRemediation\VBStargets.csv"
$outputFile = "C:\tmp\SecurityRemediation\SanitizedVBStargets.csv"

$targets = Import-Csv $inputFile
$myArray = @()

foreach($machine in $targets){
    $target = $machine.machines.Split(".")[0]
    Write-Progress -Activity $target
    $myArray += (Get-AzureADDevice -SearchString $target | Where-Object {$_.DeviceTrustType -eq "ServerAd"}).objectid
}

"version:v1.0"|Out-file $outputFile
"Member object ID or user principal name [memberObjectIdOrUpn] Required"|Out-file $outputFile -Append
$myArray | Out-File $outputFile -Append