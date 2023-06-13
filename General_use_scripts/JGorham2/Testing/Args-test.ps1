function Set-Args {
    [CmdletBinding()]
    param (
        $args
    )
    return $args
}

function PutHostsDeviceCollection {
    # feed names or file path to hostnames variable then upload them into sccm collection
    Param([Parameter(Mandatory=$true, Position=0)]$collectname, [Parameter(Mandatory=$true, Position=1, ValueFromRemainingArguments)]$hostnames)
    foreach ($item in $hostnames) {
        if ($item -like '*.txt' -or $item -like '*.csv') {
            $item = Get-Content $item
        }
        foreach($name in $item){
            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $collectname -ResourceID (Get-CMDevice -Name $name).ResourceID
        }
    }   
}