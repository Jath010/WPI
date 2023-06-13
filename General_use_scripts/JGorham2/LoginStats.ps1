
function Get-LogOnStats {
  param (
    $computer,
    $user
  )
  $filter = @"

    *[System[(EventID='4624')]
    and
    EventData[Data[@Name='LogonType'] and (Data='2' or Data='3')]
    and
    EventData[Data[@Name='TargetUserName'] and (Data='$($user)')]
    ]
"@

  $xmltext = Get-WinEvent -LogName Security -FilterXPath $filter -ComputerName $computer |
  ForEach-Object { $_.ToXML() }
  $xml = [xml]"<Events>$xmltext</Events>"
  $xml.Events.event[0].Eventdata.Data
}

function Get-LogOffStats {
  param (
    $computer,
    $user
  )
  
$filter = @"

    *[System[(EventID='4634')]
    and
    EventData[Data[@Name='LogonType'] and (Data='2' or Data='3')]
    and
    EventData[Data[@Name='TargetUserName'] and (Data='$($user)')]
    ]
"@

  $xmltext = Get-WinEvent -LogName Security -FilterXPath $filter -ComputerName $computer |
  ForEach-Object { $_.ToXML() }
  $xml = [xml]"<Events>$xmltext</Events>"
  $xml.Events.event[0].Eventdata.Data
}


# event log XPath queries.

$xmlquery = @'
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
     *[System[(EventID='4624')]
     and
     EventData[Data[@Name='LogonType'] and (Data='2' or Data='3')]
     ] 
    </Select>
  </Query>
</QueryList>
'@

$xmlquery = @'
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
     *[System[(EventID='4624') and TimeCreated[timediff(@SystemTime) &lt;= 2592000000]]
     and
     EventData[Data[@Name='LogonType'] and (Data='2' or Data='3')]] 
    </Select>
  </Query>
</QueryList>
'@
Get-WinEvent -FilterXml $xmlquery -MaxEvents 10

# Get event data as XML
#build a filter
$filter = @'
     *[System[(EventID='4624')]
     and
     EventData[Data[@Name='LogonType'] and (Data='2' or Data='3')]
     ]
'@

$xmltext = Get-WinEvent -LogName Security -FilterXPath $filter | % { $_.ToXML() }
$xml = [xml]"<Events>$xmltext</Events>"

# you now have the XML for all selected events.
$xml.Events.Event


$filter = @'
     *[System[(EventID='4624')]
     and
     EventData[Data[@Name='LogonType'] and (Data='2')]
     ]
'@

$xmltext = Get-WinEvent -LogName Security -FilterXPath $filter -MaxEvents 10 -ComputerName woodglue |
ForEach-Object { $_.ToXML() }
$xml = [xml]"<Events>$xmltext</Events>"
$xml.Events.event[0].Eventdata.Data

$filterXML = @'
<QueryList>
  <Query Id="0" Path="System">
    <Select Path="System">
		*[System[Provider[@Name='Microsoft-Windows-Kernel-General']
			and (Level=4 or Level=0) 
			and (EventID=12)]]
	</Select>
  </Query>
</QueryList>
'@
Get-WinEvent -ComputerName woodglue -MaxEvents 1 -FilterXml $filterXML