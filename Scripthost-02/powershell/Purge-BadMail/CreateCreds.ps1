$creds = get-credential #admin\exch_automation
$creds | Export-clixml -path .\exch_automation