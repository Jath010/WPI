##Creating a script to generate department groups

$Departments = get-aduser -LDAPFilter "(Department=*)" -Properties Department|Select-Object department|Sort-Object -Property Department|Get-Unique -AsString

$Departments.Count
