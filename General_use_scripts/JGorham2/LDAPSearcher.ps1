$Filter = "(objectCategory=person)"
$GroupFilter = "((objectCategory=group))"
$RootOU = "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$GroupRootOU = "CN=AdobeUserSync-Student,OU=Resources,OU=Groups,DC=admin,DC=wpi,DC=edu"
$scope = "Subtree"
$Groupscope = "Base"

$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Searcher.Filter = $Filter
$Searcher.PageSize = "20000"
$Searcher.SearchScope = $Scope # Either: "Base", "OneLevel" or "Subtree"
$Users = $Searcher.FindAll()

$GroupSearcher = New-Object DirectoryServices.DirectorySearcher
$GroupSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($GroupRootOU)")
$GroupSearcher.Filter = $GroupFilter
$GroupSearcher.PageSize = "20000"
$GroupSearcher.SearchScope = $GroupScope # Either: "Base", "OneLevel" or "Subtree"
$GroupUsers = $GroupSearcher.FindAll()



###############################################

$dn = 'ou=People,dc=wpi,dc=edu'
$domain = "LDAP://ldap.wpi.edu:389/$($dn)"
$useragent = 'cn=pam,ou=access,dc=wpi,dc=edu'
$userpass = '6zWn5nS7'
$LDAPSearchString = "(uid=ajwitkin)"
$auth = [System.DirectoryServices.AuthenticationTypes]::FastBind
#$auth = [System.DirectoryServices.AuthenticationTypes]::Secure
$root = New-Object -TypeName System.DirectoryServices.DirectoryEntry($domain, $useragent, $userpass, $auth)
$query = New-Object System.DirectoryServices.DirectorySearcher($root, $LDAPSearchString)
$objClass = $query.findall()
$objClass.Count
$objClass |Get-Member