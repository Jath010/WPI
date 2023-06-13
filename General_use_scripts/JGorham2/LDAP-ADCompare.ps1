##Created 7/26/2019
##Creator: Joshua Gorham

#$ADUserList = $null
#$ADUserList = Get-ADUser -Filter * -SearchBase "OU=Accounts,DC=admin,DC=wpi,DC=edu"

<# Stackoverflow block
$Filter = "((mailNickname=id*)(whenChanged>=20170701000000.0Z))(|(userAccountControl=514)(userAccountControl=66050))(|(memberof=CN=VPN,OU=VpnAccess,OU=Domain Global,OU=Groups,OU=01,DC=em,DC=pl,DC=ad,DC=mnl)(memberof=CN=VPN-2,OU=VpnAccess,OU=Domain Global,OU=Groups,OU=01,DC=em,DC=pl,DC=ad,DC=mnl))"
$RootOU = "OU=01,DC=em,DC=pl,DC=ad,DC=mnl"

$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Searcher.Filter = $Filter
$Searcher.SearchScope = $Scope # Either: "Base", "OneLevel" or "Subtree"
$Searcher.FindAll()

##############################################
From my last AD Searcher

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
#>

#intent to grab Student, Employee, LOA, Disabled

$Users = $null
$Filter = "(objectCategory=person)"
$RootOU = "OU=Students,OU=Accounts,DC=admin,DC=wpi,DC=edu"
#$RootOU = "ldap.wpi.edu" #Try to link to unix ldap
$scope = "Subtree"

$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Searcher.Filter = $Filter
$Searcher.PageSize = "20000"
$Searcher.SearchScope = $Scope # Either: "Base", "OneLevel" or "Subtree"
$Users += $Searcher.FindAll() # this seems to run every time that you hit it. so load it into another? Looks like you can load with a +=

$RootOU = "OU=Disabled,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Users += $Searcher.FindAll() # this seems to run every time that you hit it. so load it into another? Looks like you can load with a +=

$RootOU = "OU=Leave of Absence,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Users += $Searcher.FindAll() # this seems to run every time that you hit it. so load it into another? Looks like you can load with a +=

$RootOU = "OU=Employees,OU=Accounts,DC=admin,DC=wpi,DC=edu"
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($RootOU)")
$Users += $Searcher.FindAll() # this seems to run every time that you hit it. so load it into another? Looks like you can load with a +=


<#
[System.Collections.ArrayList]$ParsedUsers = @()
$Users | ForEach-Object { $obj = [PSCustomObject]@{

    #GivenName = $_.Properties.Item('givenName')
    Mail = ($_.Properties.Item('mail')).toString()
    }
    $ParsedUsers.Add($obj)|out-null
}
#>
$SortedADUsers = @()
$Users | ForEach-Object {$SortedADUsers += $_.Properties.Item('mail')}
$SortedADUsers = $SortedADUsers|Sort-Object

###############################################
#This Works For LDAP, but has a size limit issue
########################################
<#
[System.Reflection.assembly]::LoadWithPartialName("system.directoryservices.protocols") | Out-Null
$LDAPDirectoryService = 'ldap.wpi.edu:636'
$DomainDN = 'OU=People,dc=wpi,dc=edu'
$LDAPFilter = '(&(eduPersonAffiliation=staff))'


$null = [System.Reflection.Assembly]::LoadWithPartialName('System.DirectoryServices.Protocols')
$null = [System.Reflection.Assembly]::LoadWithPartialName('System.Net')
$LDAPServer = New-Object System.DirectoryServices.Protocols.LdapConnection $LDAPDirectoryService
$LDAPServer.AuthType = [System.DirectoryServices.Protocols.AuthType]::Basic
#$creds = Get-Credential
$LDAPServer.Bind()
#$LDAPServer.Credential.Password = 
$LDAPServer.SessionOptions.ProtocolVersion = 3
$LDAPServer.SessionOptions.SecureSocketLayer =$false
 
$Scope = [System.DirectoryServices.Protocols.SearchScope]::Subtree
$AttributeList = @('mail')

$SearchRequest = New-Object System.DirectoryServices.Protocols.SearchRequest -ArgumentList $DomainDN,$LDAPFilter,$Scope,$AttributeList
$prc = new-object System.DirectoryServices.Protocols.PageResultRequestControl(200);
$soc = New-Object System.DirectoryServices.Protocols.SearchOptionsControl([System.DirectoryServices.Protocols.SearchOption]::DomainScope)
$SearchRequest.Controls.Add($prc)
$SearchRequest.Controls.Add($soc)

[System.DirectoryServices.Protocols.SearchResponse]$groups = $LDAPServer.SendRequest($SearchRequest)
[int]$count = 0
while ($true)
{
    $searchResponse = $LDAPServer.SendRequest($searchRequest)
    $pageResponse = $searchResponse.Controls[0]
    $count = $count + $searchResponse.entries.count
    # display the entries within this page.
    foreach($entry in $searchResponse.entries){$entry.DistinguishedName}
    # Check if there are more pages.
    if ($pageResponse.Cookie.Length -eq 0){write-Host $count;break}
    $pagedRequest.Cookie = $pageResponse.Cookie
}


foreach ($group in $groups.Entries) 
{
  $users=$group.attributes['mail'].GetValues('string')
  foreach ($user in $users) {
    Write-Host $user
  }
}
#>
##############################################################
#Quickbind Attempts
##############################################################

$objClass = $null
$dn = 'ou=People,dc=wpi,dc=edu'
$domain = "LDAP://ldap.wpi.edu:389/$($dn)"
$useragent = 'cn=pam,ou=access,dc=wpi,dc=edu'
$userpass = '6zWn5nS7'
$LDAPSearchString = "(objectclass=Person)"
$auth = [System.DirectoryServices.AuthenticationTypes]::FastBind
#$auth = [System.DirectoryServices.AuthenticationTypes]::Secure
$root = New-Object -TypeName System.DirectoryServices.DirectoryEntry($domain, $useragent, $userpass, $auth)
$query = New-Object System.DirectoryServices.DirectorySearcher($root, $LDAPSearchString)
$objClass = $query.findall()
$objClass.Count
$objClass |Get-Member

<# 
[System.Collections.ArrayList]$ParsedLDAPUsers = @()
$objClass| ForEach-Object { $obj = [PSCustomObject]@{

    #GivenName = $_.Properties.Item('uid')
    Mail = $_.Properties.Item('mail')
    }
    $ParsedLDAPUsers.Add($obj)|out-null
}
 #>

$SortedLDAPusers = @()
$objClass | ForEach-Object {$SortedLDAPUsers += $_.Properties.Item('mail')}
$SortedLDAPUsers = $SortedLDAPUsers|Sort-Object


$ComparedUsers = Compare-Object -ReferenceObject $SortedADUsers -DifferenceObject $SortedLDAPUsers


$LDAPOnly = @()
$ADOnly = @()
ForEach($difference in $ComparedUsers){
    switch ($difference.SideIndicator) {
        #In LDAP but not in AD
        "=>" {
            $LDAPOnly += $difference.InputObject
        }
        #In AD but not LDAP
        "<=" {
            $ADOnly += $difference.InputObject
        }
    }
}