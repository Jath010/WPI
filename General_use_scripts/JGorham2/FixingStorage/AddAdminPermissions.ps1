#Core function

$acl = Get-Acl $Folder
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("ADMIN\storage_admins","FullControl","Allow")
$acl.SetAccessRule($AccessRule)
$acl | Set-Acl $folder

