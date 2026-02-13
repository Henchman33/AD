# This Script will get the Shared folder permissions - just change the lines you need to and run as Administrator in PowerShell or PowerShell ISE.

$FolderPath = dir -Directory -Path "\\SERVER\Share" -Recurse -Force
$ReportPath = "C:\Report\FolderPermissions.csv"
$Report = @()
Foreach ($Folder in $FolderPath) {
$Acl = Get-Acl -Path $Folder.FullName
foreach ($Access in $acl.Access)
{
$Properties = [ordered]@{'FolderName'=$Folder.FullName;'AD
Group or
User'=$Access.IdentityReference;'Permissions'=$Access.FileSystemRights;'Inherited'=$Access.IsInherited}
$Report += New-Object -TypeName PSObject -Property $Properties
}
}
$Report | Export-Csv -path $ReportPath 
