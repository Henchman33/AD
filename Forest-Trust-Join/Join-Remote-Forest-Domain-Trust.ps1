# Don't forget to change these first
$strRemoteForest = "<domainB>"
$strRemoteAdmin = "domainB\<username>"
$strRemoteAdminPassword = "<password>"

$remoteContext = New-Object -TypeName "System.DirectoryServices.ActiveDirectory.DirectoryContext" `
    -ArgumentList @("Forest", $strRemoteForest, $strRemoteAdmin, $strRemoteAdminPassword)

try {
    $remoteForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($remoteContext)
    Write-Host "GetRemoteForest: Succeeded for domain $($remoteForest.Name)"
}
catch {
    Write-Warning "GetRemoteForest: Failed:`n`tError: $($_.Exception.Message)"
}

Write-Host "Connected to Remote forest: $($remoteForest.Name)"

# Get local forest
$localForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
Write-Host "Connected to Local forest: $($localForest.Name)"

try {
    $localForest.CreateTrustRelationship($remoteForest, "Bidirectional")
    Write-Host "CreateTrustRelationship: Succeeded for domain $($remoteForest.Name)"
}
catch {
    Write-Warning "CreateTrustRelationship: Failed for domain $($remoteForest.Name)`n`tError: $($_.Exception.Message)"
}
