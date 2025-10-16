# =====================================================================
# Script Name: Get-DomainControllersDetails.ps1
# Purpose:     Retrieves all Domain Controllers in the current domain,
#              including name, IP, GC status, read-only flag,
#              OS, and AD site name. Exports to CSV.
# Author:      Stephen McKee
# =====================================================================

# Define export path
$ExportPath = "C:\Temp\DomainControllers"
$ExportFile = "$ExportPath\DomainControllersDetails.csv"

# Ensure export directory exists
if (!(Test-Path -Path $ExportPath)) {
    Write-Host "Creating export directory at $ExportPath..." -ForegroundColor Yellow
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

# Import Active Directory module
Import-Module ActiveDirectory

Write-Host "Retrieving Domain Controller information..." -ForegroundColor Cyan

# Retrieve all domain controllers
$DomainControllers = Get-ADDomainController -Filter * 

# Collect required information
$DCList = $DomainControllers | Select-Object `
    @{Name='DCName'; Expression = {$_.Hostname}}, `
    @{Name='IPAddress'; Expression = {
        try {
            [System.Net.Dns]::GetHostAddresses($_.Hostname) |
                Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                Select-Object -ExpandProperty IPAddressToString -First 1
        } catch {
            "Unable to resolve"
        }
    }}, `
    @{Name='IsGlobalCatalog'; Expression = {$_.IsGlobalCatalog}}, `
    @{Name='IsReadOnly'; Expression = {$_.IsReadOnly}}, `
    @{Name='OperatingSystem'; Expression = {$_.OperatingSystem}}, `
    @{Name='SiteName'; Expression = {$_.Site}}

# Export to CSV
$DCList | Export-Csv -Path $ExportFile -NoTypeInformation

Write-Host "`nExport complete!" -ForegroundColor Green
Write-Host "File saved to: $ExportFile" -ForegroundColor Yellow
