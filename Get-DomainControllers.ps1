# =====================================================================
# Script Name: Get-DomainControllers.ps1
# Purpose:     Retrieves all Domain Controllers in the domain, 
#              their names and IP addresses, and exports to CSV.
# Author:      Stephen McKee
# =====================================================================

# Ensure the output directory exists
$ExportPath = "C:\Temp\DomainControllers"
$ExportFile = "$ExportPath\DomainControllersDetails.csv"

if (!(Test-Path -Path $ExportPath)) {
    Write-Host "Creating output directory at $ExportPath..."
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

# Import the AD module (required for Get-ADDomainController)
Import-Module ActiveDirectory

Write-Host "Retrieving Domain Controller information..." -ForegroundColor Cyan

# Retrieve all domain controllers
$DomainControllers = Get-ADDomainController -Filter *

# Collect name and IP information
$DCList = $DomainControllers | Select-Object `
    @{Name='DCName'; Expression = {$_.Hostname}}, `
    @{Name='IPAddress'; Expression = {
        try {
            # Resolve IP address using DNS
            [System.Net.Dns]::GetHostAddresses($_.Hostname) | 
                Where-Object { $_.AddressFamily -eq 'InterNetwork' } | 
                Select-Object -First 1
        }
        catch {
            "Unable to resolve"
        }
    }}

# Export to CSV
$DCList | Export-Csv -Path $ExportFile -NoTypeInformation

Write-Host "`nExport complete!" -ForegroundColor Green
Write-Host "File saved to: $ExportFile" -ForegroundColor Yellow
