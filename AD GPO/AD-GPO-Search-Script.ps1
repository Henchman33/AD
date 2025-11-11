# GPO Search Script
# .\Find-GPOUserRights.ps1 -SearchString "*logon as a service*"
# or
# .\Find-GPOUserRights.ps1 -SearchString "*Exclude*logon as service*"
# Default behavior (search for exclusions): .\Find-GPOUserRights.ps1
# Search for any logon as a service: .\Find-GPOUserRights.ps1 -SearchString "*logon as a service*"
# Search for deny logon as a service: .\Find-GPOUserRights.ps1 -SearchString "*deny*logon as a service*"
# Author - Steve McKee - Systems Administrator II - stevemckee@outlook.com
# Exports findings to c:\Temp\GPOS\UserGPOS with a summary file with the number count and .csv and .xml file
# NOTE - It's easier to create the -path first before running.
# Updated 




<#
.SYNOPSIS
    Searches Active Directory GPOs for User Rights Assignment entries that match a given search string.

.PARAMETER SearchString
    The string to search for in GPO settings (supports wildcards). 
    Example: "*logon as a service*", "*Exclude*logon as service*"
#>

param(
    [string]$SearchString = "*Exclude*logon as service*"
)

# Requires Group Policy module + ActiveDirectory module
# Run on a Domain Controller with RSAT / GPMC installed

# Get domain name
$DomainName = (Get-ADDomain).DNSRoot

# Output folder
$OutputFolder = "C:\Temp\GPOS\UserGPOS"
If (!(Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

Write-Host "Searching GPOs in domain '$DomainName' for policies matching: $SearchString"

# Collect all GPOs
$AllGPOs = Get-GPO -All

# Container for results
$Results = @()

foreach ($GPO in $AllGPOs) {
    # Export GPO report in XML format to parse settings
    $ReportXml = Get-GPOReport -Guid $GPO.Id -ReportType Xml
    [xml]$GPOXml = $ReportXml

    # Search both Computer and User Configuration for matches
    $ComputerRights = $GPOXml.GPO.Computer.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -like $SearchString
    }
    $UserRights = $GPOXml.GPO.User.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -like $SearchString
    }

    if ($ComputerRights -or $UserRights) {
        # Get link locations
        $Links = (Get-GPOLink -Guid $GPO.Id).LinksTo | ForEach-Object { $_.Scope }

        # Collect ALL settings (Computer + User) into readable text
        $CompSettings = if ($GPOXml.GPO.Computer.ExtensionData.Extension.Policy) {
            $GPOXml.GPO.Computer.ExtensionData.Extension.Policy | ForEach-Object {
                "Computer: $($_.Name) = $($_.Setting)"
            }
        }

        $UserSettings = if ($GPOXml.GPO.User.ExtensionData.Extension.Policy) {
            $GPOXml.GPO.User.ExtensionData.Extension.Policy | ForEach-Object {
                "User: $($_.Name) = $($_.Setting)"
            }
        }

        $Settings = @($CompSettings + $UserSettings) -join "; "

        # Extract accounts/groups specifically assigned
        $AssignedAccounts = @()
        if ($ComputerRights) {
            $AssignedAccounts += $ComputerRights.Setting
        }
        if ($UserRights) {
            $AssignedAccounts += $UserRights.Setting
        }

        $AssignedAccounts = ($AssignedAccounts | Where-Object { $_ -ne $null -and $_ -ne "" }) -join ", "

        $Results += [PSCustomObject]@{
            GPOName          = $GPO.DisplayName
            GPOId            = $GPO.Id
            LinkedTo         = ($Links -join ", ")
            MatchString      = $SearchString
            ConfigScope      = @(
                if ($ComputerRights) { "Computer Configuration" }
                if ($UserRights) { "User Configuration" }
            ) -join ", "
            AssignedAccounts = $AssignedAccounts
            AllSettings      = $Settings
        }
    }
}

# Output file paths with domain name
$SafeSearch = ($SearchString -replace '[^a-zA-Z0-9]', "_").Trim("_")
$CsvPath     = Join-Path $OutputFolder "$DomainName`_GPO_Search_$SafeSearch.csv"
$XmlPath     = Join-Path $OutputFolder "$DomainName`_GPO_Search_$SafeSearch.xml"
$SummaryPath = Join-Path $OutputFolder "$DomainName`_Summary.txt"

if ($Results.Count -eq 0) {
    "Domain: $DomainName" | Out-File $SummaryPath -Encoding UTF8
    "Summary: 0 GPOs found that match search '$SearchString'." | Out-File $SummaryPath -Append -Encoding UTF8
    "No GPOs found that match search '$SearchString'." | Out-File $CsvPath -Encoding UTF8
    "No GPOs found that match search '$SearchString'." | Out-File $XmlPath -Encoding UTF8
} else {
    # Export to CSV and XML
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    $Results | Export-Clixml -Path $XmlPath -Encoding UTF8

    # Build summary content
    $SummaryLines = @()
    $SummaryLines += "Domain: $DomainName"
    $SummaryLines += "Summary: $($Results.Count) GPO(s) found that match search: $SearchString"
    $SummaryLines += ""
    $SummaryLines += "GPOs Found:"
    $SummaryLines += "-----------"
    $Results | ForEach-Object {
        $SummaryLines += "$($_.GPOName) (Linked To: $($_.LinkedTo)) [Scope: $($_.ConfigScope)]"
        if ($_.AssignedAccounts) {
            $SummaryLines += "   Accounts: $($_.AssignedAccounts)"
        }
    }

    # Save summary
    $SummaryLines | Out-File $SummaryPath -Encoding UTF8
}

Write-Host "Export completed:"
Write-Host " - CSV: $CsvPath"
Write-Host " - XML: $XmlPath"
Write-Host " - Summary: $SummaryPath"
