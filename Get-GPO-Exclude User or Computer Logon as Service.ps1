# Requires Group Policy module + ActiveDirectory module
# Run on a Domain Controller with RSAT / GPMC installed
# Uses the -like instead of -eq for searching GP's for Exclude Users or Computer GPO's to logon as a service.
# Author - Steve McKee - Systems Administrator II - stevemckee@outlook.com
# Exports findings to c:\Temp\GPOS\UserGPOS with a summary file with the number count and .csv and .xml file
# NOTE - It's easier to create the -path first before running.
# Updated 

# Get domain name
$DomainName = (Get-ADDomain).DNSRoot

# Output folder
$OutputFolder = "C:\Temp\GPOS\UserGPOS"
If (!(Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

# Collect all GPOs
$AllGPOs = Get-GPO -All

# Container for results
$Results = @()

foreach ($GPO in $AllGPOs) {
    # Export GPO report in XML format to parse settings
    $ReportXml = Get-GPOReport -Guid $GPO.Id -ReportType Xml
    [xml]$GPOXml = $ReportXml

    # Search both Computer and User Configuration for "log on as a service" (using -like)
    $ComputerRights = $GPOXml.GPO.Computer.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -like "*log on as a service*"
    }
    $UserRights = $GPOXml.GPO.User.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -like "*log on as a service*"
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

        # Extract accounts/groups specifically assigned "log on as a service"
        $AssignedAccounts = @()
        if ($ComputerRights) {
            $AssignedAccounts += $ComputerRights.Setting
        }
        if ($UserRights) {
            $AssignedAccounts += $UserRights.Setting
        }

        $AssignedAccounts = ($AssignedAccounts | Where-Object { $_ -ne $null -and $_ -ne "" }) -join ", "

        $Results += [PSCustomObject]@{
            GPOName         = $GPO.DisplayName
            GPOId           = $GPO.Id
            LinkedTo        = ($Links -join ", ")
            UserRight       = "Log on as a service (SeServiceLogonRight)"
            ConfigScope     = @(
                if ($ComputerRights) { "Computer Configuration" }
                if ($UserRights) { "User Configuration" }
            ) -join ", "
            AssignedAccounts = $AssignedAccounts
            AllSettings      = $Settings
        }
    }
}

# Output file paths with domain name
$CsvPath     = Join-Path $OutputFolder "$DomainName`_GPO_LogonAsService.csv"
$XmlPath     = Join-Path $OutputFolder "$DomainName`_GPO_LogonAsService.xml"
$SummaryPath = Join-Path $OutputFolder "$DomainName`_Summary.txt"

if ($Results.Count -eq 0) {
    "Domain: $DomainName" | Out-File $SummaryPath -Encoding UTF8
    "Summary: 0 GPOs found that configure 'Log on as a service' (SeServiceLogonRight)." | Out-File $SummaryPath -Append -Encoding UTF8
    "No GPOs found that configure 'Log on as a service' (SeServiceLogonRight)." | Out-File $CsvPath -Encoding UTF8
    "No GPOs found that configure 'Log on as a service' (SeServiceLogonRight)." | Out-File $XmlPath -Encoding UTF8
} else {
    # Export to CSV and XML
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    $Results | Export-Clixml -Path $XmlPath -Encoding UTF8

    # Build summary content
    $SummaryLines = @()
    $SummaryLines += "Domain: $DomainName"
    $SummaryLines += "Summary: $($Results.Count) GPO(s) found with 'Log on as a service' configured."
    $SummaryLines += ""
    $SummaryLines += "GPOs Found:"
    $SummaryLines += "-----------"
    $Results | ForEach-Object {
        $SummaryLines += "$($_.GPOName) (Linked To: $($_.LinkedTo)) [Scope: $($_.ConfigScope)]"
        if ($_.AssignedAccounts) {
            $SummaryLines += "   Assigned Accounts: $($_.AssignedAccounts)"
        }
    }

    # Save summary
    $SummaryLines | Out-File $SummaryPath -Encoding UTF8
}

Write-Host "Export completed:"
Write-Host " - CSV: $CsvPath"
Write-Host " - XML: $XmlPath"
Write-Host " - Summary: $SummaryPath"
