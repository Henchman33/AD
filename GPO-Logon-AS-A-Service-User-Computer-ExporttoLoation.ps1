# Requires Group Policy module + ActiveDirectory module
# Run on a Domain Controller with RSAT / GPMC installed

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

    # Search both Computer and User Configuration for "Log on as a service"
    $ComputerRights = $GPOXml.GPO.Computer.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -eq "Log on as a service"
    }
    $UserRights = $GPOXml.GPO.User.ExtensionData.Extension.Policy | Where-Object {
        $_.Name -eq "Log on as a service"
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

        $Results += [PSCustomObject]@{
            GPOName     = $GPO.DisplayName
            GPOId       = $GPO.Id
            LinkedTo    = ($Links -join ", ")
            UserRight   = "Log on as a service (SeServiceLogonRight)"
            ConfigScope = @(
                if ($ComputerRights) { "Computer Configuration" }
                if ($UserRights) { "User Configuration" }
            ) -join ", "
            AllSettings = $Settings
        }
    }
}

# Output file paths with domain name
$CsvPath     = Join-Path $OutputFolder "$DomainName`_GPO_LogonAsService.csv"
$XmlPath     = Join-Path $OutputFolder "$DomainName`_GPO_LogonAsService.xml"
$SummaryPath = Join-Path $OutputFolder "$DomainName`_Summary.txt"

if ($Results.Count -eq 0) {
    "No GPOs found that configure 'Log on as a service' (SeServiceLogonRight)." | Out-File $CsvPath -Encoding UTF8
    "No GPOs found that configure 'Log on as a service' (SeServiceLogonRight)." | Out-File $XmlPath -Encoding UTF8
    "Domain: $DomainName`r`nSummary: 0 GPOs found." | Out-File $SummaryPath -Encoding UTF8
} else {
    # Export to CSV and XML
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    $Results | Export-Clixml -Path $XmlPath -Encoding UTF8
    "Domain: $DomainName`r`nSummary: $($Results.Count) GPO(s) found with 'Log on as a service' configured." | Out-File $SummaryPath -Encoding UTF8
}

Write-Host "Export completed:"
Write-Host " - CSV: $CsvPath"
Write-Host " - XML: $XmlPath"
Write-Host " - Summary: $SummaryPath"
