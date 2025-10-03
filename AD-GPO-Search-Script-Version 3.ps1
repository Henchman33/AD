<#
.SYNOPSIS
    Searches Active Directory GPOs for User Rights Assignment entries that match one or more search strings.

.PARAMETER SearchString
    One or more strings to search for in GPO settings (supports wildcards). 
    Example: "*logon as a service*", "*Exclude*logon as service*", "*deny*logon as service*"

# GPOs with settings like “Exclude User or Computer logon as service”, not just any “Log on as a service”.

This searches only for "log on as a service", we’ll broaden the match so it catches both include and exclude rules. For example:

“Log on as a service” (SeServiceLogonRight)

“Deny log on as a service” (SeDenyServiceLogonRight)

“Exclude User or Computer logon as service” (custom naming or mislabeling in some reports)

I’ll adjust the search to look for -like "*Exclude*logon as service*" in both Computer and User Configuration.

“Log on as a service” (SeServiceLogonRight)

“Deny log on as a service” (SeDenyServiceLogonRight)

“Exclude User or Computer logon as service” (custom naming or mislabeling in some reports)

I’ll adjust the search to look for -like "*Exclude*logon as service*" in both Computer and User Configuration.
#>

param(
    [string[]]$SearchString = @("*Exclude*logon as service*")  # Default
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

Write-Host "Searching GPOs in domain '$DomainName' for policies matching: $($SearchString -join ', ')"

# Collect all GPOs
$AllGPOs = Get-GPO -All

# Container for results
$Results = @()

foreach ($GPO in $AllGPOs) {
    # Export GPO report in XML format to parse settings
    $ReportXml = Get-GPOReport -Guid $GPO.Id -ReportType Xml
    [xml]$GPOXml = $ReportXml

    # Search both Computer and User Configuration for matches against ANY search string
    $ComputerRights = $GPOXml.GPO.Computer.ExtensionData.Extension.Policy | Where-Object {
        foreach ($s in $SearchString) {
            if ($_.Name -like $s) { return $true }
        }
    }

    $UserRights = $GPOXml.GPO.User.ExtensionData.Extension.Policy | Where-Object {
        foreach ($s in $SearchString) {
            if ($_.Name -like $s) { return $true }
        }
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

        # Record which search term(s) matched
        $MatchedTerms = @()
        foreach ($s in $SearchString) {
            if ($ComputerRights | Where-Object { $_.Name -like $s }) { $MatchedTerms += $s }
            if ($UserRights | Where-Object { $_.Name -like $s }) { $MatchedTerms += $s }
        }
        $MatchedTerms = ($MatchedTerms | Select-Object -Unique) -join ", "

        $Results += [PSCustomObject]@{
            GPOName          = $GPO.DisplayName
            GPOId            = $GPO.Id
            LinkedTo         = ($Links -join ", ")
            MatchedSearch    = $MatchedTerms
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
$SafeSearch = ($SearchString -join "_") -replace '[^a-zA-Z0-9]', "_"
$CsvPath     = Join-Path $OutputFolder "$DomainName`_GPO_Search_$SafeSearch.csv"
$XmlPath     = Join-Path $OutputFolder "$DomainName`_GPO_Search_$SafeSearch.xml"
$SummaryPath = Join-Path $OutputFolder "$DomainName`_Summary.txt"

if ($Results.Count -eq 0) {
    "Domain: $DomainName" | Out-File $SummaryPath -Encoding UTF8
    "Summary: 0 GPOs found that match search '$($SearchString -join ", ")'." | Out-File $SummaryPath -Append -Encoding UTF8
    "No GPOs found that match search '$($SearchString -join ", ")'." | Out-File $CsvPath -Encoding UTF8
    "No GPOs found that match search '$($SearchString -join ", ")'." | Out-File $XmlPath -Encoding UTF8
} else {
    # Export to CSV and XML
    $Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    $Results | Export-Clixml -Path $XmlPath -Encoding UTF8

    # Build summary content
    $SummaryLines = @()
    $SummaryLines += "Domain: $DomainName"
    $SummaryLines += "Summary: $($Results.Count) GPO(s) found that match search: $($SearchString -join ', ')"
    $SummaryLines += ""
    $SummaryLines += "GPOs Found:"
    $SummaryLines += "-----------"
    $Results | ForEach-Object {
        $SummaryLines += "$($_.GPOName) (Linked To: $($_.LinkedTo)) [Scope: $($_.ConfigScope)]"
        $SummaryLines += "   Matched Search Term(s): $($_.MatchedSearch)"
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
