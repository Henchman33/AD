<#
.SYNOPSIS
    Active Directory Server Inventory Report Generator

.DESCRIPTION
    Recursively searches all Organizational Units (OUs) in the current
    Active Directory domain for Windows Server computer objects.

    Outputs:
        - CSV Report
        - Executive HTML Report

    Reports are automatically saved to:
        Desktop\All_Servers_Report

.NOTES
    Author: ChatGPT
    Version: 1.0
    Requirements:
        - RSAT ActiveDirectory Module
        - Domain connectivity
        - PowerShell ISE recommended
        - Run with Domain User privileges

    Compatible:
        - Windows Server 2012+
        - Windows 10/11 with RSAT
#>

# =========================
# INITIAL SETUP
# =========================

Clear-Host

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Active Directory Server Inventory Tool" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Import AD Module
Try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
Catch {
    Write-Host "ERROR: ActiveDirectory module not found." -ForegroundColor Red
    Write-Host "Install RSAT tools before running this script." -ForegroundColor Yellow
    Break
}

# =========================
# CREATE REPORT DIRECTORY
# =========================

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ReportFolder = Join-Path $DesktopPath "All_Servers_Report"

If (!(Test-Path $ReportFolder)) {
    New-Item -ItemType Directory -Path $ReportFolder | Out-Null
}

# Timestamp
$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"

# Output Files
$CSVReport  = Join-Path $ReportFolder "All_Servers_Report_$DateStamp.csv"
$HTMLReport = Join-Path $ReportFolder "All_Servers_Report_$DateStamp.html"

# =========================
# GET DOMAIN INFORMATION
# =========================

Try {
    $DomainInfo = Get-ADDomain
    $DomainFQDN = $DomainInfo.DNSRoot
}
Catch {
    Write-Host "Failed to retrieve domain information." -ForegroundColor Red
    Break
}

Write-Host "Connected Domain: $DomainFQDN" -ForegroundColor Green
Write-Host ""

# =========================
# FIND ALL SERVERS
# =========================

Write-Host "Searching Active Directory recursively for servers..." -ForegroundColor Yellow
Write-Host ""

# Get all Windows Server OS computer objects
$Servers = Get-ADComputer -Filter {
    OperatingSystem -like "*Server*"
} -Properties OperatingSystem, IPv4Address, DistinguishedName, DNSHostName

# =========================
# PROCESS RESULTS
# =========================

$Results = foreach ($Server in $Servers) {

    # Extract OU Path
    $OU = ($Server.DistinguishedName -replace '^CN=.*?,','')

    # Attempt DNS Resolution if IPv4 missing
    $IPv4 = $Server.IPv4Address

    if ([string]::IsNullOrWhiteSpace($IPv4)) {

        Try {
            $DNSResult = Resolve-DnsName $Server.DNSHostName -ErrorAction Stop |
                Where-Object {$_.Type -eq "A"} |
                Select-Object -First 1

            $IPv4 = $DNSResult.IPAddress
        }
        Catch {
            $IPv4 = "Unavailable"
        }
    }

    [PSCustomObject]@{
        ServerName     = $Server.Name
        IPv4Address    = $IPv4
        DomainFQDN     = $DomainFQDN
        OperatingSystem= $Server.OperatingSystem
        OrganizationalUnit = $OU
    }
}

# =========================
# EXPORT CSV
# =========================

$Results |
    Sort-Object ServerName |
    Export-Csv -Path $CSVReport -NoTypeInformation -Encoding UTF8

Write-Host "CSV Report Created:" -ForegroundColor Green
Write-Host $CSVReport -ForegroundColor White
Write-Host ""

# =========================
# CREATE HTML REPORT
# =========================

$ServerCount = $Results.Count
$GeneratedOn = Get-Date

$HTMLHeader = @"
<html>
<head>
<title>All Servers Report</title>

<style>
body {
    font-family: Segoe UI;
    background-color: #f5f5f5;
    margin: 20px;
}

h1 {
    color: #003366;
}

h2 {
    color: #444444;
}

table {
    border-collapse: collapse;
    width: 100%;
    background-color: white;
}

th {
    background-color: #003366;
    color: white;
    padding: 10px;
    border: 1px solid #dcdcdc;
}

td {
    padding: 8px;
    border: 1px solid #dcdcdc;
}

tr:nth-child(even) {
    background-color: #f2f2f2;
}

.summary {
    margin-bottom: 20px;
    padding: 10px;
    background-color: white;
    border: 1px solid #dcdcdc;
}
</style>

</head>
<body>

<h1>Executive Server Inventory Report</h1>

<div class='summary'>
<h2>Report Summary</h2>

<p><strong>Domain:</strong> $DomainFQDN</p>
<p><strong>Total Servers Found:</strong> $ServerCount</p>
<p><strong>Generated:</strong> $GeneratedOn</p>

</div>

"@

$HTMLFooter = @"
</body>
</html>
"@

$Results |
    Sort-Object ServerName |
    ConvertTo-Html `
        -Property ServerName,
                  IPv4Address,
                  DomainFQDN,
                  OperatingSystem,
                  OrganizationalUnit `
        -Head $HTMLHeader `
        -PostContent $HTMLFooter |
    Out-File $HTMLReport -Encoding UTF8

Write-Host "HTML Executive Report Created:" -ForegroundColor Green
Write-Host $HTMLReport -ForegroundColor White
Write-Host ""

# =========================
# COMPLETION SUMMARY
# =========================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " REPORT GENERATION COMPLETE" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total Servers Found: $ServerCount" -ForegroundColor Green
Write-Host ""
Write-Host "Reports saved to:" -ForegroundColor Yellow
Write-Host $ReportFolder -ForegroundColor White
Write-Host ""

# Automatically open report folder
Invoke-Item $ReportFolder
