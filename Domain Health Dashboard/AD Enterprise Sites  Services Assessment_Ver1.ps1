
<#
AD Enterprise Sites & Services Assessment Version 1
Requirements:
- RSAT ActiveDirectory
- ImportExcel module
- Run with rights to query AD
#>

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportRoot = Join-Path $env:USERPROFILE "Desktop\AD_Assessment_$TimeStamp"
New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null

$LogFile = Join-Path $ReportRoot "Execution.log"

function Write-Log {
    param([string]$Message)
    $line = "{0} : {1}" -f (Get-Date), $Message
    $line | Tee-Object -FilePath $LogFile -Append
}

Write-Log "Starting AD assessment"

$Forest = Get-ADForest
$Domains = $Forest.Domains

$Sites = @()
$Subnets = @()
$SiteLinks = @()
$SiteLinkBridges = @()
$DCs = @()
$FSMO = @()

try {
    $Sites = Get-ADReplicationSite -Filter * -Properties *
    $Subnets = Get-ADReplicationSubnet -Filter * -Properties *
    $SiteLinks = Get-ADReplicationSiteLink -Filter * -Properties *
    $SiteLinkBridges = Get-ADReplicationSiteLinkBridge -Filter * -Properties *
}
catch {
    Write-Log $_
}

foreach ($Domain in $Domains) {
    Write-Log "Processing $Domain"

    try {
        $DomainInfo = Get-ADDomain -Server $Domain

        $FSMO += [pscustomobject]@{
            Domain = $Domain
            PDCEmulator = $DomainInfo.PDCEmulator
            RIDMaster = $DomainInfo.RIDMaster
            InfrastructureMaster = $DomainInfo.InfrastructureMaster
        }

        Get-ADDomainController -Filter * -Server $Domain | ForEach-Object {
            $DCs += [pscustomobject]@{
                Domain = $Domain
                HostName = $_.HostName
                Site = $_.Site
                IPv4Address = $_.IPv4Address
                IsGlobalCatalog = $_.IsGlobalCatalog
                OperatingSystem = $_.OperatingSystem
            }
        }
    }
    catch {
        Write-Log $_
    }
}

$ForestFSMO = [pscustomobject]@{
    Forest = $Forest.Name
    DomainNamingMaster = $Forest.DomainNamingMaster
    SchemaMaster = $Forest.SchemaMaster
}

$FSMO += $ForestFSMO

# CSV Exports
$Sites | Export-Csv "$ReportRoot\Sites.csv" -NoTypeInformation
$Subnets | Export-Csv "$ReportRoot\Subnets.csv" -NoTypeInformation
$SiteLinks | Export-Csv "$ReportRoot\SiteLinks.csv" -NoTypeInformation
$SiteLinkBridges | Export-Csv "$ReportRoot\SiteLinkBridges.csv" -NoTypeInformation
$DCs | Export-Csv "$ReportRoot\DomainControllers.csv" -NoTypeInformation
$FSMO | Export-Csv "$ReportRoot\FSMO.csv" -NoTypeInformation

# Excel
$Xlsx = "$ReportRoot\AD_Assessment.xlsx"

[pscustomobject]@{
    Forest = $Forest.Name
    Domains = ($Domains -join ";")
    Sites = $Sites.Count
    Subnets = $Subnets.Count
    SiteLinks = $SiteLinks.Count
    SiteLinkBridges = $SiteLinkBridges.Count
    DomainControllers = $DCs.Count
} | Export-Excel $Xlsx -WorksheetName Summary -AutoSize

$Sites | Export-Excel $Xlsx -WorksheetName Sites -AutoSize -AutoFilter -Append
$Subnets | Export-Excel $Xlsx -WorksheetName Subnets -AutoSize -AutoFilter -Append
$SiteLinks | Export-Excel $Xlsx -WorksheetName SiteLinks -AutoSize -AutoFilter -Append
$SiteLinkBridges | Export-Excel $Xlsx -WorksheetName SiteLinkBridges -AutoSize -AutoFilter -Append
$DCs | Export-Excel $Xlsx -WorksheetName DomainControllers -AutoSize -AutoFilter -Append
$FSMO | Export-Excel $Xlsx -WorksheetName FSMO -AutoSize -AutoFilter -Append

# Repadmin
try {
    repadmin /replsummary > "$ReportRoot\repadmin_replsummary.txt"
}
catch {}

# DCDiag
try {
    dcdiag /e /v > "$ReportRoot\dcdiag.txt"
}
catch {}

# SVG Topology (basic)
$Svg = @"
<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800">
<text x="20" y="30" font-size="20">AD Topology Summary</text>
<text x="20" y="60">Forest: $($Forest.Name)</text>
<text x="20" y="90">Sites: $($Sites.Count)</text>
<text x="20" y="120">DCs: $($DCs.Count)</text>
</svg>
"@

$Svg | Out-File "$ReportRoot\Topology.svg" -Encoding utf8

# HTML
$style = @"
<style>
body {font-family:Segoe UI;}
table {border-collapse:collapse;width:100%;}
th,td {border:1px solid #ccc;padding:4px;}
th {background:#003366;color:white;}
</style>
"@

$summary = @"
<h1>AD Assessment Report</h1>
<p>Forest: $($Forest.Name)</p>
<p>Generated: $(Get-Date)</p>
<p>Domains: $($Domains.Count)</p>
<p>Sites: $($Sites.Count)</p>
<p>DCs: $($DCs.Count)</p>
"@

$body = $summary +
($Sites | Select Name | ConvertTo-Html -Fragment -PreContent "<h2>Sites</h2>") +
($Subnets | Select Name,Site | ConvertTo-Html -Fragment -PreContent "<h2>Subnets</h2>") +
($DCs | ConvertTo-Html -Fragment -PreContent "<h2>Domain Controllers</h2>")

ConvertTo-Html -Head $style -Body $body | Out-File "$ReportRoot\AD_Assessment.html"

Compress-Archive -Path $ReportRoot\* -DestinationPath "$ReportRoot.zip" -Force

Write-Host "Completed: $ReportRoot"
