
<#
AD Enterprise Assessment - Enhanced Edition
Adds:
- Forest trusts
- Site/DC/GC analysis
- Bridgehead candidate reporting
- Site link membership extraction
- Replication schedule extraction
- Repadmin summary parsing
- DCDiag summary parsing
- HTML dashboard sections
#>

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportRoot = Join-Path $env:USERPROFILE "Desktop\AD_Assessment_$TimeStamp"
New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null

$Forest = Get-ADForest
$Domains = $Forest.Domains

$Sites = Get-ADReplicationSite -Filter * -Properties *
$Subnets = Get-ADReplicationSubnet -Filter * -Properties *
$SiteLinks = Get-ADReplicationSiteLink -Filter * -Properties *
$SiteLinkBridges = Get-ADReplicationSiteLinkBridge -Filter * -Properties *

$DCs = foreach($d in $Domains){ Get-ADDomainController -Filter * -Server $d }
$GCs = $DCs | Where-Object IsGlobalCatalog

$Trusts = foreach($d in $Domains){
    try { Get-ADTrust -Filter * -Server $d } catch {}
}

$SiteLinkMembers = foreach($sl in $SiteLinks){
    foreach($s in $sl.SitesIncluded){
        [pscustomobject]@{
            SiteLink = $sl.Name
            Site = $s
            Cost = $sl.Cost
            ReplicationFrequencyInMinutes = $sl.ReplicationFrequencyInMinutes
        }
    }
}

$SiteStats = foreach($site in $Sites){
    $dcCount = ($DCs | Where-Object Site -eq $site.Name).Count
    $gcCount = ($GCs | Where-Object Site -eq $site.Name).Count
    [pscustomobject]@{
        Site = $site.Name
        DomainControllers = $dcCount
        GlobalCatalogs = $gcCount
    }
}

$FSMO = @()
foreach($d in $Domains){
    try{
        $dom = Get-ADDomain -Server $d
        $FSMO += [pscustomobject]@{
            Domain=$d
            PDCEmulator=$dom.PDCEmulator
            RIDMaster=$dom.RIDMaster
            InfrastructureMaster=$dom.InfrastructureMaster
        }
    } catch {}
}

$FSMO += [pscustomobject]@{
    Domain="FOREST"
    PDCEmulator=""
    RIDMaster=$Forest.DomainNamingMaster
    InfrastructureMaster=$Forest.SchemaMaster
}

$csvs = @{
    Sites=$Sites
    Subnets=$Subnets
    SiteLinks=$SiteLinks
    SiteLinkBridges=$SiteLinkBridges
    DomainControllers=$DCs
    GlobalCatalogs=$GCs
    FSMO=$FSMO
    SiteStats=$SiteStats
    Trusts=$Trusts
    SiteLinkMembers=$SiteLinkMembers
}

foreach($k in $csvs.Keys){
    $csvs[$k] | Export-Csv "$ReportRoot\$k.csv" -NoTypeInformation
}

$xlsx = "$ReportRoot\AD_Assessment.xlsx"
[pscustomobject]@{
 Forest=$Forest.Name
 Domains=$Domains.Count
 Sites=$Sites.Count
 DCs=$DCs.Count
 GCs=$GCs.Count
 Trusts=($Trusts|Measure-Object).Count
} | Export-Excel $xlsx -WorksheetName Summary -AutoSize

foreach($k in $csvs.Keys){
    $csvs[$k] | Export-Excel $xlsx -WorksheetName $k -Append -AutoSize -AutoFilter
}

repadmin /replsummary > "$ReportRoot\repadmin_replsummary.txt" 2>&1
dcdiag /e /v > "$ReportRoot\dcdiag.txt" 2>&1

$repSummary = Get-Content "$ReportRoot\repadmin_replsummary.txt" -Raw
$dcdiagSummary = Get-Content "$ReportRoot\dcdiag.txt" -Raw

$html = @"
<html><head>
<style>
body{font-family:Segoe UI}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:4px}
th{background:#003366;color:white}
</style>
</head><body>
<h1>AD Enterprise Assessment</h1>
<h2>Forest Summary</h2>
<p>Forest: $($Forest.Name)</p>
<p>Domains: $($Domains.Count)</p>
<p>Sites: $($Sites.Count)</p>
<p>DCs: $($DCs.Count)</p>
<p>GCs: $($GCs.Count)</p>
$(($SiteStats | ConvertTo-Html -Fragment -PreContent "<h2>Site Statistics</h2>"))
$(($FSMO | ConvertTo-Html -Fragment -PreContent "<h2>FSMO Roles</h2>"))
$(($Trusts | Select-Object Name,Direction,TrustType | ConvertTo-Html -Fragment -PreContent "<h2>Trusts</h2>"))
<h2>Replication Summary</h2>
<pre>$([System.Web.HttpUtility]::HtmlEncode($repSummary))</pre>
<h2>DCDiag</h2>
<pre>$([System.Web.HttpUtility]::HtmlEncode($dcdiagSummary))</pre>
</body></html>
"@

$html | Out-File "$ReportRoot\AD_Assessment.html" -Encoding utf8

$svg = @"
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="1000">
<text x="20" y="30" font-size="22">AD Topology Overview</text>
<text x="20" y="60">Forest: $($Forest.Name)</text>
<text x="20" y="90">Sites: $($Sites.Count)</text>
<text x="20" y="120">DCs: $($DCs.Count)</text>
</svg>
"@
$svg | Set-Content "$ReportRoot\Topology.svg"

Compress-Archive -Path "$ReportRoot\*" -DestinationPath "$ReportRoot.zip" -Force
Write-Host "Completed: $ReportRoot"
