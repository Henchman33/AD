#	   
<#
.SYNOPSIS
  AD forest inventory compatible with PowerShell ISE.
  Collects DCs, DNS, DHCP, replication (repadmin parsing), GPOs, privileged/service accounts, Exchange discovery (implicit remoting optional),
  writes CSVs to the calling user's Desktop in a timestamped folder, builds an Excel workbook (uses ImportExcel or Excel COM), and creates an HTML summary.
  NOTE: Graphviz rendering and auto-install removed for ISE compatibility — DOT file is still produced for manual rendering.

.PARAMETER OutputFolder
  Optional path. Default: $env:USERPROFILE\Desktop\AD-Inventories\<timestamp>

.PARAMETER TierMappingCsv
  Optional CSV with columns GroupName,Tier for custom Tier mapping.

.PARAMETER ExchangeServer
  Optional: hostname of an Exchange server to use for implicit remoting if Exchange module isn't installed locally.

.EXAMPLE
  .\AD-Forest-Inventory-ISE.ps1

.Author 
 Stephen McKee Server Administrator 2
#>

param(
    [string]$OutputFolder,
    [string]$TierMappingCsv,
    [string]$ExchangeServer
)

# --- Bootstrap ---
$ts = (Get-Date).ToString("yyyyMMdd-HHmmss")
if (-not $OutputFolder) {
    $OutputFolder = Join-Path $env:USERPROFILE "Desktop\AD-Inventories\$ts"
}
New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

function Ensure-Module {
    param($Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Verbose "Could not install module {$Name}: $_"
        }
    }
    Import-Module $Name -ErrorAction SilentlyContinue
}

# Required and helpful modules
Ensure-Module -Name ActiveDirectory
Ensure-Module -Name GroupPolicy
Ensure-Module -Name ImportExcel    # used when Excel COM not available
# optional DHCP module
Import-Module DhcpServer -ErrorAction SilentlyContinue

# Helper: safe invoke against remote computer with timeout
function Test-RemoteService {
    param($Computer, $ServiceName)
    try {
        $svc = Get-Service -ComputerName $Computer -Name $ServiceName -ErrorAction Stop
        return $svc.Status
    } catch {
        return $null
    }
}

function Safe-Run {
    param($ScriptBlock)
    try { & $ScriptBlock } catch { Write-Verbose "Safe-Run error: $_"; return $null }
}

# --- Detect ISE and environment notes ---
$RunningInISE = ($Host.Name -match 'ISE')
if ($RunningInISE) {
    Write-Verbose "Running inside PowerShell ISE host - script adjusted for ISE compatibility."
}

# --- Forest & Domain info ---
$forest = Safe-Run { Get-ADForest }
$domain = Safe-Run { Get-ADDomain }

$forest | Select-Object * | Export-Csv (Join-Path $OutputFolder "ForestInfo.csv") -NoTypeInformation
$domain | Select-Object * | Export-Csv (Join-Path $OutputFolder "DomainInfo.csv") -NoTypeInformation

# FSMO
$fsmo = [PSCustomObject]@{
    SchemaMaster = $forest.SchemaMaster
    DomainNamingMaster = $forest.DomainNamingMaster
    PDCEmulator = (Get-ADDomainController -Filter {OperationMasterRoles -contains 'PDCEmulator'} -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -First 1)
    RIDMaster = $forest.RIDMaster
    InfrastructureMaster = $forest.InfrastructureMaster
}
$fsmo | Export-Csv (Join-Path $OutputFolder "FSMORoles.csv") -NoTypeInformation

# --- DCs & DNS ---
$DCs = Get-ADDomainController -Filter * | Select-Object Name,HostName,Site,IPv4Address,OperatingSystem,IsGlobalCatalog,IsReadOnly
$DCs | Export-Csv (Join-Path $OutputFolder "DomainControllers.csv") -NoTypeInformation

$dnsResults = foreach ($dc in $DCs.Name) {
    try {
        if (Get-Module -ListAvailable -Name DnsServer) {
            Import-Module DnsServer -ErrorAction SilentlyContinue
            $zones = Get-DnsServerZone -ComputerName $dc -ErrorAction Stop | Select-Object ZoneName, ZoneType, IsDsIntegrated
            foreach ($z in $zones) {
                [PSCustomObject]@{DC = $dc; Zone = $z.ZoneName; ZoneType = $z.ZoneType; IsDsIntegrated = $z.IsDsIntegrated}
            }
        } else {
            [PSCustomObject]@{DC = $dc; Zone = '<DnsServer module unavailable>'; ZoneType='N/A'; IsDsIntegrated=$false}
        }
    } catch {
        [PSCustomObject]@{DC = $dc; Zone = '<error>'; ZoneType='error'; IsDsIntegrated=$false}
    }
}
$dnsResults | Export-Csv (Join-Path $OutputFolder "DC-DNS-Zones.csv") -NoTypeInformation

# --- DHCP servers (authorized in AD) ---
try {
    if (Get-Command -Name Get-DhcpServerInDC -ErrorAction SilentlyContinue) {
        Get-DhcpServerInDC | Select-Object Name, IpAddress, ExpirationTime | Export-Csv (Join-Path $OutputFolder "AuthorizedDHCPServers.csv") -NoTypeInformation
    } else {
        [PSCustomObject]@{ Name = '<DhcpServer module not installed>'; IpAddress = ''; ExpirationTime = '' } | Export-Csv (Join-Path $OutputFolder "AuthorizedDHCPServers.csv") -NoTypeInformation
    }
} catch {
    [PSCustomObject]@{ Name = '<Error enumerating DHCP>'; IpAddress = ''; ExpirationTime = '' } | Export-Csv (Join-Path $OutputFolder "AuthorizedDHCPServers.csv") -NoTypeInformation
}

# --- Replication health (AD cmdlets + repadmin parsing) ---
$repFailures = Safe-Run { Get-ADReplicationFailure -Scope Site -ErrorAction SilentlyContinue }
if (-not $repFailures) {
    $repFailures = Safe-Run { Get-ADReplicationFailure -Target $domain.DNSRoot -ErrorAction SilentlyContinue }
}
if ($repFailures) { $repFailures | Export-Csv (Join-Path $OutputFolder "ReplicationFailures.csv") -NoTypeInformation } else { "" | Out-File (Join-Path $OutputFolder "ReplicationFailures.csv") }

# repadmin showrepl parsing (best-effort)
$repadminRaw = @()
$repAdmins = @()
$repDotEdges = @()
$repDotNodes = @{}
try {
    $repadminOut = & repadmin /showrepl * 2>&1
    $repadminOut | Out-File (Join-Path $OutputFolder "repadmin-showrepl.txt")
    foreach ($line in $repadminOut) {
        if ($line -match 'Source:\s*(?<src>\S+)') {
            $src = $Matches.src
            # attempt to get destination from nearby lines
            $idx = [Array]::IndexOf($repadminOut, $line)
            $dest = $null
            for ($i = $idx - 1; $i -ge 0; $i--) {
                if ($repadminOut[$i] -match '^\s*(?<dhost>\S+)\s+:\s*') { $dest = $Matches.dhost; break }
            }
            if (-not $dest) { $dest = $env:COMPUTERNAME }
            $edge = [PSCustomObject]@{
                Source = $src
                Destination = $dest
                Raw = $line
            }
            $repDotEdges += $edge
            $repDotNodes[$src] = $true
            $repDotNodes[$dest] = $true
            $repAdmins += $edge
        }
    }
} catch {
    Write-Verbose "repadmin not available or parse failed: $_"
}
$repAdmins | Export-Csv (Join-Path $OutputFolder "Repadmin-Parsed.csv") -NoTypeInformation

# Create DOT file (no rendering here; compatible with ISE — save for manual rendering)
$dotPath = Join-Path $OutputFolder "replication-topology.dot"
try {
    $dotLines = @("digraph ADReplication {","rankdir=LR;","node [shape=box, style=filled, fillcolor=lightblue];")
    foreach ($node in $repDotNodes.Keys) {
        $safe = ($node -replace '[^a-zA-Z0-9_]','_')
        $dotLines += "`"$safe`" [label=`"$node`"];"
    }
    foreach ($e in $repDotEdges) {
        $s = ($e.Source -replace '[^a-zA-Z0-9_]','_')
        $d = ($e.Destination -replace '[^a-zA-Z0-9_]','_')
        $lblEsc = ($e.Raw -replace '"','\"')
        $dotLines += "`"$s`" -> `"$d`" [label=`"$lblEsc`" color=`"black`"];"
    }
    $dotLines += "}"
    $dotLines | Out-File -FilePath $dotPath -Encoding utf8
} catch {
    Write-Verbose "Failed to write DOT file: $_"
}

# --- GPOs and per-GPO reports ---
$gpos = Safe-Run { Get-GPO -All -ErrorAction SilentlyContinue }
$gpos | Select-Object DisplayName, Id, Owner, GpoStatus | Export-Csv (Join-Path $OutputFolder "GPOs.csv") -NoTypeInformation

$gpoReportsFolder = Join-Path $OutputFolder "GPO-Reports"
New-Item -Path $gpoReportsFolder -ItemType Directory -Force | Out-Null
foreach ($g in $gpos) {
    $safe = ($g.DisplayName -replace '[\\/:*?"<>|]','_')
    $xmlPath = Join-Path $gpoReportsFolder "$safe-$($g.Id).xml"
    try { Get-GPOReport -Guid $g.Id -ReportType Xml -Path $xmlPath -ErrorAction Stop } catch { "$($g.DisplayName) - failed" | Out-File (Join-Path $gpoReportsFolder "errors.txt") -Append }
}

# --- Privileged accounts and tier mapping ---
$defaultTierGroups = @(
    @{Group='Enterprise Admins'; Tier=0},
    @{Group='Domain Admins'; Tier=0},
    @{Group='Schema Admins'; Tier=0},
    @{Group='Administrators'; Tier=0}
)
$tierMap = @{}
if ($TierMappingCsv -and (Test-Path $TierMappingCsv)) {
    Import-Csv $TierMappingCsv | ForEach-Object { $tierMap[$_.GroupName] = [int]$_.Tier }
} else {
    foreach ($m in $defaultTierGroups) { $tierMap[$m.Group] = $m.Tier }
}

$accountRows = @{}
foreach ($grpName in $tierMap.Keys) {
    try {
        $members = Get-ADGroupMember -Identity $grpName -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
        foreach ($m in $members) {
            $acct = Get-ADUser -Identity $m.SamAccountName -Properties MemberOf, ServicePrincipalName, Enabled, DistinguishedName -ErrorAction SilentlyContinue
            if (-not $acct) { continue }
            $key = $acct.SamAccountName
            if (-not $accountRows.ContainsKey($key)) {
                $accountRows[$key] = [ordered]@{
                    SamAccountName = $acct.SamAccountName
                    Name = $acct.Name
                    DistinguishedName = $acct.DistinguishedName
                    Enabled = $acct.Enabled
                    ServicePrincipalName = ($acct.ServicePrincipalName -join ';')
                    MemberOf = ($acct.MemberOf -join ';')
                    DomainAdmin = $false
                    Tier = ''
                    IsServiceAccount = $false
                }
            }
            $existingTier = $accountRows[$key].Tier
            $newTier = $tierMap[$grpName]
            if ($existingTier -eq '' -or ($newTier -ne $null -and [int]$newTier -lt [int]$existingTier)) {
                $accountRows[$key].Tier = $newTier
            }
            if ($grpName -eq 'Domain Admins') { $accountRows[$key].DomainAdmin = $true }
        }
    } catch {
        Write-Verbose "Failed to enumerate group {$grpName}: $_"
    }
}

# service accounts (MSA + SPN users)
try { $msas = Get-ADServiceAccount -Filter * -ErrorAction SilentlyContinue } catch { $msas = $null }
$spnUsers = Get-ADUser -Filter { ServicePrincipalName -like "*" } -Properties ServicePrincipalName -ErrorAction SilentlyContinue

if ($msas) {
    foreach ($m in $msas) {
        $key = $m.SamAccountName
        if (-not $accountRows.ContainsKey($key)) {
            $accountRows[$key] = [ordered]@{
                SamAccountName = $m.SamAccountName
                Name = $m.Name
                DistinguishedName = ''
                Enabled = ''
                ServicePrincipalName = ''
                MemberOf = ''
                DomainAdmin = $false
                Tier = ''
                IsServiceAccount = $true
            }
        } else {
            $accountRows[$key].IsServiceAccount = $true
        }
    }
}
if ($spnUsers) {
    foreach ($u in $spnUsers) {
        $key = $u.SamAccountName
        if (-not $accountRows.ContainsKey($key)) {
            $accountRows[$key] = [ordered]@{
                SamAccountName = $u.SamAccountName
                Name = $u.Name
                DistinguishedName = ''
                Enabled = ''
                ServicePrincipalName = ($u.ServicePrincipalName -join ';')
                MemberOf = ''
                DomainAdmin = $false
                Tier = ''
                IsServiceAccount = $true
            }
        } else {
            $accountRows[$key].IsServiceAccount = $true
            if (-not $accountRows[$key].ServicePrincipalName) {
                $accountRows[$key].ServicePrincipalName = ($u.ServicePrincipalName -join ';')
            }
        }
    }
}

$accountRows.Values | Export-Csv (Join-Path $OutputFolder "Account-Tiers-ServiceFlags.csv") -NoTypeInformation
$accountRows.Values | Where-Object { $_.DomainAdmin -or ($_.Tier -ne '') } | Export-Csv (Join-Path $OutputFolder "Privileged-Accounts-ByTier.csv") -NoTypeInformation

# --- Exchange discovery & optional implicit remoting ---
$exchangeDetected = @()
try {
    $exchangeCandidates = Get-ADComputer -Filter { Name -like "*EXCH*" -or Name -like "*MAIL*" } -Properties Description, OperatingSystem -ErrorAction SilentlyContinue
    foreach ($c in $exchangeCandidates) {
        $svc = Test-RemoteService -Computer $c.Name -ServiceName 'MSExchangeIS'
        $exchangeDetected += [PSCustomObject]@{Computer=$c.Name; OperatingSystem=$c.OperatingSystem; Description=$c.Description; MSExchangeIS=$svc}
    }
    $exchByAD = @()
    try { $exchByAD = Get-ADComputer -LDAPFilter "(msExchServerName=*)" -ErrorAction SilentlyContinue } catch {}
    foreach ($e in $exchByAD) {
        $exchangeDetected += [PSCustomObject]@{Computer=$e.Name; OperatingSystem=$e.OperatingSystem; Description=$e.Description; MSExchangeIS='Detected by msExch attribute'}
    }
} catch {
    Write-Verbose "Exchange candidate search failed: $_"
}

$exServersDetailed = @()
if (Get-Module -ListAvailable -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue -or (Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Microsoft.Exchange' })) {
    try { $exServersDetailed = Safe-Run { Get-ExchangeServer -ErrorAction SilentlyContinue | Select-Object Name, Edition, AdminDisplayVersion } } catch {}
} elseif ($ExchangeServer) {
    Write-Host "Attempting implicit remoting to Exchange server $ExchangeServer (Get-Credential may appear)."
    try {
        $cred = Get-Credential -Message "Credentials to connect to Exchange ($ExchangeServer)"
        $uri = "http://$ExchangeServer/PowerShell/"
        $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop
        Import-PSSession $s -AllowClobber -ErrorAction Stop | Out-Null
        $exServersDetailed = Safe-Run { Get-ExchangeServer -ErrorAction SilentlyContinue | Select-Object Name, Edition, AdminDisplayVersion }
        Remove-PSSession $s
    } catch {
        Write-Verbose "Implicit remoting to Exchange failed: $_"
    }
} else {
    Write-Verbose "No Exchange module or server supplied; using heuristics only."
}

$exchangeDetected | Sort-Object Computer -Unique | Export-Csv (Join-Path $OutputFolder "ExchangeServers-Detected.csv") -NoTypeInformation
if ($exServersDetailed) { $exServersDetailed | Export-Csv (Join-Path $OutputFolder "ExchangeServers-Remote.csv") -NoTypeInformation }

# --- Additional inventories ---
Get-ADComputer -Filter * -Properties OperatingSystem,OperatingSystemVersion,DistinguishedName |
    Select-Object Name,OperatingSystem,OperatingSystemVersion,DistinguishedName |
    Export-Csv (Join-Path $OutputFolder "AllComputers.csv") -NoTypeInformation

Get-ADDomainController -Filter * | Select-Object Name,Site,IPv4Address,OperatingSystem | Export-Csv (Join-Path $OutputFolder "DCsSites.csv") -NoTypeInformation
Safe-Run { Get-ADReplicationPartnerMetadata -Target * -ErrorAction SilentlyContinue } | Export-Csv (Join-Path $OutputFolder "ReplicationPartnerMetadata.csv") -NoTypeInformation

# --- Combine CSVs into Excel workbook ---
$workbook = Join-Path $OutputFolder "AD-Inventory-$ts.xlsx"
$csvs = Get-ChildItem -Path $OutputFolder -Filter *.csv -Recurse

# Prefer ImportExcel if Excel COM not available
$excelGrouped = $false
$excelComAvailable = $false
try {
    $excelCom = New-Object -ComObject Excel.Application -ErrorAction Stop
    # if this succeeds, immediately quit to avoid leaving Excel open
    $excelCom.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelCom) | Out-Null
    $excelComAvailable = $true
} catch {
    $excelComAvailable = $false
}

if (-not (Get-Command -Name Export-Excel -ErrorAction SilentlyContinue)) {
    # ImportExcel attempted earlier; if not available we already tried to install
    Ensure-Module -Name ImportExcel
}

foreach ($csv in $csvs) {
    $sheetName = ($csv.BaseName) -replace '[\\/:*?"<>|]','_'
    try {
        Import-Csv $csv.FullName | Export-Excel -Path $workbook -WorksheetName $sheetName -AutoSize -AutoFilter -Append -Verbose:$false
    } catch {
        Write-Verbose "Export-Excel failed for $($csv.Name): $_"
    }
}

# Try grouping columns with Excel COM if available (works in ISE when Excel is accessible; may require elevation)
if ($excelComAvailable) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Open($workbook)
        foreach ($ws in $wb.Worksheets) {
            $used = $ws.UsedRange
            if ($used -ne $null) {
                $colCount = $used.Columns.Count
                for ($c = 1; $c -le $colCount; $c++) {
                    $rng = $ws.Columns.Item($c)
                    $rng.Group()
                }
            }
        }
        $wb.Save()
        $wb.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excelGrouped = $true
    } catch {
        Write-Verbose "Excel COM grouping failed or not permitted in this session: $_"
    }
} else {
    Write-Verbose "Excel COM not available; grouping not applied. ImportExcel-created workbook still saved."
}

# --- Build condensed HTML summary ---
$htmlPath = Join-Path $OutputFolder "AD-Inventory-Summary.html"
$csvLinks = Get-ChildItem -Path $OutputFolder -Filter *.csv -Recurse | Sort-Object FullName
$gpoXmls = Get-ChildItem -Path $gpoReportsFolder -Filter *.xml -Recurse -ErrorAction SilentlyContinue
$repDotNote = ''
if (Test-Path $dotPath) {
    $repDotNote = "<p>Replication DOT file saved as <code>replication-topology.dot</code> in the output folder. Render manually with Graphviz if desired.</p>"
}

$privRows = Import-Csv (Join-Path $OutputFolder "Account-Tiers-ServiceFlags.csv") -ErrorAction SilentlyContinue

$html = @"
<!doctype html>
<html>
<head><meta charset='utf-8'><title>AD Inventory Summary - $ts</title>
<style>
  body{font-family:Segoe UI, Arial; margin:16px}
  table{border-collapse:collapse;width:100%}
  th,td{border:1px solid #ddd;padding:6px}
  .author { font-size:0.85em; color:#666; margin-top:-6px; margin-bottom:12px; }
  .small { font-size:0.85em; color:#666; }
</style>
</head>
<body>
<h1>AD Inventory Summary</h1>
<p class="author">Author: Stephen McKeee - IGTPLC</p>
<p class="small">Output folder: $OutputFolder</p>
<h2>Files</h2>
<ul>
"@
foreach ($c in $csvLinks) {
    $rel = $c.FullName -replace [regex]::Escape($OutputFolder + "\"), ''
    $html += "  <li><a href=`"$rel`">$rel</a></li>`n"
}
$html += @"
</ul>
<h2>GPO reports</h2>
<ul>
"@
foreach ($g in $gpoXmls) {
    $relg = $g.FullName -replace [regex]::Escape($OutputFolder + "\"), ''
    $html += "  <li><a href=`"$relg`">$($g.Name)</a></li>`n"
}
$html += @"
</ul>
$repDotNote
<h2>Privileged accounts (sample)</h2>
<table><thead><tr><th>SamAccountName</th><th>Name</th><th>DomainAdmin</th><th>Tier</th><th>IsServiceAccount</th></tr></thead><tbody>
"@
foreach ($r in $privRows | Select-Object -First 200) {
    $html += "  <tr><td>$($r.SamAccountName)</td><td>$($r.Name)</td><td>$($r.DomainAdmin)</td><td>$($r.Tier)</td><td>$($r.IsServiceAccount)</td></tr>`n"
}
$html += @"
</tbody></table>
<p class="small">Generated: $ts</p>
</body></html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8

# Summary CSV
[PSCustomObject]@{
    OutputFolder = $OutputFolder
    Workbook = $workbook
    ExcelColumnGroupingApplied = $excelGrouped
    HtmlSummary = $htmlPath
    DotFile = $dotPath
} | Export-Csv (Join-Path $OutputFolder "Summary.csv") -NoTypeInformation

Write-Host "Inventory complete. Files saved to: $OutputFolder"
if ($excelGrouped) { Write-Host "Excel workbook created and columns grouped: $workbook" } else { Write-Host "Excel workbook created (grouping not applied)." }
if (Test-Path $htmlPath) { Write-Host "HTML summary: $htmlPath" }
if (Test-Path $dotPath) { Write-Host "Replication DOT: $dotPath (no auto-render in ISE)"; }
