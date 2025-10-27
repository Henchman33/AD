<#
.SYNOPSIS
  Enterprise Active Directory / DNS / DHCP / GPO / Exchange / NTLM / Network Policy Inventory
.DESCRIPTION
  Builds detailed HTML and Excel reports, optimized for PS 5.1 with concurrency and scope controls.
  Adds a Summary dashboard at the top of the HTML report.
#>

#region Performance & Scope Settings
$RestrictOUs = @()                 # e.g. "OU=Servers,DC=contoso,DC=com"
$FastMode = $false                 # true = skip DNS record enumeration
$MaxDnsRecordsPerZone = 500        # 0 = unlimited
$IncludeSecondaryZones = $false    # include secondary zones
#endregion

#region Setup
$OutputPath = "$env:USERPROFILE\Desktop\AD_Inventory_$(Get-Date -Format yyyyMMdd_HHmmss)"
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
$global:FullReportBuilder = New-Object System.Text.StringBuilder
Import-Module ImportExcel -ErrorAction SilentlyContinue

function Add-ContentReport {
    param([string]$html,[switch]$LineBreak)
    [void]$global:FullReportBuilder.AppendLine($html)
    if ($LineBreak) { [void]$global:FullReportBuilder.AppendLine("<br/>") }
}

function HtmlEncode([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

$ReportHeaderHTML = @"
<html><head>
<title>AD Inventory Report</title>
<style>
body{font-family:'Segoe UI',sans-serif;background:#f8f8f8;color:#333;margin:20px;}
h1,h2{color:#003366;}
table{border-collapse:collapse;width:98%;margin:8px 0;}
th,td{border:1px solid #ccc;padding:5px;font-size:13px;}
th{background:#004080;color:#fff;}
tr:nth-child(even){background:#f2f2f2;}
tr:hover{background:#e6f3ff;}
.redFlag{background:#ffcccc;}
.details{background:#fff;padding:8px;border-radius:6px;border:1px solid #ccc;margin:8px 0}
.summaryBox{padding:10px;background:#fff;border:1px solid #ccc;border-radius:6px;margin-bottom:12px}
a.anchor-link{color:#004080;text-decoration:none}
</style></head><body>
<h1>Active Directory Inventory Report</h1>
<p><b>Generated:</b> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@

$ReportFooterHTML="</body></html>"

function Import-ModuleSafe {
  param([string]$Name)
  try {
    if (-not (Get-Module -ListAvailable -Name $Name)) {
      Write-Verbose "Module ${Name} not found."
      return $false
    }
    Import-Module -Name $Name -ErrorAction Stop
    return $true
  } catch {
    Write-Warning "Import-Module failed for ${Name}: $($_)"
    return $false
  }
}
#endregion

#region Import Modules
Import-ModuleSafe ActiveDirectory | Out-Null
Import-ModuleSafe DnsServer | Out-Null
Import-ModuleSafe DhcpServer | Out-Null
Import-ModuleSafe GroupPolicy | Out-Null
#endregion

#region Forest/Domain
try {
  Write-Output "Collecting Forest/Domain info..."
  $forest=Get-ADForest -ErrorAction SilentlyContinue
  $domain=Get-ADDomain -ErrorAction SilentlyContinue
  $forestObj=[PSCustomObject]@{
    ForestRoot = if ($forest) { $forest.RootDomain } else { '' }
    FunctionalLevel = if ($forest) { $forest.ForestMode } else { '' }
    Domains = if ($forest) { ($forest.Domains -join ', ') } else { '' }
    RecycleBin = if ($forest) { $forest.RecycleBinEnabled } else { $null }
  }
  $forestObj | Export-Csv (Join-Path $OutputPath 'forest_info.csv') -NoTypeInformation -Force
  Add-ContentReport "<details open class='details'><summary>Forest Information</summary>"
  Add-ContentReport ($forestObj | ConvertTo-Html -Fragment)
  Add-ContentReport "</details>"

  $domainObj=[PSCustomObject]@{
    Domain = if ($domain) { $domain.DNSRoot } else { '' }
    NetBIOS = if ($domain) { $domain.NetBIOSName } else { '' }
    FunctionalLevel = if ($domain) { $domain.DomainMode } else { '' }
  }
  $domainObj | Export-Csv (Join-Path $OutputPath 'domain_info.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>Domain Information</summary>"
  Add-ContentReport ($domainObj | ConvertTo-Html -Fragment)
  Add-ContentReport "</details>"
} catch { Write-Warning "Forest/Domain info failed: $($_)" }
#endregion

#region FSMO
try{
  Write-Output "Collecting FSMO roles..."
  $fsmo=[PSCustomObject]@{
    SchemaMaster = (Get-ADForest -ErrorAction SilentlyContinue).SchemaMaster
    DomainNamingMaster = (Get-ADForest -ErrorAction SilentlyContinue).DomainNamingMaster
    PDCEmulator = (Get-ADDomain -ErrorAction SilentlyContinue).PDCEmulator
    RIDMaster = (Get-ADDomain -ErrorAction SilentlyContinue).RIDMaster
    InfrastructureMaster = (Get-ADDomain -ErrorAction SilentlyContinue).InfrastructureMaster
  }
  $fsmo | Export-Csv (Join-Path $OutputPath 'fsmo.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>FSMO Roles</summary>"
  Add-ContentReport ($fsmo | ConvertTo-Html -Fragment)
  Add-ContentReport "</details>"
} catch { Write-Warning "FSMO failed: $($_)" }
#endregion

#region Domain Controllers
try{
  Write-Output "Collecting DCs..."
  $dcs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
  $dcData = $dcs | Select-Object HostName, @{n='IPv4Address';e={$_.IPv4Address -join ', '}}, Site, IsGlobalCatalog, IsReadOnly, OperatingSystem, OperatingSystemVersion
  if ($dcData.Count -eq 0) {
    $dcData = @([PSCustomObject]@{ HostName = '<No DCs>'; IPv4Address=''; Site=''; IsGlobalCatalog=''; IsReadOnly=''; OperatingSystem=''; OperatingSystemVersion='' })
  }
  $dcData | Export-Csv (Join-Path $OutputPath 'dcs.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>Domain Controllers</summary>"
  Add-ContentReport ($dcData | ConvertTo-Html -Fragment)
  Add-ContentReport "</details>"
} catch { Write-Warning "DCs failed: $($_)" }
#endregion

#region DNS (Concurrent via Jobs with throttling and fast mode)
$dnsRecordsAll = @()
$dnsZonesAll = @()
try {
  Write-Output "Collecting DNS zones and records (concurrent jobs)..."
  if ($FastMode) {
    # Only list zones (fast)
    if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
      $zones = Get-DnsServerZone -ErrorAction SilentlyContinue
      foreach ($z in $zones) {
        if (-not $IncludeSecondaryZones -and $z.ZoneType -ne 'Primary') { continue }
        $dnsZonesAll += [PSCustomObject]@{ ZoneName = $z.ZoneName; ZoneType = $z.ZoneType; MasterServer = 'Local' }
      }
      $dnsZonesAll | Export-Csv (Join-Path $OutputPath 'dns_zones.csv') -NoTypeInformation -Force
      Add-ContentReport "<details class='details'><summary>DNS Zones (FastMode)</summary>"
      Add-ContentReport (($dnsZonesAll | ConvertTo-Html -Fragment))
      Add-ContentReport "</details>"
    } else {
      Add-ContentReport "<details class='details'><summary>DNS</summary><p>DNS cmdlets not available on this host.</p></details>"
    }
  } else {
    # Build list of servers to query (Local + DC hostnames)
    $dnsServers = @()
    if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) { $dnsServers += 'Local' }
    foreach ($dc in $dcData) {
      if ($dc.HostName -and $dc.HostName -ne '<No DCs>') { $dnsServers += $dc.HostName }
    }
    $dnsServers = $dnsServers | Sort-Object -Unique

    $zoneJobs = @()
    foreach ($server in $dnsServers) {
      try {
        if ($server -eq 'Local') {
          $zones = Get-DnsServerZone -ErrorAction SilentlyContinue
        } else {
          $zones = Get-DnsServerZone -ComputerName $server -ErrorAction SilentlyContinue
        }
        if (-not $zones) { continue }

        foreach ($z in $zones) {
          if (-not $IncludeSecondaryZones -and $z.ZoneType -ne 'Primary') { continue }
          $dnsZonesAll += [PSCustomObject]@{ ZoneName = $z.ZoneName; ZoneType = $z.ZoneType; MasterServer = $server }

          # start job per zone
          $sb = {
            param($zoneName, $serverName, $maxRecords)
            $results = @()
            try {
              if ($serverName -eq 'Local') {
                $records = Get-DnsServerResourceRecord -ZoneName $zoneName -ErrorAction Stop
              } else {
                $records = Get-DnsServerResourceRecord -ZoneName $zoneName -ComputerName $serverName -ErrorAction Stop
              }
              $count = 0
              foreach ($r in $records) {
                $data = ''
                switch ($r.RecordType) {
                  'A' { if ($r.RecordData) { $data = $r.RecordData.IPv4Address.ToString() } }
                  'AAAA' { if ($r.RecordData) { $data = $r.RecordData.IPv6Address.ToString() } }
                  'CNAME' { if ($r.RecordData) { $data = $r.RecordData.HostNameAlias } }
                  'NS' { if ($r.RecordData) { $data = $r.RecordData.NameServer } }
                  'SRV' { if ($r.RecordData) { $data = "$($r.RecordData.DomainNameTarget):$($r.RecordData.Port)" } }
                  'PTR' { if ($r.RecordData) { $data = $r.RecordData.PtrDomainName } }
                  default { $data = ($r.RecordData | Out-String).Trim() }
                }
                $results += [PSCustomObject]@{ Zone = $zoneName; Host = $r.HostName; Type = $r.RecordType; Data = $data; Server = $serverName }
                $count++
                if ($maxRecords -gt 0 -and $count -ge $maxRecords) { break }
              }
            } catch {
              $results += [PSCustomObject]@{ Zone = $zoneName; Host = '<ERROR>'; Type = 'ERROR'; Data = $_.Exception.Message; Server = $serverName }
            }
            return $results
          }

          $job = Start-Job -ScriptBlock $sb -ArgumentList $z.ZoneName, $server, $MaxDnsRecordsPerZone
          $zoneJobs += $job
        }
      } catch {
        Write-Warning "Failed to enumerate zones on server ${server}: $($_)"
      }
    }

    # Throttle and collect jobs
    if ($zoneJobs.Count -gt 0) {
      Write-Output "Launched $($zoneJobs.Count) DNS jobs. Collecting results..."
      while ($zoneJobs.Count -gt 0) {
        # Wait for any job to complete with timeout
        $finished = Wait-Job -Any -Timeout 5 -ErrorAction SilentlyContinue
        $completed = $zoneJobs | Where-Object { $_.State -in @('Completed','Failed','Stopped') }
        foreach ($cj in $completed) {
          try {
            $res = Receive-Job -Job $cj -ErrorAction SilentlyContinue
            if ($res) { $dnsRecordsAll += $res }
          } catch { Write-Warning "Receive-Job failed for job Id $($cj.Id): $($_)" }
          finally {
            Remove-Job -Job $cj -Force -ErrorAction SilentlyContinue
            $zoneJobs = $zoneJobs | Where-Object { $_.Id -ne $cj.Id }
          }
        }
        Start-Sleep -Seconds 1
      }
    }

    # Export collected DNS info
    $dnsZonesAll | Export-Csv (Join-Path $OutputPath 'dns_zones.csv') -NoTypeInformation -Force
    $dnsRecordsAll | Export-Csv (Join-Path $OutputPath 'dns_records.csv') -NoTypeInformation -Force

    Add-ContentReport "<details class='details'><summary>DNS Zones</summary>"
    Add-ContentReport (($dnsZonesAll | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"

    Add-ContentReport "<details class='details'><summary>DNS Records (sample)</summary>"
    Add-ContentReport "<p>All DNS records exported to CSV. A sample (first 200 rows) shown below.</p>"
    Add-ContentReport (($dnsRecordsAll | Select-Object -First 200 | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
  }
} catch { Write-Warning "DNS enumeration failed: $($_)" }
#endregion

#region DHCP
try{
  Write-Output "Collecting DHCP servers..."
  $dhcp = @()
  if (Get-Command Get-DhcpServerInDC -ErrorAction SilentlyContinue) {
    $servers = Get-DhcpServerInDC -ErrorAction SilentlyContinue
    foreach ($s in $servers) {
      $dhcp += [PSCustomObject]@{ Name = $s.DnsName; IPAddress = $s.IPAddress; Source = 'Get-DhcpServerInDC' }
    }
  }
  if ($dhcp.Count -eq 0 -and (Get-Command Get-ADObject -ErrorAction SilentlyContinue)) {
    try {
      $cfg = (Get-ADRootDSE -ErrorAction SilentlyContinue).ConfigurationNamingContext
      $dhRoot = "CN=DhcpRoot,CN=NetServices,CN=Services,$cfg"
      $exists = Get-ADObject -Identity $dhRoot -ErrorAction SilentlyContinue
      if ($exists) {
        $children = Get-ADObject -SearchBase $dhRoot -Filter * -SearchScope OneLevel -Properties name,distinguishedName -ErrorAction SilentlyContinue
        foreach ($c in $children) { $dhcp += [PSCustomObject]@{ Name = $c.Name; IPAddress = ''; Source = 'AD DhcpRoot'; DN = $c.DistinguishedName } }
      }
    } catch {}
  }
  if ($dhcp.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
    try {
      $cands = Get-ADComputer -Filter 'Name -like "*dhcp*" -or Description -like "*dhcp*"' -Properties IPv4Address,Description -ErrorAction SilentlyContinue
      foreach ($c in $cands) { $dhcp += [PSCustomObject]@{ Name = $c.Name; IPAddress = $c.IPv4Address; Source = 'AD Name/Desc'; Description = $c.Description } }
    } catch {}
  }
  if ($dhcp.Count -eq 0) { $dhcp += [PSCustomObject]@{ Name = 'No DHCP servers found'; IPAddress = ''; Source = 'None' } }

  $dhcp | Export-Csv (Join-Path $OutputPath 'dhcp.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>DHCP Servers</summary>"
  Add-ContentReport (($dhcp | ConvertTo-Html -Fragment))
  Add-ContentReport "</details>"
} catch { Write-Warning "DHCP failed: $($_)" }
#endregion

#region Exchange (On-Prem)
try{
  Write-Output "Collecting Exchange servers (on-prem)..."
  $exchangeRecords = @()
  if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
    try { $exs = Get-ExchangeServer -ErrorAction SilentlyContinue | Select-Object Name,Edition,AdminDisplayVersion } catch { $exs = @() }
    foreach ($e in $exs) { $exchangeRecords += [PSCustomObject]@{ Name = $e.Name; Edition = $e.Edition; Version = $e.AdminDisplayVersion; Source = 'Get-ExchangeServer' } }
  }
  if ($exchangeRecords.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
    try { $exAD = Get-ADComputer -LDAPFilter "(msExchVersion=*)" -Properties msExchVersion,OperatingSystem -ErrorAction SilentlyContinue } catch { $exAD = @() }
    foreach ($c in $exAD) { $exchangeRecords += [PSCustomObject]@{ Name = $c.Name; Edition = ''; Version = $c.msExchVersion; Source = 'AD msExchVersion' } }
  }
  if ($exchangeRecords.Count -eq 0) { $exchangeRecords += [PSCustomObject]@{ Name = 'No Exchange servers found'; Edition = ''; Version = ''; Source = 'None' } }

  $exchangeRecords | Export-Csv (Join-Path $OutputPath 'exchange_servers.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>Exchange Servers</summary>"
  Add-ContentReport (($exchangeRecords | ConvertTo-Html -Fragment))
  Add-ContentReport "</details>"
} catch { Write-Warning "Exchange failed: $($_)" }
#endregion

#region AD Sites/Subnets
try{
  Write-Output "Collecting AD Sites/Subnets..."
  $sites = Get-ADReplicationSite -Filter * -ErrorAction SilentlyContinue
  $subs = Get-ADReplicationSubnet -Filter * -ErrorAction SilentlyContinue
  $siteData = @()
  foreach ($s in $sites) {
    $sSubs = ($subs | Where-Object { $_.Site -eq $s.Name }).Name -join ', '
    if ([string]::IsNullOrEmpty($sSubs)) { $sSubs = 'None' }
    $siteData += [PSCustomObject]@{ SiteName = $s.Name; Subnets = $sSubs }
  }
  if ($siteData.Count -eq 0) { $siteData += [PSCustomObject]@{ SiteName = '<No Sites>'; Subnets = '' } }
  $siteData | Export-Csv (Join-Path $OutputPath 'ad_sites_subnets.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>AD Sites and Subnets</summary>"
  Add-ContentReport (($siteData | ConvertTo-Html -Fragment))
  Add-ContentReport "</details>"
} catch { Write-Warning "Sites/Subnets failed: $($_)" }
#endregion

#region GPOs (NTLM + Network Audit) - Computer config only
try{
  Write-Output "Collecting GPOs and auditing NTLM/network computer settings..."
  $allGPOs = Get-GPO -All -ErrorAction SilentlyContinue
  if (-not $allGPOs) { $allGPOs = @() }
  $ntlmGpos = @()
  $netGpos = @()
  foreach ($g in $allGPOs) {
    try {
      $xml = [xml](Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction SilentlyContinue)
      if (-not $xml) { continue }
      $compNodes = $xml.SelectNodes("//Computer") 2>$null
      foreach ($cn in $compNodes) {
        # registry policies
        $regNodes = $cn.SelectNodes(".//RegistryPolicy") 2>$null
        foreach ($rn in $regNodes) {
          $name = ''
          $value = ''
          if ($rn.SelectSingleNode("Name")) { $name = $rn.SelectSingleNode("Name").InnerText }
          if ($rn.SelectSingleNode("Value")) { $value = $rn.SelectSingleNode("Value").InnerText }
          if ($name -and ($name -match 'NTLM' -or $name -match 'LmCompatibilityLevel' -or $name -match 'Restrict' )) {
            $ntlmGpos += [PSCustomObject]@{ GPO = $g.DisplayName; Setting = $name; Value = $value; GPOEnabled = $g.GpoStatus }
          }
          if ($name -and ($name -match '(?i)Network')) {
            $netGpos += [PSCustomObject]@{ GPO = $g.DisplayName; Setting = $name; Value = $value; GPOEnabled = $g.GpoStatus }
          }
        }
        # generic name/value pairs
        $nameNodes = $cn.SelectNodes(".//Name") 2>$null
        foreach ($nn in $nameNodes) {
          try {
            $parent = $nn.ParentNode
            $valNode = $parent.SelectSingleNode("Value")
            if (-not $valNode) { $valNode = $parent.SelectSingleNode("Properties/Value") }
            $nameText = ($nn.InnerText -as [string]).Trim()
            $valText = if ($valNode) { ($valNode.InnerText -as [string]).Trim() } else { '' }
            if ($nameText -and ($nameText -match 'NTLM' -or $nameText -match 'LmCompatibilityLevel' -or $nameText -match 'Restrict')) {
              $ntlmGpos += [PSCustomObject]@{ GPO = $g.DisplayName; Setting = $nameText; Value = $valText; GPOEnabled = $g.GpoStatus }
            }
            if ($nameText -and ($nameText -match '(?i)Network')) {
              $netGpos += [PSCustomObject]@{ GPO = $g.DisplayName; Setting = $nameText; Value = $valText; GPOEnabled = $g.GpoStatus }
            }
          } catch {}
        }
      }
    } catch { Write-Warning "Parse GPO $($g.DisplayName) failed: $($_)" }
  }

  if ($ntlmGpos.Count -eq 0) { $ntlmGpos += [PSCustomObject]@{ GPO = '<No NTLM GPOs found>'; Setting=''; Value=''; GPOEnabled='' } }
  if ($netGpos.Count -eq 0) { $netGpos += [PSCustomObject]@{ GPO = '<No Network GPOs found>'; Setting=''; Value=''; GPOEnabled='' } }

  $ntlmGpos | Export-Csv (Join-Path $OutputPath 'gpo_ntlm.csv') -NoTypeInformation -Force
  $netGpos | Export-Csv (Join-Path $OutputPath 'gpo_network.csv') -NoTypeInformation -Force

  Add-ContentReport "<details class='details'><summary>NTLM-related Computer GPO Settings</summary>"
  Add-ContentReport (($ntlmGpos | ConvertTo-Html -Fragment))
  Add-ContentReport "</details>"

  Add-ContentReport "<details class='details'><summary>Network-related Computer GPO Settings</summary>"
  Add-ContentReport (($netGpos | ConvertTo-Html -Fragment))
  Add-ContentReport "</details>"

} catch { Write-Warning "GPO Audit failed: $($_)" }
#endregion

#region Privileged Groups
try {
  $ent = Get-ADGroupMember "Enterprise Admins" -Recursive -ErrorAction SilentlyContinue | Select Name,SamAccountName
  $dom = Get-ADGroupMember "Domain Admins" -Recursive -ErrorAction SilentlyContinue | Select Name,SamAccountName
  $schema = Get-ADGroupMember "Schema Admins" -Recursive -ErrorAction SilentlyContinue | Select Name,SamAccountName
  $ent | Export-Csv (Join-Path $OutputPath 'enterprise_admins.csv') -NoTypeInformation -Force
  $dom | Export-Csv (Join-Path $OutputPath 'domain_admins.csv') -NoTypeInformation -Force
  $schema | Export-Csv (Join-Path $OutputPath 'schema_admins.csv') -NoTypeInformation -Force
  Add-ContentReport "<details class='details'><summary>Privileged Groups</summary>"
  Add-ContentReport "<h3>Enterprise Admins</h3>"
  Add-ContentReport ($ent | ConvertTo-Html -Fragment)
  Add-ContentReport "<h3>Domain Admins</h3>"
  Add-ContentReport ($dom | ConvertTo-Html -Fragment)
  Add-ContentReport "<h3>Schema Admins</h3>"
  Add-ContentReport ($schema | ConvertTo-Html -Fragment)
  Add-ContentReport "</details>"
} catch { Write-Warning "Privileged groups failed: $($_)" }
#endregion

#region Build Summary Dashboard (top of HTML)
# compute counts
$dcCount = if ($dcData) { ($dcData | Where-Object { $_.HostName -ne '<No DCs>' }).Count } else { 0 }
$dhcpCount = if ($dhcp) { ($dhcp | Where-Object { $_.Name -ne 'No DHCP servers found' }).Count } else { 0 }
$exchangeCount = if ($exchangeRecords) { ($exchangeRecords | Where-Object { $_.Name -ne 'No Exchange servers found' }).Count } else { 0 }
$gpoCount = if ($allGPOs) { ($allGPOs | Measure-Object).Count } else { 0 }

# count weak NTLM settings (LmCompatibilityLevel <= 2)
$ntlmWeakCount = 0
$weakList = @()
foreach ($row in $ntlmGpos) {
  if ($row.Value -match '^\s*(\d+)\s*$') {
    try {
      $v = [int]$Matches[1]
      if ($v -le 2) {
        $ntlmWeakCount++
        $weakList += $row
      }
    } catch {}
  }
}

$summaryHtml = @"
<div class='summaryBox'>
<h2>Summary</h2>
<table>
<tr><th>Item</th><th>Count / Note</th></tr>
<tr><td>Domain Controllers discovered</td><td>$dcCount</td></tr>
<tr><td>DHCP servers discovered</td><td>$dhcpCount</td></tr>
<tr><td>Exchange servers discovered</td><td>$exchangeCount</td></tr>
<tr><td>Total GPOs scanned</td><td>$gpoCount</td></tr>
<tr><td>NTLM-weak GPO settings flagged (LmCompatibilityLevel ≤ 2)</td><td>$ntlmWeakCount</td></tr>
</table>
</div>
"@

# If weak items exist, add quick links
if ($ntlmWeakCount -gt 0) {
  $summaryHtml += "<div class='details'><h3>Flagged NTLM Issues</h3><ul>"
  foreach ($w in $weakList) {
    $label = HtmlEncode("$($w.GPO) - $($w.Setting) = $($w.Value)")
    # no anchors created earlier, so just list text (we can add anchors per-row in a later iteration)
    $summaryHtml += "<li>$label</li>"
  }
  $summaryHtml += "</ul></div>"
}

# Prepend summary to the report builder (so it appears at top)
$global:FullReportBuilder.Insert(0, $summaryHtml)
#endregion

#region Finalize HTML + Excel export
$final = $ReportHeaderHTML + $global:FullReportBuilder.ToString() + $ReportFooterHTML
$reportFile = Join-Path $OutputPath 'FullReport.html'
$final | Out-File -FilePath $reportFile -Encoding UTF8 -Force
Write-Output "HTML report saved: $reportFile"

# Excel export if ImportExcel available
if (Get-Module -ListAvailable -Name ImportExcel) {
  try {
    $xlsx = Join-Path $OutputPath 'AD_Inventory_Report.xlsx'
    if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
    # assemble and write sheets (only if they exist)
    $summaryObj = @()
    $summaryObj += [PSCustomObject]@{ Item = 'Domain Controllers discovered'; Count = $dcCount }
    $summaryObj += [PSCustomObject]@{ Item = 'DHCP servers discovered'; Count = $dhcpCount }
    $summaryObj += [PSCustomObject]@{ Item = 'Exchange servers discovered'; Count = $exchangeCount }
    $summaryObj += [PSCustomObject]@{ Item = 'Total GPOs scanned'; Count = $gpoCount }
    $summaryObj += [PSCustomObject]@{ Item = 'NTLM-weak settings flagged'; Count = $ntlmWeakCount }
    $summaryObj | Export-Excel -Path $xlsx -WorksheetName 'Summary' -AutoSize -BoldTopRow

    if (Test-Path (Join-Path $OutputPath 'dcs.csv')) { Import-Csv (Join-Path $OutputPath 'dcs.csv') | Export-Excel -Path $xlsx -WorksheetName 'DomainControllers' -AutoSize -Append }
    if ($dnsZonesAll.Count -gt 0) { $dnsZonesAll | Export-Excel -Path $xlsx -WorksheetName 'DNS_Zones' -AutoSize -Append }
    if ($dnsRecordsAll.Count -gt 0) { $dnsRecordsAll | Export-Excel -Path $xlsx -WorksheetName 'DNS_Records' -AutoSize -Append }
    if (Test-Path (Join-Path $OutputPath 'dhcp.csv')) { Import-Csv (Join-Path $OutputPath 'dhcp.csv') | Export-Excel -Path $xlsx -WorksheetName 'DHCP_Servers' -AutoSize -Append }
    if (Test-Path (Join-Path $OutputPath 'exchange_servers.csv')) { Import-Csv (Join-Path $OutputPath 'exchange_servers.csv') | Export-Excel -Path $xlsx -WorksheetName 'Exchange_Servers' -AutoSize -Append }
    if (Test-Path (Join-Path $OutputPath 'ad_sites_subnets.csv')) { Import-Csv (Join-Path $OutputPath 'ad_sites_subnets.csv') | Export-Excel -Path $xlsx -WorksheetName 'AD_Sites' -AutoSize -Append }
    if (Test-Path (Join-Path $OutputPath 'gpo_ntlm.csv')) { Import-Csv (Join-Path $OutputPath 'gpo_ntlm.csv') | Export-Excel -Path $xlsx -WorksheetName 'GPO_NTLM' -AutoSize -Append }
    if (Test-Path (Join-Path $OutputPath 'gpo_network.csv')) { Import-Csv (Join-Path $OutputPath 'gpo_network.csv') | Export-Excel -Path $xlsx -WorksheetName 'GPO_Network' -AutoSize -Append }

    Write-Output "Excel workbook saved: $xlsx"
  } catch {
    Write-Warning "Excel export failed: $($_)"
  }
} else {
  Write-Warning "ImportExcel not available in this session; Excel export skipped."
}
#endregion

Write-Output "Report generation complete. Files in: $OutputPath"
Write-Output "Open HTML: $reportFile"
