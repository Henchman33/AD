<#
.SYNOPSIS
    AD Inventory + DNS + GPO Audit (NTLM/Network Computer settings) + Excel export
.DESCRIPTION
    Full inventory: Forest/Domain/FSMO/DCs/DNS/DHCP/Exchange/AD Sites/GPOs/Privileged Groups
    NTLM (computer config only) & Network-related computer GPO audits with link/enforced info,
    red-flag insecure NTLM settings, HTML + CSV + Excel outputs.
.NOTES
    Target: PowerShell 5.1 (Windows Server 2016/2019)
    Requires: ActiveDirectory, GroupPolicy, ImportExcel modules
    Author: Stephen McKee
#>

#region Configuration
$OutputPath = Join-Path $env:USERPROFILE ("Desktop\AD_Inventory_$(Get-Date -Format yyyyMMdd_HHmmss)")
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

$ExcelFile = Join-Path $OutputPath 'AD_Inventory_Report.xlsx'

# Canonical NTLM registry names (whitelist)
$NTLMRegistryKeys = @(
    "LmCompatibilityLevel",
    "RestrictIncomingNTLMTraffic",
    "RestrictOutgoingNTLM",
    "RestrictSendingNTLMTraffic",
    "AuditNTLMInDomain",
    "AuditNTLMInWorkgroup"
)

$NTLMPatterns = @('NTLM','NTLMv2','LmCompatibility','Restrict NTLM','Audit NTLM')
$NetworkPattern = 'Network'
$LmCompatibilityWarningThreshold = 2
#endregion

#region Helpers & HTML builder
$global:FullReportBuilder = New-Object System.Text.StringBuilder

function Add-ContentReport {
    param([Parameter(Mandatory=$true)]$html, [switch]$LineBreak)
    if ($html -is [System.Array]) { $html = ($html -join "`r`n") }
    [void]$global:FullReportBuilder.AppendLine($html)
    if ($LineBreak) { [void]$global:FullReportBuilder.AppendLine("<br/>") }
}

# No Add-Type for System.Net.WebUtility on PS5.1 — it's available from the runtime
function HtmlEncode([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

$ReportHeader = @"
<html><head><meta charset='utf-8'/>
<title>Active Diretory Inventory Report Version 9</title>
<title>Author: Stephen McKee </title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f8f8f8;color:#333;margin:20px}
h1,h2,h3{color:#003366}
table{border-collapse:collapse;width:98%;margin:8px 0}
th,td{border:1px solid #ccc;padding:6px 8px;font-size:13px}
th{background:#004080;color:#fff;text-align:left}
tr:nth-child(even){background:#f2f2f2}
tr:hover{background:#e6f3ff}
.bad{background:#ffd6d6}
.details{background:#fff;padding:8px;border-radius:6px;border:1px solid #ccc;margin:8px 0}
.summary{padding:8px;background:#fff;border:1px solid #ccc;border-radius:6px;margin-bottom:12px}
a.anchor-link{color:#004080;text-decoration:none}
</style></head><body>
<h1>Active Directory Inventory Report</h1>
<p><b>Generated:</b> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@

$ReportFooter = "</body></html>"

function Import-ModuleSafe {
    param([string]$Name)
    try {
        if (Get-Module -ListAvailable -Name $Name) {
            Import-Module -Name $Name -ErrorAction Stop
            return $true
        } else { return $false }
    } catch {
        return $false
    }
}
#endregion

#region Load modules
Import-ModuleSafe ActiveDirectory | Out-Null
Import-ModuleSafe GroupPolicy | Out-Null
Import-ModuleSafe DhcpServer | Out-Null
Import-ModuleSafe DnsServer | Out-Null
# ImportExcel assumed installed by you — load if available
Import-ModuleSafe ImportExcel | Out-Null
#endregion

#region Forest / Domain / FSMO
Write-Output "Collecting Forest/Domain/FSMO..."
try { $forest = Get-ADForest -ErrorAction Stop } catch { $forest = $null }
try { $domain = Get-ADDomain -ErrorAction Stop } catch { $domain = $null }

$forestObj = [PSCustomObject]@{
    ForestRootDomain = if ($forest) { $forest.RootDomain } else { '' }
    ForestFunctionalLevel = if ($forest) { $forest.ForestMode } else { '' }
    ForestDomains = if ($forest) { ($forest.Domains -join ', ') } else { '' }
    ADRecycleBinEnabled = if ($forest) { $forest.RecycleBinEnabled } else { $null }
}
$forestObj | Export-Csv (Join-Path $OutputPath 'forest_info.csv') -NoTypeInformation -Force
Add-ContentReport "<details class='details'><summary>Forest Information</summary>"
Add-ContentReport (($forestObj | ConvertTo-Html -Fragment))
Add-ContentReport "</details>"

try {
    $fsmo = [PSCustomObject]@{
        DomainNamingMaster = (Get-ADForest).DomainNamingMaster
        SchemaMaster = (Get-ADForest).SchemaMaster
        PDCEmulator = (Get-ADDomain).PDCEmulator
        RIDMaster = (Get-ADDomain).RIDMaster
        InfrastructureMaster = (Get-ADDomain).InfrastructureMaster
    }
    $fsmo | Export-Csv (Join-Path $OutputPath 'fsmo_roles.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>FSMO Roles</summary>"
    Add-ContentReport (($fsmo | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {}
#endregion

#region Domain Controllers
Write-Output "Collecting Domain Controllers..."
$dcRecords = @()
try {
    $dcCmd = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
} catch { $dcCmd = @() }
try {
    $dcAD = Get-ADComputer -Filter 'PrimaryGroupID -eq 516' -Properties IPv4Address,OperatingSystem -ErrorAction SilentlyContinue
} catch { $dcAD = @() }

$seen = @{}
foreach ($d in $dcCmd) {
    $key = ($d.Name -as [string]).ToLower()
    if ($key -and -not $seen.ContainsKey($key)) {
        $seen[$key] = $true
        $dcRecords += [PSCustomObject]@{
            Name = $d.Name
            HostName = $d.HostName
            Site = $d.Site
            IPv4 = ($d.IPv4Address -join ', ')
            IsGlobalCatalog = $d.IsGlobalCatalog
            IsReadOnly = $d.IsReadOnly
            OperatingSystem = $d.OperatingSystem
        }
    }
}
foreach ($c in $dcAD) {
    $key = ($c.Name -as [string]).ToLower()
    if ($key -and -not $seen.ContainsKey($key)) {
        $seen[$key] = $true
        $dcRecords += [PSCustomObject]@{
            Name = $c.Name
            HostName = $c.DNSHostName
            Site = ''
            IPv4 = ($c.IPv4Address -as [string])
            IsGlobalCatalog = $null
            IsReadOnly = $null
            OperatingSystem = $c.OperatingSystem
        }
    }
}
if ($dcRecords.Count -eq 0) {
    $dcRecords += [PSCustomObject]@{ Name = '<No Domain Controllers found>'; HostName=''; Site=''; IPv4=''; IsGlobalCatalog=''; IsReadOnly=''; OperatingSystem='' }
}
$dcRecords | Export-Csv (Join-Path $OutputPath 'domain_controllers.csv') -NoTypeInformation -Force
Add-ContentReport "<details class='details'><summary>Domain Controllers</summary>"
Add-ContentReport (($dcRecords | ConvertTo-Html -Fragment))
Add-ContentReport "</details>"
#endregion

#region DNS (zones + all records)
Write-Output "Collecting DNS zones and records..."
$dnsRecordsAll = @()
$dnsZonesAll = @()
$dnsServersTried = @()

try {
    if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
        # Local zones
        try {
            $localZones = Get-DnsServerZone -ErrorAction SilentlyContinue
        } catch { $localZones = @() }
        foreach ($z in $localZones) {
            $dnsZonesAll += [PSCustomObject]@{ ZoneName = $z.ZoneName; MasterServer = 'Local' }
        }
        foreach ($z in $localZones) {
            try {
                $recs = Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ErrorAction SilentlyContinue
                foreach ($r in $recs) {
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
                    $dnsRecordsAll += [PSCustomObject]@{ Zone = $z.ZoneName; Host = $r.HostName; Type = $r.RecordType; Data = $data; Server = 'Local' }
                }
            } catch {}
        }
        $dnsServersTried += 'Local'
    }

    # Try DCs as DNS servers
    foreach ($dc in $dcRecords) {
        $server = $dc.HostName
        if (-not $server) { $server = $dc.Name }
        if ($server -and -not ($dnsServersTried -contains $server)) {
            try {
                $zones = Get-DnsServerZone -ComputerName $server -ErrorAction SilentlyContinue
                if ($zones) {
                    foreach ($z in $zones) {
                        $dnsZonesAll += [PSCustomObject]@{ ZoneName = $z.ZoneName; MasterServer = $server }
                    }
                    foreach ($z in $zones) {
                        try {
                            $recs = Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ComputerName $server -ErrorAction SilentlyContinue
                            foreach ($r in $recs) {
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
                                $dnsRecordsAll += [PSCustomObject]@{ Zone = $z.ZoneName; Host = $r.HostName; Type = $r.RecordType; Data = $data; Server = $server }
                            }
                        } catch {}
                    }
                    $dnsServersTried += $server
                }
            } catch {}
        }
    }

    $dnsZonesAll = $dnsZonesAll | Sort-Object ZoneName -Unique
    $dnsZonesAll | Export-Csv (Join-Path $OutputPath 'dns_zones.csv') -NoTypeInformation -Force
    $dnsRecordsAll | Export-Csv (Join-Path $OutputPath 'dns_records.csv') -NoTypeInformation -Force

    Add-ContentReport "<details class='details'><summary>DNS Zones</summary>"
    Add-ContentReport (($dnsZonesAll | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"

    Add-ContentReport "<details class='details'><summary>DNS Records (sample)</summary>"
    Add-ContentReport "<p>All DNS records exported to CSV. A sample (first 200) is shown below.</p>"
    Add-ContentReport (($dnsRecordsAll | Select-Object -First 200 | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "DNS collection failed: $_"
    Add-ContentReport "<details class='details'><summary>DNS Information</summary><p>Error collecting DNS data: $($_.Exception.Message)</p></details>"
}
#endregion

#region DHCP
Write-Output "Collecting DHCP servers..."
$dhcpRecords = @()
try {
    if (Get-Command Get-DhcpServerInDC -ErrorAction SilentlyContinue) {
        try { $dh = Get-DhcpServerInDC -ErrorAction Stop } catch { $dh = @() }
        foreach ($s in $dh) { $dhcpRecords += [PSCustomObject]@{ Name = $s.DnsName; IPAddress = $s.IPAddress; Source='Get-DhcpServerInDC' } }
    }
    if ($dhcpRecords.Count -eq 0 -and (Get-Command Get-ADObject -ErrorAction SilentlyContinue)) {
        try {
            $cfg = (Get-ADRootDSE).ConfigurationNamingContext
            $dhRoot = "CN=DhcpRoot,CN=NetServices,CN=Services,$cfg"
            $exists = Get-ADObject -Identity $dhRoot -ErrorAction SilentlyContinue
            if ($exists) {
                $children = Get-ADObject -SearchBase $dhRoot -Filter * -SearchScope OneLevel -Properties name,distinguishedName -ErrorAction SilentlyContinue
                foreach ($c in $children) { $dhcpRecords += [PSCustomObject]@{ Name=$c.Name; IPAddress=''; Source='AD DhcpRoot'; DN=$c.DistinguishedName } }
            }
        } catch {}
    }
    if ($dhcpRecords.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
        try {
            $cands = Get-ADComputer -Filter 'Name -like "*dhcp*" -or Description -like "*dhcp*"' -Properties IPv4Address,Description -ErrorAction SilentlyContinue
            foreach ($c in $cands) { $dhcpRecords += [PSCustomObject]@{ Name=$c.Name; IPAddress=$c.IPv4Address; Source='AD Name/Desc'; Description=$c.Description } }
        } catch {}
    }
    if ($dhcpRecords.Count -eq 0) { $dhcpRecords += [PSCustomObject]@{ Name='No DHCP servers found'; IPAddress=''; Source='None' } }

    $dhcpRecords | Export-Csv (Join-Path $OutputPath 'dhcp_servers.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>DHCP Servers</summary>"
    Add-ContentReport (($dhcpRecords | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "DHCP collection failed: $_"
}
#endregion

#region Exchange
Write-Output "Collecting Exchange servers..."
$exchangeRecords = @()
try {
    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
        try { $exs = Get-ExchangeServer -ErrorAction Stop | Select-Object Name,Edition,AdminDisplayVersion } catch { $exs = @() }
        foreach ($e in $exs) { $exchangeRecords += [PSCustomObject]@{ Name=$e.Name; Edition=$e.Edition; Version=$e.AdminDisplayVersion; Source='Get-ExchangeServer' } }
    }
    if ($exchangeRecords.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
        try { $exAD = Get-ADComputer -LDAPFilter "(msExchVersion=*)" -Properties msExchVersion,OperatingSystem -ErrorAction SilentlyContinue } catch { $exAD = @() }
        foreach ($c in $exAD) { $exchangeRecords += [PSCustomObject]@{ Name=$c.Name; Edition=''; Version=$c.msExchVersion; Source='AD msExchVersion' } }
    }
    if ($exchangeRecords.Count -eq 0) { $exchangeRecords += [PSCustomObject]@{ Name='No Exchange servers found'; Edition=''; Version=''; Source='None' } }

    $exchangeRecords | Export-Csv (Join-Path $OutputPath 'exchange_servers.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>Exchange Servers</summary>"
    Add-ContentReport (($exchangeRecords | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "Exchange collection failed: $_"
}
#endregion

#region AD Sites + Subnets
Write-Output "Collecting AD Sites & Subnets..."
$siteData = @()
try {
    $sites = Get-ADReplicationSite -Filter * -ErrorAction SilentlyContinue
    $subnets = Get-ADReplicationSubnet -Filter * -ErrorAction SilentlyContinue
    foreach ($s in $sites) {
        $siteSubnets = $subnets | Where-Object { $_.Site -eq $s.Name }
        if ($siteSubnets -and $siteSubnets.Count -gt 0) { $subsVal = ($siteSubnets.Name -join '; ') } else { $subsVal = 'None' }
        $siteData += [PSCustomObject]@{ SiteName = $s.Name; Subnets = $subsVal; SubnetCount = if ($siteSubnets) { $siteSubnets.Count } else { 0 } }
    }
    if ($siteData.Count -eq 0) { $siteData += [PSCustomObject]@{ SiteName = '<No Sites>'; Subnets = ''; SubnetCount = 0 } }

    $siteData | Export-Csv (Join-Path $OutputPath 'ad_sites.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>AD Sites & Subnets</summary>"
    Add-ContentReport (($siteData | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "Sites/Subnets failed: $_"
}
#endregion

#region GPO Summary
Write-Output "Collecting GPO summary..."
try {
    $gpos = Get-GPO -All -ErrorAction SilentlyContinue | Select-Object DisplayName,Id,GpoStatus,ModificationTime
    if (-not $gpos) { $gpos = @() }
    $gpos | Export-Csv (Join-Path $OutputPath 'gpos.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>GPO Summary</summary>"
    Add-ContentReport (($gpos | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "GPO summary failed: $_"
}
#endregion

#region Function: Get-GpoLinkInfo
function Get-GpoLinkInfo {
    param([Parameter(Mandatory=$true)][string]$GpoId, [xml]$GpoXml)
    $links = @()
    try {
        if (Get-Command Get-GPInheritance -ErrorAction SilentlyContinue) {
            $ous = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName -ErrorAction SilentlyContinue
            foreach ($ou in $ous) {
                try {
                    $inh = Get-GPInheritance -Target $ou.DistinguishedName -ErrorAction SilentlyContinue
                    if ($inh -and $inh.GpoLinks) {
                        foreach ($gl in $inh.GpoLinks) {
                            if ($gl.GpoID -and ($gl.GpoID -eq $GpoId.Trim('{}'))) {
                                $links += [PSCustomObject]@{ LinkTarget = $ou.DistinguishedName; LinkDisabled = -not $gl.Enabled; LinkEnforced = $gl.Enforced; Source='Get-GPInheritance' }
                            }
                        }
                    }
                } catch {}
            }
        }
    } catch {}

    if ($links.Count -eq 0) {
        try {
            if (Get-Command Get-ADOrganizationalUnit -ErrorAction SilentlyContinue) {
                $plainGuid = $GpoId.Trim('{}')
                $escaped = [regex]::Escape($plainGuid)
                $ous = Get-ADOrganizationalUnit -Filter * -Properties gPLink -ErrorAction SilentlyContinue
                foreach ($ou in $ous) {
                    if ($ou.gPLink -and ($ou.gPLink -match $escaped)) {
                        $matches = [regex]::Matches($ou.gPLink, '\[(.*?)\]')
                        foreach ($m in $matches) {
                            $entry = $m.Groups[1].Value
                            $parts = $entry -split ';'
                            $opt = if ($parts.Count -ge 3) { $parts[-1] } else { '' }
                            $linkDisabled = $false
                            $linkEnforced = $false
                            if ($opt -match '^\s*2\s*$') { $linkDisabled = $true }
                            if ($opt -match '^\s*1\s*$') { $linkEnforced = $true }
                            $links += [PSCustomObject]@{ LinkTarget = $ou.DistinguishedName; LinkDisabled = $linkDisabled; LinkEnforced = $linkEnforced; Source='gPLink' }
                        }
                    }
                }
            }
        } catch {}
    }

    try {
        if ($GpoXml) {
            $lnodes = $GpoXml.SelectNodes("//LinksTo/Link") 2>$null
            foreach ($ln in $lnodes) {
                $pathNode = $ln.SelectSingleNode("Properties/Path")
                $enabledNode = $ln.SelectSingleNode("Properties/Enabled")
                $enforcedNode = $ln.SelectSingleNode("Properties/Enforced")
                $path = if ($pathNode) { $pathNode.InnerText } else { '' }
                $enabled = if ($enabledNode) { $enabledNode.InnerText } else { '' }
                $enforced = if ($enforcedNode) { $enforcedNode.InnerText } else { '' }
                $linkDisabled = $false
                if ($enabled -ne '' -and ($enabled -in @('false','0','False'))) { $linkDisabled = $true }
                $linkEnforced = $false
                if ($enforced -ne '' -and ($enforced -in @('true','1','True'))) { $linkEnforced = $true }
                $links += [PSCustomObject]@{ LinkTarget = $path; LinkDisabled = $linkDisabled; LinkEnforced = $linkEnforced; Source='LinksTo' }
            }
        }
    } catch {}

    if ($links.Count -eq 0) { $links += [PSCustomObject]@{ LinkTarget = 'NotLinked/Unknown'; LinkDisabled = $false; LinkEnforced = $false; Source='None' } }
    return $links
}
#endregion

#region NTLM & Network GPO audits (computer settings only)
Write-Output "Analyzing GPOs for NTLM and Network computer settings..."
$ntlmGpoAudit = @()
$networkGpoAudit = @()
$weakAnchors = @()

try {
    $allGPOs = Get-GPO -All -ErrorAction SilentlyContinue
    if (-not $allGPOs) { $allGPOs = @() }

    foreach ($gpo in $allGPOs) {
        try {
            $xmlText = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
            if (-not $xmlText) { continue }
            [xml]$xml = $xmlText

            $nameValuePairs = @()
            $compNodes = $xml.SelectNodes("//Computer") 2>$null
            foreach ($cn in $compNodes) {
                $regNodes = $cn.SelectNodes(".//RegistryPolicy") 2>$null
                foreach ($rn in $regNodes) {
                    $n = $null; $v = $null
                    if ($rn.SelectSingleNode("Name")) { $n = $rn.SelectSingleNode("Name").InnerText }
                    elseif ($rn.PSObject.Properties['Name']) { $n = $rn.Name }
                    if ($rn.SelectSingleNode("Value")) { $v = $rn.SelectSingleNode("Value").InnerText }
                    elseif ($rn.PSObject.Properties['Value']) { $v = $rn.Value }
                    if ($n) { $nameValuePairs += [PSCustomObject]@{ Name=$n; Value=$v } }
                }

                $nameNodes = $cn.SelectNodes(".//Name") 2>$null
                foreach ($nn in $nameNodes) {
                    try {
                        $parent = $nn.ParentNode
                        if ($parent -ne $null) {
                            $valNode = $parent.SelectSingleNode("Value")
                            if (-not $valNode) { $valNode = $parent.SelectSingleNode("Properties/Value") }
                            $nameText = ($nn.InnerText -as [string]).Trim()
                            $valText = if ($valNode) { ($valNode.InnerText -as [string]).Trim() } else { '' }
                            if ($nameText) { $nameValuePairs += [PSCustomObject]@{ Name=$nameText; Value=$valText } }
                        }
                    } catch {}
                }
            }

            $nameValuePairs = $nameValuePairs | Sort-Object Name -Unique
            $links = Get-GpoLinkInfo -GpoId ([string]$gpo.Id) -GpoXml $xml

            foreach ($nv in $nameValuePairs) {
                $name = $nv.Name; $value = $nv.Value
                $isNtlm = ($NTLMRegistryKeys -contains $name) -or ($NTLMPatterns | Where-Object { $name -match [regex]::Escape($_) })
                if ($isNtlm) {
                    foreach ($ln in $links) {
                        $isWeak = $false
                        if ($name -eq 'LmCompatibilityLevel') {
                            if ($value -match '^\d+$') {
                                try { if ([int]$value -le $LmCompatibilityWarningThreshold) { $isWeak = $true } } catch {}
                            }
                        }
                        $anchor = if ($isWeak) { "weak_$( [guid]::NewGuid().ToString() )" } else { '' }
                        $ntlmGpoAudit += [PSCustomObject]@{
                            Anchor = $anchor
                            GPOName = $gpo.DisplayName
                            GPOId = [string]$gpo.Id
                            Setting = $name
                            Value = $value
                            GPOEnabled = $gpo.GpoStatus
                            LinkTarget = $ln.LinkTarget
                            LinkEnabled = if ($ln.LinkDisabled) { 'No' } else { 'Yes' }
                            Enforced = if ($ln.LinkEnforced) { 'Yes' } else { 'No' }
                            IsWeak = $isWeak
                            Source = $ln.Source
                        }
                        if ($isWeak) { $weakAnchors += [PSCustomObject]@{ Anchor=$anchor; GPOName=$gpo.DisplayName; Setting=$name; Value=$value } }
                    }
                }

                $isNetwork = ($name -match "(?i)\b$NetworkPattern\b") -or ($value -match "(?i)\b$NetworkPattern\b")
                if ($isNetwork) {
                    foreach ($ln in $links) {
                        $networkGpoAudit += [PSCustomObject]@{
                            GPOName = $gpo.DisplayName
                            GPOId = [string]$gpo.Id
                            Setting = $name
                            Value = $value
                            GPOEnabled = $gpo.GpoStatus
                            LinkTarget = $ln.LinkTarget
                            LinkEnabled = if ($ln.LinkDisabled) { 'No' } else { 'Yes' }
                            Enforced = if ($ln.LinkEnforced) { 'Yes' } else { 'No' }
                            Source = $ln.Source
                        }
                    }
                }
            }

        } catch { Write-Warning "GPO parse error ($($gpo.DisplayName)): $($_.Exception.Message)" }
    }

} catch { Write-Warning "NTLM/Network audit overall error: $_" }

if ($ntlmGpoAudit.Count -eq 0) {
    $ntlmGpoAudit += [PSCustomObject]@{ Anchor=''; GPOName='<No NTLM GPOs found>'; GPOId=''; Setting=''; Value=''; GPOEnabled=''; LinkTarget=''; LinkEnabled=''; Enforced=''; IsWeak=$false; Source='' }
}
if ($networkGpoAudit.Count -eq 0) {
    $networkGpoAudit += [PSCustomObject]@{ GPOName='<No Network GPOs found>'; GPOId=''; Setting=''; Value=''; GPOEnabled=''; LinkTarget=''; LinkEnabled=''; Enforced=''; Source='' }
}

$ntlmGpoAudit | Export-Csv (Join-Path $OutputPath 'gpos_ntlm_computer_details.csv') -NoTypeInformation -Force
$networkGpoAudit | Export-Csv (Join-Path $OutputPath 'gpos_network_computer_details.csv') -NoTypeInformation -Force
#endregion

#region HTML: Summary, Issues, NTLM & Network sections, Suggestions
$dcCount = ($dcRecords | Where-Object { $_.Name -ne '<No Domain Controllers found>' }).Count
$dhcpCount = ($dhcpRecords | Where-Object { $_.Name -ne 'No DHCP servers found' }).Count
$exchangeCount = ($exchangeRecords | Where-Object { $_.Name -ne 'No Exchange servers found' }).Count
$gpoCount = ($gpos | Measure-Object).Count
$ntlmWeakCount = ($ntlmGpoAudit | Where-Object { $_.IsWeak -eq $true }).Count

$summaryHtml = @"
<div class='summary'>
<h2>Summary</h2>
<table>
<tr><th>Item</th><th>Count / Note</th></tr>
<tr><td>Domain Controllers discovered</td><td>$dcCount</td></tr>
<tr><td>DHCP servers discovered</td><td>$dhcpCount</td></tr>
<tr><td>Exchange servers discovered</td><td>$exchangeCount</td></tr>
<tr><td>Total GPOs</td><td>$gpoCount</td></tr>
<tr><td>NTLM-weak GPO settings flagged</td><td>$ntlmWeakCount</td></tr>
</table>
</div>
"@
Add-ContentReport $summaryHtml

# Anchors for flagged items
if ($ntlmWeakCount -gt 0) {
    $anchorsHtml = "<div class='details'><h3>Flagged Insecure NTLM Settings</h3><ul>"
    foreach ($a in ($ntlmGpoAudit | Where-Object { $_.IsWeak -eq $true })) {
        $label = HtmlEncode("$($a.GPOName) - $($a.Setting) = $($a.Value)")
        if ($a.Anchor -and $a.Anchor -ne '') {
            $anchorsHtml += "<li><a class='anchor-link' href='#$($a.Anchor)'>$label</a></li>"
        } else {
            $anchorsHtml += "<li>$label</li>"
        }
    }
    $anchorsHtml += "</ul></div>"
    Add-ContentReport $anchorsHtml
}

# NTLM table
$ntlmHtml = "<details class='details'><summary>NTLM & NTLMv2 (Computer Configuration) GPO Audit</summary><p>Only Computer config. Red rows = insecure LmCompatibilityLevel (<= $LmCompatibilityWarningThreshold).</p><table><tr><th>GPO Name</th><th>GPOId</th><th>Setting</th><th>Value</th><th>GPOEnabled</th><th>LinkTarget</th><th>LinkEnabled</th><th>Enforced</th><th>Source</th></tr>"
foreach ($r in $ntlmGpoAudit) {
    $cls = if ($r.IsWeak) { " class='bad'" } else { "" }
    $anchorAttr = if ($r.IsWeak -and $r.Anchor) { " id='$($r.Anchor)'" } else { "" }
    $ntlmHtml += "<tr$cls$anchorAttr><td>$(HtmlEncode($r.GPOName))</td><td>$(HtmlEncode($r.GPOId))</td><td>$(HtmlEncode($r.Setting))</td><td>$(HtmlEncode($r.Value))</td><td>$(HtmlEncode($r.GPOEnabled))</td><td>$(HtmlEncode($r.LinkTarget))</td><td>$(HtmlEncode($r.LinkEnabled))</td><td>$(HtmlEncode($r.Enforced))</td><td>$(HtmlEncode($r.Source))</td></tr>"
}
$ntlmHtml += "</table></details>"
Add-ContentReport $ntlmHtml

# Network table
$netHtml = "<details class='details'><summary>Network-related Computer Configuration GPO Audit</summary><table><tr><th>GPO Name</th><th>GPOId</th><th>Setting</th><th>Value</th><th>GPOEnabled</th><th>LinkTarget</th><th>LinkEnabled</th><th>Enforced</th><th>Source</th></tr>"
foreach ($r in $networkGpoAudit) {
    $netHtml += "<tr><td>$(HtmlEncode($r.GPOName))</td><td>$(HtmlEncode($r.GPOId))</td><td>$(HtmlEncode($r.Setting))</td><td>$(HtmlEncode($r.Value))</td><td>$(HtmlEncode($r.GPOEnabled))</td><td>$(HtmlEncode($r.LinkTarget))</td><td>$(HtmlEncode($r.LinkEnabled))</td><td>$(HtmlEncode($r.Enforced))</td><td>$(HtmlEncode($r.Source))</td></tr>"
}
$netHtml += "</table></details>"
Add-ContentReport $netHtml

# Suggestions section
$suggestionsHtml = @"
<details class='details'><summary>Suggestions & Remediation (quick)</summary>
<h3>Top suggestions</h3>
<ul>
<li>Raise <b>LmCompatibilityLevel</b> to 5 in GPOs where safe (NTLMv2-only). Verify application compatibility first.</li>
<li>Enable auditing (AuditNTLM*) before enforcing restrictions so you can see impact.</li>
<li>Review GPO links that are disabled or not enforced; remove stale or conflicting links.</li>
<li>Authorize DHCP servers in AD if using Windows DHCP; verify DHCP servers listed in the DhcpRoot container.</li>
<li>Ensure Exchange servers are discovered by running this on a host with Exchange management tools or via implicit remoting to Exchange.</li>
<li>Use the Excel sheet 'NTLM_GPOs' to filter and triage flagged items; remediation guidance included in CSV/Excel exports.</li>
</ul>
</details>
"@
Add-ContentReport $suggestionsHtml
#endregion

#region Privileged groups
Write-Output "Collecting privileged groups..."
try {
    $ent = Get-ADGroupMember "Enterprise Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name,SamAccountName
    $dom = Get-ADGroupMember "Domain Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name,SamAccountName
    $schema = Get-ADGroupMember "Schema Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name,SamAccountName
    $ent | Export-Csv (Join-Path $OutputPath 'enterprise_admins.csv') -NoTypeInformation -Force
    $dom | Export-Csv (Join-Path $OutputPath 'domain_admins.csv') -NoTypeInformation -Force
    $schema | Export-Csv (Join-Path $OutputPath 'schema_admins.csv') -NoTypeInformation -Force
    Add-ContentReport "<details class='details'><summary>Privileged Groups</summary>"
    Add-ContentReport "<h3>Enterprise Admins</h3>"
    Add-ContentReport (($ent | ConvertTo-Html -Fragment))
    Add-ContentReport "<h3>Domain Admins</h3>"
    Add-ContentReport (($dom | ConvertTo-Html -Fragment))
    Add-ContentReport "<h3>Schema Admins</h3>"
    Add-ContentReport (($schema | ConvertTo-Html -Fragment))
    Add-ContentReport "</details>"
} catch {
    Write-Warning "Privileged groups collection failed: $_"
}
#endregion

#region Finalize HTML file
$FullHtml = $ReportHeader + $global:FullReportBuilder.ToString() + $ReportFooter
$FullHtmlPath = Join-Path $OutputPath 'FullReport.html'
$FullHtml | Out-File -FilePath $FullHtmlPath -Encoding UTF8 -Force
Write-Output "HTML report saved to: $FullHtmlPath"
#endregion

#region Excel export (ImportExcel)
if (Get-Module -ListAvailable -Name ImportExcel) {
    try {
        if (Test-Path $ExcelFile) { Remove-Item $ExcelFile -Force }
        # Build summary data
        $summaryObj = @()
        $summaryObj += [PSCustomObject]@{ Item='Domain Controllers discovered'; Count = $dcCount }
        $summaryObj += [PSCustomObject]@{ Item='DHCP servers discovered'; Count = $dhcpCount }
        $summaryObj += [PSCustomObject]@{ Item='Exchange servers discovered'; Count = $exchangeCount }
        $summaryObj += [PSCustomObject]@{ Item='Total GPOs'; Count = $gpoCount }
        $summaryObj += [PSCustomObject]@{ Item='NTLM-weak settings flagged'; Count = $ntlmWeakCount }

        $summaryObj | Export-Excel -Path $ExcelFile -WorksheetName 'Summary' -AutoSize -BoldTopRow
        $dcRecords | Export-Excel -Path $ExcelFile -WorksheetName 'DomainControllers' -AutoSize -Append
        $dnsZonesAll | Export-Excel -Path $ExcelFile -WorksheetName 'DNS_Zones' -AutoSize -Append
        $dnsRecordsAll | Export-Excel -Path $ExcelFile -WorksheetName 'DNS_Records' -AutoSize -Append
        $dhcpRecords | Export-Excel -Path $ExcelFile -WorksheetName 'DHCP_Servers' -AutoSize -Append
        $exchangeRecords | Export-Excel -Path $ExcelFile -WorksheetName 'Exchange_Servers' -AutoSize -Append
        $siteData | Export-Excel -Path $ExcelFile -WorksheetName 'AD_Sites' -AutoSize -Append
        $gpos | Export-Excel -Path $ExcelFile -WorksheetName 'GPO_Summary' -AutoSize -Append
        $ntlmGpoAudit | Export-Excel -Path $ExcelFile -WorksheetName 'NTLM_GPOs' -AutoSize -Append
        $networkGpoAudit | Export-Excel -Path $ExcelFile -WorksheetName 'Network_GPOs' -AutoSize -Append

        Write-Output "Excel workbook written to: $ExcelFile"
    } catch {
        Write-Warning "Excel export failed: $($_.Exception.Message)"
    }
} else {
    Write-Warning "ImportExcel module not detected in this session — Excel export skipped. CSVs and HTML were generated."
}
#endregion

Write-Output "`nAll outputs saved to: $OutputPath"
Write-Output "Full HTML report: $FullHtmlPath"
if (Test-Path $ExcelFile) { Write-Output "Excel workbook: $ExcelFile" }
