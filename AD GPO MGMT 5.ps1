<#
.SYNOPSIS
    AD GPOS Domain Report - enumerates all GPOs in domain, Domain Controllers list, Domain Controllers OU GPOs,
    NTLM/NTLMv2-related GPOs and settings. Exports CSV, XLSX and HTML (collapsible sections, NTLM highlight).

.DESCRIPTION
    Designed for PowerShell (ISE) 5.1. Uses GroupPolicy and ActiveDirectory modules. Produces:
      - AD_GPOS_Domain_Report_AllGPOs.csv
      - AD_GPOS_Domain_Report_AllGPOs.xlsx
      - AD_GPOS_Domain_Report.html

.NOTES
    Run elevated. Requires RSAT: GroupPolicy and ActiveDirectory modules.

#>

### ---------- Parameters / output location ----------
$ReportNameBase = "AD GPOS Domain Report"
$TimeStamp = (Get-Date).ToString("yyyyMMdd_HHmm")
$OutFolder = Join-Path -Path (Get-Location) -ChildPath "AD_GPO_Report_$TimeStamp"
New-Item -Path $OutFolder -ItemType Directory -Force | Out-Null

$CsvPath = Join-Path $OutFolder "$($ReportNameBase -replace ' ','_')_AllGPOs.csv"
$ExcelPath = Join-Path $OutFolder "$($ReportNameBase -replace ' ','_')_AllGPOs.xlsx"
$HtmlPath = Join-Path $OutFolder "$($ReportNameBase -replace ' ','_').html"

### ---------- Basic prerequisite checks ----------
function Ensure-Module {
    param($ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Throw "Module '$ModuleName' is not available. Install RSAT/appropriate module and re-run. (`Install-WindowsFeature` / `Add-WindowsCapability` or install RSAT for your OS). Module: $ModuleName"
    }
}
# Need GroupPolicy and ActiveDirectory
Ensure-Module -ModuleName "GroupPolicy"
Ensure-Module -ModuleName "ActiveDirectory"

Import-Module GroupPolicy -ErrorAction Stop
Import-Module ActiveDirectory -ErrorAction Stop

### ---------- Helper functions ----------
function Get-GPO-XmlReport {
    param([string]$GPOId)
    # returns [xml] doc or $null
    try {
        $xmlString = Get-GPOReport -Guid $GPOId -ReportType Xml -ErrorAction Stop
        [xml]$doc = $xmlString
        return $doc
    } catch {
        Write-Warning "Failed to get XML report for GPO $GPOId : $_"
        return $null
    }
}

function Parse-GPO-LinksFromXml {
    param([xml]$GpoXml)
    # Returns array of PSCustomObjects: @{SOM; Enforced; LinkEnabled; LinkOrder}
    $links = @()
    if (-not $gpoXml) { return $links }
    # XML contains //LinksTo/Link elements
    $linkNodes = $gpoXml.SelectNodes("//LinksTo/Link")
    if ($linkNodes) {
        foreach ($ln in $linkNodes) {
            # Link attributes differ between schema versions; attempt to gather common ones
            $sOM = $ln.SOM -as [string]
            if (-not $sOM) { $sOM = $ln.GetAttribute("SOM") 2>$null }
            $enabled = $ln.enabled
            if ($enabled -eq $null) { $enabled = $ln.GetAttribute("Enabled") 2>$null }
            $enforced = $ln.enforced
            if ($enforced -eq $null) { $enforced = $ln.GetAttribute("Enforced") 2>$null }
            $linkOrder = $ln.linkOrder
            if ($linkOrder -eq $null) { $linkOrder = $ln.GetAttribute("LinkOrder") 2>$null }

            $links += [PSCustomObject]@{
                SOM = $sOM
                LinkEnabled = if ($enabled -eq $null) { $true } else { [bool]([string]$enabled -eq "true") }
                Enforced = if ($enforced -eq $null) { $false } else { [bool]([string]$enforced -eq "true") }
                LinkOrder = ($linkOrder -as [int])
            }
        }
    }
    return $links
}

function Get-WmiFilterFromXml {
    param([xml]$GpoXml)
    if (-not $gpoXml) { return $null }
    $node = $gpoXml.SelectSingleNode("//WmiFilter")
    if ($node) {
        return [PSCustomObject]@{
            Id = ($node.id -as [string])
            Name = ($node.name -as [string])
            Query = ($node.query -as [string])
        }
    } else { return $null }
}

function Find-NTLM-Settings-In-GPOXml {
    param([xml]$GpoXml)
    # Search for common NTLM/NTLMv2 settings inside the GPO XML
    # We'll look for registry values named LmCompatibilityLevel and textual policy names that mention "LAN Manager" or "Restrict NTLM"
    $results = @()
    if (-not $gpoXml) { return $results }

    # 1) RegistryValue LmCompatibilityLevel
    $regNodes = $gpoXml.SelectNodes("//RegistryValue[translate(name,'[]','')]")
    if ($regNodes) {
        foreach ($r in $regNodes) {
            $name = $r.Name
            if ($name -match "LmCompatibilityLevel") {
                $val = $r.Value
                $hive = $r.Hive
                $key = $r.KeyName
                $results += [PSCustomObject]@{
                    SettingType = "Registry"
                    SettingName = $name
                    Value = $val
                    Hive = $hive
                    Key = $key
                    FoundIn = "RegistryValue"
                }
            }
        }
    }

    # 2) Look for Security Options textual names in the XML (policy strings)
    $textNodes = $gpoXml.SelectNodes("//Policy[@name] | //Setting")
    if ($textNodes) {
        foreach ($t in $textNodes) {
            $txt = ($t.Name -as [string]) + " " + ($t.Text -as [string])
            if ($txt -match "LAN Manager" -or $txt -match "LANManager" -or $txt -match "Restrict NTLM" -or $txt -match "LMCompatibility" -or $txt -match "Network security:.*LAN Manager") {
                # try to extract a value
                $value = $t.Value
                $results += [PSCustomObject]@{
                    SettingType = "PolicySetting"
                    SettingName = ($t.Name -as [string])
                    Value = ($t.Value -as [string])
                    Hive = ""
                    Key = ""
                    FoundIn = "Policy/Setting"
                }
            }
        }
    }

    # 3) Advanced: look anywhere in XML text content for known phrases
    $allText = $gpoXml.InnerXml
    if ($allText -match "LmCompatibilityLevel") {
        # If not already present, add a marker
        if (-not ($results | Where-Object { $_.SettingName -match "LmCompatibilityLevel" })) {
            $results += [PSCustomObject]@{
                SettingType = "Registry"
                SettingName = "LmCompatibilityLevel"
                Value = "(found text, value may be in registry nodes)"
                Hive = ""
                Key = ""
                FoundIn = "InnerXml"
            }
        }
    }

    return $results
}

### ---------- Get domain & domain controllers info ----------
$ADDomain = Get-ADDomain -ErrorAction Stop
$DomainDN = $ADDomain.DistinguishedName
$DomainName = $ADDomain.DNSRoot

# Domain controllers
$DCs = Get-ADDomainController -Filter * | Select-Object Name,HostName,IPv4Address,OperatingSystem,IsGlobalCatalog,Site,IsReadOnly,Domain

# FSMO role holders
$DomainObj = Get-ADDomain
$ForestObj = Get-ADForest
$FSMO = [PSCustomObject]@{
    PDCEmulator = $DomainObj.PDCEmulator
    RIDMaster = $DomainObj.RIDMaster
    InfrastructureMaster = $DomainObj.InfrastructureMaster
    DomainNamingMaster = $ForestObj.DomainNamingMaster
    SchemaMaster = $ForestObj.SchemaMaster
}

### ---------- Gather all GPOs & details ----------
Write-Host "Collecting all GPOs and parsing their XML..." -ForegroundColor Cyan
$GPOs = Get-GPO -All

$AllGpoDetails = @()
$GpoNtlmHits = @()

foreach ($g in $GPOs) {
    $xml = Get-GPO-XmlReport -GPOId $g.Id
    $links = Parse-GPO-LinksFromXml -GpoXml $xml
    $wmi = Get-WmiFilterFromXml -GpoXml $xml
    $ntlmSettings = Find-NTLM-Settings-In-GPOXml -GpoXml $xml

    # Build textual summary of links
    $linksSummary = if ($links.Count -gt 0) {
        ($links | ForEach-Object {
            "$($_.SOM) [Enabled:$($_.LinkEnabled); Enforced:$($_.Enforced); Order:$($_.LinkOrder)]"
        }) -join " ; "
    } else { "" }

    $gpoObj = [PSCustomObject]@{
        DisplayName = $g.DisplayName
        Id = $g.Id.Guid
        Domain = $g.DomainName
        Owner = $g.Owner
        CreationTime = $g.CreationTime
        ModificationTime = $g.ModificationTime
        GpoStatus = $g.GpoStatus
        SysvolPath = $g.SysvolPath
        WmiFilterId = if ($wmi) { $wmi.Id } else { $null }
        WmiFilterName = if ($wmi) { $wmi.Name } else { $null }
        WmiFilterQuery = if ($wmi) { ($wmi.Query -replace "\r?\n","\n") } else { $null }
        Links = $linksSummary
        LinksObject = $links
        NTLMSettingsFound = if ($ntlmSettings.Count -gt 0) { $true } else { $false }
        NTLMSettings = ($ntlmSettings | ForEach-Object {
            # compact textual
            ("{0} | {1} = {2}" -f $_.SettingType, $_.SettingName, $_.Value)
        }) -join " ; "
    }
    $AllGpoDetails += $gpoObj

    if ($ntlmSettings.Count -gt 0) {
        $GpoNtlmHits += [PSCustomObject]@{
            DisplayName = $g.DisplayName
            Id = $g.Id.Guid
            Domain = $g.DomainName
            NTLMSettings = ($ntlmSettings | ForEach-Object { ("{0} | {1} = {2}" -f $_.SettingType, $_.SettingName, $_.Value) }) -join " ; "
            LinksSummary = $linksSummary
            WmiFilterName = if ($wmi) { $wmi.Name } else { $null }
        }
    }
}

### ---------- Get GPO links specifically for "top of domain" (domain root) and "Domain Controllers OU" ----------
# Top of domain: target the domain DN
Write-Host "Retrieving GP inheritance for the domain root and the Domain Controllers OU..." -ForegroundColor Cyan
try {
    $DomainInheritance = Get-GPInheritance -Target $DomainDN -ErrorAction Stop
} catch {
    Write-Warning "Get-GPInheritance for domain root failed: $_"
    $DomainInheritance = $null
}

# Domain Controllers OU DN
$DCDN = "OU=Domain Controllers,$DomainDN"
try {
    $DCInheritance = Get-GPInheritance -Target $DCDN -ErrorAction Stop
} catch {
    Write-Warning "Get-GPInheritance for Domain Controllers OU ($DCDN) failed: $_"
    $DCInheritance = $null
}

# Helper to normalize Get-GPInheritance output into objects
function Convert-GPInheritanceToTable {
    param($gpInheritance)
    $out = @()
    if (-not $gpInheritance) { return $out }
    foreach ($ap in $gpInheritance.AppliedGpo) {
        $out += [PSCustomObject]@{
            DisplayName = $ap.DisplayName
            Id = $ap.GpoId.Guid
            LinkEnabled = $ap.LinkEnabled
            Enforced = $ap.Enforced
            GpoStatus = $ap.GpoStatus
            EnforcedLink = $ap.Enforced
            # the AppliedGpo object contains other useful properties, but above are the main ones
        }
    }
    return $out
}

$TopDomainGpoLinks = Convert-GPInheritanceToTable -gpInheritance $DomainInheritance
$DCouGpoLinks = Convert-GPInheritanceToTable -gpInheritance $DCInheritance

### ---------- Prepare CSV export - flattened ----------
# For CSV/Excel we want one row per GPO with a textified Links field
$CsvRows = $AllGpoDetails | Select-Object DisplayName,Id,Domain,Owner,CreationTime,ModificationTime,GpoStatus,SysvolPath,WmiFilterName,WmiFilterId,Links,NTLMSettingsFound,NTLMSettings

Write-Host "Exporting CSV to $CsvPath" -ForegroundColor Green
$CsvRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

### ---------- Excel export - try to use ImportExcel, otherwise COM ----------
function Export-ToExcel {
    param($DataTable, [string]$Path)
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Try {
            Import-Module ImportExcel -ErrorAction Stop
            $DataTable | Export-Excel -Path $Path -WorksheetName 'AllGPOs' -AutoSize -AutoFilter
            Write-Host "Excel exported via ImportExcel to $Path" -ForegroundColor Green
        } Catch {
            Write-Warning "Export-Excel failed: $_. Falling back to COM method."
            Export-ToExcel-Com -DataTable $DataTable -Path $Path
        }
    } else {
        Export-ToExcel-Com -DataTable $DataTable -Path $Path
    }
}

function Export-ToExcel-Com {
    param($DataTable, [string]$Path)
    # Minimal COM-based Excel creation - will require Excel on the machine.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    # write headers
    $col = 1
    foreach ($h in $DataTable[0].psobject.properties.name) {
        $ws.Cells.Item(1,$col).Value2 = $h
        $col++
    }
    $row = 2
    foreach ($r in $DataTable) {
        $col = 1
        foreach ($p in $r.psobject.properties) {
            $ws.Cells.Item($row,$col).Value2 = ($p.Value -as [string])
            $col++
        }
        $row++
    }
    $wb.SaveAs((Resolve-Path $Path).ProviderPath)
    $wb.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Write-Host "Excel exported via COM to $Path" -ForegroundColor Green
}

# Call it
if ($CsvRows.Count -gt 0) {
    Export-ToExcel -DataTable $CsvRows -Path $ExcelPath
} else {
    Write-Warning "No GPO rows to export to Excel."
}

### ---------- Build HTML report with collapsible sections ----------
Write-Host "Building HTML report..." -ForegroundColor Cyan

# CSS and JS for collapsible sections and highlighting
$css = @"
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px;}
h1 { font-size: 1.6em; }
.section { margin-bottom: 1em; border:1px solid #ddd; border-radius:6px; padding:10px; }
.section-header { cursor:pointer; padding:6px 0; display:flex; justify-content:space-between; align-items:center; }
.section-content { padding:8px 4px; display:block; }
.table { border-collapse:collapse; width:100%; }
.table th, .table td { border:1px solid #ddd; padding:6px; text-align:left; font-size:0.9em; }
.table th { background:#f5f5f5; }
.ntlm { background-color: #fff2cc; } /* highlight NTLM GPO rows */
.small { font-size:0.85em; color:#666; }
.toggle-btn { background:#0078d4; color:white; padding:6px 10px; border-radius:4px; text-decoration:none; }
"@

$js = @"
function toggle(id) {
    var el = document.getElementById(id);
    if (!el) return;
    if (el.style.display === 'none') { el.style.display = 'block'; } else { el.style.display = 'none'; }
}
function toggleAll(open) {
    var c = document.getElementsByClassName('section-content');
    for (var i=0;i<c.length;i++){ c[i].style.display = open ? 'block' : 'none'; }
}
"@

# Helper to create HTML tables from objects
function Convert-ObjectsToHtmlTable {
    param([Parameter(Mandatory=$true)][array]$Objects, [string[]]$Columns)
    if (-not $Objects -or $Objects.Count -eq 0) {
        return "<div class='small'>No items</div>"
    }
    if (-not $Columns) { $Columns = $Objects[0].psobject.Properties.Name }
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("<table class='table'>") | Out-Null
    $sb.AppendLine("<thead><tr>") | Out-Null
    foreach ($c in $Columns) { $sb.AppendLine("<th>" + [System.Web.HttpUtility]::HtmlEncode($c) + "</th>") | Out-Null }
    $sb.AppendLine("</tr></thead>") | Out-Null
    $sb.AppendLine("<tbody>") | Out-Null
    foreach ($o in $Objects) {
        $rowClass = ""
        if ($o.NTLMSettingsFound -eq $true -or ($o.DisplayName -and $o.DisplayName -match "NTLM|LMCompatibility|LAN Manager")) { $rowClass = "class='ntlm'" }
        $sb.AppendLine("<tr $rowClass>") | Out-Null
        foreach ($c in $Columns) {
            $val = $null
            try { $val = $o.$c } catch { $val = $null }
            if ($val -eq $null) { $val = "" }
            $enc = [System.Web.HttpUtility]::HtmlEncode($val.ToString())
            $sb.AppendLine("<td>$enc</td>") | Out-Null
        }
        $sb.AppendLine("</tr>") | Out-Null
    }
    $sb.AppendLine("</tbody></table>") | Out-Null
    return $sb.ToString()
}

# Build sections
# 1) Overview header
$header = "<h1>$ReportNameBase</h1><div class='small'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;| Domain: $DomainName</div><div style='margin-top:8px'><a class='toggle-btn' href='javascript:toggleAll(true)'>Expand All</a> <a style='margin-left:6px' class='toggle-btn' href='javascript:toggleAll(false)'>Collapse All</a></div><hr/>"

# 2) Domain Controllers info table
$dcCols = @("Name","HostName","IPv4Address","OperatingSystem","IsGlobalCatalog","Site","IsReadOnly")
$dcHtml = Convert-ObjectsToHtmlTable -Objects $DCs -Columns $dcCols
$fsmHtml = "<div class='small'><strong>FSMO Role holders</strong><br/>PDCEmulator: $($FSMO.PDCEmulator) &nbsp;|&nbsp; RIDMaster: $($FSMO.RIDMaster) &nbsp;|&nbsp; InfrastructureMaster: $($FSMO.InfrastructureMaster) <br/>DomainNamingMaster: $($FSMO.DomainNamingMaster) &nbsp;|&nbsp; SchemaMaster: $($FSMO.SchemaMaster)</div>"

# 3) Top-of-domain GPO links table
$TopDomainTableHtml = if ($TopDomainGpoLinks.Count -gt 0) {
    Convert-ObjectsToHtmlTable -Objects $TopDomainGpoLinks -Columns @("DisplayName","Id","LinkEnabled","Enforced","GpoStatus")
} else { "<div class='small'>No GPO links applied at domain root.</div>" }

# 4) Domain Controllers OU GPO links table
$DCTableHtml = if ($DCouGpoLinks.Count -gt 0) {
    Convert-ObjectsToHtmlTable -Objects $DCouGpoLinks -Columns @("DisplayName","Id","LinkEnabled","Enforced","GpoStatus")
} else { "<div class='small'>No GPO links applied to Domain Controllers OU ($DCDN)</div>" }

# 5) All GPOs table (flattened)
$AllGposHtml = Convert-ObjectsToHtmlTable -Objects $AllGpoDetails -Columns @("DisplayName","Id","GpoStatus","WmiFilterName","Links","NTLMSettingsFound")

# 6) NTLM hits table
$NtlmHtml = if ($GpoNtlmHits.Count -gt 0) {
    Convert-ObjectsToHtmlTable -Objects $GpoNtlmHits -Columns @("DisplayName","Id","NTLMSettings","LinksSummary","WmiFilterName")
} else { "<div class='small'>No NTLM/NTLMv2-related settings found in any GPO XML report.</div>" }

# Compose full HTML
$fullHtml = @"
<!doctype html>
<html>
<head>
<meta charset='utf-8' />
<title>$ReportNameBase</title>
<style>
$css
</style>
<script>
$js
</script>
</head>
<body>
$header

<div class='section'>
  <div class='section-header' onclick='toggle(""sec-dc"")'>
    <div><strong>Domain Controllers (list)</strong></div>
    <div class='small'>Click to expand/collapse</div>
  </div>
  <div id='sec-dc' class='section-content'>
    $fsmHtml
    $dcHtml
  </div>
</div>

<div class='section'>
  <div class='section-header' onclick='toggle(""sec-domainroot"")'>
    <div><strong>GPOs linked at Top of Domain (Domain Root)</strong></div>
    <div class='small'>Shows GPOs directly linked at the domain root</div>
  </div>
  <div id='sec-domainroot' class='section-content'>
    $TopDomainTableHtml
  </div>
</div>

<div class='section'>
  <div class='section-header' onclick='toggle(""sec-dcou"")'>
    <div><strong>GPOs linked to Domain Controllers OU ($DCDN)</strong></div>
    <div class='small'>Shows GPOs applied to Domain Controllers OU</div>
  </div>
  <div id='sec-dcou' class='section-content'>
    $DCTableHtml
  </div>
</div>

<div class='section'>
  <div class='section-header' onclick='toggle(""sec-allgpos"")'>
    <div><strong>All GPOs in domain</strong></div>
    <div class='small'>All GPOs with WMI filter, links summary, NTLM hit flag</div>
  </div>
  <div id='sec-allgpos' class='section-content'>
    $AllGposHtml
  </div>
</div>

<div class='section'>
  <div class='section-header' onclick='toggle(""sec-ntlm"")'>
    <div><strong>NTLM / NTLMv2 related GPOs and settings (highlighted)</strong></div>
    <div class='small'>Search performed in GPO XML for LmCompatibilityLevel, LAN Manager, Restrict NTLM text</div>
  </div>
  <div id='sec-ntlm' class='section-content'>
    $NtlmHtml
  </div>
</div>

<div class='small' style='margin-top:10px;'>Report files written to: <strong>$OutFolder</strong></div>

</body>
</html>
"@

Set-Content -Path $HtmlPath -Value $fullHtml -Encoding UTF8
Write-Host "HTML report exported to: $HtmlPath" -ForegroundColor Green

### ---------- Save summary of generated files ----------
$summary = [PSCustomObject]@{
    Csv = (Resolve-Path $CsvPath).ProviderPath
    Excel = (Resolve-Path $ExcelPath).ProviderPath
    Html = (Resolve-Path $HtmlPath).ProviderPath
    Folder = (Resolve-Path $OutFolder).ProviderPath
}
Write-Host "`nReport generation complete. Files created:" -ForegroundColor Cyan
$summary | Format-List

# End of script
