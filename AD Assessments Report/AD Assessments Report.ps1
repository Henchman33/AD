# AD Assessments Report
# CLd
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    IGT Active Directory Assessment Script
.DESCRIPTION
    Comprehensive AD environment assessment and documentation tool.
    Generates a modern HTML report with collapsible sections, search, and export capabilities.
    Designed for tiered Microsoft model environments (Tier 0 / Tier 1 / Tier 2).
.AUTHOR
    Steve McKee - IGT PLC - Server Administrator II
.NOTES
    Requires: ActiveDirectory PowerShell module, Domain Admin rights
    Output: Desktop\AD_Assessment\ folder
#>

# ============================================================
# CONFIGURATION
# ============================================================
$ReportTitle   = "XXX AD Assessment"
$ReportAuthor  = "Steve McKee"
$ReportOrg     = "XXX"
$ReportRole    = "Server Administrator II"
$OutputFolder  = "$env:USERPROFILE\Desktop\AD_Assessment"
$Timestamp     = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportFile    = "$OutputFolder\IGT_AD_Assessment_$Timestamp.html"

# ============================================================
# SETUP OUTPUT FOLDER
# ============================================================
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-Host "[+] Created output folder: $OutputFolder" -ForegroundColor Green
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  IGT Active Directory Assessment Tool" -ForegroundColor Cyan
Write-Host "  Author: $ReportAuthor | $ReportOrg" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================
# HELPER FUNCTIONS
# ============================================================
function Get-ADSafe {
    param([scriptblock]$ScriptBlock, [string]$Label = "Query")
    try {
        & $ScriptBlock
    } catch {
        Write-Warning "[$Label] Error: $($_.Exception.Message)"
        return $null
    }
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else { return "$Bytes B" }
}

function ConvertTo-HtmlTable {
    param(
        [Parameter(Mandatory)][string]$SectionId,
        [Parameter(Mandatory)][array]$Data,
        [string[]]$Properties
    )
    if (-not $Data -or $Data.Count -eq 0) {
        return '<div class="empty-state"><span class="empty-icon">&#8709;</span> No data found for this section.</div>'
    }

    if (-not $Properties) { $Properties = $Data[0].PSObject.Properties.Name }

    $rows = foreach ($item in $Data) {
        $cells = foreach ($prop in $Properties) {
            $val = $item.$prop
            if ($null -eq $val) { $val = "" }
            "<td>$([System.Web.HttpUtility]::HtmlEncode($val.ToString()))</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $headers = ($Properties | ForEach-Object { "<th>$_</th>" }) -join ''

    return @"
<div class="table-wrapper" id="table-$SectionId">
  <table class="data-table" data-section="$SectionId">
    <thead><tr>$headers</tr></thead>
    <tbody>$($rows -join '')</tbody>
  </table>
</div>
"@
}

# ============================================================
# DATA COLLECTION
# ============================================================
Write-Host "[1/12] Collecting Domain & Forest info..." -ForegroundColor Yellow
$Domain        = Get-ADSafe { Get-ADDomain } "Domain"
$Forest        = Get-ADSafe { Get-ADForest } "Forest"
$DomainDN      = $Domain.DistinguishedName
$DomainName    = $Domain.DNSRoot
$NetBIOS       = $Domain.NetBIOSName
$ForestName    = $Forest.Name
$FunctionalLvl = $Domain.DomainMode
$ForestFuncLvl = $Forest.ForestMode

Write-Host "[2/12] Collecting Domain Controllers..." -ForegroundColor Yellow
$DCs = Get-ADSafe {
    Get-ADDomainController -Filter * | Select-Object `
        Name, HostName, IPv4Address, Site, IsGlobalCatalog, IsReadOnly,
        OperatingSystem, OperatingSystemVersion,
        @{N='FSMO Roles';E={
            $roles = @()
            if ($_.OperationMasterRoles) { $roles += $_.OperationMasterRoles }
            $roles -join ', '
        }}
} "DCs"

Write-Host "[3/12] Collecting Sites & Subnets..." -ForegroundColor Yellow
$Sites = Get-ADSafe {
    Get-ADReplicationSite -Filter * | Select-Object Name, Description,
        @{N='Replication Schedule';E={ "Default" }}
} "Sites"

$Subnets = Get-ADSafe {
    Get-ADReplicationSubnet -Filter * | Select-Object Name, Site, Location, Description
} "Subnets"

$SiteLinks = Get-ADSafe {
    Get-ADReplicationSiteLink -Filter * | Select-Object Name, Cost, ReplicationFrequencyInMinutes,
        @{N='Sites';E={ $_.SitesIncluded -join ', ' }}
} "SiteLinks"

Write-Host "[4/12] Collecting OUs..." -ForegroundColor Yellow
$OUs = Get-ADSafe {
    Get-ADOrganizationalUnit -Filter * -Properties Description, ProtectedFromAccidentalDeletion |
    Select-Object Name, DistinguishedName, Description, ProtectedFromAccidentalDeletion,
        @{N='Parent';E={ ($_.DistinguishedName -split ',',2)[1] }}
} "OUs"

Write-Host "[5/12] Collecting all Users..." -ForegroundColor Yellow
$AllUsers = Get-ADSafe {
    Get-ADUser -Filter * -Properties SamAccountName, DisplayName, EmailAddress, Department,
        Title, Enabled, PasswordNeverExpires, PasswordLastSet, LastLogonDate,
        DistinguishedName, MemberOf, Description, WhenCreated, LockedOut,
        'msDS-PrincipalName' |
    Select-Object SamAccountName, DisplayName, EmailAddress, Department, Title,
        Enabled, PasswordNeverExpires,
        @{N='PasswordLastSet';E={ if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" } }},
        @{N='LastLogon';E={ if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" } }},
        LockedOut, Description,
        @{N='OU';E={ ($_.DistinguishedName -split ',',2)[1] }},
        @{N='WhenCreated';E={ $_.WhenCreated.ToString("yyyy-MM-dd") }},
        @{N='Tier';E={
            $sam = $_.SamAccountName.ToLower()
            if ($sam -match '^t0-' -or $sam -match '-t0') { 'Tier 0' }
            elseif ($sam -match '^t1-' -or $sam -match '-t1') { 'Tier 1' }
            elseif ($sam -match '^t2-' -or $sam -match '-t2') { 'Tier 2' }
            else { 'Standard' }
        }}
} "AllUsers"

# Tier Splits
$T0Users  = $AllUsers | Where-Object { $_.Tier -eq 'Tier 0' }
$T1Users  = $AllUsers | Where-Object { $_.Tier -eq 'Tier 1' }
$T2Users  = $AllUsers | Where-Object { $_.Tier -eq 'Tier 2' }
$StdUsers = $AllUsers | Where-Object { $_.Tier -eq 'Standard' }

Write-Host "[6/12] Collecting Service & Managed Service Accounts..." -ForegroundColor Yellow
$ServiceAccounts = $AllUsers | Where-Object {
    $_.SamAccountName -match 'svc|service|sa-|_sa|msvc' -or
    $_.Description -match 'service account'
}

$MSAs = Get-ADSafe {
    Get-ADServiceAccount -Filter * -Properties SamAccountName, DisplayName, DNSHostName,
        Enabled, Description, PrincipalsAllowedToRetrieveManagedPassword, WhenCreated |
    Select-Object SamAccountName, DisplayName, DNSHostName, Enabled, Description,
        @{N='AllowedHosts';E={ $_.PrincipalsAllowedToRetrieveManagedPassword -join ', ' }},
        @{N='WhenCreated';E={ $_.WhenCreated.ToString("yyyy-MM-dd") }},
        @{N='Type';E={
            if ($_.ObjectClass -eq 'msDS-GroupManagedServiceAccount') { 'gMSA' } else { 'sMSA' }
        }}
} "MSAs"

Write-Host "[7/12] Collecting Groups..." -ForegroundColor Yellow
$Groups = Get-ADSafe {
    Get-ADGroup -Filter * -Properties Description, MemberOf, Members, WhenCreated, GroupScope, GroupCategory |
    Select-Object Name, SamAccountName, GroupScope, GroupCategory, Description,
        @{N='MemberCount';E={ ($_.Members | Measure-Object).Count }},
        @{N='WhenCreated';E={ $_.WhenCreated.ToString("yyyy-MM-dd") }},
        @{N='OU';E={ ($_.DistinguishedName -split ',',2)[1] }}
} "Groups"

Write-Host "[8/12] Collecting Computers..." -ForegroundColor Yellow
$Computers = Get-ADSafe {
    Get-ADComputer -Filter * -Properties Name, OperatingSystem, OperatingSystemVersion,
        IPv4Address, Enabled, LastLogonDate, Description, DistinguishedName, WhenCreated |
    Select-Object Name, OperatingSystem, OperatingSystemVersion, IPv4Address, Enabled,
        @{N='LastLogon';E={ if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" } }},
        Description,
        @{N='OU';E={ ($_.DistinguishedName -split ',',2)[1] }},
        @{N='WhenCreated';E={ $_.WhenCreated.ToString("yyyy-MM-dd") }},
        @{N='Tier';E={
            $dn = $_.DistinguishedName.ToLower()
            if ($dn -match 'tier.?0|t0') { 'Tier 0' }
            elseif ($dn -match 'tier.?1|t1') { 'Tier 1' }
            elseif ($dn -match 'tier.?2|t2') { 'Tier 2' }
            else { 'Standard' }
        }}
} "Computers"

Write-Host "[9/12] Collecting GPOs..." -ForegroundColor Yellow
$GPOs = Get-ADSafe {
    Get-GPO -All | Select-Object DisplayName, GpoStatus, CreationTime, ModificationTime,
        @{N='WMIFilter';E={ if ($_.WmiFilter) { $_.WmiFilter.Name } else { "None" } }},
        @{N='CreationTime';E={ $_.CreationTime.ToString("yyyy-MM-dd") }},
        @{N='ModificationTime';E={ $_.ModificationTime.ToString("yyyy-MM-dd") }}
} "GPOs"

Write-Host "[10/12] Collecting GPO Links..." -ForegroundColor Yellow
$GPOLinks = Get-ADSafe {
    Get-ADOrganizationalUnit -Filter * | ForEach-Object {
        $ou = $_
        try {
            $links = (Get-GPInheritance -Target $ou.DistinguishedName).GpoLinks
            foreach ($link in $links) {
                [PSCustomObject]@{
                    OU       = $ou.Name
                    OUDN     = $ou.DistinguishedName
                    GPO      = $link.DisplayName
                    Enabled  = $link.Enabled
                    Enforced = $link.Enforced
                    Order    = $link.Order
                }
            }
        } catch {}
    }
} "GPOLinks"

Write-Host "[11/12] Collecting Trusts..." -ForegroundColor Yellow
$Trusts = Get-ADSafe {
    Get-ADTrust -Filter * | Select-Object Name, TrustType, TrustDirection, TrustAttributes,
        SelectiveAuthentication, IntraForest, ForestTransitive
} "Trusts"

Write-Host "[12/12] Collecting Password & Security Policies..." -ForegroundColor Yellow
$DefaultPWPolicy = Get-ADSafe {
    Get-ADDefaultDomainPasswordPolicy | Select-Object ComplexityEnabled, MinPasswordLength,
        MaxPasswordAge, MinPasswordAge, PasswordHistoryCount, LockoutThreshold,
        LockoutDuration, LockoutObservationWindow, ReversibleEncryptionEnabled
} "PWPolicy"

$FGPPs = Get-ADSafe {
    Get-ADFineGrainedPasswordPolicy -Filter * | Select-Object Name, Precedence,
        MinPasswordLength, MaxPasswordAge, MinPasswordAge, PasswordHistoryCount,
        LockoutThreshold, ComplexityEnabled,
        @{N='AppliesTo';E={ $_.AppliesTo -join ', ' }}
} "FGPPs"

$PrivUsers = Get-ADSafe {
    Get-ADGroupMember -Identity "Domain Admins" -Recursive |
    Get-ADUser -Properties SamAccountName, DisplayName, LastLogonDate, Enabled |
    Select-Object SamAccountName, DisplayName, Enabled,
        @{N='LastLogon';E={ if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" } }}
} "PrivUsers"

# ============================================================
# STATISTICS
# ============================================================
$Stats = @{
    TotalUsers         = ($AllUsers | Measure-Object).Count
    EnabledUsers       = ($AllUsers | Where-Object { $_.Enabled -eq $true } | Measure-Object).Count
    DisabledUsers      = ($AllUsers | Where-Object { $_.Enabled -eq $false } | Measure-Object).Count
    T0Users            = ($T0Users | Measure-Object).Count
    T1Users            = ($T1Users | Measure-Object).Count
    T2Users            = ($T2Users | Measure-Object).Count
    StaleUsers         = ($AllUsers | Where-Object { $_.LastLogon -ne "Never" -and ([datetime]::ParseExact($_.LastLogon,"yyyy-MM-dd",$null)) -lt (Get-Date).AddDays(-90) } | Measure-Object).Count
    TotalComputers     = ($Computers | Measure-Object).Count
    TotalGroups        = ($Groups | Measure-Object).Count
    TotalOUs           = ($OUs | Measure-Object).Count
    TotalGPOs          = ($GPOs | Measure-Object).Count
    TotalDCs           = ($DCs | Measure-Object).Count
    TotalSites         = ($Sites | Measure-Object).Count
    TotalMSAs          = ($MSAs | Measure-Object).Count
    TotalTrusts        = ($Trusts | Measure-Object).Count
}

# ============================================================
# BUILD HTML REPORT
# ============================================================
Write-Host ""
Write-Host "[HTML] Building report..." -ForegroundColor Cyan

# Pre-build table HTML
$tblDCs          = ConvertTo-HtmlTable -SectionId "dcs"          -Data $DCs
$tblSites        = ConvertTo-HtmlTable -SectionId "sites"        -Data $Sites
$tblSubnets      = ConvertTo-HtmlTable -SectionId "subnets"      -Data $Subnets
$tblSiteLinks    = ConvertTo-HtmlTable -SectionId "sitelinks"    -Data $SiteLinks
$tblOUs          = ConvertTo-HtmlTable -SectionId "ous"          -Data $OUs
$tblT0           = ConvertTo-HtmlTable -SectionId "tier0"        -Data $T0Users
$tblT1           = ConvertTo-HtmlTable -SectionId "tier1"        -Data $T1Users
$tblT2           = ConvertTo-HtmlTable -SectionId "tier2"        -Data $T2Users
$tblStdUsers     = ConvertTo-HtmlTable -SectionId "stdusers"     -Data $StdUsers
$tblSvcAccts     = ConvertTo-HtmlTable -SectionId "svcaccounts"  -Data $ServiceAccounts
$tblMSAs         = ConvertTo-HtmlTable -SectionId "msas"         -Data $MSAs
$tblGroups       = ConvertTo-HtmlTable -SectionId "groups"       -Data $Groups
$tblComputers    = ConvertTo-HtmlTable -SectionId "computers"    -Data $Computers
$tblGPOs         = ConvertTo-HtmlTable -SectionId "gpos"         -Data $GPOs
$tblGPOLinks     = ConvertTo-HtmlTable -SectionId "gpolinks"     -Data $GPOLinks
$tblTrusts       = ConvertTo-HtmlTable -SectionId "trusts"       -Data $Trusts
$tblPrivUsers    = ConvertTo-HtmlTable -SectionId "privusers"    -Data $PrivUsers
$tblFGPPs        = ConvertTo-HtmlTable -SectionId "fgpps"        -Data $FGPPs

# Build JSON data for export (escaping for JS)
function ConvertTo-JSArray {
    param([array]$Data)
    if (-not $Data -or $Data.Count -eq 0) { return "[]" }
    $json = $Data | ConvertTo-Json -Compress -Depth 3
    return $json
}

$jsonDCs       = (ConvertTo-JSArray $DCs)       -replace "'", "\'"
$jsonSites     = (ConvertTo-JSArray $Sites)     -replace "'", "\'"
$jsonSubnets   = (ConvertTo-JSArray $Subnets)   -replace "'", "\'"
$jsonT0        = (ConvertTo-JSArray $T0Users)   -replace "'", "\'"
$jsonT1        = (ConvertTo-JSArray $T1Users)   -replace "'", "\'"
$jsonT2        = (ConvertTo-JSArray $T2Users)   -replace "'", "\'"
$jsonSvc       = (ConvertTo-JSArray $ServiceAccounts) -replace "'", "\'"
$jsonMSAs      = (ConvertTo-JSArray $MSAs)      -replace "'", "\'"
$jsonGroups    = (ConvertTo-JSArray $Groups)    -replace "'", "\'"
$jsonComputers = (ConvertTo-JSArray $Computers) -replace "'", "\'"
$jsonGPOs      = (ConvertTo-JSArray $GPOs)      -replace "'", "\'"
$jsonTrusts    = (ConvertTo-JSArray $Trusts)    -replace "'", "\'"
$jsonPrivUsers = (ConvertTo-JSArray $PrivUsers) -replace "'", "\'"
$jsonOUs       = (ConvertTo-JSArray $OUs)       -replace "'", "\'"

$GeneratedDate = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"

$HTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$ReportTitle</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg:          #0d0f14;
    --surface:     #141720;
    --surface2:    #1c2030;
    --surface3:    #242840;
    --border:      #2a2f45;
    --accent:      #00c2ff;
    --accent2:     #7b5cfa;
    --accent3:     #00e5a0;
    --warn:        #ffb84d;
    --danger:      #ff5c72;
    --text:        #e8ecf4;
    --text-muted:  #7a8099;
    --text-dim:    #4a5168;
    --t0:          #ff5c72;
    --t1:          #ffb84d;
    --t2:          #00c2ff;
    --std:         #00e5a0;
    --radius:      8px;
    --radius-lg:   14px;
    --shadow:      0 4px 24px rgba(0,0,0,0.5);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'IBM Plex Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    font-size: 14px;
    line-height: 1.6;
  }
  /* ── TOP HEADER ── */
  .report-header {
    background: linear-gradient(135deg, #0d0f14 0%, #141a2e 50%, #0d1220 100%);
    border-bottom: 1px solid var(--border);
    padding: 32px 40px 24px;
    position: relative;
    overflow: hidden;
  }
  .report-header::before {
    content: '';
    position: absolute; inset: 0;
    background: radial-gradient(ellipse 600px 200px at 80% 50%, rgba(0,194,255,0.06), transparent);
    pointer-events: none;
  }
  .header-grid { display: grid; grid-template-columns: 1fr auto; gap: 24px; align-items: start; }
  .report-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: rgba(0,194,255,0.1); border: 1px solid rgba(0,194,255,0.25);
    border-radius: 4px; padding: 4px 12px; margin-bottom: 12px;
    font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--accent);
    letter-spacing: 0.08em; text-transform: uppercase;
  }
  .report-badge::before { content: '●'; font-size: 8px; }
  .report-title {
    font-size: 32px; font-weight: 700; letter-spacing: -0.5px;
    background: linear-gradient(135deg, #fff 30%, var(--accent));
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin-bottom: 6px;
  }
  .report-meta { color: var(--text-muted); font-size: 13px; }
  .report-meta span { margin-right: 20px; }
  .report-meta strong { color: var(--text); }
  .header-stats { display: flex; gap: 12px; flex-wrap: wrap; }
  .hstat {
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 12px 18px; text-align: center; min-width: 90px;
  }
  .hstat-val { font-size: 22px; font-weight: 700; font-family: 'IBM Plex Mono', monospace; line-height: 1; }
  .hstat-lbl { font-size: 10px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.06em; margin-top: 4px; }
  /* ── DOMAIN MAP ── */
  .domain-map {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius-lg); margin: 24px 40px;
    padding: 24px; overflow-x: auto;
  }
  .domain-map-title {
    font-size: 11px; text-transform: uppercase; letter-spacing: 0.1em;
    color: var(--text-muted); margin-bottom: 16px;
    display: flex; align-items: center; gap: 8px;
  }
  .domain-map-title::before { content: ''; display: block; height: 1px; width: 24px; background: var(--border); }
  .dm-tree { display: flex; flex-direction: column; gap: 8px; }
  .dm-row { display: flex; align-items: center; gap: 12px; }
  .dm-node {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 6px 14px; border-radius: 6px; font-size: 13px;
    border: 1px solid; white-space: nowrap;
  }
  .dm-node.forest { background: rgba(123,92,250,0.12); border-color: rgba(123,92,250,0.35); color: var(--accent2); }
  .dm-node.domain { background: rgba(0,194,255,0.1); border-color: rgba(0,194,255,0.3); color: var(--accent); }
  .dm-node.dc     { background: rgba(0,229,160,0.08); border-color: rgba(0,229,160,0.25); color: var(--accent3); }
  .dm-node.site   { background: rgba(255,184,77,0.08); border-color: rgba(255,184,77,0.25); color: var(--warn); }
  .dm-node.tier0  { background: rgba(255,92,114,0.1); border-color: rgba(255,92,114,0.3); color: var(--t0); }
  .dm-node.tier1  { background: rgba(255,184,77,0.08); border-color: rgba(255,184,77,0.25); color: var(--t1); }
  .dm-node.tier2  { background: rgba(0,194,255,0.08); border-color: rgba(0,194,255,0.25); color: var(--t2); }
  .dm-connector { color: var(--text-dim); font-size: 18px; line-height: 1; }
  .dm-indent { padding-left: 32px; border-left: 1px dashed var(--border); margin-left: 20px; }
  /* ── TOOLBAR ── */
  .toolbar {
    display: flex; align-items: center; gap: 12px;
    padding: 16px 40px; background: var(--surface);
    border-bottom: 1px solid var(--border); flex-wrap: wrap; position: sticky; top: 0; z-index: 100;
  }
  .search-wrap {
    flex: 1; min-width: 220px; max-width: 400px;
    display: flex; align-items: center; gap: 8px;
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 8px 14px;
    transition: border-color .2s;
  }
  .search-wrap:focus-within { border-color: var(--accent); }
  .search-wrap svg { color: var(--text-muted); flex-shrink: 0; }
  .search-wrap input {
    background: transparent; border: none; outline: none;
    color: var(--text); font-family: 'IBM Plex Sans', sans-serif;
    font-size: 13px; width: 100%;
  }
  .search-wrap input::placeholder { color: var(--text-dim); }
  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 8px 14px; border-radius: var(--radius); border: 1px solid var(--border);
    background: var(--surface2); color: var(--text-muted); font-size: 12px;
    cursor: pointer; font-family: 'IBM Plex Sans', sans-serif;
    transition: all .15s; white-space: nowrap; text-transform: uppercase; letter-spacing: 0.05em;
  }
  .btn:hover { background: var(--surface3); color: var(--text); border-color: var(--accent); }
  .btn-accent { background: rgba(0,194,255,0.1); border-color: rgba(0,194,255,0.35); color: var(--accent); }
  .btn-accent:hover { background: rgba(0,194,255,0.2); }
  .toolbar-sep { width: 1px; height: 24px; background: var(--border); flex-shrink: 0; }
  .search-count { font-size: 11px; color: var(--text-muted); white-space: nowrap; }
  /* ── MAIN CONTENT ── */
  .main { padding: 24px 40px; }
  /* ── SECTIONS ── */
  .section {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius-lg); margin-bottom: 16px; overflow: hidden;
  }
  .section-header {
    display: flex; align-items: center; gap: 12px;
    padding: 16px 20px; cursor: pointer; user-select: none;
    transition: background .15s; border-radius: var(--radius-lg) var(--radius-lg) 0 0;
  }
  .section-header:hover { background: var(--surface2); }
  .section-icon {
    width: 32px; height: 32px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 15px; flex-shrink: 0;
  }
  .section-title-group { flex: 1; }
  .section-title { font-size: 15px; font-weight: 600; }
  .section-sub { font-size: 11px; color: var(--text-muted); }
  .section-count {
    font-family: 'IBM Plex Mono', monospace; font-size: 12px;
    background: var(--surface3); border: 1px solid var(--border);
    padding: 2px 8px; border-radius: 4px; color: var(--text-muted);
  }
  .section-chevron {
    font-size: 12px; color: var(--text-dim); transition: transform .2s;
  }
  .section.open .section-chevron { transform: rotate(90deg); }
  .section-body { display: none; padding: 0 20px 20px; }
  .section.open .section-body { display: block; }
  /* Color coding */
  .sec-forest .section-icon { background: rgba(123,92,250,0.15); color: var(--accent2); }
  .sec-dc     .section-icon { background: rgba(0,229,160,0.12); color: var(--accent3); }
  .sec-sites  .section-icon { background: rgba(255,184,77,0.12); color: var(--warn); }
  .sec-ous    .section-icon { background: rgba(0,194,255,0.1); color: var(--accent); }
  .sec-t0     .section-icon { background: rgba(255,92,114,0.12); color: var(--t0); }
  .sec-t1     .section-icon { background: rgba(255,184,77,0.12); color: var(--t1); }
  .sec-t2     .section-icon { background: rgba(0,194,255,0.1); color: var(--t2); }
  .sec-std    .section-icon { background: rgba(0,229,160,0.1); color: var(--std); }
  .sec-gpo    .section-icon { background: rgba(123,92,250,0.12); color: var(--accent2); }
  .sec-sec    .section-icon { background: rgba(255,92,114,0.1); color: var(--t0); }
  /* ── TIER BADGE ── */
  .tier-badge {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 2px 10px; border-radius: 4px; font-size: 10px;
    font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; margin-left: 4px;
  }
  .tier-badge.t0 { background: rgba(255,92,114,0.15); color: var(--t0); border: 1px solid rgba(255,92,114,0.3); }
  .tier-badge.t1 { background: rgba(255,184,77,0.12); color: var(--t1); border: 1px solid rgba(255,184,77,0.3); }
  .tier-badge.t2 { background: rgba(0,194,255,0.1); color: var(--t2); border: 1px solid rgba(0,194,255,0.25); }
  /* ── EXPORT BAR ── */
  .export-bar {
    display: flex; align-items: center; gap: 8px;
    padding: 12px 0 16px; flex-wrap: wrap;
  }
  .export-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.06em; margin-right: 4px; }
  .btn-export {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 5px 12px; border-radius: 5px; border: 1px solid var(--border);
    background: var(--surface2); color: var(--text-muted); font-size: 11px;
    cursor: pointer; font-family: 'IBM Plex Mono', monospace;
    transition: all .15s; text-transform: uppercase; letter-spacing: 0.06em;
  }
  .btn-export:hover { color: var(--text); border-color: var(--accent); background: rgba(0,194,255,0.06); }
  .btn-export.csv:hover  { border-color: var(--accent3); color: var(--accent3); }
  .btn-export.xlsx:hover { border-color: var(--warn); color: var(--warn); }
  .btn-export.txt:hover  { border-color: var(--text-muted); color: var(--text); }
  .btn-export.pdf:hover  { border-color: var(--t0); color: var(--t0); }
  /* ── TABLE ── */
  .table-wrapper { overflow-x: auto; border-radius: var(--radius); border: 1px solid var(--border); margin-top: 8px; }
  .data-table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .data-table thead tr { background: var(--surface3); }
  .data-table th {
    padding: 10px 14px; text-align: left; font-weight: 600;
    font-size: 11px; text-transform: uppercase; letter-spacing: 0.07em;
    color: var(--text-muted); white-space: nowrap; border-bottom: 1px solid var(--border);
  }
  .data-table td {
    padding: 9px 14px; border-bottom: 1px solid rgba(42,47,69,0.5);
    color: var(--text); vertical-align: top; max-width: 300px;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  .data-table tr:last-child td { border-bottom: none; }
  .data-table tbody tr:hover { background: rgba(255,255,255,0.025); }
  .data-table tr.hidden { display: none; }
  .data-table td mark {
    background: rgba(0,194,255,0.25); color: var(--accent);
    border-radius: 2px; padding: 0 2px;
  }
  /* ── POLICY BOX ── */
  .policy-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px,1fr)); gap: 12px; margin-top: 8px; }
  .policy-card {
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: var(--radius); padding: 16px;
  }
  .policy-card-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 4px; }
  .policy-card-val { font-size: 18px; font-weight: 600; font-family: 'IBM Plex Mono', monospace; }
  /* ── EMPTY STATE ── */
  .empty-state {
    padding: 32px; text-align: center; color: var(--text-dim);
    font-size: 13px; border: 1px dashed var(--border); border-radius: var(--radius);
    margin-top: 8px;
  }
  /* ── SUBSECTION ── */
  .subsection { margin-top: 20px; }
  .subsection-title {
    font-size: 12px; text-transform: uppercase; letter-spacing: 0.08em;
    color: var(--text-muted); margin-bottom: 8px;
    padding-bottom: 6px; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; gap: 8px;
  }
  /* ── FOOTER ── */
  footer {
    text-align: center; padding: 24px 40px; color: var(--text-dim);
    font-size: 11px; border-top: 1px solid var(--border); margin-top: 40px;
  }
  footer strong { color: var(--text-muted); }
  /* ── SCROLLBAR ── */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: var(--surface); }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--text-dim); }
</style>
</head>
<body>

<!-- ═══════════════════════════ HEADER ═══════════════════════════ -->
<div class="report-header">
  <div class="header-grid">
    <div>
      <div class="report-badge">Active Directory Assessment</div>
      <div class="report-title">$ReportTitle</div>
      <div class="report-meta">
        <span>&#128100; <strong>$ReportAuthor</strong></span>
        <span>&#127970; <strong>$ReportOrg</strong></span>
        <span>&#129333; <strong>$ReportRole</strong></span>
        <span>&#128197; Generated: <strong>$GeneratedDate</strong></span>
      </div>
    </div>
    <div class="header-stats">
      <div class="hstat"><div class="hstat-val" style="color:var(--accent)">$($Stats.TotalUsers)</div><div class="hstat-lbl">Users</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--accent3)">$($Stats.TotalComputers)</div><div class="hstat-lbl">Computers</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--accent2)">$($Stats.TotalGroups)</div><div class="hstat-lbl">Groups</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--warn)">$($Stats.TotalGPOs)</div><div class="hstat-lbl">GPOs</div></div>
      <div class="hstat"><div class="hstat-val" style="color:var(--t0)">$($Stats.TotalDCs)</div><div class="hstat-lbl">DCs</div></div>
    </div>
  </div>
</div>

<!-- ═══════════════════════════ DOMAIN MAP ═══════════════════════════ -->
<div class="domain-map">
  <div class="domain-map-title">Domain Architecture Map</div>
  <div class="dm-tree">
    <div class="dm-row">
      <div class="dm-node forest">&#127758; Forest: $ForestName</div>
      <div class="dm-connector">&#x2500;&#x2500;</div>
      <div class="dm-node domain">&#129428; Domain: $DomainName ($NetBIOS)</div>
      <div class="dm-connector">&#x2500;&#x2500;</div>
      <div class="dm-node dc">&#128421; $($Stats.TotalDCs) Domain Controllers</div>
    </div>
    <div class="dm-indent">
      <div class="dm-row" style="margin-top:8px">
        <div class="dm-node tier0">&#128274; Tier 0 — Domain Admins (t0-*) — $($Stats.T0Users) accounts</div>
      </div>
      <div class="dm-row" style="margin-top:6px">
        <div class="dm-node tier1">&#128296; Tier 1 — Server Operators (t1-*) — $($Stats.T1Users) accounts</div>
      </div>
      <div class="dm-row" style="margin-top:6px">
        <div class="dm-node tier2">&#128187; Tier 2 — Client Operators (t2-*) — $($Stats.T2Users) accounts</div>
      </div>
      <div class="dm-row" style="margin-top:6px">
        <div class="dm-node site">&#127968; $($Stats.TotalSites) AD Sites &nbsp;|&nbsp; $($Stats.TotalOUs) OUs &nbsp;|&nbsp; $($Stats.TotalGPOs) GPOs &nbsp;|&nbsp; $($Stats.TotalTrusts) Trusts &nbsp;|&nbsp; $($Stats.TotalMSAs) gMSAs</div>
      </div>
    </div>
  </div>
</div>

<!-- ═══════════════════════════ TOOLBAR ═══════════════════════════ -->
<div class="toolbar">
  <div class="search-wrap">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    <input type="text" id="globalSearch" placeholder="Search all tables..." oninput="globalFilter(this.value)">
  </div>
  <span class="search-count" id="searchCount"></span>
  <div class="toolbar-sep"></div>
  <button class="btn btn-accent" onclick="expandAll()">&#x25BC; Expand All</button>
  <button class="btn" onclick="collapseAll()">&#x25B2; Collapse All</button>
</div>

<!-- ═══════════════════════════ MAIN ═══════════════════════════ -->
<div class="main">

<!-- ──────────── DOMAIN INFO ──────────── -->
<div class="section sec-forest open" id="s-domaininfo">
  <div class="section-header" onclick="toggleSection('s-domaininfo')">
    <div class="section-icon">&#129428;</div>
    <div class="section-title-group">
      <div class="section-title">Domain &amp; Forest Information</div>
      <div class="section-sub">Core domain details, functional levels, FSMO roles</div>
    </div>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="policy-grid">
      <div class="policy-card"><div class="policy-card-label">Domain Name (FQDN)</div><div class="policy-card-val" style="font-size:14px">$DomainName</div></div>
      <div class="policy-card"><div class="policy-card-label">NetBIOS Name</div><div class="policy-card-val" style="font-size:14px">$NetBIOS</div></div>
      <div class="policy-card"><div class="policy-card-label">Forest Name</div><div class="policy-card-val" style="font-size:14px">$ForestName</div></div>
      <div class="policy-card"><div class="policy-card-label">Domain Functional Level</div><div class="policy-card-val" style="font-size:13px">$FunctionalLvl</div></div>
      <div class="policy-card"><div class="policy-card-label">Forest Functional Level</div><div class="policy-card-val" style="font-size:13px">$ForestFuncLvl</div></div>
      <div class="policy-card"><div class="policy-card-label">Distinguished Name</div><div class="policy-card-val" style="font-size:11px;word-break:break-all">$DomainDN</div></div>
    </div>
  </div>
</div>

<!-- ──────────── DOMAIN CONTROLLERS ──────────── -->
<div class="section sec-dc" id="s-dcs">
  <div class="section-header" onclick="toggleSection('s-dcs')">
    <div class="section-icon">&#128421;</div>
    <div class="section-title-group">
      <div class="section-title">Domain Controllers</div>
      <div class="section-sub">All DCs, FSMO role holders, OS, GC status</div>
    </div>
    <span class="section-count">$($Stats.TotalDCs)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('dcs')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('dcs','Domain Controllers')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('dcs')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-dcs')">PDF</button>
    </div>
    $tblDCs
  </div>
</div>

<!-- ──────────── SITES & SUBNETS ──────────── -->
<div class="section sec-sites" id="s-sites">
  <div class="section-header" onclick="toggleSection('s-sites')">
    <div class="section-icon">&#127968;</div>
    <div class="section-title-group">
      <div class="section-title">Sites, Subnets &amp; Site Links</div>
      <div class="section-sub">AD replication topology, subnet mappings</div>
    </div>
    <span class="section-count">$($Stats.TotalSites) sites</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="subsection">
      <div class="subsection-title">AD Sites</div>
      <div class="export-bar">
        <span class="export-label">Export:</span>
        <button class="btn-export csv"  onclick="exportCSV('sites')">CSV</button>
        <button class="btn-export xlsx" onclick="exportXLSX('sites','AD Sites')">XLSX</button>
        <button class="btn-export txt"  onclick="exportTXT('sites')">TXT</button>
      </div>
      $tblSites
    </div>
    <div class="subsection">
      <div class="subsection-title">Subnets</div>
      <div class="export-bar">
        <span class="export-label">Export:</span>
        <button class="btn-export csv"  onclick="exportCSV('subnets')">CSV</button>
        <button class="btn-export xlsx" onclick="exportXLSX('subnets','Subnets')">XLSX</button>
        <button class="btn-export txt"  onclick="exportTXT('subnets')">TXT</button>
      </div>
      $tblSubnets
    </div>
    <div class="subsection">
      <div class="subsection-title">Site Links</div>
      $tblSiteLinks
    </div>
  </div>
</div>

<!-- ──────────── OUs ──────────── -->
<div class="section sec-ous" id="s-ous">
  <div class="section-header" onclick="toggleSection('s-ous')">
    <div class="section-icon">&#128193;</div>
    <div class="section-title-group">
      <div class="section-title">Organizational Units</div>
      <div class="section-sub">Complete OU structure, deletion protection status</div>
    </div>
    <span class="section-count">$($Stats.TotalOUs)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('ous')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('ous','Organizational Units')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('ous')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-ous')">PDF</button>
    </div>
    $tblOUs
  </div>
</div>

<!-- ──────────── TIER 0 ACCOUNTS ──────────── -->
<div class="section sec-t0" id="s-tier0">
  <div class="section-header" onclick="toggleSection('s-tier0')">
    <div class="section-icon">&#128274;</div>
    <div class="section-title-group">
      <div class="section-title">Tier 0 Accounts <span class="tier-badge t0">T0 — Domain Admins</span></div>
      <div class="section-sub">Privileged admin accounts (t0-*), Domain Admin tier — highest security boundary</div>
    </div>
    <span class="section-count">$($Stats.T0Users)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('tier0')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('tier0','Tier 0 Accounts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('tier0')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-tier0')">PDF</button>
    </div>
    $tblT0
  </div>
</div>

<!-- ──────────── TIER 1 ACCOUNTS ──────────── -->
<div class="section sec-t1" id="s-tier1">
  <div class="section-header" onclick="toggleSection('s-tier1')">
    <div class="section-icon">&#128296;</div>
    <div class="section-title-group">
      <div class="section-title">Tier 1 Accounts <span class="tier-badge t1">T1 — Server Operators</span></div>
      <div class="section-sub">Server admin accounts (t1-*), Tier 1 server access group</div>
    </div>
    <span class="section-count">$($Stats.T1Users)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('tier1')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('tier1','Tier 1 Accounts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('tier1')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-tier1')">PDF</button>
    </div>
    $tblT1
  </div>
</div>

<!-- ──────────── TIER 2 ACCOUNTS ──────────── -->
<div class="section sec-t2" id="s-tier2">
  <div class="section-header" onclick="toggleSection('s-tier2')">
    <div class="section-icon">&#128187;</div>
    <div class="section-title-group">
      <div class="section-title">Tier 2 Accounts <span class="tier-badge t2">T2 — Client Operators</span></div>
      <div class="section-sub">Workstation/client admin accounts (t2-*)</div>
    </div>
    <span class="section-count">$($Stats.T2Users)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('tier2')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('tier2','Tier 2 Accounts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('tier2')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-tier2')">PDF</button>
    </div>
    $tblT2
  </div>
</div>

<!-- ──────────── STANDARD USERS ──────────── -->
<div class="section sec-std" id="s-stdusers">
  <div class="section-header" onclick="toggleSection('s-stdusers')">
    <div class="section-icon">&#128100;</div>
    <div class="section-title-group">
      <div class="section-title">Standard Users</div>
      <div class="section-sub">Non-tiered user accounts — standard domain users</div>
    </div>
    <span class="section-count">$($Stats.TotalUsers - $Stats.T0Users - $Stats.T1Users - $Stats.T2Users)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('stdusers')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('stdusers','Standard Users')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('stdusers')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-stdusers')">PDF</button>
    </div>
    $tblStdUsers
  </div>
</div>

<!-- ──────────── SERVICE ACCOUNTS ──────────── -->
<div class="section sec-t1" id="s-svc">
  <div class="section-header" onclick="toggleSection('s-svc')">
    <div class="section-icon">&#9881;&#65039;</div>
    <div class="section-title-group">
      <div class="section-title">Service Accounts</div>
      <div class="section-sub">All service accounts across all tiers (svc-*, sa-*, *service*)</div>
    </div>
    <span class="section-count">$($ServiceAccounts.Count)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('svcaccounts')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('svcaccounts','Service Accounts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('svcaccounts')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-svc')">PDF</button>
    </div>
    $tblSvcAccts
  </div>
</div>

<!-- ──────────── gMSAs / sMSAs ──────────── -->
<div class="section sec-dc" id="s-msas">
  <div class="section-header" onclick="toggleSection('s-msas')">
    <div class="section-icon">&#128273;</div>
    <div class="section-title-group">
      <div class="section-title">Managed Service Accounts (gMSA / sMSA)</div>
      <div class="section-sub">Group and Standalone Managed Service Accounts</div>
    </div>
    <span class="section-count">$($Stats.TotalMSAs)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('msas')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('msas','Managed Service Accounts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('msas')">TXT</button>
    </div>
    $tblMSAs
  </div>
</div>

<!-- ──────────── GROUPS ──────────── -->
<div class="section sec-std" id="s-groups">
  <div class="section-header" onclick="toggleSection('s-groups')">
    <div class="section-icon">&#128101;</div>
    <div class="section-title-group">
      <div class="section-title">Security &amp; Distribution Groups</div>
      <div class="section-sub">All AD groups, scope, category, member counts</div>
    </div>
    <span class="section-count">$($Stats.TotalGroups)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('groups')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('groups','Groups')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('groups')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-groups')">PDF</button>
    </div>
    $tblGroups
  </div>
</div>

<!-- ──────────── COMPUTERS ──────────── -->
<div class="section sec-sites" id="s-computers">
  <div class="section-header" onclick="toggleSection('s-computers')">
    <div class="section-icon">&#128421;&#65039;</div>
    <div class="section-title-group">
      <div class="section-title">Computer Objects</div>
      <div class="section-sub">All domain-joined computers, OS, last logon, tier placement</div>
    </div>
    <span class="section-count">$($Stats.TotalComputers)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('computers')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('computers','Computers')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('computers')">TXT</button>
      <button class="btn-export pdf"  onclick="exportPDF('s-computers')">PDF</button>
    </div>
    $tblComputers
  </div>
</div>

<!-- ──────────── GPOs ──────────── -->
<div class="section sec-gpo" id="s-gpos">
  <div class="section-header" onclick="toggleSection('s-gpos')">
    <div class="section-icon">&#128220;</div>
    <div class="section-title-group">
      <div class="section-title">Group Policy Objects</div>
      <div class="section-sub">All GPOs, status, WMI filters, modification dates</div>
    </div>
    <span class="section-count">$($Stats.TotalGPOs)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="subsection">
      <div class="subsection-title">All GPOs</div>
      <div class="export-bar">
        <span class="export-label">Export:</span>
        <button class="btn-export csv"  onclick="exportCSV('gpos')">CSV</button>
        <button class="btn-export xlsx" onclick="exportXLSX('gpos','GPOs')">XLSX</button>
        <button class="btn-export txt"  onclick="exportTXT('gpos')">TXT</button>
      </div>
      $tblGPOs
    </div>
    <div class="subsection">
      <div class="subsection-title">GPO Links (OU Assignments)</div>
      $tblGPOLinks
    </div>
  </div>
</div>

<!-- ──────────── TRUSTS ──────────── -->
<div class="section sec-sec" id="s-trusts">
  <div class="section-header" onclick="toggleSection('s-trusts')">
    <div class="section-icon">&#128257;</div>
    <div class="section-title-group">
      <div class="section-title">Domain &amp; Forest Trusts</div>
      <div class="section-sub">All trust relationships, direction, transitivity</div>
    </div>
    <span class="section-count">$($Stats.TotalTrusts)</span>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="export-bar">
      <span class="export-label">Export:</span>
      <button class="btn-export csv"  onclick="exportCSV('trusts')">CSV</button>
      <button class="btn-export xlsx" onclick="exportXLSX('trusts','Trusts')">XLSX</button>
      <button class="btn-export txt"  onclick="exportTXT('trusts')">TXT</button>
    </div>
    $tblTrusts
  </div>
</div>

<!-- ──────────── SECURITY & PASSWORD POLICIES ──────────── -->
<div class="section sec-sec" id="s-security">
  <div class="section-header" onclick="toggleSection('s-security')">
    <div class="section-icon">&#128737;</div>
    <div class="section-title-group">
      <div class="section-title">Security &amp; Password Policies</div>
      <div class="section-sub">Default domain policy, fine-grained policies, privileged account inventory</div>
    </div>
    <span class="section-chevron">&#9658;</span>
  </div>
  <div class="section-body">
    <div class="subsection">
      <div class="subsection-title">Default Domain Password Policy</div>
      <div class="policy-grid">
        <div class="policy-card"><div class="policy-card-label">Min Password Length</div><div class="policy-card-val">$($DefaultPWPolicy.MinPasswordLength)</div></div>
        <div class="policy-card"><div class="policy-card-label">Max Password Age</div><div class="policy-card-val" style="font-size:14px">$($DefaultPWPolicy.MaxPasswordAge)</div></div>
        <div class="policy-card"><div class="policy-card-label">Password History</div><div class="policy-card-val">$($DefaultPWPolicy.PasswordHistoryCount)</div></div>
        <div class="policy-card"><div class="policy-card-label">Complexity Enabled</div><div class="policy-card-val" style="font-size:14px">$($DefaultPWPolicy.ComplexityEnabled)</div></div>
        <div class="policy-card"><div class="policy-card-label">Lockout Threshold</div><div class="policy-card-val">$($DefaultPWPolicy.LockoutThreshold)</div></div>
        <div class="policy-card"><div class="policy-card-label">Lockout Duration</div><div class="policy-card-val" style="font-size:14px">$($DefaultPWPolicy.LockoutDuration)</div></div>
        <div class="policy-card"><div class="policy-card-label">Lockout Window</div><div class="policy-card-val" style="font-size:14px">$($DefaultPWPolicy.LockoutObservationWindow)</div></div>
        <div class="policy-card"><div class="policy-card-label">Reversible Encryption</div><div class="policy-card-val" style="font-size:14px">$($DefaultPWPolicy.ReversibleEncryptionEnabled)</div></div>
      </div>
    </div>
    <div class="subsection">
      <div class="subsection-title">Fine-Grained Password Policies</div>
      <div class="export-bar">
        <span class="export-label">Export:</span>
        <button class="btn-export csv"  onclick="exportCSV('fgpps')">CSV</button>
        <button class="btn-export xlsx" onclick="exportXLSX('fgpps','Fine-Grained Password Policies')">XLSX</button>
      </div>
      $tblFGPPs
    </div>
    <div class="subsection">
      <div class="subsection-title">Domain Admins Group Members (Privileged Accounts)</div>
      <div class="export-bar">
        <span class="export-label">Export:</span>
        <button class="btn-export csv"  onclick="exportCSV('privusers')">CSV</button>
        <button class="btn-export xlsx" onclick="exportXLSX('privusers','Privileged Users')">XLSX</button>
        <button class="btn-export txt"  onclick="exportTXT('privusers')">TXT</button>
      </div>
      $tblPrivUsers
    </div>
  </div>
</div>

</div><!-- /main -->

<footer>
  <strong>$ReportTitle</strong> &nbsp;|&nbsp; $ReportAuthor &nbsp;|&nbsp; $ReportOrg &nbsp;|&nbsp; $ReportRole &nbsp;|&nbsp; Generated: $GeneratedDate
</footer>

<!-- ═══════════════════════════ JAVASCRIPT ═══════════════════════════ -->
<script>
// ── Section toggle ──
function toggleSection(id) {
  const el = document.getElementById(id);
  el.classList.toggle('open');
}
function expandAll()   { document.querySelectorAll('.section').forEach(s => s.classList.add('open')); }
function collapseAll() { document.querySelectorAll('.section').forEach(s => s.classList.remove('open')); }

// ── Global search ──
function globalFilter(q) {
  const term = q.trim().toLowerCase();
  let totalMatch = 0, totalRows = 0;
  document.querySelectorAll('.data-table').forEach(tbl => {
    const rows = tbl.querySelectorAll('tbody tr');
    rows.forEach(row => {
      totalRows++;
      if (!term) {
        row.classList.remove('hidden');
        row.querySelectorAll('td').forEach(td => {
          td.innerHTML = td.textContent;
        });
        totalMatch++;
      } else {
        const text = row.textContent.toLowerCase();
        if (text.includes(term)) {
          row.classList.remove('hidden');
          totalMatch++;
          row.querySelectorAll('td').forEach(td => {
            const raw = td.textContent;
            const idx = raw.toLowerCase().indexOf(term);
            if (idx >= 0) {
              td.innerHTML = raw.substring(0,idx) + '<mark>' + raw.substring(idx, idx+term.length) + '</mark>' + raw.substring(idx+term.length);
            } else {
              td.innerHTML = raw;
            }
          });
        } else {
          row.classList.add('hidden');
        }
      }
    });
  });
  const cnt = document.getElementById('searchCount');
  if (term) {
    cnt.textContent = totalMatch + ' matching rows';
    // auto-expand sections with results
    document.querySelectorAll('.data-table').forEach(tbl => {
      const hasVisible = tbl.querySelectorAll('tbody tr:not(.hidden)').length > 0;
      if (hasVisible) {
        const sec = tbl.closest('.section');
        if (sec) sec.classList.add('open');
      }
    });
  } else {
    cnt.textContent = '';
  }
}

// ── DATA STORE ──
const DATA = {
  dcs:         $jsonDCs,
  sites:       $jsonSites,
  subnets:     $jsonSubnets,
  tier0:       $jsonT0,
  tier1:       $jsonT1,
  tier2:       $jsonT2,
  svcaccounts: $jsonSvc,
  msas:        $jsonMSAs,
  groups:      $jsonGroups,
  computers:   $jsonComputers,
  gpos:        $jsonGPOs,
  trusts:      $jsonTrusts,
  privusers:   $jsonPrivUsers,
  ous:         $jsonOUs
};

// ── CSV Export ──
function exportCSV(key) {
  const rows = DATA[key];
  if (!rows || rows.length === 0) { alert('No data to export.'); return; }
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(','), ...rows.map(r => headers.map(h => {
    const v = r[h] == null ? '' : String(r[h]).replace(/"/g,'""');
    return '"' + v + '"';
  }).join(','))].join('\r\n');
  downloadBlob(csv, key + '_export.csv', 'text/csv;charset=utf-8;');
}

// ── XLSX Export ──
function exportXLSX(key, sheetName) {
  const rows = DATA[key];
  if (!rows || rows.length === 0) { alert('No data to export.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName || key);
  XLSX.writeFile(wb, key + '_export.xlsx');
}

// ── TXT Export ──
function exportTXT(key) {
  const rows = DATA[key];
  if (!rows || rows.length === 0) { alert('No data to export.'); return; }
  const headers = Object.keys(rows[0]);
  const colW = headers.map(h => Math.max(h.length, ...rows.map(r => String(r[h] ?? '').length)));
  const line = colW.map(w => '-'.repeat(w + 2)).join('+');
  const fmt = row => '| ' + headers.map((h,i) => String(row[h] ?? '').padEnd(colW[i])).join(' | ') + ' |';
  const txt = ['IGT AD Assessment - ' + key.toUpperCase(), '=' .repeat(60), '',
    '| ' + headers.map((h,i) => h.padEnd(colW[i])).join(' | ') + ' |',
    line, ...rows.map(fmt), ''].join('\r\n');
  downloadBlob(txt, key + '_export.txt', 'text/plain;charset=utf-8;');
}

// ── PDF Export (print section) ──
function exportPDF(sectionId) {
  const sec = document.getElementById(sectionId);
  if (!sec) return;
  const title = sec.querySelector('.section-title')?.textContent || sectionId;
  const tbl = sec.querySelector('.table-wrapper')?.outerHTML || '<p>No data</p>';
  const win = window.open('','_blank','width=1100,height=700');
  win.document.write('<html><head><title>' + title + '</title>');
  win.document.write('<style>body{font-family:Arial,sans-serif;font-size:12px;color:#111;background:#fff;padding:20px}');
  win.document.write('h2{font-size:16px;margin-bottom:12px}');
  win.document.write('table{width:100%;border-collapse:collapse;font-size:11px}');
  win.document.write('th{background:#e8ecf4;padding:7px 10px;text-align:left;border:1px solid #ccc;font-size:10px;text-transform:uppercase}');
  win.document.write('td{padding:6px 10px;border:1px solid #e0e0e0;vertical-align:top}');
  win.document.write('tr:nth-child(even){background:#f9fafb}');
  win.document.write('.footer{margin-top:20px;font-size:10px;color:#888;border-top:1px solid #e0e0e0;padding-top:10px}</style></head><body>');
  win.document.write('<h2>IGT AD Assessment — ' + title + '</h2>');
  win.document.write(tbl);
  win.document.write('<div class="footer">IGT AD Assessment | $ReportAuthor | $ReportOrg | Generated: $GeneratedDate</div>');
  win.document.write('</body></html>');
  win.document.close();
  setTimeout(() => { win.print(); }, 400);
}

// ── Blob download helper ──
function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
</script>
</body>
</html>
"@

# ============================================================
# WRITE REPORT
# ============================================================
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  [DONE] Report generated successfully!" -ForegroundColor Green
Write-Host "  Path: $ReportFile" -ForegroundColor White
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""

# Auto-open report
try {
    Start-Process $ReportFile
    Write-Host "[+] Opening report in default browser..." -ForegroundColor Cyan
} catch {
    Write-Host "[!] Could not auto-open. Please open manually: $ReportFile" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "[Summary]" -ForegroundColor Cyan
Write-Host "  Total Users:     $($Stats.TotalUsers) ($($Stats.EnabledUsers) enabled, $($Stats.DisabledUsers) disabled)" -ForegroundColor White
Write-Host "  Tier 0 (DA):     $($Stats.T0Users)" -ForegroundColor Red
Write-Host "  Tier 1 (Srvr):   $($Stats.T1Users)" -ForegroundColor Yellow
Write-Host "  Tier 2 (Client): $($Stats.T2Users)" -ForegroundColor Cyan
Write-Host "  Computers:       $($Stats.TotalComputers)" -ForegroundColor White
Write-Host "  Groups:          $($Stats.TotalGroups)" -ForegroundColor White
Write-Host "  GPOs:            $($Stats.TotalGPOs)" -ForegroundColor White
Write-Host "  Domain Ctrls:    $($Stats.TotalDCs)" -ForegroundColor White
Write-Host "  Stale Users:     $($Stats.StaleUsers) (no logon 90+ days)" -ForegroundColor $(if ($Stats.StaleUsers -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""
