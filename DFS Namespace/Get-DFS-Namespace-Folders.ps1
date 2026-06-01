#Requires -Version 5.1

<#
.SYNOPSIS
    DFS Namespace Audit Tool - Documents all DFS Namespaces, folder targets,
    hosting servers, and NTFS security group assignments.

.DESCRIPTION
    Enumerates every DFS Namespace root hosted on this server (or a specified server),
    then for each DFS folder documents:
      - Folder target UNC path(s)
      - Hosting server(s) parsed from target paths
      - All NTFS ACL entries (security groups, users, rights, inherited flag)
      - AD group type resolution (if ActiveDirectory module is present)

    Exports three deliverables to %USERPROFILE%\Desktop\DFS_Export (auto-created):
      1. CSV  - raw flat data, one row per ACL entry
      2. XLSX - two-sheet workbook (Summary + Details) via raw Open XML / ZipArchive
      3. HTML - dark-themed interactive report with live search and collapsible sections

    Run from the primary DFS Namespace server.
    Required  : RSAT DFS Management Tools  (DFSN module)
    Optional  : RSAT AD DS Tools           (ActiveDirectory module - enables group type resolution)

.PARAMETER ExportPath
    Destination folder for all output files.
    Default: %USERPROFILE%\Desktop\DFS_Export

.PARAMETER NamespaceServer
    DFS server to query. Defaults to the local computer ($env:COMPUTERNAME).

.PARAMETER SkipACL
    Skip NTFS ACL collection entirely. Useful if running without share read rights
    or to speed up a quick namespace structure dump.

.EXAMPLE
    .\IGT-DFS-NamespaceAudit.ps1
    Run against the local server; export to Desktop\DFS_Export.

.EXAMPLE
    .\IGT-DFS-NamespaceAudit.ps1 -NamespaceServer "RNOP-DFSR01.ad.igt.com"
    Run against a specific DFS server by FQDN.

.EXAMPLE
    .\IGT-DFS-NamespaceAudit.ps1 -SkipACL -ExportPath "C:\Exports\DFS"
    Skip ACL collection; write output to a custom path.

.NOTES
    Author   : Stephen McKee - Server Administration
    Version  : 1.0.0
    Platform : PowerShell 5.1
    Modules  : DFSN (required), ActiveDirectory (optional)
    No external module dependencies - XLSX written via raw Open XML / ZipArchive.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExportPath = (Join-Path $env:USERPROFILE 'Desktop\DFS_Export'),

    [Parameter()]
    [string]$NamespaceServer = $env:COMPUTERNAME,

    [Parameter()]
    [switch]$SkipACL
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ============================================================
#  REGION 1 - INITIALIZATION
# ============================================================
$ScriptVersion  = '1.0.0'
$StartTime      = Get-Date
$Timestamp      = $StartTime | Get-Date -Format 'yyyyMMdd_HHmmss'
$ReportDate     = $StartTime | Get-Date -Format 'MMMM dd, yyyy HH:mm:ss'
$RunningUser    = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$LogFile  = Join-Path $ExportPath "DFS_Audit_$Timestamp.log"
$CsvFile  = Join-Path $ExportPath "DFS_Audit_$Timestamp.csv"
$XlsxFile = Join-Path $ExportPath "DFS_Audit_$Timestamp.xlsx"
$HtmlFile = Join-Path $ExportPath "DFS_Audit_$Timestamp.html"

# Create export directory if missing
if (-not (Test-Path -Path $ExportPath -PathType Container)) {
    try {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Error "[FATAL] Cannot create export directory '$ExportPath': $_"
        exit 1
    }
}

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    switch ($Level) {
        'INFO'    { Write-Host "  $entry" -ForegroundColor Cyan    }
        'WARN'    { Write-Host "  $entry" -ForegroundColor Yellow  }
        'ERROR'   { Write-Host "  $entry" -ForegroundColor Red     }
        'SUCCESS' { Write-Host "  $entry" -ForegroundColor Green   }
    }
}

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host '     IGT DFS Namespace Audit Tool  v' -NoNewline -ForegroundColor Cyan
Write-Host $ScriptVersion -ForegroundColor White
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host "  Server  : $NamespaceServer"   -ForegroundColor White
Write-Host "  RunAs   : $RunningUser"        -ForegroundColor White
Write-Host "  Export  : $ExportPath"         -ForegroundColor White
Write-Host "  Started : $ReportDate"         -ForegroundColor White
Write-Host "  SkipACL : $($SkipACL.IsPresent)" -ForegroundColor White
Write-Host '  ================================================================' -ForegroundColor Cyan
Write-Host ''

Write-Log "Script started | Server=$NamespaceServer | User=$RunningUser | SkipACL=$($SkipACL.IsPresent)"

# Check elevation (advisory only)
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log 'Not running as Administrator. ACL reads may be incomplete on restricted shares.' -Level WARN
}

# ============================================================
#  REGION 2 - MODULE IMPORTS
# ============================================================
Add-Type -AssemblyName 'System.IO.Compression'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# DFSN module (required)
if (-not (Get-Module -Name DFSN -ErrorAction SilentlyContinue)) {
    try {
        Import-Module -Name DFSN -ErrorAction Stop
        Write-Log 'DFSN module loaded.'
    }
    catch {
        Write-Log "DFSN module unavailable. Install RSAT: Add-WindowsFeature FS-DFS-Namespace,RSAT-DFS-Mgmt-Con" -Level ERROR
        exit 1
    }
}
else {
    Write-Log 'DFSN module already loaded.'
}

# ActiveDirectory module (optional - group type resolution)
$ADModuleAvailable = $false
if (-not $SkipACL) {
    if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
        try {
            Import-Module -Name ActiveDirectory -ErrorAction Stop
            $ADModuleAvailable = $true
            Write-Log 'ActiveDirectory module loaded. Group type resolution enabled.'
        }
        catch {
            Write-Log 'ActiveDirectory module import failed. Group types will show as Unknown.' -Level WARN
        }
    }
    else {
        Write-Log 'ActiveDirectory module not found. Group types will show as Unknown.' -Level WARN
    }
}

# ============================================================
#  REGION 3 - HELPER FUNCTIONS
# ============================================================

#-- XML special-character escaper (for XLSX cell values)
function Escape-Xml {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $v = $Value -replace '&',  '&amp;'
    $v = $v     -replace '<',  '&lt;'
    $v = $v     -replace '>',  '&gt;'
    $v = $v     -replace '"',  '&quot;'
    $v = $v     -replace "'",  '&apos;'
    return $v
}

#-- HTML special-character escaper (for HTML report)
function Escape-Html {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $v = $Value -replace '&',  '&amp;'
    $v = $v     -replace '<',  '&lt;'
    $v = $v     -replace '>',  '&gt;'
    $v = $v     -replace '"',  '&quot;'
    return $v
}

#-- Extract server hostname from UNC path  \\SERVER\share\... -> SERVER
function Get-UNCServer {
    param([string]$UNCPath)
    if ($UNCPath -match '^\\\\([^\\]+)') { return $Matches[1] }
    return $UNCPath
}

#-- Collect NTFS ACL entries for a given UNC path
function Get-FolderACLEntries {
    param([string]$Path)

    $entries = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($SkipACL) {
        $entries.Add([PSCustomObject]@{
            Identity    = '(ACL collection skipped)'
            Domain      = ''
            SamAccount  = ''
            Rights      = ''
            AccessType  = ''
            IsInherited = $false
            GroupType   = 'N/A'
        })
        return $entries
    }

    try {
        $acl = Get-Acl -Path $Path -ErrorAction Stop

        foreach ($ace in $acl.Access) {
            $identity = $ace.IdentityReference.Value

            # Skip well-known built-in / system identities
            if ($identity -match '^(NT AUTHORITY\\|BUILTIN\\|Creator Owner|Everyone|S-1-5-)') { continue }

            # Parse domain\samaccount
            $domain     = ''
            $samAccount = $identity
            if ($identity -match '^(.+)\\(.+)$') {
                $domain     = $Matches[1]
                $samAccount = $Matches[2]
            }

            # AD group type resolution (optional)
            $groupType = 'Unknown'
            if ($ADModuleAvailable) {
                try {
                    $cleanSam = $samAccount -replace "'", "''"
                    $adObj = Get-ADObject -Filter "SamAccountName -eq '$cleanSam'" `
                                 -Properties objectClass -ErrorAction SilentlyContinue
                    if ($null -ne $adObj) {
                        if ($adObj.objectClass -eq 'group') {
                            $grp = Get-ADGroup -Identity $samAccount `
                                       -Properties GroupScope, GroupCategory `
                                       -ErrorAction SilentlyContinue
                            if ($null -ne $grp) {
                                $groupType = "$($grp.GroupScope) $($grp.GroupCategory) Group"
                            }
                            else {
                                $groupType = 'AD Group'
                            }
                        }
                        elseif ($adObj.objectClass -eq 'user')     { $groupType = 'AD User Account'     }
                        elseif ($adObj.objectClass -eq 'computer') { $groupType = 'AD Computer Account' }
                        else                                        { $groupType = $adObj.objectClass    }
                    }
                }
                catch {
                    $groupType = 'Lookup Error'
                }
            }

            $entries.Add([PSCustomObject]@{
                Identity    = $identity
                Domain      = $domain
                SamAccount  = $samAccount
                Rights      = $ace.FileSystemRights.ToString()
                AccessType  = $ace.AccessControlType.ToString()
                IsInherited = $ace.IsInherited
                GroupType   = $groupType
            })
        }
    }
    catch {
        $entries.Add([PSCustomObject]@{
            Identity    = "ACL read failed: $($_.Exception.Message)"
            Domain      = ''
            SamAccount  = ''
            Rights      = 'ERROR'
            AccessType  = 'ERROR'
            IsInherited = $false
            GroupType   = 'ERROR'
        })
    }

    if ($entries.Count -eq 0) {
        $entries.Add([PSCustomObject]@{
            Identity    = '(No non-system ACL entries found)'
            Domain      = ''
            SamAccount  = ''
            Rights      = ''
            AccessType  = ''
            IsInherited = $false
            GroupType   = ''
        })
    }

    return $entries
}

# ============================================================
#  REGION 4 - XLSX WRITER  (raw Open XML / ZipArchive - no Excel required)
# ============================================================
function Export-XlsxReport {
    param(
        [string]$FilePath,
        [System.Collections.Generic.List[PSCustomObject]]$SummaryRows,
        [System.Collections.Generic.List[PSCustomObject]]$DetailRows
    )

    # ---- Summary sheet XML ----
    $sbS = [System.Text.StringBuilder]::new(65536)
    [void]$sbS.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sbS.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sbS.Append('<sheetData>')

    $sHdrs = @('Namespace Path','Type','Description','State','Namespace Servers','Folder Count')
    [void]$sbS.Append('<row r="1">')
    for ($i = 0; $i -lt $sHdrs.Count; $i++) {
        $col = [char](65 + $i)
        [void]$sbS.Append("<c r=`"${col}1`" t=`"inlineStr`" s=`"1`"><is><t>$(Escape-Xml $sHdrs[$i])</t></is></c>")
    }
    [void]$sbS.Append('</row>')

    $rn = 2
    foreach ($r in $SummaryRows) {
        [void]$sbS.Append("<row r=`"$rn`">")
        [void]$sbS.Append("<c r=`"A$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespacePath)</t></is></c>")
        [void]$sbS.Append("<c r=`"B$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceType)</t></is></c>")
        [void]$sbS.Append("<c r=`"C$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.Description)</t></is></c>")
        [void]$sbS.Append("<c r=`"D$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.State)</t></is></c>")
        [void]$sbS.Append("<c r=`"E$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceServers)</t></is></c>")
        [void]$sbS.Append("<c r=`"F$rn`" t=`"n`"><v>$($r.FolderCount)</v></c>")
        [void]$sbS.Append('</row>')
        $rn++
    }
    [void]$sbS.Append('</sheetData></worksheet>')

    # ---- Details sheet XML (columns A-S = 19 cols, all < 26) ----
    $sbD = [System.Text.StringBuilder]::new(1048576)
    [void]$sbD.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sbD.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sbD.Append('<sheetData>')

    $dHdrs = @(
        'Namespace','Namespace Type','NS Description','NS State','NS Servers',
        'DFS Folder Path','Folder Description','Folder State',
        'Target Path','Target Server','Target State','Target Priority',
        'Security Identity','Domain','SAM Account','Rights','Access Type','Is Inherited','Group Type'
    )
    [void]$sbD.Append('<row r="1">')
    for ($i = 0; $i -lt $dHdrs.Count; $i++) {
        $col = [char](65 + $i)
        [void]$sbD.Append("<c r=`"${col}1`" t=`"inlineStr`" s=`"1`"><is><t>$(Escape-Xml $dHdrs[$i])</t></is></c>")
    }
    [void]$sbD.Append('</row>')

    $rn = 2
    foreach ($r in $DetailRows) {
        [void]$sbD.Append("<row r=`"$rn`">")
        [void]$sbD.Append("<c r=`"A$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.Namespace)</t></is></c>")
        [void]$sbD.Append("<c r=`"B$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceType)</t></is></c>")
        [void]$sbD.Append("<c r=`"C$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceDesc)</t></is></c>")
        [void]$sbD.Append("<c r=`"D$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceState)</t></is></c>")
        [void]$sbD.Append("<c r=`"E$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.NamespaceServers)</t></is></c>")
        [void]$sbD.Append("<c r=`"F$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.FolderPath)</t></is></c>")
        [void]$sbD.Append("<c r=`"G$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.FolderDesc)</t></is></c>")
        [void]$sbD.Append("<c r=`"H$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.FolderState)</t></is></c>")
        [void]$sbD.Append("<c r=`"I$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.TargetPath)</t></is></c>")
        [void]$sbD.Append("<c r=`"J$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.TargetServer)</t></is></c>")
        [void]$sbD.Append("<c r=`"K$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.TargetState)</t></is></c>")
        [void]$sbD.Append("<c r=`"L$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.TargetPriority)</t></is></c>")
        [void]$sbD.Append("<c r=`"M$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.SecurityIdentity)</t></is></c>")
        [void]$sbD.Append("<c r=`"N$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.Domain)</t></is></c>")
        [void]$sbD.Append("<c r=`"O$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.SamAccount)</t></is></c>")
        [void]$sbD.Append("<c r=`"P$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.Rights)</t></is></c>")
        [void]$sbD.Append("<c r=`"Q$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.AccessType)</t></is></c>")
        [void]$sbD.Append("<c r=`"R$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.IsInherited)</t></is></c>")
        [void]$sbD.Append("<c r=`"S$rn`" t=`"inlineStr`"><is><t>$(Escape-Xml $r.GroupType)</t></is></c>")
        [void]$sbD.Append('</row>')
        $rn++
    }
    [void]$sbD.Append('</sheetData></worksheet>')

    # ---- Open XML package manifest files ----
    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

    $relsRoot = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

    $wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/><sheet name="DFS Details" sheetId="2" r:id="rId2"/></sheets></workbook>'

    # Styles: index 0 = normal, index 1 = bold white on dark-blue (header rows)
    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><sz val="10"/><b/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1E3A5F"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>'

    # ---- Write XLSX (ZIP archive) ----
    $fs  = $null
    $zip = $null
    try {
        $fs  = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $enc = [System.Text.Encoding]::UTF8

        $zipEntries = [ordered]@{
            '[Content_Types].xml'        = $contentTypes
            '_rels/.rels'                = $relsRoot
            'xl/workbook.xml'            = $workbook
            'xl/_rels/workbook.xml.rels' = $wbRels
            'xl/styles.xml'              = $styles
            'xl/worksheets/sheet1.xml'   = $sbS.ToString()
            'xl/worksheets/sheet2.xml'   = $sbD.ToString()
        }

        foreach ($entryName in $zipEntries.Keys) {
            $entry  = $zip.CreateEntry($entryName)
            $stream = $entry.Open()
            $bytes  = $enc.GetBytes($zipEntries[$entryName])
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush()
            $stream.Dispose()
        }
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        if ($null -ne $fs)  { $fs.Dispose()  }
    }
}

# ============================================================
#  REGION 5 - HTML REPORT BUILDER
# ============================================================
function Build-HtmlReport {
    param(
        [string]$OutputPath,
        [System.Collections.Generic.List[PSCustomObject]]$Records,
        [System.Collections.Generic.List[PSCustomObject]]$Summary,
        [string]$Server,
        [string]$RunUser,
        [string]$GeneratedAt,
        [string]$Version
    )

    # Compute statistics
    $totalNS      = ($Records | Select-Object -ExpandProperty Namespace -Unique | Measure-Object).Count
    $totalFolders = ($Records | Where-Object { $_.FolderPath -ne '' -and $_.FolderPath -notlike '(No*' } |
                     Select-Object -Property Namespace, FolderPath -Unique | Measure-Object).Count
    $totalTargets = ($Records | Where-Object { $_.TargetPath -ne '' -and $_.TargetPath -notlike '(No*' } |
                     Select-Object -ExpandProperty TargetPath -Unique | Measure-Object).Count
    $totalACL     = ($Records | Where-Object {
                         $_.SecurityIdentity -ne '' -and
                         $_.SecurityIdentity -notlike '(No*' -and
                         $_.SecurityIdentity -notlike '(ACL*'
                     } | Measure-Object).Count

    $sb = [System.Text.StringBuilder]::new(2097152)

    [void]$sb.Append(@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>IGT DFS Namespace Audit Report</title>
<style>
:root{
  --bg0:#0d1117;--bg1:#161b22;--bg2:#1c2128;--bg3:#21262d;
  --blue:#58a6ff;--blue-dark:#1e3a5f;--blue-mid:#255075;
  --green:#3fb950;--green-bg:#1a3a1e;--green-bd:#2d5a31;
  --amber:#d29922;--amber-bg:#3a2a1a;
  --red:#f85149;--red-bg:#3a1a1a;--red-bd:#5a2d2d;
  --purple:#bc8cff;--purple-bg:#2a1a3a;--purple-bd:#4a2d6a;
  --teal:#39d353;
  --txt:#e6edf3;--txt2:#8b949e;
  --bd:#30363d;--bd2:#58a6ff44;
  --shadow:0 4px 24px rgba(0,0,0,.5);
  --r:8px;--r2:4px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;background:var(--bg0);color:var(--txt);line-height:1.5}

/* ── HEADER ── */
.pg-hdr{
  background:linear-gradient(135deg,#0d1117 0%,var(--blue-dark) 100%);
  border-bottom:1px solid var(--bd);padding:18px 32px;
  position:sticky;top:0;z-index:100;
  display:flex;align-items:center;justify-content:space-between;
  flex-wrap:wrap;gap:12px;box-shadow:var(--shadow)
}
.hdr-title{font-size:1.3rem;font-weight:700;color:var(--blue)}
.hdr-title span{color:var(--txt);font-weight:400}
.hdr-meta{font-size:11px;color:var(--txt2);margin-top:3px}
.hdr-meta b{color:var(--txt)}
.srch-wrap{display:flex;align-items:center;gap:8px}
#sBox{
  background:var(--bg1);border:1px solid var(--bd);border-radius:20px;
  color:var(--txt);font-size:13px;padding:7px 16px 7px 36px;width:300px;outline:none;
  transition:border .2s,box-shadow .2s;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='%238b949e' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:12px center
}
#sBox:focus{border-color:var(--blue);box-shadow:0 0 0 3px #58a6ff22}
#sBox::placeholder{color:var(--txt2)}
#sClear{
  background:transparent;border:1px solid var(--bd);border-radius:var(--r2);
  color:var(--txt2);cursor:pointer;font-size:12px;padding:6px 10px;
  transition:all .15s;display:none
}
#sClear:hover{border-color:var(--blue);color:var(--blue)}
#sClear.vis{display:inline-block}

/* ── STATS BAR ── */
.stats-bar{
  background:var(--bg1);border-bottom:1px solid var(--bd);
  padding:10px 32px;display:flex;gap:24px;flex-wrap:wrap;align-items:center
}
.stat-item{display:flex;align-items:baseline;gap:5px;font-size:12px}
.stat-n{font-size:20px;font-weight:700;color:var(--blue);line-height:1}
.stat-lbl{color:var(--txt2)}
.stat-sep{color:var(--bd)}
#sStatus{font-size:11px;color:var(--amber);font-style:italic;margin-left:auto}

/* ── MAIN ── */
.main{padding:20px 32px 60px;max-width:1700px;margin:0 auto}

/* ── TOOLBAR ── */
.toolbar{display:flex;gap:8px;margin-bottom:16px;align-items:center}
.tbar-lbl{color:var(--txt2);font-size:11px}
.btn{
  background:var(--bg1);border:1px solid var(--bd);border-radius:var(--r2);
  color:var(--txt);cursor:pointer;font-size:11px;padding:5px 12px;transition:all .15s
}
.btn:hover{border-color:var(--blue);color:var(--blue);background:#1a2a3a}

/* ── NAMESPACE CARD ── */
.ns-card{
  background:var(--bg2);border:1px solid var(--bd);border-radius:var(--r);
  margin-bottom:16px;overflow:hidden;box-shadow:var(--shadow);
  transition:border-color .2s
}
.ns-card:hover{border-color:var(--bd2)}
.ns-hdr{
  background:linear-gradient(90deg,var(--blue-dark),var(--bg2));
  padding:12px 18px;cursor:pointer;display:flex;align-items:center;
  gap:10px;user-select:none;border-bottom:1px solid var(--bd);
  transition:background .2s
}
.ns-hdr:hover{background:linear-gradient(90deg,var(--blue-mid),var(--bg3))}
.ti{color:var(--blue);font-size:11px;width:14px;text-align:center;flex-shrink:0;transition:transform .2s}
.ns-ttl{font-size:14px;font-weight:600;color:var(--blue);font-family:'Consolas','Courier New',monospace}
.ns-meta{color:var(--txt2);font-size:11px;margin-left:auto;white-space:nowrap}
.ns-info{
  display:flex;gap:24px;padding:9px 18px;background:var(--bg1);
  border-bottom:1px solid var(--bd);flex-wrap:wrap
}
.ii{font-size:11px}
.il{color:var(--txt2)}
.iv{color:var(--txt);font-weight:500}

/* ── FOLDER SECTION ── */
.f-sec{border-bottom:1px solid var(--bd)}
.f-sec:last-child{border-bottom:none}
.f-hdr{
  padding:9px 18px 9px 34px;cursor:pointer;display:flex;align-items:center;
  gap:8px;background:var(--bg1);transition:background .15s;user-select:none
}
.f-hdr:hover{background:var(--bg3)}
.fi{color:var(--amber);font-size:11px;width:12px;text-align:center;flex-shrink:0}
.f-ttl{font-size:12px;font-weight:500;color:var(--txt);font-family:'Consolas','Courier New',monospace}
.f-meta{color:var(--txt2);font-size:11px;margin-left:auto}
.f-body{padding:12px 18px 12px 42px}
.f-body.bc{display:none}

/* ── TARGET BLOCK ── */
.t-blk{
  background:var(--bg0);border:1px solid var(--bd);border-radius:var(--r2);
  margin-bottom:10px;overflow:hidden
}
.t-blk:last-child{margin-bottom:0}
.t-hdr{
  display:flex;align-items:center;gap:8px;padding:7px 12px;
  background:var(--bg3);border-bottom:1px solid var(--bd);flex-wrap:wrap
}
.t-path{font-family:'Consolas','Courier New',monospace;font-size:11px;color:var(--teal);font-weight:500}
.tag{font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;letter-spacing:.3px;white-space:nowrap}
.t-on {background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.t-off{background:var(--red-bg);color:var(--red);border:1px solid var(--red-bd)}
.t-srv{background:#1a2a3a;color:#6ab4ff;border:1px solid #2d4a6a}
.t-pri{background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd)}

/* ── ACL TABLE ── */
.acl-tbl{width:100%;border-collapse:collapse;font-size:11px}
.acl-tbl th{
  background:var(--blue-dark);color:#a8c8f0;padding:6px 10px;
  text-align:left;font-weight:600;font-size:10px;text-transform:uppercase;
  letter-spacing:.5px;white-space:nowrap;border-bottom:1px solid var(--bd)
}
.acl-tbl td{padding:5px 10px;border-bottom:1px solid var(--bd);vertical-align:middle;color:var(--txt)}
.acl-tbl tr:last-child td{border-bottom:none}
.acl-tbl tr:nth-child(even) td{background:var(--bg2)}
.acl-tbl tr:hover td{background:var(--bg3)}
.c-id{font-family:'Consolas','Courier New',monospace;color:#c9d1d9}
.c-sam{color:var(--blue);font-weight:500}
.c-rts{color:var(--txt2);max-width:220px;word-break:break-word}
.c-allow{color:var(--green);font-weight:600}
.c-deny {color:var(--red);font-weight:600}
.c-inh  {color:var(--txt2)}
.c-ninh {color:var(--blue)}

/* group-type badges */
.gb{font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;display:inline-block}
.gt-dl {background:#1a2a3a;color:#6ab4ff;border:1px solid #2d4a6a}
.gt-gl {background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.gt-un {background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd)}
.gt-usr{background:var(--amber-bg);color:var(--amber);border:1px solid #6a4a1a}
.gt-err{background:var(--red-bg);color:var(--red);border:1px solid var(--red-bd)}
.gt-unk{background:var(--bg3);color:var(--txt2);border:1px solid var(--bd)}

/* search/filter visibility */
.rh{display:none!important}
.bh{display:none!important}
.sh{display:none!important}
.ch{display:none!important}

/* ── EMPTY STATE ── */
.empty{padding:20px 18px;text-align:center;color:var(--txt2);font-style:italic;font-size:12px}

/* ── FOOTER ── */
.pg-ftr{
  background:var(--bg1);border-top:1px solid var(--bd);
  color:var(--txt2);font-size:11px;padding:14px 32px;text-align:center
}
.pg-ftr b{color:var(--txt)}

/* ── SCROLLBAR ── */
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:var(--bg0)}
::-webkit-scrollbar-thumb{background:var(--bd);border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:var(--txt2)}

@media(max-width:768px){
  .pg-hdr,.main,.stats-bar{padding-left:16px;padding-right:16px}
  #sBox{width:200px}
}
</style>
</head>
<body>

<!-- ═══════════ HEADER ═══════════ -->
<div class="pg-hdr">
  <div>
    <div class="hdr-title">IGT <span>DFS Namespace Audit Report</span></div>
    <div class="hdr-meta">
      Server: <b>$Server</b> &nbsp;|&nbsp;
      Generated: <b>$GeneratedAt</b> &nbsp;|&nbsp;
      Run By: <b>$RunUser</b> &nbsp;|&nbsp;
      Script v$Version
    </div>
  </div>
  <div class="srch-wrap">
    <input id="sBox" type="text" placeholder="Search namespace, folder, group, server..."
           oninput="doSearch(this.value)" autocomplete="off">
    <button id="sClear" onclick="clearSearch()">&#x2715; Clear</button>
  </div>
</div>

<!-- ═══════════ STATS BAR ═══════════ -->
<div class="stats-bar">
  <div class="stat-item"><span class="stat-n">$totalNS</span><span class="stat-lbl">Namespaces</span></div>
  <span class="stat-sep">|</span>
  <div class="stat-item"><span class="stat-n">$totalFolders</span><span class="stat-lbl">DFS Folders</span></div>
  <span class="stat-sep">|</span>
  <div class="stat-item"><span class="stat-n">$totalTargets</span><span class="stat-lbl">Folder Targets</span></div>
  <span class="stat-sep">|</span>
  <div class="stat-item"><span class="stat-n">$totalACL</span><span class="stat-lbl">ACL Entries</span></div>
  <span id="sStatus"></span>
</div>

<!-- ═══════════ MAIN CONTENT ═══════════ -->
<div class="main">
  <div class="toolbar">
    <span class="tbar-lbl">Sections:</span>
    <button class="btn" onclick="expandAll()">&#x25BC; Expand All</button>
    <button class="btn" onclick="collapseAll()">&#x25B6; Collapse All</button>
  </div>
"@)

    # ---- BUILD NAMESPACE CARDS ----
    $nsGroups = $Records | Group-Object -Property Namespace
    $nsIdx    = 0

    foreach ($nsGrp in $nsGroups) {
        $nsIdx++
        $nsPath    = Escape-Html $nsGrp.Name
        $nsFirst   = $nsGrp.Group[0]
        $nsType    = Escape-Html $nsFirst.NamespaceType
        $nsState   = Escape-Html $nsFirst.NamespaceState
        $nsDesc    = Escape-Html $nsFirst.NamespaceDesc
        $nsSrvs    = Escape-Html $nsFirst.NamespaceServers
        $nsFldCt   = ($nsGrp.Group |
                       Where-Object { $_.FolderPath -ne '' -and $_.FolderPath -notlike '(No*' } |
                       Select-Object -ExpandProperty FolderPath -Unique | Measure-Object).Count
        $nsId      = "ns$nsIdx"
        $stCls     = if ($nsState -eq 'Online') { 't-on' } else { 't-off' }

        [void]$sb.Append(@"

  <!-- === NAMESPACE: $nsPath === -->
  <div class="ns-card" id="${nsId}-card">
    <div class="ns-hdr" onclick="toggleSec('$nsId')">
      <span class="ti" id="${nsId}-ti">&#x25BC;</span>
      <span class="ns-ttl">$nsPath</span>
      <span class="tag $stCls">$nsState</span>
      <span class="ns-meta">$nsFldCt folder(s)&nbsp;&nbsp;|&nbsp;&nbsp;Type: $nsType</span>
    </div>
    <div id="${nsId}-bd">
      <div class="ns-info">
        <div class="ii"><span class="il">Type: </span><span class="iv">$nsType</span></div>
        <div class="ii"><span class="il">Namespace Servers: </span><span class="iv">$nsSrvs</span></div>
"@)
        if ($nsDesc) {
            [void]$sb.Append("        <div class=`"ii`"><span class=`"il`">Description: </span><span class=`"iv`">$nsDesc</span></div>`n")
        }
        [void]$sb.Append("      </div>`n")

        # ---- FOLDER SECTIONS ----
        $fldGroups = $nsGrp.Group | Group-Object -Property FolderPath
        $fldIdx    = 0

        foreach ($fGrp in $fldGroups) {
            $fldIdx++
            $fPath   = $fGrp.Group[0].FolderPath

            if ($fPath -eq '(No folders)') {
                [void]$sb.Append("      <div class=`"empty`">No DFS folders defined in this namespace.</div>`n")
                continue
            }

            $fPathH   = Escape-Html $fPath
            $fState   = Escape-Html $fGrp.Group[0].FolderState
            $fDesc    = Escape-Html $fGrp.Group[0].FolderDesc
            $fTgtCt   = ($fGrp.Group | Select-Object -ExpandProperty TargetPath -Unique | Measure-Object).Count
            $fId      = "${nsId}f${fldIdx}"
            $fStCls   = if ($fState -eq 'Online') { 't-on' } else { 't-off' }

            [void]$sb.Append(@"
      <div class="f-sec" id="${fId}-sec">
        <div class="f-hdr" onclick="toggleSec('$fId')">
          <span class="fi" id="${fId}-ti">&#x25BC;</span>
          <span class="f-ttl">$fPathH</span>
          <span class="tag $fStCls" style="font-size:9px">$fState</span>
"@)
            if ($fDesc) { [void]$sb.Append("          <span class=`"f-meta`">$fDesc &nbsp;|&nbsp;</span>`n") }
            [void]$sb.Append("          <span class=`"f-meta`">$fTgtCt target(s)</span>`n        </div>`n")
            [void]$sb.Append("        <div id=`"${fId}-bd`" class=`"f-body`">`n")

            # ---- TARGET BLOCKS ----
            $tgtGroups = $fGrp.Group | Group-Object -Property TargetPath
            $tgtIdx    = 0

            foreach ($tGrp in $tgtGroups) {
                $tgtIdx++
                $tPath    = $tGrp.Group[0].TargetPath

                if ($tPath -eq '(No targets)') {
                    [void]$sb.Append("          <div class=`"empty`">No targets configured for this folder.</div>`n")
                    continue
                }

                $tPathH   = Escape-Html $tPath
                $tServer  = Escape-Html $tGrp.Group[0].TargetServer
                $tState   = Escape-Html $tGrp.Group[0].TargetState
                $tPri     = Escape-Html $tGrp.Group[0].TargetPriority
                $tId      = "${fId}t${tgtIdx}"
                $tStCls   = if ($tState -eq 'Online') { 't-on' } else { 't-off' }

                [void]$sb.Append(@"
          <div class="t-blk" id="${tId}-blk">
            <div class="t-hdr">
              <span class="t-path">$tPathH</span>
              <span class="tag $tStCls">$tState</span>
              <span class="tag t-srv">$tServer</span>
"@)
                if ($tPri) {
                    [void]$sb.Append("              <span class=`"tag t-pri`">$tPri</span>`n")
                }
                [void]$sb.Append(@"
            </div>
            <table class="acl-tbl">
              <thead>
                <tr>
                  <th>Security Identity</th>
                  <th>SAM Account</th>
                  <th>Domain</th>
                  <th>Group Type</th>
                  <th>Rights</th>
                  <th>Access</th>
                  <th>Inherited</th>
                </tr>
              </thead>
              <tbody>
"@)

                # ---- ACL ROWS ----
                foreach ($ace in $tGrp.Group) {
                    $aId    = Escape-Html $ace.SecurityIdentity
                    $aSam   = Escape-Html $ace.SamAccount
                    $aDom   = Escape-Html $ace.Domain
                    $aGT    = Escape-Html $ace.GroupType
                    $aRts   = Escape-Html $ace.Rights
                    $aAcc   = Escape-Html $ace.AccessType
                    $aInh   = if ($ace.IsInherited -eq 'True') { 'Yes' } else { 'No' }
                    $aAcCls = if ($aAcc -eq 'Allow') { 'c-allow' } elseif ($aAcc -eq 'Deny') { 'c-deny' } else { '' }
                    $aInCls = if ($aInh -eq 'Yes') { 'c-inh' } else { 'c-ninh' }

                    # Group-type badge CSS class
                    $gtCls = 'gt-unk'
                    if     ($aGT -match 'DomainLocal') { $gtCls = 'gt-dl'  }
                    elseif ($aGT -match 'Global')      { $gtCls = 'gt-gl'  }
                    elseif ($aGT -match 'Universal')   { $gtCls = 'gt-un'  }
                    elseif ($aGT -match 'User')        { $gtCls = 'gt-usr' }
                    elseif ($aGT -match 'Error')       { $gtCls = 'gt-err' }

                    # Build lowercase search-data string (identity, group, path, server, namespace, folder)
                    $srchRaw = "$($ace.SecurityIdentity) $($ace.SamAccount) $($ace.Domain) $($ace.GroupType) $($ace.Rights) $($ace.AccessType) $tPath $($tGrp.Group[0].TargetServer) $fPath $($nsGrp.Name)"
                    $srchVal = Escape-Html $srchRaw.ToLower()

                    [void]$sb.Append(@"
                <tr class="acl-row" data-s="$srchVal">
                  <td class="c-id">$aId</td>
                  <td class="c-sam">$aSam</td>
                  <td>$aDom</td>
                  <td><span class="gb $gtCls">$aGT</span></td>
                  <td class="c-rts">$aRts</td>
                  <td class="$aAcCls">$aAcc</td>
                  <td class="$aInCls">$aInh</td>
                </tr>
"@)
                }

                [void]$sb.Append("              </tbody>`n            </table>`n          </div>`n")
            }  # end target loop

            [void]$sb.Append("        </div>`n      </div>`n")
        }  # end folder loop

        [void]$sb.Append("    </div>`n  </div>`n")
    }  # end namespace loop

    [void]$sb.Append(@"
</div>

<!-- ═══════════ FOOTER ═══════════ -->
<div class="pg-ftr">
  <b>IGT DFS Namespace Audit Report</b> &nbsp;|&nbsp;
  Generated by <b>$RunUser</b> &nbsp;on&nbsp; <b>$GeneratedAt</b> &nbsp;from&nbsp; <b>$Server</b> &nbsp;|&nbsp;
  Script v$Version &nbsp;|&nbsp; IGT Server Administration
</div>

<script>
// ─────────────────────────────────────────────
//  COLLAPSE / EXPAND
// ─────────────────────────────────────────────
var colMap = {};

function toggleSec(id) {
  var bd = document.getElementById(id + '-bd');
  var ti = document.getElementById(id + '-ti');
  if (!bd) return;
  if (bd.classList.contains('bc')) {
    bd.classList.remove('bc');
    if (ti) ti.innerHTML = '&#x25BC;';
    colMap[id] = false;
  } else {
    bd.classList.add('bc');
    if (ti) ti.innerHTML = '&#x25B6;';
    colMap[id] = true;
  }
}

function expandAll() {
  document.querySelectorAll('[id$="-bd"]').forEach(function(el){ el.classList.remove('bc'); });
  document.querySelectorAll('.ti,.fi').forEach(function(ic){ ic.innerHTML = '&#x25BC;'; });
  Object.keys(colMap).forEach(function(k){ colMap[k] = false; });
}

function collapseAll() {
  document.querySelectorAll('[id$="-bd"]').forEach(function(el){
    el.classList.add('bc');
    var id = el.id.replace('-bd','');
    colMap[id] = true;
  });
  document.querySelectorAll('.ti,.fi').forEach(function(ic){ ic.innerHTML = '&#x25B6;'; });
}

// ─────────────────────────────────────────────
//  SEARCH / FILTER
// ─────────────────────────────────────────────
function doSearch(term) {
  term = term.trim().toLowerCase();
  var clrBtn = document.getElementById('sClear');
  var status = document.getElementById('sStatus');

  if (clrBtn) { term ? clrBtn.classList.add('vis') : clrBtn.classList.remove('vis'); }

  if (!term) {
    // Clear: restore visibility
    document.querySelectorAll('.acl-row').forEach(function(r){ r.classList.remove('rh'); });
    document.querySelectorAll('.t-blk').forEach(function(b){ b.classList.remove('bh'); });
    document.querySelectorAll('.f-sec').forEach(function(s){ s.classList.remove('sh'); });
    document.querySelectorAll('.ns-card').forEach(function(c){ c.classList.remove('ch'); });
    // Restore pre-search collapsed state
    document.querySelectorAll('[id$="-bd"]').forEach(function(el){
      var id = el.id.replace('-bd','');
      if (colMap[id]) { el.classList.add('bc'); } else { el.classList.remove('bc'); }
    });
    document.querySelectorAll('.ti,.fi').forEach(function(ic){
      var bdId = ic.id ? ic.id.replace('-ti','') : '';
      ic.innerHTML = (colMap[bdId]) ? '&#x25B6;' : '&#x25BC;';
    });
    if (status) status.textContent = '';
    return;
  }

  // During search: expand everything first
  document.querySelectorAll('[id$="-bd"]').forEach(function(el){ el.classList.remove('bc'); });

  var total = 0, visible = 0;
  document.querySelectorAll('.acl-row').forEach(function(row){
    total++;
    var data = (row.getAttribute('data-s') || '');
    if (data.indexOf(term) !== -1) {
      row.classList.remove('rh'); visible++;
    } else {
      row.classList.add('rh');
    }
  });

  // Hide target blocks with no visible rows
  document.querySelectorAll('.t-blk').forEach(function(blk){
    var ok = blk.querySelector('.acl-row:not(.rh)');
    ok ? blk.classList.remove('bh') : blk.classList.add('bh');
  });

  // Hide folder sections with no visible target blocks
  document.querySelectorAll('.f-sec').forEach(function(sec){
    var ok = sec.querySelector('.t-blk:not(.bh)');
    ok ? sec.classList.remove('sh') : sec.classList.add('sh');
  });

  // Hide namespace cards with no visible folder sections
  document.querySelectorAll('.ns-card').forEach(function(card){
    var bd = card.querySelector('[id$="-bd"]');
    var ok = bd ? bd.querySelector('.f-sec:not(.sh)') : null;
    ok ? card.classList.remove('ch') : card.classList.add('ch');
  });

  if (status) {
    status.textContent = 'Showing ' + visible + ' of ' + total + ' ACL entries for: "' + term + '"';
  }
}

function clearSearch() {
  var box = document.getElementById('sBox');
  if (box) box.value = '';
  doSearch('');
}
</script>
</body>
</html>
"@)

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}

# ============================================================
#  REGION 6 - DATA COLLECTION
# ============================================================
Write-Host ''
Write-Log "Enumerating DFS Namespace roots on: $NamespaceServer"

$AllRecords       = [System.Collections.Generic.List[PSCustomObject]]::new()
$NamespaceSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
$CollectionErrors = [System.Collections.Generic.List[string]]::new()

# Get all DFS roots hosted on the target server
try {
    $DFSRoots = @(Get-DfsnRoot -ComputerName $NamespaceServer -ErrorAction Stop)
}
catch {
    Write-Log "Failed to enumerate DFS roots on '$NamespaceServer': $_" -Level ERROR
    exit 1
}

if ($DFSRoots.Count -eq 0) {
    Write-Log "No DFS namespace roots found on '$NamespaceServer'." -Level WARN
    exit 0
}

Write-Log "Found $($DFSRoots.Count) namespace root(s)." -Level SUCCESS

foreach ($root in $DFSRoots) {
    $rootPath  = $root.Path
    $rootType  = if ($null -ne $root.Type)        { $root.Type.ToString()        } else { 'Unknown' }
    $rootDesc  = if ($root.Description)            { $root.Description            } else { ''        }
    $rootState = if ($null -ne $root.State)        { $root.State.ToString()       } else { 'Unknown' }

    Write-Host ''
    Write-Host "  [NS] $rootPath" -ForegroundColor Yellow

    # Get root targets (server(s) hosting this namespace)
    $rsvList = [System.Collections.Generic.List[string]]::new()
    try {
        $rootTargets = @(Get-DfsnRootTarget -Path $rootPath -ErrorAction Stop)
        foreach ($rt in $rootTargets) {
            $rsvList.Add((Get-UNCServer $rt.TargetPath)) | Out-Null
        }
    }
    catch {
        Write-Log "  Root targets unavailable for ${rootPath}: $_" -Level WARN
        $CollectionErrors.Add("Root targets [$rootPath]: $_")
    }
    $rootSrvStr = ($rsvList -join '; ')

    # Get all DFS folders in this namespace
    $folders = @()
    try {
        $folders = @(Get-DfsnFolder -Path "$rootPath\*" -ErrorAction Stop)
    }
    catch {
        Write-Log "  Folders unavailable for ${rootPath}: $_" -Level WARN
        $CollectionErrors.Add("Folders [$rootPath]: $_")
    }

    Write-Host "       Folders : $($folders.Count)" -ForegroundColor DarkGray

    # Add namespace summary row
    $NamespaceSummary.Add([PSCustomObject]@{
        NamespacePath    = $rootPath
        NamespaceType    = $rootType
        Description      = $rootDesc
        State            = $rootState
        NamespaceServers = $rootSrvStr
        FolderCount      = $folders.Count
    })

    # Empty namespace placeholder
    if ($folders.Count -eq 0) {
        $AllRecords.Add([PSCustomObject]@{
            Namespace        = $rootPath
            NamespaceType    = $rootType
            NamespaceDesc    = $rootDesc
            NamespaceState   = $rootState
            NamespaceServers = $rootSrvStr
            FolderPath       = '(No folders)'
            FolderDesc       = ''
            FolderState      = ''
            TargetPath       = ''
            TargetServer     = ''
            TargetState      = ''
            TargetPriority   = ''
            SecurityIdentity = ''
            Domain           = ''
            SamAccount       = ''
            Rights           = ''
            AccessType       = ''
            IsInherited      = ''
            GroupType        = ''
        })
        continue
    }

    foreach ($folder in $folders) {
        $fPath  = $folder.Path
        $fDesc  = if ($folder.Description) { $folder.Description } else { '' }
        $fState = if ($null -ne $folder.State) { $folder.State.ToString() } else { 'Unknown' }

        Write-Host "       [FOLDER] $fPath" -ForegroundColor White

        # Get folder targets
        $fTargets = @()
        try {
            $fTargets = @(Get-DfsnFolderTarget -Path $fPath -ErrorAction Stop)
        }
        catch {
            Write-Log "  Targets unavailable for ${fPath}: $_" -Level WARN
            $CollectionErrors.Add("Targets [$fPath]: $_")
        }

        if ($fTargets.Count -eq 0) {
            $AllRecords.Add([PSCustomObject]@{
                Namespace        = $rootPath
                NamespaceType    = $rootType
                NamespaceDesc    = $rootDesc
                NamespaceState   = $rootState
                NamespaceServers = $rootSrvStr
                FolderPath       = $fPath
                FolderDesc       = $fDesc
                FolderState      = $fState
                TargetPath       = '(No targets)'
                TargetServer     = ''
                TargetState      = ''
                TargetPriority   = ''
                SecurityIdentity = ''
                Domain           = ''
                SamAccount       = ''
                Rights           = ''
                AccessType       = ''
                IsInherited      = ''
                GroupType        = ''
            })
            continue
        }

        foreach ($target in $fTargets) {
            $tPath   = $target.TargetPath
            $tState  = if ($null -ne $target.State)                 { $target.State.ToString()                 } else { 'Unknown' }
            $tPri    = if ($null -ne $target.ReferralPriorityClass) { $target.ReferralPriorityClass.ToString() } else { '' }
            $tSrv    = Get-UNCServer $tPath

            Write-Host "              [TARGET] $tPath  ($tState)" -ForegroundColor DarkGray

            # Collect NTFS ACL entries
            $aceList = Get-FolderACLEntries -Path $tPath

            foreach ($ace in $aceList) {
                $AllRecords.Add([PSCustomObject]@{
                    Namespace        = $rootPath
                    NamespaceType    = $rootType
                    NamespaceDesc    = $rootDesc
                    NamespaceState   = $rootState
                    NamespaceServers = $rootSrvStr
                    FolderPath       = $fPath
                    FolderDesc       = $fDesc
                    FolderState      = $fState
                    TargetPath       = $tPath
                    TargetServer     = $tSrv
                    TargetState      = $tState
                    TargetPriority   = $tPri
                    SecurityIdentity = $ace.Identity
                    Domain           = $ace.Domain
                    SamAccount       = $ace.SamAccount
                    Rights           = $ace.Rights
                    AccessType       = $ace.AccessType
                    IsInherited      = $ace.IsInherited.ToString()
                    GroupType        = $ace.GroupType
                })
            }
        }
    }
}

Write-Host ''
Write-Log "Collection complete. Total records: $($AllRecords.Count)" -Level SUCCESS

# ============================================================
#  REGION 7 - CSV EXPORT
# ============================================================
Write-Log 'Writing CSV...'
try {
    $AllRecords | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Log "CSV saved: $CsvFile" -Level SUCCESS
}
catch {
    Write-Log "CSV export failed: $_" -Level ERROR
    $CollectionErrors.Add("CSV: $_")
}

# ============================================================
#  REGION 8 - XLSX EXPORT
# ============================================================
Write-Log 'Writing XLSX (Open XML / ZipArchive)...'
try {
    Export-XlsxReport -FilePath $XlsxFile -SummaryRows $NamespaceSummary -DetailRows $AllRecords
    Write-Log "XLSX saved: $XlsxFile" -Level SUCCESS
}
catch {
    Write-Log "XLSX export failed: $_" -Level ERROR
    $CollectionErrors.Add("XLSX: $_")
}

# ============================================================
#  REGION 9 - HTML EXPORT
# ============================================================
Write-Log 'Writing HTML report...'
try {
    Build-HtmlReport `
        -OutputPath  $HtmlFile `
        -Records     $AllRecords `
        -Summary     $NamespaceSummary `
        -Server      $NamespaceServer `
        -RunUser     $RunningUser `
        -GeneratedAt $ReportDate `
        -Version     $ScriptVersion
    Write-Log "HTML saved: $HtmlFile" -Level SUCCESS
}
catch {
    Write-Log "HTML export failed: $_" -Level ERROR
    $CollectionErrors.Add("HTML: $_")
}

# ============================================================
#  REGION 10 - COMPLETION SUMMARY
# ============================================================
$EndTime  = Get-Date
$Elapsed  = [Math]::Round(($EndTime - $StartTime).TotalSeconds, 1)

Write-Host ''
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host '                       AUDIT COMPLETE                            ' -ForegroundColor Green
Write-Host '  ================================================================' -ForegroundColor Green
Write-Host "  Duration   : $Elapsed seconds"                                    -ForegroundColor White
Write-Host "  Namespaces : $($NamespaceSummary.Count)"                          -ForegroundColor White
Write-Host "  Records    : $($AllRecords.Count) total ACL entries"              -ForegroundColor White

$errColor = if ($CollectionErrors.Count -gt 0) { 'Yellow' } else { 'White' }
Write-Host "  Errors     : $($CollectionErrors.Count)" -ForegroundColor $errColor
Write-Host ''
Write-Host '  Output files:' -ForegroundColor Cyan
Write-Host "    LOG  -> $LogFile"  -ForegroundColor DarkGray
Write-Host "    CSV  -> $CsvFile"  -ForegroundColor White
Write-Host "    XLSX -> $XlsxFile" -ForegroundColor White
Write-Host "    HTML -> $HtmlFile" -ForegroundColor White

if ($CollectionErrors.Count -gt 0) {
    Write-Host ''
    Write-Host '  Errors encountered (non-fatal):' -ForegroundColor Yellow
    foreach ($e in $CollectionErrors) { Write-Host "    - $e" -ForegroundColor Yellow }
}

Write-Host '  ================================================================' -ForegroundColor Green
Write-Log "Script completed. Elapsed: ${Elapsed}s | Records: $($AllRecords.Count) | Errors: $($CollectionErrors.Count)"

# Open export folder in Explorer
try { Start-Process -FilePath 'explorer.exe' -ArgumentList $ExportPath } catch { }

Write-Host ''
