#Requires -Version 5.1
<#
.SYNOPSIS - Claude output version
    Enterprise Active Directory user security export (DUMPSEC-style) with CSV, XLSX, and HTML deliverables.

.DESCRIPTION
    Queries all AD user accounts (single domain, a specified list of domains, or every domain in the current
    forest) and produces a security-focused export covering identity, logon/password state, extension
    attributes, and Tier 0/1/2 classification. Output is written to a timestamped DUMPSEC folder on the
    Desktop of the account running the script:

        <Desktop>\DUMPSEC_yyyy-MM-dd_HH-mm-ss\

    Three report artifacts are generated in that folder:
        - AD_User_Security_Export_<stamp>.csv   (raw data, Excel/Power BI friendly)
        - AD_User_Security_Export_<stamp>.xlsx  (native Excel workbook, built via raw Open XML - no
                                                   ImportExcel/Excel COM dependency required)
        - AD_User_Security_Export_<stamp>.html  (interactive dark-themed dashboard: KPI cards, search,
                                                   sortable columns, per-domain breakdown)

    Run this ON a Domain Controller (or a management host with RSAT AD tools) using an account with rights
    to read user objects and enumerate privileged group membership across the domain(s) in scope.

.PARAMETER Domains
    One or more domain DNS names to query (e.g. -Domains 'ad.igt.com','igtsap.ad.igt.com'). If omitted, the
    script auto-discovers every domain in the current forest via Get-ADForest and queries all of them.

.PARAMETER OutputRoot
    Root folder under which the DUMPSEC_<timestamp> folder is created. Defaults to the current user's Desktop.

.PARAMETER InactiveThresholdDays
    Number of days without logon before an enabled account is flagged "Inactive" in the KPI summary and
    highlighted in the HTML report. Default 90.

.PARAMETER PasswordExpiringSoonDays
    Number of days remaining before password expiry that triggers the "Expiring Soon" password status.
    Default 14.

.PARAMETER PrivilegedGroups
    Override the list of groups used (recursively) to flag Tier 0 membership when an account is not already
    tiered by OU path or naming convention. Defaults to the standard AD built-in privileged groups.

.EXAMPLE
    .\Get-ADSecurityExport.ps1
    Auto-discovers every domain in the forest and exports to the local Desktop.

.EXAMPLE
    .\Get-ADSecurityExport.ps1 -Domains 'ad.igt.com','igtsap.ad.igt.com' -InactiveThresholdDays 60

.NOTES
    Author:  SysAdmin (Steve McKee)
    Version: 1.0
    Requires: RSAT Active Directory PowerShell module (ActiveDirectory)
    PowerShell 5.1 compatible - no ?., ??, ??= operators, no external modules, ASCII-only log output.

    KNOWN DATA LIMITATION - "Last Logon Time":
    True per-DC LastLogon is not replicated between domain controllers, so an accurate value requires
    querying every DC in the domain and taking the maximum - expensive at scale. This report uses
    LastLogonDate (sourced from the replicated lastLogonTimestamp attribute), which is accurate to within
    the domain's LastLogonTimestamp sync window (14 days by default in most forests). Both "Last Logon
    Time" and "Last Logon Time Stamp" columns are populated from this same replicated value. If you need
    exact last-logon-anywhere data, that is a separate, heavier multi-DC pass and is not included here.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$Domains,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = [Environment]::GetFolderPath('Desktop'),

    [Parameter(Mandatory = $false)]
    [int]$InactiveThresholdDays = 90,

    [Parameter(Mandatory = $false)]
    [int]$PasswordExpiringSoonDays = 14,

    [Parameter(Mandatory = $false)]
    [string[]]$PrivilegedGroups = @(
        'Domain Admins','Enterprise Admins','Schema Admins','Administrators',
        'Account Operators','Backup Operators','Server Operators',
        'Group Policy Creator Owners','DnsAdmins'
    )
)

$ErrorActionPreference = 'Stop'
$ScriptStartTime = Get-Date

# ============================================================================
# REGION: Setup - output folder, logging
# ============================================================================

$FolderStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$FolderName  = "DUMPSEC_$FolderStamp"
$OutputFolder = Join-Path -Path $OutputRoot -ChildPath $FolderName

try {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}
catch {
    Write-Host "FATAL: Could not create output folder '$OutputFolder': $($_.Exception.Message)" -ForegroundColor Red
    return
}

$LogPath = Join-Path -Path $OutputFolder -ChildPath "DUMPSEC_Log_$FolderStamp.log"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding ASCII
    switch ($Level) {
        'INFO'  { Write-Host $entry -ForegroundColor Cyan }
        'WARN'  { Write-Host $entry -ForegroundColor Yellow }
        'ERROR' { Write-Host $entry -ForegroundColor Red }
    }
}

Write-Log "DUMPSEC AD Security Export started by $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"
Write-Log "Output folder: $OutputFolder"

# ============================================================================
# REGION: Module / prerequisite checks
# ============================================================================

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ActiveDirectory PowerShell module not found. Install RSAT: Active Directory Domain Services and LDS Tools." 'ERROR'
    return
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Log "Failed to import ActiveDirectory module: $($_.Exception.Message)" 'ERROR'
    return
}

# ============================================================================
# REGION: Domain discovery
# ============================================================================

if (-not $Domains -or $Domains.Count -eq 0) {
    Write-Log "No -Domains specified, auto-discovering all domains in the current forest..."
    try {
        $forest = Get-ADForest
        $Domains = $forest.Domains
        Write-Log "Forest '$($forest.Name)' contains $($Domains.Count) domain(s): $($Domains -join ', ')"
    }
    catch {
        Write-Log "Get-ADForest failed ($($_.Exception.Message)); falling back to current domain only." 'WARN'
        try {
            $Domains = @((Get-ADDomain).DNSRoot)
        }
        catch {
            Write-Log "Could not determine current domain either: $($_.Exception.Message)" 'ERROR'
            return
        }
    }
}
else {
    Write-Log "Domains specified: $($Domains -join ', ')"
}

# ============================================================================
# REGION: Helper functions
# ============================================================================

function Get-ExcelColumnLetter {
    param([Parameter(Mandatory = $true)][int]$Index)
    $dividend = $Index
    $columnName = ''
    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $columnName = [char](65 + $modulo) + $columnName
        $dividend = [int][math]::Floor(($dividend - $modulo) / 26)
    }
    return $columnName
}

function Escape-XmlText {
    param($Text)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    $s = $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
    # Strip control characters that are illegal in XML 1.0
    $s = ($s.ToCharArray() | Where-Object { ([int]$_) -ge 32 -or [int]$_ -eq 9 -or [int]$_ -eq 10 -or [int]$_ -eq 13 }) -join ''
    return $s
}

function Escape-JsonText {
    param($Text)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", '').Replace("`n", ' ').Replace("`t", ' ')
    return $s
}

function Get-PrivilegedMemberSidSet {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string[]]$Groups
    )
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($g in $Groups) {
        try {
            $members = Get-ADGroupMember -Identity $g -Server $Server -Recursive -ErrorAction Stop
            foreach ($m in $members) {
                if ($m.objectClass -eq 'user' -and $m.SID) {
                    [void]$set.Add($m.SID.Value)
                }
            }
        }
        catch {
            Write-Log "  Could not enumerate privileged group '$g' on $Server (may not exist / access denied): $($_.Exception.Message)" 'WARN'
        }
    }
    return $set
}

function Get-AccountTier {
    param(
        [string]$DistinguishedName,
        [string]$SamAccountName,
        [System.Collections.Generic.HashSet[string]]$PrivilegedSidSet,
        [string]$Sid
    )

    if ($DistinguishedName -match '(?i)OU=Tier ?0|OU=T0(,|$)') { return 'Tier 0' }
    if ($DistinguishedName -match '(?i)OU=Tier ?1|OU=T1(,|$)') { return 'Tier 1' }
    if ($DistinguishedName -match '(?i)OU=Tier ?2|OU=T2(,|$)') { return 'Tier 2' }

    if ($SamAccountName -match '(?i)^adm[._-]?t0|^t0[._-]?adm') { return 'Tier 0' }
    if ($SamAccountName -match '(?i)^adm[._-]?t1|^t1[._-]?adm') { return 'Tier 1' }
    if ($SamAccountName -match '(?i)^adm[._-]?t2|^t2[._-]?adm') { return 'Tier 2' }

    if ($PrivilegedSidSet -and $Sid -and $PrivilegedSidSet.Contains($Sid)) {
        return 'Tier 0 (Privileged Group)'
    }

    return 'Standard / Untiered'
}

function New-SimpleXlsx {
    <#
        Builds a minimal single-sheet .xlsx workbook from an array of PSCustomObjects using raw Open XML
        (System.IO.Compression.ZipArchive). No ImportExcel module or Excel COM interop required.
        All cell values are written as shared strings for simplicity/robustness; the header row is bold
        with a fill color.
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Data,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$SheetName = 'AD Users'
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path $Path) { Remove-Item $Path -Force }
    if (-not $Data -or $Data.Count -eq 0) {
        Write-Log "New-SimpleXlsx: no data rows supplied, skipping xlsx generation." 'WARN'
        return
    }

    $columns = $Data[0].PSObject.Properties.Name

    # ---- Shared strings table ----
    $sharedStrings = New-Object System.Collections.Generic.List[string]
    $sharedStringIndex = New-Object 'System.Collections.Generic.Dictionary[string,int]'

    function Get-SSIndex {
        param($Value)
        $s = if ($null -eq $Value) { '' } else { [string]$Value }
        if ($sharedStringIndex.ContainsKey($s)) { return $sharedStringIndex[$s] }
        $idx = $sharedStrings.Count
        $sharedStrings.Add($s)
        $sharedStringIndex[$s] = $idx
        return $idx
    }

    $sheetXmlBuilder = New-Object System.Text.StringBuilder
    [void]$sheetXmlBuilder.Append('<sheetData>')

    # Header row (style index 1 = bold header style, defined in styles.xml below)
    [void]$sheetXmlBuilder.Append('<row r="1">')
    for ($c = 0; $c -lt $columns.Count; $c++) {
        $colLetter = Get-ExcelColumnLetter -Index ($c + 1)
        $idx = Get-SSIndex -Value $columns[$c]
        [void]$sheetXmlBuilder.Append('<c r="' + $colLetter + '1" t="s" s="1"><v>' + $idx + '</v></c>')
    }
    [void]$sheetXmlBuilder.Append('</row>')

    $rowNum = 2
    foreach ($row in $Data) {
        [void]$sheetXmlBuilder.Append('<row r="' + $rowNum + '">')
        for ($c = 0; $c -lt $columns.Count; $c++) {
            $colLetter = Get-ExcelColumnLetter -Index ($c + 1)
            $val = $row.($columns[$c])
            $idx = Get-SSIndex -Value $val
            [void]$sheetXmlBuilder.Append('<c r="' + $colLetter + $rowNum + '" t="s"><v>' + $idx + '</v></c>')
        }
        [void]$sheetXmlBuilder.Append('</row>')
        $rowNum++
    }
    [void]$sheetXmlBuilder.Append('</sheetData>')

    $lastCol = Get-ExcelColumnLetter -Index $columns.Count
    $dimension = "A1:$lastCol$($rowNum - 1)"

    # Freeze header row + basic column widths
    $colsXml = New-Object System.Text.StringBuilder
    [void]$colsXml.Append('<cols>')
    for ($c = 0; $c -lt $columns.Count; $c++) {
        [void]$colsXml.Append('<col min="' + ($c + 1) + '" max="' + ($c + 1) + '" width="22" customWidth="1"/>')
    }
    [void]$colsXml.Append('</cols>')

    $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<dimension ref="' + $dimension + '"/>' +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        $colsXml.ToString() +
        $sheetXmlBuilder.ToString() +
        '</worksheet>'

    $sharedStringsXml = New-Object System.Text.StringBuilder
    [void]$sharedStringsXml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sharedStringsXml.Append('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $sharedStrings.Count + '" uniqueCount="' + $sharedStrings.Count + '">')
    foreach ($s in $sharedStrings) {
        [void]$sharedStringsXml.Append('<si><t xml:space="preserve">' + (Escape-XmlText $s) + '</t></si>')
    }
    [void]$sharedStringsXml.Append('</sst>')

    $contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '</Types>'

    $rootRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'

    $workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + (Escape-XmlText $SheetName) + '" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>'

    $workbookRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
        '</Relationships>'

    # Style 0 = default, Style 1 = bold white text on dark mauve fill (header)
    $stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2">' +
        '<font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><b/></font>' +
        '</fonts>' +
        '<fills count="3">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF313244"/><bgColor indexed="64"/></patternFill></fill>' +
        '</fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
        '</cellXfs>' +
        '</styleSheet>'

    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

        function Add-ZipEntry {
            param($Archive, [string]$EntryName, [string]$Content)
            $entry = $Archive.CreateEntry($EntryName)
            $stream = $entry.Open()
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
            $writer.Write($Content)
            $writer.Flush()
            $writer.Close()
            $stream.Close()
        }

        Add-ZipEntry -Archive $zip -EntryName '[Content_Types].xml' -Content $contentTypesXml
        Add-ZipEntry -Archive $zip -EntryName '_rels/.rels' -Content $rootRelsXml
        Add-ZipEntry -Archive $zip -EntryName 'xl/workbook.xml' -Content $workbookXml
        Add-ZipEntry -Archive $zip -EntryName 'xl/_rels/workbook.xml.rels' -Content $workbookRelsXml
        Add-ZipEntry -Archive $zip -EntryName 'xl/styles.xml' -Content $stylesXml
        Add-ZipEntry -Archive $zip -EntryName 'xl/sharedStrings.xml' -Content $sharedStringsXml.ToString()
        Add-ZipEntry -Archive $zip -EntryName 'xl/worksheets/sheet1.xml' -Content $sheetXml

        $zip.Dispose()
        $fs.Dispose()
        Write-Log "XLSX workbook written: $Path"
    }
    catch {
        Write-Log "New-SimpleXlsx failed: $($_.Exception.Message)" 'ERROR'
        if ($zip) { $zip.Dispose() }
        if ($fs) { $fs.Dispose() }
    }
}

# ============================================================================
# REGION: AD query - property list
# ============================================================================

$AdProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'SamAccountName', 'Name', 'Description',
    'DistinguishedName', 'EmailAddress', 'mailNickname', 'UserPrincipalName',
    'AccountExpirationDate', 'Enabled', 'LockedOut', 'PasswordNeverExpires', 'PasswordExpired',
    'PasswordLastSet', 'LastLogonDate', 'whenChanged', 'whenCreated', 'userWorkstations',
    'extensionAttribute8', 'extensionAttribute9', 'extensionAttribute10',
    'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15',
    'msDS-UserPasswordExpiryTimeComputed', 'ObjectSID'
)

$AllUsers = New-Object System.Collections.Generic.List[object]
$DomainSummaries = New-Object System.Collections.Generic.List[object]
$Now = Get-Date

foreach ($Domain in $Domains) {
    Write-Log "----------------------------------------------------------------"
    Write-Log "Processing domain: $Domain"

    try {
        $null = Get-ADDomain -Server $Domain -ErrorAction Stop
    }
    catch {
        Write-Log "  Could not contact domain '$Domain': $($_.Exception.Message)" 'ERROR'
        continue
    }

    Write-Log "  Enumerating privileged group membership for Tier 0 detection..."
    $PrivSet = Get-PrivilegedMemberSidSet -Server $Domain -Groups $PrivilegedGroups
    Write-Log "  Privileged accounts found: $($PrivSet.Count)"

    Write-Log "  Querying all user objects (this may take a while on large domains)..."
    try {
        $Users = Get-ADUser -Filter * -Server $Domain -Properties $AdProperties -ErrorAction Stop
    }
    catch {
        Write-Log "  Get-ADUser failed for domain '$Domain': $($_.Exception.Message)" 'ERROR'
        continue
    }

    Write-Log "  Retrieved $($Users.Count) user object(s) from $Domain"

    $domainUserCount = 0
    foreach ($u in $Users) {

        $daysSinceLogon = $null
        if ($u.LastLogonDate) { $daysSinceLogon = [math]::Round(($Now - $u.LastLogonDate).TotalDays) }

        $daysSincePwd = $null
        if ($u.PasswordLastSet) { $daysSincePwd = [math]::Round(($Now - $u.PasswordLastSet).TotalDays) }

        $pwdExpiryDate = $null
        $rawExpiry = $u.'msDS-UserPasswordExpiryTimeComputed'
        if ($rawExpiry -and $rawExpiry -gt 0 -and $rawExpiry -lt [long]::MaxValue) {
            try { $pwdExpiryDate = [DateTime]::FromFileTime($rawExpiry) } catch { $pwdExpiryDate = $null }
        }

        $pwdExpiresInDays = $null
        if ($pwdExpiryDate -and -not $u.PasswordNeverExpires) {
            $pwdExpiresInDays = [math]::Round(($pwdExpiryDate - $Now).TotalDays)
        }

        $pwdStatus = 'Active'
        if ($u.PasswordNeverExpires) { $pwdStatus = 'Never Expires' }
        elseif ($u.PasswordExpired) { $pwdStatus = 'Expired' }
        elseif ($null -ne $pwdExpiresInDays -and $pwdExpiresInDays -le $PasswordExpiringSoonDays) { $pwdStatus = 'Expiring Soon' }
        elseif (-not $u.PasswordLastSet) { $pwdStatus = 'Unknown' }

        $sidValue = $null
        if ($u.ObjectSID) { $sidValue = $u.ObjectSID.Value }

        $tier = Get-AccountTier -DistinguishedName $u.DistinguishedName -SamAccountName $u.SamAccountName -PrivilegedSidSet $PrivSet -Sid $sidValue

        $cn = $u.Name
        if ($u.DistinguishedName -match '^CN=([^,]+),') { $cn = $matches[1] }

        $parentOU = ''
        if ($u.DistinguishedName -match '^CN=[^,]+,(.+)$') { $parentOU = $matches[1] }

        [void]$AllUsers.Add([PSCustomObject]@{
            'Display Name'                  = $u.DisplayName
            'Common Name'                   = $cn
            'SAM Account Name'              = $u.SamAccountName
            'Domain Name'                   = $Domain
            'First Name'                    = $u.GivenName
            'Last Name'                     = $u.Surname
            'Full Name'                     = $u.Name
            'Name'                          = $u.Name
            'Description'                   = $u.Description
            'Distinguished Name'            = $u.DistinguishedName
            'Logon To OU'                   = $parentOU
            'Logon Workstations'            = $u.userWorkstations
            'E-mail'                        = $u.EmailAddress
            'Alias'                         = $u.mailNickname
            'UPN'                           = $u.UserPrincipalName
            'Ext8'                          = $u.extensionAttribute8
            'Ext9'                          = $u.extensionAttribute9
            'Ext10'                         = $u.extensionAttribute10
            'Ext13'                         = $u.extensionAttribute13
            'Ext14'                         = $u.extensionAttribute14
            'Ext15'                         = $u.extensionAttribute15
            'Account Status'                = if ($u.Enabled) { 'Enabled' } else { 'Disabled' }
            'Account Expiry Time'           = $u.AccountExpirationDate
            'Account Locked'                = $u.LockedOut
            'Last Logon Time'               = $u.LastLogonDate
            'Last Logon Time Stamp'         = $u.LastLogonDate
            'Days Since Last Logon'         = $daysSinceLogon
            'Password Last Set'             = $u.PasswordLastSet
            'Days Since Password Last Set'  = $daysSincePwd
            'Pwd Never Expires'             = $u.PasswordNeverExpires
            'Password Expiry Date'          = $pwdExpiryDate
            'Password Expires In (Days)'    = $pwdExpiresInDays
            'Password Status'               = $pwdStatus
            'When Created'                  = $u.whenCreated
            'When Changed'                  = $u.whenChanged
            'Security Tier'                 = $tier
        })
        $domainUserCount++
    }

    [void]$DomainSummaries.Add([PSCustomObject]@{
        Domain    = $Domain
        UserCount = $domainUserCount
        Privileged = $PrivSet.Count
    })

    Write-Log "  Finished domain $Domain ($domainUserCount users processed)"
}

Write-Log "----------------------------------------------------------------"
Write-Log "Total users collected across all domains: $($AllUsers.Count)"

if ($AllUsers.Count -eq 0) {
    Write-Log "No user data collected - aborting before export." 'ERROR'
    return
}

# ============================================================================
# REGION: KPI summary calculations (used by HTML dashboard)
# ============================================================================

$Kpi = [PSCustomObject]@{
    TotalUsers          = $AllUsers.Count
    Enabled             = ($AllUsers | Where-Object { $_.'Account Status' -eq 'Enabled' }).Count
    Disabled            = ($AllUsers | Where-Object { $_.'Account Status' -eq 'Disabled' }).Count
    Locked              = ($AllUsers | Where-Object { $_.'Account Locked' -eq $true }).Count
    PwdNeverExpires     = ($AllUsers | Where-Object { $_.'Pwd Never Expires' -eq $true }).Count
    PwdExpired          = ($AllUsers | Where-Object { $_.'Password Status' -eq 'Expired' }).Count
    PwdExpiringSoon     = ($AllUsers | Where-Object { $_.'Password Status' -eq 'Expiring Soon' }).Count
    InactiveEnabled     = ($AllUsers | Where-Object { $_.'Account Status' -eq 'Enabled' -and $_.'Days Since Last Logon' -ne $null -and $_.'Days Since Last Logon' -ge $InactiveThresholdDays }).Count
    NeverLoggedOn       = ($AllUsers | Where-Object { -not $_.'Last Logon Time' }).Count
    Tier0               = ($AllUsers | Where-Object { $_.'Security Tier' -like 'Tier 0*' }).Count
    Tier1               = ($AllUsers | Where-Object { $_.'Security Tier' -eq 'Tier 1' }).Count
    Tier2               = ($AllUsers | Where-Object { $_.'Security Tier' -eq 'Tier 2' }).Count
    DomainCount         = $Domains.Count
}

Write-Log "KPI - Total: $($Kpi.TotalUsers) | Enabled: $($Kpi.Enabled) | Disabled: $($Kpi.Disabled) | Locked: $($Kpi.Locked)"
Write-Log "KPI - Pwd Never Expires: $($Kpi.PwdNeverExpires) | Expired: $($Kpi.PwdExpired) | Expiring Soon: $($Kpi.PwdExpiringSoon)"
Write-Log "KPI - Inactive (>= $InactiveThresholdDays days, enabled): $($Kpi.InactiveEnabled) | Never Logged On: $($Kpi.NeverLoggedOn)"
Write-Log "KPI - Tier 0: $($Kpi.Tier0) | Tier 1: $($Kpi.Tier1) | Tier 2: $($Kpi.Tier2)"

# ============================================================================
# REGION: CSV export
# ============================================================================

$CsvPath = Join-Path -Path $OutputFolder -ChildPath "AD_User_Security_Export_$FolderStamp.csv"
try {
    $AllUsers | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV export written: $CsvPath"
}
catch {
    Write-Log "CSV export failed: $($_.Exception.Message)" 'ERROR'
}

# ============================================================================
# REGION: XLSX export
# ============================================================================

$XlsxPath = Join-Path -Path $OutputFolder -ChildPath "AD_User_Security_Export_$FolderStamp.xlsx"
try {
    New-SimpleXlsx -Data $AllUsers.ToArray() -Path $XlsxPath -SheetName 'AD Users'
}
catch {
    Write-Log "XLSX export failed: $($_.Exception.Message)" 'ERROR'
}

# ============================================================================
# REGION: HTML dashboard export (Catppuccin Mocha dark theme)
# ============================================================================

$HtmlPath = Join-Path -Path $OutputFolder -ChildPath "AD_User_Security_Export_$FolderStamp.html"

Write-Log "Building HTML dashboard..."

# Columns shown in the interactive table (kept to the most operationally relevant set;
# full attribute list is always available in the CSV/XLSX)
$HtmlColumns = @(
    'Display Name', 'SAM Account Name', 'Domain Name', 'Account Status', 'Security Tier',
    'Account Locked', 'Days Since Last Logon', 'Password Status', 'Days Since Password Last Set',
    'Pwd Never Expires', 'Account Expiry Time', 'Description', 'Distinguished Name'
)

# Build JSON rows (hand-rolled, PS5.1-safe - ConvertTo-Json is also fine but this keeps output compact)
$jsonRows = New-Object System.Text.StringBuilder
[void]$jsonRows.Append('[')
$first = $true
foreach ($row in $AllUsers) {
    if (-not $first) { [void]$jsonRows.Append(',') }
    $first = $false
    [void]$jsonRows.Append('{')
    $fFirst = $true
    foreach ($col in $HtmlColumns) {
        if (-not $fFirst) { [void]$jsonRows.Append(',') }
        $fFirst = $false
        $val = $row.$col
        $jsonVal = if ($null -eq $val) { '' } elseif ($val -is [bool]) { if ($val) {'Yes'} else {'No'} } else { Escape-JsonText $val }
        $keySafe = $col.Replace(' ', '_').Replace('-', '_')
        [void]$jsonRows.Append('"' + $keySafe + '":"' + $jsonVal + '"')
    }
    [void]$jsonRows.Append('}')
}
[void]$jsonRows.Append(']')

$jsonColumnKeys = ($HtmlColumns | ForEach-Object { $_.Replace(' ', '_').Replace('-', '_') })
$columnKeysJs = '[' + (($jsonColumnKeys | ForEach-Object { '"' + $_ + '"' }) -join ',') + ']'
$columnLabelsJs = '[' + (($HtmlColumns | ForEach-Object { '"' + (Escape-JsonText $_) + '"' }) -join ',') + ']'

$domainSummaryRowsHtml = New-Object System.Text.StringBuilder
foreach ($ds in $DomainSummaries) {
    [void]$domainSummaryRowsHtml.Append("<tr><td>$(Escape-XmlText $ds.Domain)</td><td>$($ds.UserCount)</td><td>$($ds.Privileged)</td></tr>")
}

$generatedOn = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$runBy = "$env:USERDOMAIN\$env:USERNAME"
$runFrom = $env:COMPUTERNAME

$htmlHead = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>AD User Security Export - DUMPSEC</title>
<style>
  :root {
    --base: #1e1e2e; --mantle: #181825; --crust: #11111b;
    --surface0: #313244; --surface1: #45475a; --surface2: #585b70;
    --text: #cdd6f4; --subtext1: #bac2de; --subtext0: #a6adc8;
    --blue: #89b4fa; --lavender: #b4befe; --mauve: #cba6f7;
    --red: #f38ba8; --maroon: #eba0ac; --peach: #fab387;
    --yellow: #f9e2af; --green: #a6e3a1; --teal: #94e2d5; --sky: #89dceb;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; background: var(--base); color: var(--text);
    font-family: 'Segoe UI', Consolas, Roboto, Arial, sans-serif; font-size: 14px;
  }
  header {
    background: var(--mantle); padding: 20px 28px; border-bottom: 1px solid var(--surface0);
  }
  header h1 { margin: 0 0 4px 0; font-size: 22px; color: var(--mauve); }
  header .meta { color: var(--subtext0); font-size: 12.5px; }
  .container { padding: 24px 28px; }
  .kpi-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px; margin-bottom: 24px;
  }
  .kpi-card {
    background: var(--mantle); border: 1px solid var(--surface0); border-radius: 10px;
    padding: 14px 16px;
  }
  .kpi-card .val { font-size: 26px; font-weight: 700; }
  .kpi-card .lbl { font-size: 11.5px; color: var(--subtext0); text-transform: uppercase; letter-spacing: .04em; margin-top: 2px; }
  .kpi-blue .val { color: var(--blue); }
  .kpi-green .val { color: var(--green); }
  .kpi-red .val { color: var(--red); }
  .kpi-yellow .val { color: var(--yellow); }
  .kpi-peach .val { color: var(--peach); }
  .kpi-mauve .val { color: var(--mauve); }
  .kpi-teal .val { color: var(--teal); }
  section { margin-bottom: 28px; }
  h2 { color: var(--lavender); font-size: 16px; border-bottom: 1px solid var(--surface0); padding-bottom: 8px; }
  table { width: 100%; border-collapse: collapse; background: var(--mantle); border-radius: 8px; overflow: hidden; }
  th, td { padding: 8px 10px; text-align: left; border-bottom: 1px solid var(--surface0); white-space: nowrap; }
  th { background: var(--surface0); color: var(--subtext1); font-size: 12px; text-transform: uppercase; letter-spacing: .03em; cursor: pointer; position: sticky; top: 0; }
  th:hover { color: var(--sky); }
  tr:hover td { background: var(--surface0); }
  .summary-table td, .summary-table th { white-space: normal; }
  .controls { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  input#searchBox {
    background: var(--mantle); border: 1px solid var(--surface1); color: var(--text);
    padding: 9px 14px; border-radius: 8px; width: 320px; font-size: 13px;
  }
  input#searchBox:focus { outline: none; border-color: var(--mauve); }
  .rowcount { color: var(--subtext0); font-size: 12.5px; }
  .badge { padding: 2px 9px; border-radius: 999px; font-size: 11.5px; font-weight: 600; display: inline-block; }
  .badge-green { background: rgba(166,227,161,0.15); color: var(--green); }
  .badge-red { background: rgba(243,139,168,0.15); color: var(--red); }
  .badge-yellow { background: rgba(249,226,175,0.15); color: var(--yellow); }
  .badge-peach { background: rgba(250,179,135,0.15); color: var(--peach); }
  .badge-mauve { background: rgba(203,166,247,0.15); color: var(--mauve); }
  .badge-grey { background: rgba(166,173,200,0.15); color: var(--subtext0); }
  .table-wrap { max-height: 640px; overflow: auto; border: 1px solid var(--surface0); border-radius: 8px; }
  footer { color: var(--subtext0); font-size: 11.5px; padding: 10px 28px 30px; }
  .tier0-row td { background: rgba(243,139,168,0.06); }
</style>
</head>
'@

$htmlBodyTop = @"
<body>
<header>
  <h1>Active Directory User Security Export</h1>
  <div class="meta">Generated $generatedOn by $runBy on $runFrom &nbsp;|&nbsp; Domains: $(($Domains -join ', '))</div>
</header>
<div class="container">

<section>
  <div class="kpi-grid">
    <div class="kpi-card kpi-blue"><div class="val">$($Kpi.TotalUsers)</div><div class="lbl">Total Users</div></div>
    <div class="kpi-card kpi-green"><div class="val">$($Kpi.Enabled)</div><div class="lbl">Enabled</div></div>
    <div class="kpi-card kpi-red"><div class="val">$($Kpi.Disabled)</div><div class="lbl">Disabled</div></div>
    <div class="kpi-card kpi-red"><div class="val">$($Kpi.Locked)</div><div class="lbl">Locked Out</div></div>
    <div class="kpi-card kpi-yellow"><div class="val">$($Kpi.PwdNeverExpires)</div><div class="lbl">Pwd Never Expires</div></div>
    <div class="kpi-card kpi-red"><div class="val">$($Kpi.PwdExpired)</div><div class="lbl">Pwd Expired</div></div>
    <div class="kpi-card kpi-peach"><div class="val">$($Kpi.PwdExpiringSoon)</div><div class="lbl">Pwd Expiring Soon</div></div>
    <div class="kpi-card kpi-peach"><div class="val">$($Kpi.InactiveEnabled)</div><div class="lbl">Inactive >= $InactiveThresholdDays d</div></div>
    <div class="kpi-card kpi-grey"><div class="val">$($Kpi.NeverLoggedOn)</div><div class="lbl">Never Logged On</div></div>
    <div class="kpi-card kpi-mauve"><div class="val">$($Kpi.Tier0)</div><div class="lbl">Tier 0 Accounts</div></div>
    <div class="kpi-card kpi-mauve"><div class="val">$($Kpi.Tier1)</div><div class="lbl">Tier 1 Accounts</div></div>
    <div class="kpi-card kpi-mauve"><div class="val">$($Kpi.Tier2)</div><div class="lbl">Tier 2 Accounts</div></div>
  </div>
</section>

<section>
  <h2>Domains Queried</h2>
  <table class="summary-table">
    <tr><th>Domain</th><th>Users Found</th><th>Privileged Accounts (recursive)</th></tr>
    $($domainSummaryRowsHtml.ToString())
  </table>
</section>

<section>
  <h2>User Detail</h2>
  <div class="controls">
    <input type="text" id="searchBox" placeholder="Search name, SAM, domain, description...">
    <span class="rowcount" id="rowCount"></span>
  </div>
  <div class="table-wrap">
    <table id="userTable">
      <thead><tr id="tableHeadRow"></tr></thead>
      <tbody id="tableBody"></tbody>
    </table>
  </div>
</section>

</div>
<footer>
  DUMPSEC-style AD security export. "Last Logon Time" / "Last Logon Time Stamp" are sourced from the replicated
  lastLogonTimestamp attribute (accurate to within the domain's replication window, typically up to 14 days) -
  not a live per-DC value. Security Tier is inferred from OU path, account naming convention, and privileged
  group membership; verify against your authoritative tiering documentation before acting on it. Full attribute
  set (including extension attributes 8/9/10/13/14/15) is available in the accompanying CSV and XLSX files.
</footer>
"@

$htmlScript = @"
<script>
const columnKeys = $columnKeysJs;
const columnLabels = $columnLabelsJs;
const rows = $($jsonRows.ToString());

function badge(col, val) {
  if (col === 'Account_Status') {
    return val === 'Enabled' ? '<span class="badge badge-green">Enabled</span>' : '<span class="badge badge-red">Disabled</span>';
  }
  if (col === 'Account_Locked') {
    return val === 'Yes' ? '<span class="badge badge-red">Locked</span>' : '<span class="badge badge-grey">No</span>';
  }
  if (col === 'Password_Status') {
    if (val === 'Expired') return '<span class="badge badge-red">Expired</span>';
    if (val === 'Expiring_Soon' || val === 'Expiring Soon') return '<span class="badge badge-peach">Expiring Soon</span>';
    if (val === 'Never Expires') return '<span class="badge badge-yellow">Never Expires</span>';
    if (val === 'Unknown') return '<span class="badge badge-grey">Unknown</span>';
    return '<span class="badge badge-green">Active</span>';
  }
  if (col === 'Security_Tier') {
    if (val && val.indexOf('Tier 0') === 0) return '<span class="badge badge-mauve">' + val + '</span>';
    if (val === 'Tier 1') return '<span class="badge badge-peach">Tier 1</span>';
    if (val === 'Tier 2') return '<span class="badge badge-yellow">Tier 2</span>';
    return '<span class="badge badge-grey">' + (val || 'Untiered') + '</span>';
  }
  if (col === 'Pwd_Never_Expires') {
    return val === 'Yes' ? '<span class="badge badge-yellow">Yes</span>' : 'No';
  }
  return val === undefined || val === null || val === '' ? '<span style="color:var(--subtext0)">-</span>' : val;
}

const theadRow = document.getElementById('tableHeadRow');
columnLabels.forEach((label, i) => {
  const th = document.createElement('th');
  th.textContent = label;
  th.dataset.col = columnKeys[i];
  th.addEventListener('click', () => sortByColumn(columnKeys[i]));
  theadRow.appendChild(th);
});

let sortState = { col: null, dir: 1 };

function sortByColumn(col) {
  if (sortState.col === col) { sortState.dir *= -1; } else { sortState = { col: col, dir: 1 }; }
  rows.sort((a, b) => {
    const av = (a[col] || '').toString().toLowerCase();
    const bv = (b[col] || '').toString().toLowerCase();
    const an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn) && av !== '' && bv !== '') return (an - bn) * sortState.dir;
    return av.localeCompare(bv) * sortState.dir;
  });
  renderRows(rows);
}

function renderRows(dataset) {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';
  const frag = document.createDocumentFragment();
  dataset.forEach(r => {
    const tr = document.createElement('tr');
    if (r.Security_Tier && r.Security_Tier.indexOf('Tier 0') === 0) tr.className = 'tier0-row';
    columnKeys.forEach(col => {
      const td = document.createElement('td');
      td.innerHTML = badge(col, r[col]);
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
  document.getElementById('rowCount').textContent = dataset.length + ' of ' + rows.length + ' rows';
}

document.getElementById('searchBox').addEventListener('input', function (e) {
  const term = e.target.value.toLowerCase().trim();
  if (!term) { renderRows(rows); return; }
  const filtered = rows.filter(r => columnKeys.some(col => (r[col] || '').toString().toLowerCase().indexOf(term) !== -1));
  renderRows(filtered);
});

renderRows(rows);
</script>
</body>
</html>
"@

$fullHtml = $htmlHead + $htmlBodyTop + $htmlScript

try {
    Set-Content -Path $HtmlPath -Value $fullHtml -Encoding UTF8
    Write-Log "HTML dashboard written: $HtmlPath"
}
catch {
    Write-Log "HTML export failed: $($_.Exception.Message)" 'ERROR'
}

# ============================================================================
# REGION: Wrap-up
# ============================================================================

$duration = (Get-Date) - $ScriptStartTime
Write-Log "----------------------------------------------------------------"
Write-Log "DUMPSEC export complete in $([math]::Round($duration.TotalSeconds,1)) seconds."
Write-Log "Output folder: $OutputFolder"
Write-Log "  CSV : $CsvPath"
Write-Log "  XLSX: $XlsxPath"
Write-Log "  HTML: $HtmlPath"

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Green
Write-Host " DUMPSEC AD Security Export complete" -ForegroundColor Green
Write-Host " Folder: $OutputFolder" -ForegroundColor Green
Write-Host " Users exported: $($AllUsers.Count) across $($Domains.Count) domain(s)" -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Green

Invoke-Item $OutputFolder
