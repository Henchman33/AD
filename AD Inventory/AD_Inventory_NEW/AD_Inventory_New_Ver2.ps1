#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    AD Forest Inventory, Health Dashboard & Topology Map Generator — v3.0

.DESCRIPTION
    Performs a full deep-inventory of an Active Directory forest. Automatically
    classifies every OU and container by purpose (Tier 0/1/2, Domain Controllers,
    Servers, Workstations, Service Accounts, Users, Groups, etc.) and enumerates
    every object inside. Also collects DCs, FSMO roles, replication, GPOs, DNS,
    DHCP, PKI/certificates, password policies, and security alerts.

    Exports five output types, all to the current user's Desktop:
      1. HTML  — Interactive collapsible dashboard with live search
      2. CSV   — One .csv file per data category (12 files total)
      3. XLSX  — Multi-sheet Excel workbook via Excel COM (requires Excel)
      4. SVG   — Full AD forest / OU topology map (always generated)
      5. VSD   — Visio diagram via COM (generated only if Visio is installed)

    Requires : RSAT-AD-PowerShell  (ActiveDirectory module)
    Optional : RSAT-GPMC, RSAT-DNS-Server, RSAT-DHCP, Microsoft Visio, Microsoft Excel
    Run As   : Domain Administrator (or equivalent read rights)
    Platform : PowerShell ISE / PowerShell Console — Windows only

.PARAMETER DomainFQDN
    Target domain FQDN. Defaults to the current machine's joined domain.

.PARAMETER OutputPath
    Destination folder for all exports. Defaults to the current user's Desktop.

.PARAMETER Author
    Name stamped in all report headers.

.PARAMETER MaxOUDepth
    Maximum OU depth shown in the SVG/Visio map (default 4). Deeper OUs still
    appear in the HTML and CSV — only the visual map is capped.

.PARAMETER OpenOnComplete
    If set, opens the HTML report in the default browser when the script finishes.

.EXAMPLE
    .\AD-ForestInventory.ps1
    Inventory of current domain, all output to Desktop.

.EXAMPLE
    .\AD-ForestInventory.ps1 -DomainFQDN corp.contoso.com -OpenOnComplete
    Targets a specific domain and auto-opens the HTML report on completion.

.NOTES
    Author  : Stephen McKee — Server Administrator 2  (Enhanced v3.0)
    Version : 3.0
    Fixes from v2.0:
      - Broken PKI cert-template expression (?.Format syntax error)
      - Mismatched quotes in security alert action strings
      - GPO status inline-if with side-effects now properly handled
      - $ErrorActionPreference changed to 'Continue' (ISE-safe)
      - Added explicit ErrorAction per cmdlet instead of global Stop
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DomainFQDN = $env:USERDNSDOMAIN,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = [Environment]::GetFolderPath('Desktop'),

    [Parameter(Mandatory = $false)]
    [string]$Author = "$env:USERNAME — Server Administrator 2",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1,8)]
    [int]$MaxOUDepth = 4,

    [Parameter(Mandatory = $false)]
    [switch]$OpenOnComplete
)

$ErrorActionPreference = 'Continue'

# ══════════════════════════════════════════════════════════════════
#  REGION 1 — INITIALISE FOLDERS & LOGGING
# ══════════════════════════════════════════════════════════════════
$Timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
$ShortDomain  = ($DomainFQDN -split '\.')[0].ToUpper()
$ExportFolder = Join-Path $OutputPath "ADInventory_${ShortDomain}_${Timestamp}"

if (-not (Test-Path $ExportFolder)) {
    New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
}

$LogPath = Join-Path $ExportFolder "00_ADInventory_Run.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -ErrorAction SilentlyContinue
    switch ($Level) {
        'INFO'    { Write-Host $entry -ForegroundColor Cyan    }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow  }
        'ERROR'   { Write-Host $entry -ForegroundColor Red     }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green   }
    }
}

Write-Log "═══════════════════════════════════════════════════════" -Level INFO
Write-Log "  AD Forest Inventory v3.0 — Starting" -Level SUCCESS
Write-Log "  Running as : $env:USERNAME on $env:COMPUTERNAME"
Write-Log "  Target     : $DomainFQDN"
Write-Log "  Output     : $ExportFolder"
Write-Log "═══════════════════════════════════════════════════════" -Level INFO

# ══════════════════════════════════════════════════════════════════
#  REGION 2 — PREREQUISITES CHECK
# ══════════════════════════════════════════════════════════════════
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Script must be run as Administrator. Re-launch elevated." -Level ERROR
    exit 1
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory module loaded." -Level SUCCESS
} catch {
    Write-Log "ActiveDirectory module not found. Install RSAT-AD-PowerShell or run on a DC." -Level ERROR
    exit 1
}

$optionalModules = [ordered]@{
    'GroupPolicy' = 'GPMC / RSAT-GPMC'
    'DnsServer'   = 'RSAT-DNS-Server'
    'DHCPServer'  = 'RSAT-DHCP'
}
$modOK = @{}
foreach ($mod in $optionalModules.Keys) {
    if (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue) {
        Import-Module $mod -ErrorAction SilentlyContinue
        $modOK[$mod] = $true
        Write-Log "Optional module '$mod' loaded." -Level INFO
    } else {
        $modOK[$mod] = $false
        Write-Log "Optional module '$mod' not available ($($optionalModules[$mod])). That section will show N/A." -Level WARN
    }
}

# ══════════════════════════════════════════════════════════════════
#  REGION 3 — HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════

function Escape-Html {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return '—' }
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function Get-ParentDN {
    param([string]$DN)
    $parts = $DN -split '(?<!\\),', 2
    if ($parts.Count -ge 2) { return $parts[1] } else { return '' }
}

function Get-OUDepth {
    param([string]$DN, [string]$DomainDN)
    $relative = $DN -replace [regex]::Escape($DomainDN), ''
    $ouCount  = ([regex]::Matches($relative, '(?i),OU=')).Count
    return $ouCount
}

# Classify a container/OU by its name — returns type, hex color, short icon label
function Get-OUPurpose {
    param([string]$OUName)
    $n = $OUName.ToLower().Trim()

    if ($n -match 'domain.?controller|domain.?ctrl')    { return [PSCustomObject]@{ Type='Domain Controllers';          Color='#C0392B'; TxtColor='#FFFFFF'; Icon='DC'  } }
    if ($n -match '\btier.?0\b|\bt0\b')                 { return [PSCustomObject]@{ Type='Tier 0 — Privileged Admin';   Color='#6C3483'; TxtColor='#FFFFFF'; Icon='T0'  } }
    if ($n -match '\btier.?1\b|\bt1\b')                 { return [PSCustomObject]@{ Type='Tier 1 — Servers';            Color='#1A5276'; TxtColor='#FFFFFF'; Icon='T1'  } }
    if ($n -match '\btier.?2\b|\bt2\b')                 { return [PSCustomObject]@{ Type='Tier 2 — Workstations';       Color='#1E8449'; TxtColor='#FFFFFF'; Icon='T2'  } }
    if ($n -match 'paw|privileged.?access')             { return [PSCustomObject]@{ Type='Privileged Access (PAW)';     Color='#7D3C98'; TxtColor='#FFFFFF'; Icon='PAW' } }
    if ($n -match 'service.?account|svc.?account|\bsvc\b') { return [PSCustomObject]@{ Type='Service Accounts';         Color='#B7770D'; TxtColor='#FFFFFF'; Icon='SVC' } }
    if ($n -match 'managed.?service|gmsa|\bgmsa\b')     { return [PSCustomObject]@{ Type='Managed Svc Accounts (gMSA)'; Color='#D4AC0D'; TxtColor='#333333'; Icon='gMSA'} }
    if ($n -match '\bserver\b')                         { return [PSCustomObject]@{ Type='Servers';                     Color='#21618C'; TxtColor='#FFFFFF'; Icon='SRV' } }
    if ($n -match 'workstation|desktop|laptop|endpoint') { return [PSCustomObject]@{ Type='Workstations';               Color='#196F3D'; TxtColor='#FFFFFF'; Icon='WKS' } }
    if ($n -match '^users?$|people|staff|employee')     { return [PSCustomObject]@{ Type='Users';                       Color='#117A65'; TxtColor='#FFFFFF'; Icon='USR' } }
    if ($n -match 'security.?group|\bgroup\b')          { return [PSCustomObject]@{ Type='Security Groups';             Color='#9A7D0A'; TxtColor='#FFFFFF'; Icon='GRP' } }
    if ($n -match '\badmin\b|admins')                   { return [PSCustomObject]@{ Type='Administrative';              Color='#922B21'; TxtColor='#FFFFFF'; Icon='ADM' } }
    if ($n -match 'builtin')                            { return [PSCustomObject]@{ Type='Built-In Container';          Color='#424949'; TxtColor='#FFFFFF'; Icon='BLT' } }
    if ($n -match '\bcomputer\b')                       { return [PSCustomObject]@{ Type='Computers';                   Color='#1A5276'; TxtColor='#FFFFFF'; Icon='CMP' } }
    if ($n -match 'disabled|decommission|archive')      { return [PSCustomObject]@{ Type='Disabled / Archived';         Color='#717D7E'; TxtColor='#FFFFFF'; Icon='DIS' } }
    if ($n -match 'quarantine|staging|new.?obj|new.?hire') { return [PSCustomObject]@{ Type='Staging / Quarantine';    Color='#CA6F1E'; TxtColor='#FFFFFF'; Icon='STG' } }
    if ($n -match 'test|dev|lab|sandbox')               { return [PSCustomObject]@{ Type='Test / Dev / Lab';            Color='#1F618D'; TxtColor='#FFFFFF'; Icon='LAB' } }
    if ($n -match 'foreign.?security')                  { return [PSCustomObject]@{ Type='Foreign Security Principals'; Color='#5D6D7E'; TxtColor='#FFFFFF'; Icon='FSP' } }
    if ($n -match 'contact')                            { return [PSCustomObject]@{ Type='Contacts';                    Color='#148F77'; TxtColor='#FFFFFF'; Icon='CON' } }
    if ($n -match 'printer|print')                      { return [PSCustomObject]@{ Type='Printers';                    Color='#5F6A6A'; TxtColor='#FFFFFF'; Icon='PRT' } }
    return                                                       [PSCustomObject]@{ Type='General Container';           Color='#515A5A'; TxtColor='#FFFFFF'; Icon='OU'  }
}

function Get-ComputerRole {
    param([string]$OS, [string]$Name, [string[]]$DCNames)
    if ($DCNames -contains $Name) { return 'Domain Controller' }
    if ($OS -match 'Server') { return 'Member Server' }
    if ($OS -match 'Windows 1[01]|Windows 8|Windows 7|Windows XP|Windows Vista') { return 'Workstation' }
    return 'Unknown'
}

# ══════════════════════════════════════════════════════════════════
#  REGION 4 — DATA COLLECTION
# ══════════════════════════════════════════════════════════════════

# ── 4a. Forest & Domain ──────────────────────────────────────────
Write-Log "Collecting Forest & Domain info..."
Write-Progress -Activity "AD Inventory" -Status "Forest & Domain" -PercentComplete 2

$domain = $null; $forest = $null; $domainDN = ''; $domainMode = ''; $forestMode = ''; $pdcEmulator = ''
try {
    $domain      = Get-ADDomain -Server $DomainFQDN -ErrorAction Stop
    $forest      = Get-ADForest -Server $DomainFQDN -ErrorAction Stop
    $domainDN    = $domain.DistinguishedName
    $domainMode  = $domain.DomainMode.ToString()
    $forestMode  = $forest.ForestMode.ToString()
    $pdcEmulator = $domain.PDCEmulator
    Write-Log "Domain: $domainDN  |  Mode: $domainMode  |  Forest: $($forest.Name)" -Level SUCCESS
} catch {
    Write-Log "Failed to connect to domain '$DomainFQDN': $_" -Level ERROR
    exit 1
}

# All domains in the forest
$forestDomains = @()
try {
    foreach ($d in $forest.Domains) {
        try {
            $fd = Get-ADDomain -Identity $d -Server $d -ErrorAction SilentlyContinue
            if ($fd) {
                $forestDomains += [PSCustomObject]@{
                    DomainName  = $fd.DNSRoot
                    NetBIOS     = $fd.NetBIOSName
                    DN          = $fd.DistinguishedName
                    DomainMode  = $fd.DomainMode.ToString()
                    PDC         = $fd.PDCEmulator
                    RIDMaster   = $fd.RIDMaster
                    InfMaster   = $fd.InfrastructureMaster
                }
            }
        } catch { }
    }
    Write-Log "Forest contains $($forestDomains.Count) domain(s)." -Level SUCCESS
} catch {
    Write-Log "Could not enumerate all forest domains: $_" -Level WARN
}

# ── 4b. Domain Controllers ───────────────────────────────────────
Write-Log "Collecting Domain Controller data..."
Write-Progress -Activity "AD Inventory" -Status "Domain Controllers" -PercentComplete 8

$dcData = @()
$dcNames = @()
try {
    $dcs = Get-ADDomainController -Filter * -Server $DomainFQDN -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($dc in $dcs) {
        $dcNames += $dc.Name

        $rolesStr = if ($dc.OperationMasterRoles -and $dc.OperationMasterRoles.Count -gt 0) {
            ($dc.OperationMasterRoles | ForEach-Object { $_.ToString() }) -join ', '
        } else { 'None' }

        $replOK     = $true
        $replDetail = 'OK'
        try {
            $replStatus = Get-ADReplicationPartnerMetadata -Target $dc.HostName -Scope Server -ErrorAction SilentlyContinue
            if ($replStatus) {
                $failedPartners = @($replStatus | Where-Object { $_.LastReplicationResult -ne 0 })
                if ($failedPartners.Count -gt 0) { $replOK = $false; $replDetail = "ERR($($failedPartners.Count))" }
            }
        } catch { $replDetail = 'N/A' }

        $uptime = 'N/A'
        try {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $dc.HostName `
                                  -ErrorAction SilentlyContinue -OperationTimeoutSec 10
            if ($os) {
                $ts     = (Get-Date) - $os.LastBootUpTime
                $uptime = "$([int]$ts.TotalDays)d $($ts.Hours)h"
            }
        } catch { }

        $dcData += [PSCustomObject]@{
            Name         = $dc.Name
            HostName     = $dc.HostName
            Site         = $dc.Site
            OSVersion    = $dc.OperatingSystem
            OSBuild      = $dc.OperatingSystemVersion
            IsGC         = $dc.IsGlobalCatalog
            IsRODC       = $dc.IsReadOnly
            FSMORoles    = $rolesStr
            IPv4         = $dc.IPv4Address
            ReplStatus   = $replDetail
            ReplOK       = $replOK
            Uptime       = $uptime
        }
    }
    Write-Log "Collected $($dcData.Count) Domain Controllers." -Level SUCCESS
} catch {
    Write-Log "DC collection error: $_" -Level WARN
}

# ── 4c. Sites & Replication ──────────────────────────────────────
Write-Log "Collecting Sites & Replication data..."
Write-Progress -Activity "AD Inventory" -Status "Sites & Replication" -PercentComplete 14

$sitesData = @()
try {
    $sites = Get-ADReplicationSite -Filter * -Server $DomainFQDN -ErrorAction SilentlyContinue
    foreach ($site in $sites | Sort-Object Name) {
        $siteSubnets = @()
        try {
            $siteSubnets = Get-ADReplicationSubnet -Filter "Site -eq '$($site.DistinguishedName)'" `
                            -Server $DomainFQDN -ErrorAction SilentlyContinue |
                            Select-Object -ExpandProperty Name
        } catch { }
        $siteDCs = @($dcData | Where-Object { $_.Site -eq $site.Name })
        $sitesData += [PSCustomObject]@{
            SiteName    = $site.Name
            Description = $site.Description
            Subnets     = if ($siteSubnets.Count -gt 0) { $siteSubnets -join '; ' } else { '(none)' }
            DCCount     = $siteDCs.Count
            DCs         = ($siteDCs | Select-Object -ExpandProperty Name) -join ', '
        }
    }
    Write-Log "Collected $($sitesData.Count) AD Sites." -Level SUCCESS
} catch {
    Write-Log "Sites collection error: $_" -Level WARN
}

$replData = @()
try {
    $replConns = Get-ADReplicationConnection -Filter * -Server $DomainFQDN -Properties * -ErrorAction SilentlyContinue
    foreach ($conn in $replConns) {
        try {
            $src = ($conn.ReplicateFromDirectoryServer -split ',')[0] -replace '^CN=',''
            $dst = $conn.ReplicateToDirectoryServer   -replace 'CN=NTDS Settings,CN=','' -replace ',.*',''

            $meta = Get-ADReplicationPartnerMetadata -Target $src -Scope Server -ErrorAction SilentlyContinue |
                    Where-Object { $_.Partner -like "*$dst*" } | Select-Object -First 1

            $lastSuccess  = if ($meta -and $meta.LastReplicationSuccess) { $meta.LastReplicationSuccess.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
            $failures     = if ($meta -and $meta.ConsecutiveReplicationFailures) { $meta.ConsecutiveReplicationFailures } else { 0 }
            $result       = if ($meta -and $meta.LastReplicationResult -eq 0) { 'Success' } else { "Error ($($meta.LastReplicationResult))" }
            $statusClass  = if ($failures -eq 0) { 'ok' } elseif ($failures -lt 5) { 'warn' } else { 'crit' }

            $replData += [PSCustomObject]@{
                SourceDC    = $src
                DestDC      = $dst
                NC          = $conn.ReplicatedNamingContexts -join '; '
                LastSuccess = $lastSuccess
                Failures    = $failures
                Status      = $result
                StatusClass = $statusClass
            }
        } catch { }
    }
    Write-Log "Collected $($replData.Count) replication connections." -Level SUCCESS
} catch {
    Write-Log "Replication collection error: $_" -Level WARN
}

# ── 4d. PRE-FETCH ALL OBJECTS FOR EFFICIENT OU MAPPING ──────────
Write-Log "Pre-fetching all Users, Computers, and Groups for container analysis..."
Write-Progress -Activity "AD Inventory" -Status "Pre-fetching all AD objects" -PercentComplete 20

$allUsersRaw       = @()
$allComputersRaw   = @()
$allGroupsRaw      = @()

try {
    $allUsersRaw = @(Get-ADUser -Filter * -Server $DomainFQDN -Properties `
        DistinguishedName, SamAccountName, DisplayName, Name, Enabled,
        LastLogonDate, PasswordLastSet, PasswordNeverExpires, LockedOut,
        Department, Title, Manager, Description,
        'msDS-UserPasswordExpiryTimeComputed' -ErrorAction SilentlyContinue)
    Write-Log "Pre-fetched $($allUsersRaw.Count) users." -Level SUCCESS
} catch {
    Write-Log "Error pre-fetching users: $_" -Level WARN
}

try {
    $allComputersRaw = @(Get-ADComputer -Filter * -Server $DomainFQDN -Properties `
        DistinguishedName, Name, Enabled, OperatingSystem, OperatingSystemVersion,
        LastLogonDate, IPv4Address, DNSHostName, Description -ErrorAction SilentlyContinue)
    Write-Log "Pre-fetched $($allComputersRaw.Count) computers." -Level SUCCESS
} catch {
    Write-Log "Error pre-fetching computers: $_" -Level WARN
}

try {
    $allGroupsRaw = @(Get-ADGroup -Filter * -Server $DomainFQDN -Properties `
        DistinguishedName, Name, GroupCategory, GroupScope, Description,
        Members, ManagedBy -ErrorAction SilentlyContinue)
    Write-Log "Pre-fetched $($allGroupsRaw.Count) groups." -Level SUCCESS
} catch {
    Write-Log "Error pre-fetching groups: $_" -Level WARN
}

# Build parent-DN lookup tables for fast per-OU counting
$usersByParent     = @{}
$computersByParent = @{}
$groupsByParent    = @{}

foreach ($u in $allUsersRaw) {
    $p = Get-ParentDN $u.DistinguishedName
    if (-not $usersByParent.ContainsKey($p)) { $usersByParent[$p] = [System.Collections.Generic.List[object]]::new() }
    $usersByParent[$p].Add($u)
}
foreach ($c in $allComputersRaw) {
    $p = Get-ParentDN $c.DistinguishedName
    if (-not $computersByParent.ContainsKey($p)) { $computersByParent[$p] = [System.Collections.Generic.List[object]]::new() }
    $computersByParent[$p].Add($c)
}
foreach ($g in $allGroupsRaw) {
    $p = Get-ParentDN $g.DistinguishedName
    if (-not $groupsByParent.ContainsKey($p)) { $groupsByParent[$p] = [System.Collections.Generic.List[object]]::new() }
    $groupsByParent[$p].Add($g)
}

# ── 4e. DEEP OU / CONTAINER INVENTORY ───────────────────────────
Write-Log "Performing deep OU and container inventory..."
Write-Progress -Activity "AD Inventory" -Status "OU & Container Inventory" -PercentComplete 30

$ouInventory = @()

# Get all OUs
$allOUs = @()
try {
    $allOUs = @(Get-ADOrganizationalUnit -Filter * -Server $DomainFQDN -Properties `
        Name, DistinguishedName, Description, LinkedGroupPolicyObjects -ErrorAction SilentlyContinue |
        Sort-Object DistinguishedName)
} catch {
    Write-Log "OU enumeration error: $_" -Level WARN
}

# Also include built-in CN= containers
$builtinCNs = @('CN=Users', 'CN=Computers', 'CN=Builtin',
                'CN=ForeignSecurityPrincipals', 'CN=Managed Service Accounts')
$cnContainers = @()
foreach ($cn in $builtinCNs) {
    try {
        $obj = Get-ADObject -Identity "$cn,$domainDN" -Properties Name, DistinguishedName, Description `
               -Server $DomainFQDN -ErrorAction SilentlyContinue
        if ($obj) { $cnContainers += $obj }
    } catch { }
}

$allContainers = @($cnContainers) + @($allOUs)

foreach ($container in $allContainers) {
    try {
        $dn      = $container.DistinguishedName
        $ouName  = $container.Name
        $purpose = Get-OUPurpose -OUName $ouName
        $depth   = Get-OUDepth -DN $dn -DomainDN $domainDN

        # Objects directly inside this container (one level)
        $directUsers     = if ($usersByParent.ContainsKey($dn))     { @($usersByParent[$dn])     } else { @() }
        $directComputers = if ($computersByParent.ContainsKey($dn)) { @($computersByParent[$dn]) } else { @() }
        $directGroups    = if ($groupsByParent.ContainsKey($dn))    { @($groupsByParent[$dn])    } else { @() }

        # Computer role breakdown
        $dcCount          = ($directComputers | Where-Object { $dcNames -contains $_.Name }).Count
        $serverCount      = ($directComputers | Where-Object { $dcNames -notcontains $_.Name -and $_.OperatingSystem -match 'Server' }).Count
        $workstationCount = ($directComputers | Where-Object { $dcNames -notcontains $_.Name -and $_.OperatingSystem -notmatch 'Server' }).Count

        # Enabled/Disabled breakdown for users
        $enabledUsers    = ($directUsers | Where-Object { $_.Enabled }).Count
        $disabledUsers   = ($directUsers | Where-Object { -not $_.Enabled }).Count

        # GPO link count
        $gpoLinkCount = 0
        if ($container.LinkedGroupPolicyObjects) { $gpoLinkCount = @($container.LinkedGroupPolicyObjects).Count }

        # Path relative to domain root (for readability)
        $relativePath = ($dn -replace [regex]::Escape($domainDN), '') -replace '^,', '' -replace '(?i)(,OU=|,CN=)', ' > ' -replace '(?i)^(OU=|CN=)', ''

        $ouInventory += [PSCustomObject]@{
            ContainerName     = $ouName
            Purpose           = $purpose.Type
            PurposeColor      = $purpose.Color
            PurposeIcon       = $purpose.Icon
            Depth             = $depth
            RelativePath      = $relativePath
            DistinguishedName = $dn
            Description       = if ($container.Description) { $container.Description } else { '' }
            UserCount         = $directUsers.Count
            EnabledUsers      = $enabledUsers
            DisabledUsers     = $disabledUsers
            ComputerCount     = $directComputers.Count
            DCCount           = $dcCount
            ServerCount       = $serverCount
            WorkstationCount  = $workstationCount
            GroupCount        = $directGroups.Count
            GPOLinksCount     = $gpoLinkCount
            ContainerType     = if ($container.ObjectClass -eq 'organizationalUnit') { 'OU' } else { 'CN Container' }
        }
    } catch {
        Write-Log "Error processing container '$($container.DistinguishedName)': $_" -Level WARN
    }
}

Write-Log "Container inventory complete — $($ouInventory.Count) containers catalogued." -Level SUCCESS

# ── 4f. USER ACCOUNT HEALTH ──────────────────────────────────────
Write-Log "Analyzing User account health..."
Write-Progress -Activity "AD Inventory" -Status "User Account Health" -PercentComplete 42

$cutoff90    = (Get-Date).AddDays(-90)
$cutoff180   = (Get-Date).AddDays(-180)
$cutoff14exp = (Get-Date).AddDays(14)

$totalEnabled    = 0; $totalDisabled   = 0
$userStaleCount  = 0; $userLockedCount = 0
$userPwdNever    = 0; $userPwdExpiring = 0

$userTableData = @()

foreach ($u in $allUsersRaw) {
    if ($u.Enabled) { $totalEnabled++ } else { $totalDisabled++ }
    if ($u.LockedOut -and $u.Enabled) { $userLockedCount++ }
    if ($u.PasswordNeverExpires -and $u.Enabled) { $userPwdNever++ }
    if ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $cutoff90) { $userStaleCount++ }

    if ($u.Enabled -and -not $u.PasswordNeverExpires) {
        $expRaw = $u.'msDS-UserPasswordExpiryTimeComputed'
        if ($expRaw -and $expRaw -ne '0' -and $expRaw -ne '9223372036854775807') {
            try {
                $expDt = [datetime]::FromFileTime([int64]$expRaw)
                if ($expDt -gt (Get-Date) -and $expDt -lt $cutoff14exp) { $userPwdExpiring++ }
            } catch { }
        }
    }

    # Build audit row for notable accounts
    $flag = ''
    $risk = 'LOW'
    if ($u.LockedOut -and $u.Enabled)                                               { $flag = 'LOCKED';         $risk = 'HIGH'   }
    elseif ($u.PasswordNeverExpires -and $u.Enabled -and $u.SamAccountName -notmatch 'krbtgt') { $flag = 'PWD_NEVER_EXP'; $risk = 'HIGH' }
    elseif ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $cutoff180)  { $flag = 'STALE_180D';     $risk = 'HIGH'   }
    elseif ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $cutoff90)   { $flag = 'STALE_90D';      $risk = 'MEDIUM' }
    elseif (-not $u.Enabled)                                                         { $flag = 'DISABLED';       $risk = 'INFO'   }

    if ($flag -ne '') {
        $ouPath = ($u.DistinguishedName -replace 'CN=[^,]+,', '') -replace [regex]::Escape($domainDN), '' -replace '^,', ''
        $userTableData += [PSCustomObject]@{
            SamAccountName  = $u.SamAccountName
            DisplayName     = $u.Name
            Department      = $u.Department
            Title           = $u.Title
            Enabled         = $u.Enabled
            Flag            = $flag
            Risk            = $risk
            LastLogon       = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdLastSet      = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdNeverExpires = $u.PasswordNeverExpires
            LockedOut       = $u.LockedOut
            OU              = $ouPath
        }
    }
}

$userTableData = $userTableData | Sort-Object @{E={switch($_.Risk){'HIGH'{0}'MEDIUM'{1}default{2}}}}, SamAccountName | Select-Object -First 200
Write-Log "User analysis: Enabled=$totalEnabled, Disabled=$totalDisabled, Stale(90d)=$userStaleCount, Locked=$userLockedCount, PwdNeverExp=$userPwdNever" -Level SUCCESS

# ── 4g. COMPUTER INVENTORY ───────────────────────────────────────
Write-Log "Building Computer inventory..."
Write-Progress -Activity "AD Inventory" -Status "Computer Inventory" -PercentComplete 50

$computerData = @()
foreach ($c in $allComputersRaw | Sort-Object Name) {
    $role    = Get-ComputerRole -OS $c.OperatingSystem -Name $c.Name -DCNames $dcNames
    $ouPath  = ($c.DistinguishedName -replace 'CN=[^,]+,', '') -replace [regex]::Escape($domainDN), '' -replace '^,', ''
    $purpose = Get-OUPurpose -OUName (($ouPath -split '>')[0].Trim())

    $computerData += [PSCustomObject]@{
        Name         = $c.Name
        DNSHostName  = $c.DNSHostName
        OperatingSystem = $c.OperatingSystem
        OSVersion    = $c.OperatingSystemVersion
        Role         = $role
        Enabled      = $c.Enabled
        LastLogon    = if ($c.LastLogonDate) { $c.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
        IPv4         = $c.IPv4Address
        Description  = $c.Description
        OU           = $ouPath
        ContainerPurpose = $purpose.Type
    }
}
Write-Log "Computer inventory: $($computerData.Count) total — DCs: $(($computerData | Where-Object {$_.Role -eq 'Domain Controller'}).Count), Servers: $(($computerData | Where-Object {$_.Role -eq 'Member Server'}).Count), Workstations: $(($computerData | Where-Object {$_.Role -eq 'Workstation'}).Count)" -Level SUCCESS

# ── 4h. SECURITY GROUPS ──────────────────────────────────────────
Write-Log "Building Group inventory..."
Write-Progress -Activity "AD Inventory" -Status "Security Groups" -PercentComplete 56

$groupData = @()
foreach ($g in $allGroupsRaw | Sort-Object Name) {
    $memberCount = 0
    try {
        $memberCount = @(Get-ADGroupMember -Identity $g.DistinguishedName -Server $DomainFQDN -ErrorAction SilentlyContinue).Count
    } catch { }

    $ouPath = ($g.DistinguishedName -replace 'CN=[^,]+,', '') -replace [regex]::Escape($domainDN), '' -replace '^,', ''
    $groupData += [PSCustomObject]@{
        Name         = $g.Name
        Category     = $g.GroupCategory.ToString()
        Scope        = $g.GroupScope.ToString()
        MemberCount  = $memberCount
        Description  = $g.Description
        ManagedBy    = if ($g.ManagedBy) { ($g.ManagedBy -split ',')[0] -replace 'CN=','' } else { '' }
        OU           = $ouPath
    }
}
Write-Log "Group inventory: $($groupData.Count) groups." -Level SUCCESS

# ── 4i. FINE-GRAINED PASSWORD POLICIES ──────────────────────────
Write-Log "Collecting Password Policies..."
Write-Progress -Activity "AD Inventory" -Status "Password Policies" -PercentComplete 60

$psoData = @()
try {
    $psos = Get-ADFineGrainedPasswordPolicy -Filter * -Server $DomainFQDN -Properties * -ErrorAction SilentlyContinue
    foreach ($pso in $psos | Sort-Object Precedence) {
        $subjects = @()
        try {
            $subjects = @(Get-ADFineGrainedPasswordPolicySubject -Identity $pso -Server $DomainFQDN -ErrorAction SilentlyContinue |
                          Select-Object -ExpandProperty Name)
        } catch { }
        $psoData += [PSCustomObject]@{
            Name            = $pso.Name
            Precedence      = $pso.Precedence
            AppliedTo       = if ($subjects.Count -gt 0) { $subjects -join ', ' } else { '(none assigned)' }
            MinLength       = $pso.MinPasswordLength
            MaxAgeDays      = if ($pso.MaxPasswordAge.TotalDays -eq 0) { 'Never' } else { [int]$pso.MaxPasswordAge.TotalDays }
            LockoutThresh   = if ($pso.LockoutThreshold -eq 0) { 'Disabled' } else { "$($pso.LockoutThreshold) attempts" }
            LockoutWindow   = "$([int]$pso.LockoutObservationWindow.TotalMinutes) min"
            LockoutDuration = if ($pso.LockoutDuration.TotalMinutes -eq 0) { 'Manual unlock' } else { "$([int]$pso.LockoutDuration.TotalMinutes) min" }
            Complexity      = $pso.ComplexityEnabled
            Reversible      = $pso.ReversibleEncryptionEnabled
            Type            = 'Fine-Grained PSO'
        }
    }
    if ($psoData.Count -eq 0) {
        $ddp = Get-ADDefaultDomainPasswordPolicy -Server $DomainFQDN -ErrorAction SilentlyContinue
        if ($ddp) {
            $psoData += [PSCustomObject]@{
                Name            = 'Default Domain Policy'
                Precedence      = 'N/A (Domain Default)'
                AppliedTo       = 'All users (domain-level)'
                MinLength       = $ddp.MinPasswordLength
                MaxAgeDays      = if ($ddp.MaxPasswordAge.TotalDays -eq 0) { 'Never' } else { [int]$ddp.MaxPasswordAge.TotalDays }
                LockoutThresh   = if ($ddp.LockoutThreshold -eq 0) { 'Disabled' } else { "$($ddp.LockoutThreshold) attempts" }
                LockoutWindow   = "$([int]$ddp.LockoutObservationWindow.TotalMinutes) min"
                LockoutDuration = if ($ddp.LockoutDuration.TotalMinutes -eq 0) { 'Manual unlock' } else { "$([int]$ddp.LockoutDuration.TotalMinutes) min" }
                Complexity      = $ddp.ComplexityEnabled
                Reversible      = $ddp.ReversibleEncryptionEnabled
                Type            = 'Default Domain'
            }
        }
    }
    Write-Log "Collected $($psoData.Count) password policy entries." -Level SUCCESS
} catch {
    Write-Log "Password policy collection error: $_" -Level WARN
}

# ── 4j. GROUP POLICY ─────────────────────────────────────────────
Write-Log "Collecting Group Policy data..."
Write-Progress -Activity "AD Inventory" -Status "Group Policy" -PercentComplete 64

$gpoData       = @()
$gpoTotal      = 0; $gpoUnlinked = 0; $gpoLinked = 0; $gpoEnforced = 0

if ($modOK['GroupPolicy']) {
    try {
        $allGPOs = Get-GPO -All -Domain $DomainFQDN -ErrorAction SilentlyContinue
        $gpoTotal = $allGPOs.Count
        foreach ($gpo in $allGPOs | Sort-Object DisplayName) {
            try {
                $report   = [xml]($gpo | Get-GPOReport -ReportType XML -Domain $DomainFQDN -ErrorAction SilentlyContinue)
                $links     = $report.GPO.LinksTo
                $linkedTo  = if ($links) { @($links | ForEach-Object { $_.SOMPath }) -join '; ' } else { 'Not Linked' }
                $isEnforced = $false
                if ($links) {
                    $isEnforced = (@($links | Where-Object { $_.NoOverride -eq $true }).Count -gt 0)
                }
                $status = if (-not $links) { $gpoUnlinked++; 'UNLINKED' }
                          elseif ($isEnforced) { $gpoEnforced++; 'ENFORCED' }
                          else { $gpoLinked++; 'ENABLED' }

                $gpoData += [PSCustomObject]@{
                    Name         = $gpo.DisplayName
                    GUID         = $gpo.Id.ToString()
                    LinkedTo     = $linkedTo
                    Status       = $status
                    GpoStatus    = $gpo.GpoStatus.ToString()
                    WMIFilter    = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { 'None' }
                    LastModified = $gpo.ModificationTime.ToString('yyyy-MM-dd HH:mm')
                    CreatedDate  = $gpo.CreationTime.ToString('yyyy-MM-dd')
                }
            } catch { }
        }
        Write-Log "GPOs: $gpoTotal total ($gpoLinked linked, $gpoUnlinked unlinked, $gpoEnforced enforced)." -Level SUCCESS
    } catch {
        Write-Log "GPO collection error: $_" -Level WARN
    }
} else {
    $gpoData += [PSCustomObject]@{ Name='Install RSAT-GPMC to collect GPO data'; GUID='N/A'; LinkedTo='N/A'; Status='N/A'; GpoStatus='N/A'; WMIFilter='N/A'; LastModified='N/A'; CreatedDate='N/A' }
}

# ── 4k. DNS ZONES ────────────────────────────────────────────────
Write-Log "Collecting DNS Zone data..."
Write-Progress -Activity "AD Inventory" -Status "DNS Zones" -PercentComplete 70

$dnsData = @()
if ($modOK['DnsServer']) {
    try {
        $dnsServer = $pdcEmulator
        $zones = Get-DnsServerZone -ComputerName $dnsServer -ErrorAction SilentlyContinue
        foreach ($z in $zones | Where-Object { -not $z.IsAutoCreated }) {
            $rrCount = 0
            try { $rrCount = @(Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue).Count } catch { }

            $aging    = 'N/A'
            $scavOn   = $false
            if ($z.ZoneType -eq 'Primary') {
                try {
                    $za   = Get-DnsServerZoneAging -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue
                    $scavOn = $za.AgingEnabled
                    $aging  = if ($za.AgingEnabled) { 'Enabled' } else { 'Disabled' }
                } catch { $aging = 'Unknown' }
            }
            $dnsData += [PSCustomObject]@{
                ZoneName    = $z.ZoneName
                ZoneType    = "$($z.ZoneType)$(if ($z.IsADIntegrated) {' (AD-Integrated)'})"
                RepScope    = if ($z.IsADIntegrated) { $z.ReplicationScope.ToString() } else { 'File-based' }
                DynUpdate   = $z.DynamicUpdate.ToString()
                RecordCount = $rrCount
                Scavenging  = $aging
                Status      = if (-not $scavOn -and $z.ZoneType -eq 'Primary' -and $z.IsADIntegrated) { 'WARN' } else { 'HEALTHY' }
            }
        }
        Write-Log "DNS: $($dnsData.Count) zones collected." -Level SUCCESS
    } catch {
        Write-Log "DNS collection error: $_" -Level WARN
    }
} else {
    $dnsData += [PSCustomObject]@{ ZoneName='Install RSAT-DNS-Server to collect DNS data'; ZoneType='N/A'; RepScope='N/A'; DynUpdate='N/A'; RecordCount=0; Scavenging='N/A'; Status='N/A' }
}

# ── 4l. DHCP SCOPES ──────────────────────────────────────────────
Write-Log "Collecting DHCP Scope data..."
Write-Progress -Activity "AD Inventory" -Status "DHCP Scopes" -PercentComplete 74

$dhcpData = @()
if ($modOK['DHCPServer']) {
    try {
        $dhcpServers = Get-DhcpServerInDC -ErrorAction SilentlyContinue
        foreach ($srv in $dhcpServers) {
            try {
                $scopes = Get-DhcpServerv4Scope -ComputerName $srv.DnsName -ErrorAction SilentlyContinue
                foreach ($scope in $scopes) {
                    $stats   = Get-DhcpServerv4ScopeStatistics -ScopeId $scope.ScopeId -ComputerName $srv.DnsName -ErrorAction SilentlyContinue
                    $pct     = if ($stats -and $stats.PercentageInUse) { [math]::Round($stats.PercentageInUse, 1) } else { 0 }
                    $failover = $null
                    try { $failover = Get-DhcpServerv4Failover -ScopeId $scope.ScopeId -ComputerName $srv.DnsName -ErrorAction SilentlyContinue } catch { }

                    $dhcpData += [PSCustomObject]@{
                        Server      = $srv.DnsName
                        ScopeName   = $scope.Name
                        ScopeID     = $scope.ScopeId
                        Subnet      = "$($scope.ScopeId)/$($scope.SubnetMask)"
                        Total       = if ($stats) { $stats.AddressesFree + $stats.AddressesInUse } else { 0 }
                        InUse       = if ($stats) { $stats.AddressesInUse } else { 0 }
                        Free        = if ($stats) { $stats.AddressesFree }  else { 0 }
                        UtilPct     = $pct
                        Failover    = if ($failover) { $failover.Mode.ToString() } else { 'None' }
                        Status      = if ($pct -ge 90) { 'HIGH USE' } elseif ($pct -ge 75) { 'WARNING' } else { 'HEALTHY' }
                    }
                }
            } catch { }
        }
        Write-Log "DHCP: $($dhcpData.Count) scopes collected." -Level SUCCESS
    } catch {
        Write-Log "DHCP collection error: $_" -Level WARN
    }
} else {
    $dhcpData += [PSCustomObject]@{ Server='N/A'; ScopeName='Install RSAT-DHCP to collect DHCP data'; ScopeID='N/A'; Subnet='N/A'; Total=0; InUse=0; Free=0; UtilPct=0; Failover='N/A'; Status='N/A' }
}

# ── 4m. PKI / CERTIFICATES ───────────────────────────────────────
Write-Log "Collecting PKI / Certificate data..."
Write-Progress -Activity "AD Inventory" -Status "PKI & Certificates" -PercentComplete 78

$pkiData = @()
foreach ($dc in $dcData) {
    try {
        $certs = Invoke-Command -ComputerName $dc.HostName -ErrorAction SilentlyContinue -ScriptBlock {
            Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
                Select-Object Subject, Thumbprint, NotAfter, NotBefore, HasPrivateKey,
                @{N='Template'; E={
                    $ext = $_.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Certificate Template Name' }
                    if ($ext) { try { $ext.Format(0) } catch { 'Unknown' } } else { 'Unknown' }
                }}
        }
        if ($certs) {
            foreach ($cert in $certs) {
                $daysLeft = ([datetime]$cert.NotAfter - (Get-Date)).Days
                $pkiData += [PSCustomObject]@{
                    CommonName  = ($cert.Subject -replace 'CN=', '' -replace ',.*', '').Trim()
                    IssuedTo    = $dc.HostName
                    Template    = $cert.Template
                    IssuedDate  = ([datetime]$cert.NotBefore).ToString('yyyy-MM-dd')
                    Expires     = ([datetime]$cert.NotAfter).ToString('yyyy-MM-dd')
                    DaysLeft    = $daysLeft
                    HasKey      = $cert.HasPrivateKey
                    Status      = if ($daysLeft -lt 0) { 'EXPIRED' } elseif ($daysLeft -lt 14) { 'CRITICAL' } elseif ($daysLeft -lt 30) { 'EXPIRING' } else { 'HEALTHY' }
                    Source      = "DC:$($dc.Name)"
                }
            }
        }
    } catch { }
}

# Local machine certificate store
try {
    foreach ($cert in (Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { $_.HasPrivateKey })) {
        $daysLeft = ($cert.NotAfter - (Get-Date)).Days
        $template = 'Unknown'
        try {
            $ext = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Certificate Template Name' }
            if ($ext) { $template = $ext.Format(0) }
        } catch { }
        $pkiData += [PSCustomObject]@{
            CommonName  = ($cert.Subject -replace 'CN=', '' -replace ',.*', '').Trim()
            IssuedTo    = $env:COMPUTERNAME
            Template    = $template
            IssuedDate  = $cert.NotBefore.ToString('yyyy-MM-dd')
            Expires     = $cert.NotAfter.ToString('yyyy-MM-dd')
            DaysLeft    = $daysLeft
            HasKey      = $true
            Status      = if ($daysLeft -lt 0) { 'EXPIRED' } elseif ($daysLeft -lt 14) { 'CRITICAL' } elseif ($daysLeft -lt 30) { 'EXPIRING' } else { 'HEALTHY' }
            Source      = 'LocalMachine'
        }
    }
} catch { }

$pkiData = $pkiData | Sort-Object DaysLeft
Write-Log "PKI: $($pkiData.Count) certificate entries collected." -Level SUCCESS

# ── 4n. PRIVILEGED GROUPS ────────────────────────────────────────
Write-Log "Collecting Privileged Group memberships..."
Write-Progress -Activity "AD Inventory" -Status "Privileged Groups" -PercentComplete 82

$privGroups = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators',
                'Backup Operators','Account Operators','Server Operators','Print Operators',
                'Group Policy Creator Owners','DnsAdmins','DHCP Administrators','Protected Users')
$privGroupData = @()
foreach ($gname in $privGroups) {
    try {
        $grp = Get-ADGroup -Filter "Name -eq '$gname'" -Server $DomainFQDN -ErrorAction SilentlyContinue
        if (-not $grp) { continue }
        $members    = @(Get-ADGroupMember -Identity $grp -Recursive -Server $DomainFQDN -ErrorAction SilentlyContinue)
        $memberStr  = ($members | Select-Object -ExpandProperty SamAccountName) -join ', '
        $tier       = switch ($gname) {
            'Domain Admins'         { 'TIER 0' }
            'Enterprise Admins'     { 'TIER 0' }
            'Schema Admins'         { 'TIER 0' }
            'Administrators'        { 'TIER 0' }
            'Protected Users'       { 'TIER 0' }
            default                 { 'TIER 1' }
        }
        $privGroupData += [PSCustomObject]@{
            GroupName    = $gname
            Tier         = $tier
            MemberCount  = $members.Count
            Members      = if ($memberStr.Length -gt 200) { $memberStr.Substring(0, 197) + '...' } else { $memberStr }
            ReviewStatus = if ($members.Count -gt 10) { 'REVIEW REQUIRED' } else { 'OK' }
        }
    } catch { }
}
Write-Log "Privileged group inventory: $($privGroupData.Count) groups." -Level SUCCESS

# ── 4o. SECURITY ALERTS ──────────────────────────────────────────
Write-Log "Building Security Alert summary..."
Write-Progress -Activity "AD Inventory" -Status "Security Alerts" -PercentComplete 86

$secAlerts = @()
$alertId   = 1

if ($userLockedCount -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'Accounts'; Severity = 'WARNING'
        Finding  = "$userLockedCount user account(s) currently locked out"
        Object   = "$userLockedCount accounts"
        Action   = 'Review lockout source via Event ID 4740; unlock and investigate'
    }
}
if ($userStaleCount -gt 5) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'Accounts'; Severity = 'WARNING'
        Finding  = "$userStaleCount enabled accounts inactive 90+ days"
        Object   = "$userStaleCount accounts"
        Action   = 'Disable or remove; coordinate with managers before action'
    }
}
if ($userPwdNever -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'Accounts'; Severity = if ($userPwdNever -gt 20) { 'CRITICAL' } else { 'WARNING' }
        Finding  = "$userPwdNever enabled accounts with Password Never Expires set"
        Object   = "$userPwdNever accounts"
        Action   = 'Apply PSO or enforce expiry; exempt only gMSA and approved break-glass accounts'
    }
}
foreach ($cert in ($pkiData | Where-Object { $_.DaysLeft -lt 30 })) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'PKI'; Severity = if ($cert.DaysLeft -lt 7) { 'CRITICAL' } else { 'WARNING' }
        Finding  = "Certificate expiring in $($cert.DaysLeft) day(s): $($cert.CommonName)"
        Object   = $cert.IssuedTo
        Action   = 'Renew certificate immediately to prevent auth failures'
    }
}
if ($gpoUnlinked -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'GPO'; Severity = 'INFO'
        Finding  = "$gpoUnlinked unlinked GPO(s) in the domain"
        Object   = 'Group Policy'
        Action   = 'Review unlinked GPOs; link or delete to reduce namespace clutter'
    }
}
foreach ($rf in ($replData | Where-Object { $_.Failures -gt 0 })) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'Replication'; Severity = if ($rf.Failures -ge 5) { 'CRITICAL' } else { 'WARNING' }
        Finding  = "Replication failures: $($rf.SourceDC) to $($rf.DestDC) ($($rf.Failures) consecutive)"
        Object   = $rf.DestDC
        Action   = 'Run: repadmin /replsummary and repadmin /syncall /AdeP'
    }
}
$dnsWarnZones = @($dnsData | Where-Object { $_.Status -eq 'WARN' })
if ($dnsWarnZones.Count -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'DNS'; Severity = 'INFO'
        Finding  = "$($dnsWarnZones.Count) DNS zone(s) with scavenging disabled"
        Object   = ($dnsWarnZones | Select-Object -ExpandProperty ZoneName) -join ', '
        Action   = 'Enable DNS scavenging to prevent stale record buildup'
    }
}
$domAdminsRow = $privGroupData | Where-Object { $_.GroupName -eq 'Domain Admins' }
if ($domAdminsRow -and [int]$domAdminsRow.MemberCount -gt 8) {
    $secAlerts += [PSCustomObject]@{
        AlertID = "SEC-$($alertId.ToString('000'))"; $alertId++=
        Category = 'Privilege'; Severity = 'WARNING'
        Finding  = "Domain Admins has $($domAdminsRow.MemberCount) members (best practice: 5 or fewer)"
        Object   = 'Domain Admins'
        Action   = 'Review and reduce Domain Admins membership; use delegated role groups instead'
    }
}
$critAlerts = ($secAlerts | Where-Object { $_.Severity -eq 'CRITICAL' }).Count
$warnAlerts = ($secAlerts | Where-Object { $_.Severity -eq 'WARNING'  }).Count
Write-Log "Security alerts: $($secAlerts.Count) total ($critAlerts CRITICAL, $warnAlerts WARNING)." -Level SUCCESS

# ══════════════════════════════════════════════════════════════════
#  REGION 5 — EXPORT: CSV FILES
# ══════════════════════════════════════════════════════════════════
Write-Log "Exporting CSV files..."
Write-Progress -Activity "AD Inventory" -Status "Exporting CSVs" -PercentComplete 88

$csvSections = [ordered]@{
    '01_ForestDomains'      = $forestDomains
    '02_DomainControllers'  = $dcData
    '03_Sites'              = $sitesData
    '04_Replication'        = $replData
    '05_OUContainerInventory' = ($ouInventory | Select-Object ContainerName, Purpose, Depth, ContainerType, RelativePath, Description, UserCount, EnabledUsers, DisabledUsers, ComputerCount, DCCount, ServerCount, WorkstationCount, GroupCount, GPOLinksCount, DistinguishedName)
    '06_Users_Flagged'      = $userTableData
    '07_Computers'          = $computerData
    '08_Groups'             = $groupData
    '09_PasswordPolicies'   = $psoData
    '10_GroupPolicy'        = $gpoData
    '11_DNS'                = $dnsData
    '12_DHCP'               = $dhcpData
    '13_PKI_Certificates'   = $pkiData
    '14_PrivilegedGroups'   = $privGroupData
    '15_SecurityAlerts'     = $secAlerts
}

foreach ($name in $csvSections.Keys) {
    $csvPath = Join-Path $ExportFolder "${name}.csv"
    try {
        $csvSections[$name] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Log "CSV: $name.csv saved." -Level INFO
    } catch {
        Write-Log "CSV export failed for $name : $_" -Level WARN
    }
}

# ══════════════════════════════════════════════════════════════════
#  REGION 6 — EXPORT: EXCEL WORKBOOK (via COM — requires Excel)
# ══════════════════════════════════════════════════════════════════
Write-Log "Exporting Excel workbook via COM..."
Write-Progress -Activity "AD Inventory" -Status "Exporting Excel" -PercentComplete 90

$xlPath   = Join-Path $ExportFolder "ADInventory_${ShortDomain}_${Timestamp}.xlsx"
$xlApp    = $null
$xlBook   = $null

try {
    $xlApp = New-Object -ComObject Excel.Application -ErrorAction Stop
    $xlApp.Visible        = $false
    $xlApp.DisplayAlerts  = $false
    $xlBook = $xlApp.Workbooks.Add()

    # Color constants (OLE BGR format for Excel COM)
    $headerBg    = 0x9E4F1B  # deep blue  (BGR of #1B4F9E)
    $headerFg    = 0xFFFFFF  # white
    $altRow      = 0xFBF2EB  # light blue (BGR of #EBF2FB)
    $critColor   = 0x0000CC  # red
    $warnColor   = 0x00AAFF  # amber

    function Add-XLSheet {
        param(
            [object]$Workbook,
            [string]$SheetName,
            [object[]]$Data,
            [string]$HeaderNote = ''
        )
        if ($Data.Count -eq 0) { return }

        $ws   = $Workbook.Sheets.Add()
        $ws.Name = $SheetName.Substring(0, [Math]::Min($SheetName.Length, 31))

        $headers = @($Data[0].PSObject.Properties.Name)

        # Title row (row 1)
        $ws.Cells(1, 1) = "AD Inventory — $SheetName — $DomainFQDN — $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $titleRange = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $headers.Count))
        $titleRange.Merge() | Out-Null
        $titleRange.Interior.Color = $headerBg
        $titleRange.Font.Color     = $headerFg
        $titleRange.Font.Bold      = $true
        $titleRange.Font.Size      = 13
        $titleRange.HorizontalAlignment = -4108  # xlCenter

        # Header row (row 2)
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $cell = $ws.Cells(2, $i + 1)
            $cell.Value2 = $headers[$i]
            $cell.Interior.Color  = $headerBg
            $cell.Font.Color      = $headerFg
            $cell.Font.Bold       = $true
        }

        # Data rows (starting row 3)
        for ($r = 0; $r -lt $Data.Count; $r++) {
            $row = $Data[$r]
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $val = $row.($headers[$c])
                if ($null -eq $val) { $val = '' }
                $ws.Cells($r + 3, $c + 1).Value2 = $val.ToString()
            }
            if ($r % 2 -eq 1) {
                $rowRange = $ws.Range($ws.Cells($r + 3, 1), $ws.Cells($r + 3, $headers.Count))
                $rowRange.Interior.Color = $altRow
            }
            # Highlight critical/warning in security alerts
            if ($row.PSObject.Properties['Severity']) {
                $sevCell = $ws.Cells($r + 3, ($headers.IndexOf('Severity') + 1))
                if ($row.Severity -eq 'CRITICAL') { $sevCell.Interior.Color = $critColor; $sevCell.Font.Color = 0xFFFFFF }
                elseif ($row.Severity -eq 'WARNING') { $sevCell.Interior.Color = $warnColor }
            }
        }

        # Auto-fit columns
        $usedRange = $ws.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null

        # Freeze panes below header rows
        $ws.Rows(3).Select() | Out-Null
        $xlApp.ActiveWindow.FreezePanes = $true
    }

    # Add a sheet per major section
    Add-XLSheet -Workbook $xlBook -SheetName 'Forest & Domains'       -Data $forestDomains
    Add-XLSheet -Workbook $xlBook -SheetName 'Domain Controllers'     -Data $dcData
    Add-XLSheet -Workbook $xlBook -SheetName 'Sites & Replication'    -Data ($sitesData + $replData)
    Add-XLSheet -Workbook $xlBook -SheetName 'OU Container Inventory' -Data ($ouInventory | Select-Object ContainerName, Purpose, Depth, ContainerType, RelativePath, UserCount, EnabledUsers, DisabledUsers, ComputerCount, DCCount, ServerCount, WorkstationCount, GroupCount)
    Add-XLSheet -Workbook $xlBook -SheetName 'Flagged Users'          -Data $userTableData
    Add-XLSheet -Workbook $xlBook -SheetName 'Computers'              -Data $computerData
    Add-XLSheet -Workbook $xlBook -SheetName 'Groups'                 -Data $groupData
    Add-XLSheet -Workbook $xlBook -SheetName 'Password Policies'      -Data $psoData
    Add-XLSheet -Workbook $xlBook -SheetName 'Group Policy'           -Data $gpoData
    Add-XLSheet -Workbook $xlBook -SheetName 'DNS'                    -Data $dnsData
    Add-XLSheet -Workbook $xlBook -SheetName 'DHCP'                   -Data $dhcpData
    Add-XLSheet -Workbook $xlBook -SheetName 'PKI Certificates'       -Data $pkiData
    Add-XLSheet -Workbook $xlBook -SheetName 'Privileged Groups'      -Data $privGroupData
    Add-XLSheet -Workbook $xlBook -SheetName 'Security Alerts'        -Data $secAlerts

    # Remove the default empty Sheet1
    $defaultSheets = @($xlBook.Sheets | Where-Object { $_.Name -match '^Sheet\d+$' })
    foreach ($s in $defaultSheets) { $s.Delete() }

    $xlBook.SaveAs($xlPath)
    Write-Log "Excel workbook saved: $xlPath" -Level SUCCESS
} catch {
    Write-Log "Excel COM export failed (Excel may not be installed): $_" -Level WARN
} finally {
    if ($xlBook) { try { $xlBook.Close($false) } catch { } }
    if ($xlApp)  {
        try { $xlApp.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($xlApp) | Out-Null } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ══════════════════════════════════════════════════════════════════
#  REGION 7 — EXPORT: SVG FOREST TOPOLOGY MAP
# ══════════════════════════════════════════════════════════════════
Write-Log "Generating SVG AD Forest Topology Map..."
Write-Progress -Activity "AD Inventory" -Status "Generating SVG Map" -PercentComplete 93

# Build tree structure: Forest > Domains > Top-level OUs > Sub-OUs (up to MaxOUDepth)
$svgNodes   = [System.Collections.Generic.List[hashtable]]::new()
$svgEdges   = [System.Collections.Generic.List[hashtable]]::new()
$nodeId     = 0

function New-SvgNode {
    param([string]$Label, [string]$SubLabel, [string]$BgColor, [string]$TxtColor, [string]$Icon, [int]$Level, [string]$ParentId)
    $script:nodeId++
    $id = "N$($script:nodeId)"
    $script:svgNodes.Add(@{ Id=$id; Label=$label; SubLabel=$subLabel; BgColor=$BgColor; TxtColor=$TxtColor; Icon=$Icon; Level=$Level; ParentId=$ParentId; X=0; Y=0 })
    if ($ParentId) { $script:svgEdges.Add(@{ From=$ParentId; To=$id }) }
    return $id
}

# Forest node
$forestId = New-SvgNode -Label $forest.Name -SubLabel "Forest | FL: $forestMode" -BgColor '#B7950B' -TxtColor '#FFFFFF' -Icon 'FRST' -Level 0 -ParentId ''

# Domain nodes
foreach ($fd in $forestDomains) {
    $dcCnt = ($dcData | Where-Object { $_.HostName -match $fd.DomainName }).Count
    $domId = New-SvgNode -Label $fd.DomainName -SubLabel "DCs: $dcCnt | $($fd.DomainMode)" -BgColor '#1A5276' -TxtColor '#FFFFFF' -Icon 'DOM' -Level 1 -ParentId $forestId

    # Top-level OUs under this domain
    $topOUs = $ouInventory | Where-Object {
        $_.Depth -eq 1 -and $_.DistinguishedName -match [regex]::Escape($fd.DN)
    }

    foreach ($ou in $topOUs | Sort-Object ContainerName) {
        $purpose = Get-OUPurpose -OUName $ou.ContainerName
        $statsLabel = "U:$($ou.UserCount) C:$($ou.ComputerCount) G:$($ou.GroupCount)"
        $ouId = New-SvgNode -Label $ou.ContainerName -SubLabel "$($ou.Purpose) | $statsLabel" `
                  -BgColor $purpose.Color -TxtColor $purpose.TxtColor -Icon $purpose.Icon -Level 2 -ParentId $domId

        if ($MaxOUDepth -ge 3) {
            $subOUs = $ouInventory | Where-Object {
                $_.Depth -eq 2 -and $_.DistinguishedName -match [regex]::Escape($ou.DistinguishedName)
            }
            foreach ($sub in $subOUs | Sort-Object ContainerName | Select-Object -First 8) {
                $subPurpose = Get-OUPurpose -OUName $sub.ContainerName
                $subStats   = "U:$($sub.UserCount) C:$($sub.ComputerCount)"
                New-SvgNode -Label $sub.ContainerName -SubLabel "$($sub.Purpose) | $subStats" `
                    -BgColor $subPurpose.Color -TxtColor $subPurpose.TxtColor -Icon $subPurpose.Icon -Level 3 -ParentId $ouId | Out-Null
            }
        }
    }

    # Built-in CN containers
    foreach ($cn in ($ouInventory | Where-Object { $_.ContainerType -eq 'CN Container' -and $_.DistinguishedName -match [regex]::Escape($fd.DN) } | Sort-Object ContainerName)) {
        $p = Get-OUPurpose -OUName $cn.ContainerName
        New-SvgNode -Label $cn.ContainerName -SubLabel "$($cn.Purpose) | U:$($cn.UserCount) G:$($cn.GroupCount)" `
            -BgColor $p.Color -TxtColor $p.TxtColor -Icon $p.Icon -Level 2 -ParentId $domId | Out-Null
    }
}

# Layout algorithm: group nodes by level, distribute evenly
$nodeW = 220; $nodeH = 70; $hGap = 30; $vGap = 90
$byLevel = @{}
foreach ($n in $svgNodes) {
    $l = $n.Level
    if (-not $byLevel.ContainsKey($l)) { $byLevel[$l] = [System.Collections.Generic.List[hashtable]]::new() }
    $byLevel[$l].Add($n)
}

$maxLevel  = ($byLevel.Keys | Measure-Object -Maximum).Maximum
$maxWidth  = 0
foreach ($l in $byLevel.Keys) { if ($byLevel[$l].Count -gt $maxWidth) { $maxWidth = $byLevel[$l].Count } }

$svgWidth  = [Math]::Max(1200, $maxWidth * ($nodeW + $hGap) + 100)
$svgHeight = ($maxLevel + 1) * ($nodeH + $vGap) + 150

foreach ($l in ($byLevel.Keys | Sort-Object)) {
    $nodes = $byLevel[$l]
    $totalW = $nodes.Count * $nodeW + ($nodes.Count - 1) * $hGap
    $startX = ($svgWidth - $totalW) / 2
    for ($i = 0; $i -lt $nodes.Count; $i++) {
        $nodes[$i].X = $startX + $i * ($nodeW + $hGap)
        $nodes[$i].Y = 60 + $l * ($nodeH + $vGap)
    }
}

# Build node lookup for edge drawing
$nodeLookup = @{}
foreach ($n in $svgNodes) { $nodeLookup[$n.Id] = $n }

# Render SVG
$svgLines = [System.Text.StringBuilder]::new()

$svgLines.AppendLine("<?xml version=`"1.0`" encoding=`"UTF-8`"?>") | Out-Null
$svgLines.AppendLine("<svg xmlns=`"http://www.w3.org/2000/svg`" width=`"$svgWidth`" height=`"$svgHeight`" viewBox=`"0 0 $svgWidth $svgHeight`" style=`"font-family:Segoe UI,Arial,sans-serif;background:#0F1923;`">") | Out-Null
$svgLines.AppendLine("  <defs>") | Out-Null
$svgLines.AppendLine("    <filter id=`"shadow`"><feDropShadow dx=`"2`" dy=`"2`" stdDeviation=`"3`" flood-color=`"#000`" flood-opacity=`"0.5`"/></filter>") | Out-Null
$svgLines.AppendLine("    <marker id=`"arrow`" viewBox=`"0 0 10 10`" refX=`"10`" refY=`"5`" markerWidth=`"6`" markerHeight=`"6`" orient=`"auto`"><path d=`"M 0 0 L 10 5 L 0 10 z`" fill=`"#4A90D9`"/></marker>") | Out-Null
$svgLines.AppendLine("  </defs>") | Out-Null

# Background grid
$svgLines.AppendLine("  <rect width=`"$svgWidth`" height=`"$svgHeight`" fill=`"#0F1923`"/>") | Out-Null
$svgLines.AppendLine("  <text x=`"20`" y=`"35`" font-size=`"22`" font-weight=`"bold`" fill=`"#F0B429`">AD Forest Topology Map — $($forest.Name)</text>") | Out-Null
$svgLines.AppendLine("  <text x=`"20`" y=`"52`" font-size=`"12`" fill=`"#7FB3D3`">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') | $Author</text>") | Out-Null

# Edges (draw first so they appear behind nodes)
foreach ($edge in $svgEdges) {
    $fromNode = $nodeLookup[$edge.From]
    $toNode   = $nodeLookup[$edge.To]
    if (-not $fromNode -or -not $toNode) { continue }
    $x1 = $fromNode.X + $nodeW / 2
    $y1 = $fromNode.Y + $nodeH
    $x2 = $toNode.X + $nodeW / 2
    $y2 = $toNode.Y
    $cx = $x1; $cy = $y1 + ($y2 - $y1) / 2
    $svgLines.AppendLine("  <path d=`"M $x1 $y1 C $cx $cy $x2 $cy $x2 $y2`" fill=`"none`" stroke=`"#4A90D9`" stroke-width=`"1.5`" opacity=`"0.6`" marker-end=`"url(#arrow)`"/>") | Out-Null
}

# Nodes
foreach ($n in $svgNodes) {
    $rx = $n.X; $ry = $n.Y
    $cx = $rx + $nodeW / 2

    # Node box
    $svgLines.AppendLine("  <rect x=`"$rx`" y=`"$ry`" width=`"$nodeW`" height=`"$nodeH`" rx=`"8`" ry=`"8`" fill=`"$($n.BgColor)`" stroke=`"#ffffff22`" stroke-width=`"1`" filter=`"url(#shadow)`"/>") | Out-Null

    # Icon badge
    $svgLines.AppendLine("  <rect x=`"$($rx + 6)`" y=`"$($ry + 6)`" width=`"34`" height=`"24`" rx=`"4`" fill=`"#ffffff22`"/>") | Out-Null
    $svgLines.AppendLine("  <text x=`"$($rx + 23)`" y=`"$($ry + 23)`" font-size=`"9`" font-weight=`"bold`" fill=`"#FFFFFF`" text-anchor=`"middle`">$([System.Web.HttpUtility]::HtmlEncode($n.Icon))</text>") | Out-Null

    # Label
    $labelText = if ($n.Label.Length -gt 24) { $n.Label.Substring(0, 21) + '...' } else { $n.Label }
    $svgLines.AppendLine("  <text x=`"$($rx + 46)`" y=`"$($ry + 22)`" font-size=`"12`" font-weight=`"bold`" fill=`"$($n.TxtColor)`">$([System.Web.HttpUtility]::HtmlEncode($labelText))</text>") | Out-Null

    # Sub-label
    $subText = if ($n.SubLabel.Length -gt 30) { $n.SubLabel.Substring(0, 27) + '...' } else { $n.SubLabel }
    $svgLines.AppendLine("  <text x=`"$($rx + 46)`" y=`"$($ry + 38)`" font-size=`"9`" fill=`"$($n.TxtColor)`" opacity=`"0.82`">$([System.Web.HttpUtility]::HtmlEncode($subText))</text>") | Out-Null

    # Level indicator dots
    for ($d = 0; $d -le $n.Level -and $d -lt 5; $d++) {
        $svgLines.AppendLine("  <circle cx=`"$($rx + 8 + $d * 10)`" cy=`"$($ry + $nodeH - 8)`" r=`"3`" fill=`"#ffffff55`"/>") | Out-Null
    }
}

# Legend
$legendX = 20; $legendY = $svgHeight - 120
$legendItems = @(
    @{Color='#B7950B'; Label='Forest'},  @{Color='#1A5276'; Label='Domain'},
    @{Color='#C0392B'; Label='Domain Controllers'}, @{Color='#6C3483'; Label='Tier 0'},
    @{Color='#21618C'; Label='Tier 1 Servers'}, @{Color='#196F3D'; Label='Tier 2 Workstations'},
    @{Color='#117A65'; Label='Users'}, @{Color='#9A7D0A'; Label='Groups'},
    @{Color='#B7770D'; Label='Service Accounts'}, @{Color='#717D7E'; Label='Disabled/Archived'}
)
$svgLines.AppendLine("  <rect x=`"$legendX`" y=`"$legendY`" width=`"$($svgWidth - 40)`" height=`"100`" rx=`"6`" fill=`"#ffffff0A`" stroke=`"#ffffff15`"/>") | Out-Null
$svgLines.AppendLine("  <text x=`"$($legendX+10)`" y=`"$($legendY+18)`" font-size=`"11`" font-weight=`"bold`" fill=`"#F0B429`">Legend</text>") | Out-Null
for ($li = 0; $li -lt $legendItems.Count; $li++) {
    $lx = $legendX + 10 + ($li % 5) * 200
    $ly = $legendY + 35 + [Math]::Floor($li / 5) * 22
    $svgLines.AppendLine("  <rect x=`"$lx`" y=`"$($ly-10)`" width=`"14`" height=`"14`" rx=`"3`" fill=`"$($legendItems[$li].Color)`"/>") | Out-Null
    $svgLines.AppendLine("  <text x=`"$($lx+18)`" y=`"$ly`" font-size=`"10`" fill=`"#CCCCCC`">$($legendItems[$li].Label)</text>") | Out-Null
}

$svgLines.AppendLine("</svg>") | Out-Null

$svgPath = Join-Path $ExportFolder "ADForestMap_${ShortDomain}_${Timestamp}.svg"
try {
    [System.IO.File]::WriteAllText($svgPath, $svgLines.ToString(), [System.Text.Encoding]::UTF8)
    Write-Log "SVG map saved: $svgPath" -Level SUCCESS
} catch {
    Write-Log "SVG map write error: $_" -Level WARN
}

# ══════════════════════════════════════════════════════════════════
#  REGION 8 — EXPORT: VISIO DIAGRAM (COM — requires Visio)
# ══════════════════════════════════════════════════════════════════
Write-Log "Attempting Visio diagram export via COM..."
Write-Progress -Activity "AD Inventory" -Status "Attempting Visio Export" -PercentComplete 95

$visioPath = Join-Path $ExportFolder "ADForestMap_${ShortDomain}_${Timestamp}.vsdx"
$visioApp  = $null

try {
    $visioApp = New-Object -ComObject Visio.Application -ErrorAction Stop
    $visioApp.Visible = $false

    $visioDoc = $visioApp.Documents.Add('')
    $visioPage = $visioDoc.Pages.Item(1)
    $visioPage.Name = "AD Forest - $ShortDomain"
    $visioPage.PageSheet.CellsSRC(1, 43, 0).FormulaU = '14 in'  # Page width
    $visioPage.PageSheet.CellsSRC(1, 44, 0).FormulaU = '11 in'  # Page height

    # Shape dimensions
    $boxW   = 2.0  # inches
    $boxH   = 0.6
    $hSpace = 2.4
    $vSpace = 1.4

    # Track shapes by node Id for connection drawing
    $visioShapes = @{}

    foreach ($n in $svgNodes) {
        $xIn = ($n.X / 96) + 0.5  # Convert pixels to inches (96dpi assumption)
        $yIn = 10.0 - ($n.Y / 96) # Visio Y is bottom-up

        # Drop rectangle shape
        $shape = $visioPage.DrawRectangle($xIn, $yIn, $xIn + $boxW, $yIn - $boxH)
        $shape.Text = "$($n.Label)`n$($n.SubLabel)"

        # Color fill (convert hex to RGB then to Visio BGR long)
        $hex   = $n.BgColor.TrimStart('#')
        $r     = [Convert]::ToInt32($hex.Substring(0,2),16)
        $g     = [Convert]::ToInt32($hex.Substring(2,2),16)
        $b     = [Convert]::ToInt32($hex.Substring(4,2),16)
        $oleColor = $b * 65536 + $g * 256 + $r

        $shape.CellsSRC(1, 1, 0).FormulaU = "RGB($r,$g,$b)"   # FillForegnd
        $shape.CellsSRC(2, 0, 0).FormulaU = 'RGB(255,255,255)' # Line color
        $shape.CellsSRC(9, 0, 0).FormulaU = 'RGB(255,255,255)' # Char color
        $shape.CellsSRC(9, 2, 0).FormulaU = '10 pt'            # Char size

        $visioShapes[$n.Id] = $shape
    }

    # Draw connectors
    foreach ($edge in $svgEdges) {
        if ($visioShapes.ContainsKey($edge.From) -and $visioShapes.ContainsKey($edge.To)) {
            try {
                $conn = $visioPage.Drop($visioApp.ConnectorToolDataObject, 0, 0)
                $conn.CellsSRC(0, 0, 5).FormulaU = 'RGB(74,144,217)'  # Line color
                $conn.CellsSRC(0, 0, 6).FormulaU = '1.5 pt'           # Line weight
                $conn.Cells('BeginX').GlueToPos($visioShapes[$edge.From], 0.5, 0)
                $conn.Cells('EndX').GlueToPos($visioShapes[$edge.To], 0.5, 1)
            } catch { }
        }
    }

    # Title
    $titleShape = $visioPage.DrawRectangle(0.2, 10.7, 13.8, 10.2)
    $titleShape.Text = "AD Forest Topology — $($forest.Name)  |  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')  |  $Author"
    $titleShape.CellsSRC(1, 1, 0).FormulaU = 'RGB(11,19,35)'
    $titleShape.CellsSRC(2, 0, 0).FormulaU = 'RGB(240,180,41)'
    $titleShape.CellsSRC(9, 0, 0).FormulaU = 'RGB(240,180,41)'
    $titleShape.CellsSRC(9, 2, 0).FormulaU = '14 pt'

    $visioDoc.SaveAs($visioPath)
    Write-Log "Visio diagram saved: $visioPath" -Level SUCCESS

} catch {
    Write-Log "Visio not available or COM error (this is expected if Visio is not installed): $_" -Level WARN
    Write-Log "SVG map is the topology map output for this run." -Level INFO
} finally {
    if ($visioApp) {
        try { $visioApp.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($visioApp) | Out-Null } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ══════════════════════════════════════════════════════════════════
#  REGION 9 — EXPORT: HTML DASHBOARD
# ══════════════════════════════════════════════════════════════════
Write-Log "Building HTML dashboard..."
Write-Progress -Activity "AD Inventory" -Status "Building HTML Report" -PercentComplete 97

$reportDate = Get-Date -Format 'dddd, MMMM dd yyyy  HH:mm:ss'

# ── HTML helper: build a table from PSCustomObjects ──────────────
function Build-HtmlTable {
    param([object[]]$Data, [string]$Id, [string]$StatusProp = '', [hashtable]$StatusMap = @{})
    if (-not $Data -or $Data.Count -eq 0) { return '<p class="empty">No data collected.</p>' }
    $headers = @($Data[0].PSObject.Properties.Name | Where-Object { $_ -notmatch 'Color$|Class$|Icon$|TxtColor$|PurposeColor$|PurposeIcon$|StatusClass$|TierClass$|SevClass$|ReplOK$|ScavEnabled$' })
    $sb = [System.Text.StringBuilder]::new()
    $sb.Append("<table class=`"data-table`" id=`"$Id`"><thead><tr>") | Out-Null
    foreach ($h in $headers) { $sb.Append("<th>$(Escape-Html $h)</th>") | Out-Null }
    $sb.Append("</tr></thead><tbody>") | Out-Null
    foreach ($row in $Data) {
        # Determine row class
        $rowClass = ''
        if ($StatusProp -and $row.PSObject.Properties[$StatusProp]) {
            $sv = $row.$StatusProp
            if ($StatusMap.ContainsKey($sv)) { $rowClass = " class=`"row-$($StatusMap[$sv])`"" }
        }
        $sb.Append("<tr$rowClass>") | Out-Null
        foreach ($h in $headers) {
            $val = $row.$h
            if ($null -eq $val) { $val = '' }
            $cell = Escape-Html $val.ToString()
            # Badge certain columns
            if ($h -eq 'Severity' -or $h -eq 'Status' -or $h -eq 'Risk' -or $h -eq 'Tier' -or $h -eq 'Flag') {
                $cls = switch ($val.ToString().ToUpper()) {
                    'CRITICAL'  { 'badge-crit' }
                    'HIGH'      { 'badge-crit' }
                    'WARNING'   { 'badge-warn' }
                    'MEDIUM'    { 'badge-warn' }
                    'TIER 0'    { 'badge-t0'   }
                    'TIER 1'    { 'badge-t1'   }
                    'OK'        { 'badge-ok'   }
                    'HEALTHY'   { 'badge-ok'   }
                    'SUCCESS'   { 'badge-ok'   }
                    'ENABLED'   { 'badge-ok'   }
                    default     { 'badge-info' }
                }
                $cell = "<span class=`"badge $cls`">$cell</span>"
            }
            $sb.Append("<td>$cell</td>") | Out-Null
        }
        $sb.Append("</tr>") | Out-Null
    }
    $sb.Append("</tbody></table>") | Out-Null
    return $sb.ToString()
}

# ── OU Inventory with purpose color tiles ───────────────────────
function Build-OUInventoryHtml {
    $sb = [System.Text.StringBuilder]::new()
    $sb.Append('<div class="ou-grid">') | Out-Null
    foreach ($ou in $ouInventory | Sort-Object Depth, ContainerName) {
        $indentPx = $ou.Depth * 24
        $sb.Append("<div class=`"ou-card`" style=`"border-left: 5px solid $($ou.PurposeColor); margin-left:${indentPx}px`">") | Out-Null
        $sb.Append("<div class=`"ou-header`" style=`"background:$($ou.PurposeColor)`">") | Out-Null
        $sb.Append("<span class=`"ou-icon`">$($ou.PurposeIcon)</span>") | Out-Null
        $sb.Append("<span class=`"ou-name`">$(Escape-Html $ou.ContainerName)</span>") | Out-Null
        $sb.Append("<span class=`"ou-type`">$(Escape-Html $ou.Purpose)</span>") | Out-Null
        $sb.Append("</div>") | Out-Null
        $sb.Append("<div class=`"ou-body`">") | Out-Null
        if ($ou.Description) { $sb.Append("<p class=`"ou-desc`">$(Escape-Html $ou.Description)</p>") | Out-Null }
        $sb.Append("<div class=`"ou-stats`">") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.UserCount)</span><span class=`"stat-lbl`">Users</span></div>") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.ServerCount)</span><span class=`"stat-lbl`">Servers</span></div>") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.WorkstationCount)</span><span class=`"stat-lbl`">Workstations</span></div>") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.DCCount)</span><span class=`"stat-lbl`">DCs</span></div>") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.GroupCount)</span><span class=`"stat-lbl`">Groups</span></div>") | Out-Null
        $sb.Append("<div class=`"ou-stat`"><span class=`"stat-val`">$($ou.GPOLinksCount)</span><span class=`"stat-lbl`">GPOs</span></div>") | Out-Null
        $sb.Append("</div>") | Out-Null
        $sb.Append("<p class=`"ou-path`">$(Escape-Html $ou.RelativePath)</p>") | Out-Null
        $sb.Append("</div></div>") | Out-Null
    }
    $sb.Append('</div>') | Out-Null
    return $sb.ToString()
}

# Build table HTML
$dcTableHtml          = Build-HtmlTable -Data $dcData           -Id 'tbl-dc'    -StatusProp 'ReplStatus' -StatusMap @{ 'OK'='ok'; 'ERR'='crit'; 'N/A'='warn' }
$replTableHtml        = Build-HtmlTable -Data $replData         -Id 'tbl-repl'  -StatusProp 'Status'     -StatusMap @{ 'Success'='ok' }
$sitesTableHtml       = Build-HtmlTable -Data $sitesData        -Id 'tbl-sites'
$ouHtml               = Build-OUInventoryHtml
$userTableHtml        = Build-HtmlTable -Data $userTableData    -Id 'tbl-users' -StatusProp 'Risk'        -StatusMap @{ 'HIGH'='crit'; 'MEDIUM'='warn'; 'LOW'='ok' }
$computerTableHtml    = Build-HtmlTable -Data $computerData     -Id 'tbl-comp'
$groupTableHtml       = Build-HtmlTable -Data $groupData        -Id 'tbl-grp'
$psoTableHtml         = Build-HtmlTable -Data $psoData          -Id 'tbl-pso'
$gpoTableHtml         = Build-HtmlTable -Data $gpoData          -Id 'tbl-gpo'  -StatusProp 'Status'      -StatusMap @{ 'UNLINKED'='warn'; 'ENFORCED'='ok'; 'ENABLED'='ok' }
$dnsTableHtml         = Build-HtmlTable -Data $dnsData          -Id 'tbl-dns'  -StatusProp 'Status'      -StatusMap @{ 'WARN'='warn'; 'HEALTHY'='ok' }
$dhcpTableHtml        = Build-HtmlTable -Data $dhcpData         -Id 'tbl-dhcp' -StatusProp 'Status'      -StatusMap @{ 'HIGH USE'='crit'; 'WARNING'='warn'; 'HEALTHY'='ok' }
$pkiTableHtml         = Build-HtmlTable -Data $pkiData          -Id 'tbl-pki'  -StatusProp 'Status'      -StatusMap @{ 'CRITICAL'='crit'; 'EXPIRING'='warn'; 'HEALTHY'='ok'; 'EXPIRED'='crit' }
$privTableHtml        = Build-HtmlTable -Data $privGroupData    -Id 'tbl-priv'
$secTableHtml         = Build-HtmlTable -Data $secAlerts        -Id 'tbl-sec'  -StatusProp 'Severity'    -StatusMap @{ 'CRITICAL'='crit'; 'WARNING'='warn'; 'INFO'='ok' }
$forestDomainHtml     = Build-HtmlTable -Data $forestDomains    -Id 'tbl-dom'

# Summary KPIs
$totalOUs       = $ouInventory.Count
$totalComputers = $computerData.Count
$totalDCs       = ($computerData | Where-Object { $_.Role -eq 'Domain Controller' }).Count
$totalServers   = ($computerData | Where-Object { $_.Role -eq 'Member Server'    }).Count
$totalWKS       = ($computerData | Where-Object { $_.Role -eq 'Workstation'      }).Count

# Helper: build an HTML section
function Build-Section {
    param([string]$Id, [string]$Title, [string]$Icon, [string]$Badge, [string]$BadgeClass, [string]$Body, [string]$ExportId)
    return @"
<div class="section" id="sec-$Id">
  <div class="section-header" onclick="toggleSection('sec-$Id')">
    <div class="section-icon">$Icon</div>
    <span class="section-title">$Title</span>
    <span class="badge $BadgeClass" style="margin-left:10px">$Badge</span>
    <svg class="chevron" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="section-body">$Body</div>
</div>
"@
}

$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AD Inventory — $DomainFQDN</title>
<style>
  :root {
    --bg:        #0F1923;
    --panel:     #162232;
    --border:    #1E3A52;
    --accent:    #1B6CA8;
    --accent2:   #F0B429;
    --text:      #D0E4F4;
    --text-dim:  #7FB3D3;
    --crit:      #E53E3E;
    --warn:      #DD6B20;
    --ok:        #38A169;
    --info:      #4A90D9;
    --t0:        #9B59B6;
    --t1:        #2471A3;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; }
  a { color: var(--accent2); }

  /* ─ TOP BAR ─ */
  .topbar { background: #0A1118; padding: 14px 24px; display: flex; align-items: center; gap: 16px;
            border-bottom: 2px solid var(--accent); position: sticky; top: 0; z-index: 100; }
  .topbar-title { font-size: 20px; font-weight: 700; color: var(--accent2); white-space: nowrap; }
  .topbar-domain { font-size: 13px; color: var(--text-dim); }
  .topbar-spacer { flex: 1; }
  #searchBox { background: #1A2E42; border: 1px solid var(--border); color: var(--text); padding: 7px 12px;
               border-radius: 6px; font-size: 13px; width: 280px; }
  #searchBox:focus { outline: none; border-color: var(--accent); }
  #searchCount { color: var(--text-dim); font-size: 12px; min-width: 100px; text-align: right; }
  .btn-clear { background: transparent; border: none; color: var(--text-dim); cursor: pointer; font-size: 16px; padding: 4px 8px; }

  /* ─ KPI STRIP ─ */
  .kpi-strip { display: flex; flex-wrap: wrap; gap: 12px; padding: 18px 24px; background: var(--panel);
               border-bottom: 1px solid var(--border); }
  .kpi-card { background: var(--bg); border: 1px solid var(--border); border-radius: 10px;
              padding: 12px 20px; min-width: 130px; text-align: center; flex: 1; }
  .kpi-val { font-size: 28px; font-weight: 700; color: var(--accent2); line-height: 1; }
  .kpi-lbl { font-size: 11px; color: var(--text-dim); margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi-card.red   .kpi-val { color: var(--crit); }
  .kpi-card.amber .kpi-val { color: var(--warn); }
  .kpi-card.green .kpi-val { color: var(--ok);   }

  /* ─ MAIN LAYOUT ─ */
  main { max-width: 1600px; margin: 0 auto; padding: 20px 24px 60px; }

  /* ─ SECTIONS ─ */
  .section { background: var(--panel); border: 1px solid var(--border); border-radius: 10px;
             margin-bottom: 14px; overflow: hidden; }
  .section-header { display: flex; align-items: center; gap: 12px; padding: 14px 18px;
                    cursor: pointer; user-select: none; transition: background .2s; }
  .section-header:hover { background: #1E3A52; }
  .section-icon { font-size: 18px; width: 28px; text-align: center; }
  .section-title { font-size: 15px; font-weight: 600; flex: 1; }
  .chevron { transition: transform .25s; margin-left: auto; color: var(--text-dim); }
  .section.open .chevron { transform: rotate(180deg); }
  .section-body { display: none; padding: 16px 18px; border-top: 1px solid var(--border); overflow-x: auto; }
  .section.open .section-body { display: block; }

  /* ─ BADGES ─ */
  .badge { display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 11px; font-weight: 600; letter-spacing: .3px; }
  .badge-crit { background: var(--crit);  color: #fff; }
  .badge-warn { background: var(--warn);  color: #fff; }
  .badge-ok   { background: var(--ok);    color: #fff; }
  .badge-t0   { background: var(--t0);    color: #fff; }
  .badge-t1   { background: var(--t1);    color: #fff; }
  .badge-info { background: var(--info);  color: #fff; }

  /* ─ DATA TABLE ─ */
  .data-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 600px; }
  .data-table th { background: #1B4F9E; color: #fff; padding: 9px 12px; text-align: left;
                   font-size: 12px; letter-spacing: .3px; position: sticky; top: 0; }
  .data-table td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
  .data-table tbody tr:nth-child(even) { background: #122030; }
  .data-table tbody tr:hover { background: #1E3A52; }
  .row-crit td { border-left: 3px solid var(--crit); }
  .row-warn td { border-left: 3px solid var(--warn); }
  .row-ok td   { border-left: 3px solid var(--ok);   }
  .search-match td { background: #1C3A1C !important; }
  .empty { color: var(--text-dim); font-style: italic; padding: 10px 0; }

  /* ─ OU CARDS ─ */
  .ou-grid { display: flex; flex-direction: column; gap: 8px; }
  .ou-card { background: var(--bg); border-radius: 8px; overflow: hidden; border: 1px solid var(--border); }
  .ou-header { display: flex; align-items: center; gap: 10px; padding: 8px 14px; }
  .ou-icon { background: rgba(255,255,255,0.15); border-radius: 4px; padding: 2px 7px;
             font-size: 10px; font-weight: 700; color: #fff; font-family: monospace; }
  .ou-name { font-weight: 600; color: #fff; font-size: 14px; flex: 1; }
  .ou-type { font-size: 11px; color: rgba(255,255,255,0.7); margin-left: auto; }
  .ou-body { padding: 10px 14px; }
  .ou-desc { font-size: 12px; color: var(--text-dim); margin-bottom: 8px; font-style: italic; }
  .ou-stats { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 6px; }
  .ou-stat { display: flex; flex-direction: column; align-items: center; background: var(--panel);
             border: 1px solid var(--border); border-radius: 6px; padding: 4px 14px; min-width: 60px; }
  .stat-val { font-size: 18px; font-weight: 700; color: var(--accent2); }
  .stat-lbl { font-size: 10px; color: var(--text-dim); text-transform: uppercase; }
  .ou-path { font-size: 10px; color: var(--text-dim); font-family: monospace; margin-top: 4px; }

  /* ─ FOOTER ─ */
  footer { text-align: center; padding: 22px; color: var(--text-dim); font-size: 12px;
           border-top: 1px solid var(--border); background: #0A1118; }
</style>
</head>
<body>

<div class="topbar">
  <div>
    <div class="topbar-title">🏛 AD Forest Inventory Dashboard</div>
    <div class="topbar-domain">$DomainFQDN &nbsp;|&nbsp; Forest: $($forest.Name) &nbsp;|&nbsp; $reportDate</div>
  </div>
  <div class="topbar-spacer"></div>
  <button class="btn-clear" id="btnClear" title="Clear search">✕</button>
  <input id="searchBox" type="text" placeholder="🔍  Search all sections…">
  <span id="searchCount"></span>
</div>

<div class="kpi-strip">
  <div class="kpi-card"><div class="kpi-val">$($forest.Domains.Count)</div><div class="kpi-lbl">Domains in Forest</div></div>
  <div class="kpi-card"><div class="kpi-val">$($dcData.Count)</div><div class="kpi-lbl">Domain Controllers</div></div>
  <div class="kpi-card"><div class="kpi-val">$totalOUs</div><div class="kpi-lbl">OUs & Containers</div></div>
  <div class="kpi-card"><div class="kpi-val">$totalEnabled</div><div class="kpi-lbl">Enabled Users</div></div>
  <div class="kpi-card"><div class="kpi-val">$totalDisabled</div><div class="kpi-lbl">Disabled Users</div></div>
  <div class="kpi-card"><div class="kpi-val">$totalServers</div><div class="kpi-lbl">Member Servers</div></div>
  <div class="kpi-card"><div class="kpi-val">$totalWKS</div><div class="kpi-lbl">Workstations</div></div>
  <div class="kpi-card"><div class="kpi-val">$($allGroupsRaw.Count)</div><div class="kpi-lbl">Security Groups</div></div>
  <div class="kpi-card"><div class="kpi-val">$($gpoTotal)</div><div class="kpi-lbl">Group Policies</div></div>
  <div class="kpi-card $(if($critAlerts -gt 0){'red'}elseif($warnAlerts -gt 0){'amber'}else{'green'})">
    <div class="kpi-val">$($secAlerts.Count)</div><div class="kpi-lbl">Security Alerts</div>
  </div>
</div>

<main>

$(Build-Section -Id 'forest'   -Title 'Forest & Domains'            -Icon '🌐' -Badge "$($forestDomains.Count) domain(s)" -BadgeClass 'badge-info' -Body $forestDomainHtml)
$(Build-Section -Id 'dc'       -Title 'Domain Controllers & FSMO'   -Icon '🖥' -Badge "$($dcData.Count) DCs"               -BadgeClass 'badge-info' -Body $dcTableHtml)
$(Build-Section -Id 'sites'    -Title 'AD Sites'                     -Icon '🗺' -Badge "$($sitesData.Count) site(s)"        -BadgeClass 'badge-info' -Body $sitesTableHtml)
$(Build-Section -Id 'repl'     -Title 'Replication Connections'      -Icon '🔄' -Badge "$($replData.Count) connections"     -BadgeClass "$(if(($replData|Where-Object{$_.Failures -gt 0}).Count -gt 0){'badge-crit'}else{'badge-ok'})" -Body $replTableHtml)
$(Build-Section -Id 'ou'       -Title 'OU & Container Inventory'     -Icon '📁' -Badge "$totalOUs containers classified"   -BadgeClass 'badge-info' -Body $ouHtml)
$(Build-Section -Id 'users'    -Title 'Flagged User Accounts'        -Icon '👤' -Badge "$($userTableData.Count) flagged of $($allUsersRaw.Count) total" -BadgeClass "$(if(($userTableData|Where-Object{$_.Risk -eq 'HIGH'}).Count -gt 0){'badge-warn'}else{'badge-ok'})" -Body $userTableHtml)
$(Build-Section -Id 'comp'     -Title 'Computer Inventory'           -Icon '💻' -Badge "$totalComputers total ($totalDCs DC · $totalServers SVR · $totalWKS WKS)" -BadgeClass 'badge-info' -Body $computerTableHtml)
$(Build-Section -Id 'groups'   -Title 'Security Group Inventory'     -Icon '👥' -Badge "$($groupData.Count) groups"         -BadgeClass 'badge-info' -Body $groupTableHtml)
$(Build-Section -Id 'priv'     -Title 'Privileged Group Memberships' -Icon '🔐' -Badge "$($privGroupData.Count) groups audited" -BadgeClass "$(if(($privGroupData|Where-Object{$_.ReviewStatus -ne 'OK'}).Count -gt 0){'badge-warn'}else{'badge-ok'})" -Body $privTableHtml)
$(Build-Section -Id 'pso'      -Title 'Password Policies'            -Icon '🔑' -Badge "$($psoData.Count) polic(ies)"       -BadgeClass 'badge-info' -Body $psoTableHtml)
$(Build-Section -Id 'gpo'      -Title 'Group Policy Objects'         -Icon '📋' -Badge "$gpoTotal GPOs ($gpoUnlinked unlinked)" -BadgeClass "$(if($gpoUnlinked -gt 0){'badge-warn'}else{'badge-ok'})" -Body $gpoTableHtml)
$(Build-Section -Id 'dns'      -Title 'DNS Zones'                    -Icon '🌍' -Badge "$($dnsData.Count) zones"             -BadgeClass 'badge-info' -Body $dnsTableHtml)
$(Build-Section -Id 'dhcp'     -Title 'DHCP Scopes'                  -Icon '📡' -Badge "$($dhcpData.Count) scope(s)"        -BadgeClass 'badge-info' -Body $dhcpTableHtml)
$(Build-Section -Id 'pki'      -Title 'PKI & Certificates'           -Icon '🏅' -Badge "$($pkiData.Count) certificates"     -BadgeClass "$(if(($pkiData|Where-Object{$_.DaysLeft -lt 30}).Count -gt 0){'badge-warn'}else{'badge-ok'})" -Body $pkiTableHtml)
$(Build-Section -Id 'security' -Title 'Security Alerts & Findings'   -Icon '⚠' -Badge "$critAlerts CRITICAL · $warnAlerts WARNING" -BadgeClass "$(if($critAlerts -gt 0){'badge-crit'}elseif($warnAlerts -gt 0){'badge-warn'}else{'badge-ok'})" -Body $secTableHtml)

</main>

<footer>
  AD Forest Inventory v3.0 &nbsp;|&nbsp; $DomainFQDN &nbsp;|&nbsp; $Author &nbsp;|&nbsp; Generated: $reportDate
</footer>

<script>
// Section toggle
function toggleSection(id) {
  document.getElementById(id).classList.toggle('open');
}

// Expand all sections with alerts on load
document.querySelectorAll('.section').forEach(function(s) {
  var badge = s.querySelector('.badge');
  if (badge && (badge.classList.contains('badge-crit') || badge.classList.contains('badge-warn'))) {
    s.classList.add('open');
  }
});

// Live search across all tables
var searchBox   = document.getElementById('searchBox');
var searchCount = document.getElementById('searchCount');
var btnClear    = document.getElementById('btnClear');

searchBox.addEventListener('input', runSearch);
btnClear.addEventListener('click', function() { searchBox.value = ''; runSearch(); });

function runSearch() {
  var q = searchBox.value.trim().toLowerCase();
  if (!q) { clearSearch(); return; }
  var total = 0;
  document.querySelectorAll('.section').forEach(function(sec) {
    var hit = sec.querySelector('.section-title').textContent.toLowerCase().includes(q);
    sec.querySelectorAll('tbody tr').forEach(function(tr) {
      var m = tr.textContent.toLowerCase().includes(q);
      tr.classList.toggle('search-match', m);
      if (m) { hit = true; total++; }
    });
    sec.querySelectorAll('.ou-card').forEach(function(card) {
      var m = card.textContent.toLowerCase().includes(q);
      card.style.display = m ? '' : 'none';
      if (m) { hit = true; total++; }
    });
    sec.classList.toggle('search-hidden', !hit);
    if (hit) sec.classList.add('open');
  });
  searchCount.textContent = total === 0 ? 'No matches' : (total + ' match' + (total === 1 ? '' : 'es'));
}

function clearSearch() {
  document.querySelectorAll('.section').forEach(function(s) { s.classList.remove('search-hidden'); });
  document.querySelectorAll('tbody tr').forEach(function(r) { r.classList.remove('search-match'); });
  document.querySelectorAll('.ou-card').forEach(function(c) { c.style.display = ''; });
  searchCount.textContent = '';
}
</script>
</body>
</html>
"@

$htmlPath = Join-Path $ExportFolder "ADInventory_${ShortDomain}_${Timestamp}.html"
try {
    [System.IO.File]::WriteAllText($htmlPath, $htmlReport, [System.Text.Encoding]::UTF8)
    Write-Log "HTML dashboard saved: $htmlPath" -Level SUCCESS
} catch {
    Write-Log "HTML write error: $_" -Level WARN
}

# ══════════════════════════════════════════════════════════════════
#  REGION 10 — FINAL SUMMARY
# ══════════════════════════════════════════════════════════════════
Write-Progress -Activity "AD Inventory" -Status "Complete" -PercentComplete 100

Write-Log "" -Level INFO
Write-Log "═══════════════════════════════════════════════════════" -Level SUCCESS
Write-Log "  AD FOREST INVENTORY COMPLETE" -Level SUCCESS
Write-Log "═══════════════════════════════════════════════════════" -Level INFO
Write-Log "  Forest         : $($forest.Name)  ($($forest.Domains.Count) domain(s))" -Level INFO
Write-Log "  Domain         : $DomainFQDN" -Level INFO
Write-Log "  Domain Mode    : $domainMode" -Level INFO
Write-Log "  DCs Found      : $($dcData.Count)" -Level INFO
Write-Log "  OUs/Containers : $totalOUs" -Level INFO
Write-Log "  Users          : $($allUsersRaw.Count) ($totalEnabled enabled, $totalDisabled disabled)" -Level INFO
Write-Log "  Computers      : $totalComputers ($totalDCs DC, $totalServers SVR, $totalWKS WKS)" -Level INFO
Write-Log "  Groups         : $($allGroupsRaw.Count)" -Level INFO
Write-Log "  GPOs           : $gpoTotal ($gpoUnlinked unlinked)" -Level INFO
Write-Log "  Security Alerts: $($secAlerts.Count) ($critAlerts CRITICAL, $warnAlerts WARNING)" -Level INFO
Write-Log "─────────────────────────────────────────────────────── " -Level INFO
Write-Log "  Output Folder  : $ExportFolder" -Level SUCCESS
Write-Log "  HTML Report    : $(Split-Path $htmlPath -Leaf)" -Level SUCCESS
Write-Log "  SVG Map        : $(Split-Path $svgPath  -Leaf)" -Level SUCCESS
Write-Log "  Excel Workbook : $(if (Test-Path $xlPath) { Split-Path $xlPath -Leaf } else { 'Skipped (Excel not installed)' })" -Level SUCCESS
Write-Log "  Visio Diagram  : $(if (Test-Path $visioPath) { Split-Path $visioPath -Leaf } else { 'Skipped (Visio not installed)' })" -Level SUCCESS
Write-Log "  CSVs           : $($csvSections.Count) files in output folder" -Level SUCCESS
Write-Log "  Log            : $(Split-Path $LogPath -Leaf)" -Level INFO
Write-Log "═══════════════════════════════════════════════════════" -Level SUCCESS

# Open Explorer to output folder
try { Start-Process explorer.exe -ArgumentList $ExportFolder } catch { }

if ($OpenOnComplete) {
    try { Start-Process $htmlPath } catch { Write-Log "Could not auto-open HTML report." -Level WARN }
}

# Return output folder path for pipeline use
Write-Output $ExportFolder
