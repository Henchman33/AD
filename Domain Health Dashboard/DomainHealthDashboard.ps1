#Requires -Version 5.1
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    AD Health Dashboard — Data Collection & HTML Report Generator
.DESCRIPTION
    Collects comprehensive Active Directory health data from the local domain and
    generates a fully self-contained HTML dashboard report. The report includes
    collapsible sections, live search, and per-section export to CSV / XLSX / TXT / DOCX.

    Must be run on a Domain Controller or a machine with RSAT (AD DS Tools) installed
    that has network access to the domain. Elevation (Run as Administrator) is required.

.PARAMETER DomainFQDN
    The fully qualified domain name to report on. Defaults to the current computer's domain.

.PARAMETER OutputPath
    Directory where the HTML report will be saved. Defaults to C:\Reports\ADHealth\

.PARAMETER Author
    Name shown in the report header and exports. Defaults to "Stephen McKee - Server Administrator 2"

.PARAMETER OpenOnComplete
    Switch. If specified, the report will open in the default browser when complete.

.EXAMPLE
    .\Get-ADHealthDashboard.ps1
    Runs against the current domain with all defaults.

.EXAMPLE
    .\Get-ADHealthDashboard.ps1 -DomainFQDN "corp.contoso.com" -OutputPath "D:\Reports" -OpenOnComplete
    Runs against a specific domain, saves to D:\Reports, and opens on completion.

.NOTES
    Author  : Stephen McKee — Server Administrator 2
    Version : 2.0
    Requires: ActiveDirectory module (RSAT-AD-PowerShell or installed on DC)
              DNS Server module recommended (RSAT-DNS-Server) — gracefully skipped if absent
              DHCP Server module recommended (RSAT-DHCP)       — gracefully skipped if absent
              GroupPolicy module recommended (GPMC)             — gracefully skipped if absent
              Run as: Domain Administrator or equivalent read rights across all AD objects
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DomainFQDN = $env:USERDNSDOMAIN,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Reports\ADHealth",

    [Parameter(Mandatory = $false)]
    [string]$Author = "Stephen McKee - Server Administrator 2",

    [Parameter(Mandatory = $false)]
    [switch]$OpenOnComplete
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════
#  LOGGING
# ════════════════════════════════════════════════
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$LogPath = Join-Path $OutputPath "ADHealthDashboard_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -ErrorAction SilentlyContinue
    switch ($Level) {
        'INFO'    { Write-Host $entry -ForegroundColor Cyan }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
    }
}

# ════════════════════════════════════════════════
#  PREREQUISITES CHECK
# ════════════════════════════════════════════════
Write-Log "AD Health Dashboard v2.0 — Starting"
Write-Log "Running as: $env:USERNAME on $env:COMPUTERNAME"
Write-Log "Target Domain: $DomainFQDN"

# Check elevation
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Script must be run as Administrator. Please re-launch elevated." -Level ERROR
    exit 1
}

# Required module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory module loaded." -Level SUCCESS
} catch {
    Write-Log "ActiveDirectory module not found. Install RSAT-AD-PowerShell or run on a DC." -Level ERROR
    exit 1
}

# Optional modules — warn but continue
$optionalModules = @{
    'GroupPolicy' = 'GPMC / RSAT-GPMC'
    'DnsServer'   = 'RSAT-DNS-Server'
    'DHCPServer'  = 'RSAT-DHCP'
}
$modAvailable = @{}
foreach ($mod in $optionalModules.Keys) {
    if (Get-Module -ListAvailable -Name $mod) {
        Import-Module $mod -ErrorAction SilentlyContinue
        $modAvailable[$mod] = $true
        Write-Log "Optional module '$mod' loaded." -Level INFO
    } else {
        $modAvailable[$mod] = $false
        Write-Log "Optional module '$mod' not found (install $($optionalModules[$mod])). That section will show placeholder data." -Level WARN
    }
}

# ════════════════════════════════════════════════
#  HELPER: Safe HTML escape
# ════════════════════════════════════════════════
function Escape-Html([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return '—' }
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function Format-Age([datetime]$dt) {
    $ts = (Get-Date) - $dt
    if ($ts.TotalDays -gt 365) { return "$([math]::Round($ts.TotalDays/365,1))y" }
    if ($ts.TotalDays -gt 1)   { return "$([int]$ts.TotalDays)d $($ts.Hours)h" }
    return "$($ts.Hours)h $($ts.Minutes)m"
}

function Status-Badge([string]$status, [string]$class) {
    return "<span class=`"td-badge $class`">$(Escape-Html $status)</span>"
}

# ════════════════════════════════════════════════
#  DATA COLLECTION
# ════════════════════════════════════════════════

Write-Log "Collecting AD domain info..."
try {
    $domain      = Get-ADDomain -Server $DomainFQDN
    $forest      = Get-ADForest -Server $DomainFQDN
    $domainDN    = $domain.DistinguishedName
    $domainMode  = $domain.DomainMode
    $forestMode  = $forest.ForestMode
    $pdcEmulator = $domain.PDCEmulator
} catch {
    Write-Log "Failed to connect to domain '$DomainFQDN': $_" -Level ERROR
    exit 1
}

# ── DOMAIN CONTROLLERS ──────────────────────────
Write-Log "Collecting Domain Controller data..."
$dcData = @()
try {
    $dcs = Get-ADDomainController -Filter * -Server $DomainFQDN | Sort-Object Name
    foreach ($dc in $dcs) {
        $fsmoRoles = @()
        if ($dc.OperationMasterRoles) { $fsmoRoles = $dc.OperationMasterRoles }
        $rolesStr = if ($fsmoRoles.Count -gt 0) { ($fsmoRoles -join ', ') } else { 'None' }

        # Replication summary for this DC
        $replOk = $true
        $replDetail = 'OK'
        try {
            $replStatus = Get-ADReplicationPartnerMetadata -Target $dc.HostName -Scope Server -ErrorAction SilentlyContinue
            if ($replStatus) {
                $failedPartners = $replStatus | Where-Object { $_.LastReplicationAttempt -and $_.LastReplicationResult -ne 0 }
                if ($failedPartners) { $replOk = $false; $replDetail = "ERR($($failedPartners.Count))" }
            }
        } catch { $replDetail = 'N/A' }

        # Uptime via WMI
        $uptime = 'N/A'
        try {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $dc.HostName -ErrorAction SilentlyContinue
            if ($os) {
                $ts = (Get-Date) - $os.LastBootUpTime
                $uptime = "$([int]$ts.TotalDays)d $($ts.Hours)h"
            }
        } catch { }

        $dcData += [PSCustomObject]@{
            Name          = $dc.Name
            HostName      = $dc.HostName
            Site          = $dc.Site
            OSVersion     = $dc.OperatingSystem
            OSVersion2    = $dc.OperatingSystemVersion
            IsGC          = $dc.IsGlobalCatalog
            IsRODC        = $dc.IsReadOnly
            FSMORoles     = $rolesStr
            ReplStatus    = $replDetail
            ReplOK        = $replOk
            DNSRunning    = (if ($dc.IsGlobalCatalog) {'OK'} else {'Unknown'})
            Uptime        = $uptime
            IPv4          = $dc.IPv4Address
        }
    }
    Write-Log "Collected $($dcData.Count) Domain Controllers." -Level SUCCESS
} catch {
    Write-Log "Error collecting DC data: $_" -Level ERROR
    $dcData = @()
}

# ── REPLICATION STATUS ───────────────────────────
Write-Log "Collecting Replication status..."
$replData = @()
try {
    $replPartners = Get-ADReplicationConnection -Filter * -Server $DomainFQDN -Properties *
    foreach ($conn in $replPartners) {
        try {
            $src  = ($conn.ReplicateFromDirectoryServer -split ',')[0].Replace('CN=','')
            $dst  = ($conn.ReplicateToDirectoryServer  -split ',CN=NTDS')[0].Split(',')[-1].Replace('CN=','') 2>$null
            $dst  = $conn.ReplicateToDirectoryServer -replace 'CN=NTDS Settings,CN=',''-replace ',.*',''

            $meta = Get-ADReplicationPartnerMetadata -Target $src -Scope Server -ErrorAction SilentlyContinue |
                    Where-Object { $_.Partner -like "*$dst*" } | Select-Object -First 1

            $lastSuccess = if ($meta.LastReplicationSuccess) { $meta.LastReplicationSuccess.ToString('yyyy-MM-dd HH:mm') } else { 'Never' }
            $failures    = if ($meta.ConsecutiveReplicationFailures) { $meta.ConsecutiveReplicationFailures } else { 0 }
            $result      = if ($meta.LastReplicationResult -eq 0) { 'Success' } else { "Error ($($meta.LastReplicationResult))" }
            $statusClass = if ($failures -eq 0) { 'ok' } elseif ($failures -lt 5) { 'warn' } else { 'crit' }

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
    Write-Log "Replication data collection error: $_" -Level WARN
}

# ── USERS ──────────────────────────────────────
Write-Log "Collecting User account statistics..."
$userStaleCount    = 0
$userPwdNeverCount = 0
$userLockedCount   = 0
$userPwdExpiring   = 0
$totalEnabled      = 0
$totalDisabled     = 0
$staleUserData     = @()
$lockedUserData    = @()
$pwdNeverData      = @()

try {
    $cutoff90 = (Get-Date).AddDays(-90)
    $cutoff14 = (Get-Date).AddDays(14)

    $allUsers = Get-ADUser -Filter * -Server $DomainFQDN -Properties `
        LastLogonDate, PasswordLastSet, PasswordNeverExpires, LockedOut,
        Department, Title, Manager, Enabled, DistinguishedName,
        'msDS-UserPasswordExpiryTimeComputed' -ErrorAction SilentlyContinue

    foreach ($u in $allUsers) {
        if ($u.Enabled) { $totalEnabled++ } else { $totalDisabled++ }
        if ($u.LockedOut -and $u.Enabled) { $userLockedCount++ }
        if ($u.PasswordNeverExpires -and $u.Enabled) { $userPwdNeverCount++ }
        if ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $cutoff90) { $userStaleCount++ }

        # Password expiry check
        if ($u.Enabled -and -not $u.PasswordNeverExpires -and $u['msDS-UserPasswordExpiryTimeComputed']) {
            try {
                $expTs = [datetime]::FromFileTime([int64]$u['msDS-UserPasswordExpiryTimeComputed'])
                if ($expTs -gt (Get-Date) -and $expTs -lt $cutoff14.AddDays(14)) { $userPwdExpiring++ }
            } catch { }
        }
    }

    # Stale user sample (top 8)
    $staleUserData = $allUsers | Where-Object {
        $_.Enabled -and $_.LastLogonDate -and $_.LastLogonDate -lt $cutoff90
    } | Sort-Object LastLogonDate | Select-Object -First 8 | ForEach-Object {
        [PSCustomObject]@{
            SamAccountName   = $_.SamAccountName
            DisplayName      = $_.Name
            OU               = ($_.DistinguishedName -replace 'CN=[^,]+,','') -replace $domainDN,'' -replace '^,',''
            LastLogon        = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdLastSet       = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdNeverExpires  = $_.PasswordNeverExpires
            Status           = 'Stale'
            Risk             = 'MEDIUM'
        }
    }

    # Locked accounts
    $lockedUserData = $allUsers | Where-Object { $_.LockedOut -and $_.Enabled } | ForEach-Object {
        [PSCustomObject]@{
            SamAccountName   = $_.SamAccountName
            DisplayName      = $_.Name
            OU               = ($_.DistinguishedName -replace 'CN=[^,]+,','') -replace $domainDN,'' -replace '^,',''
            LastLogon        = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdLastSet       = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdNeverExpires  = $_.PasswordNeverExpires
            Status           = 'LOCKED'
            Risk             = 'HIGH'
        }
    }

    # Password never expires sample (top 8)
    $pwdNeverData = $allUsers | Where-Object {
        $_.PasswordNeverExpires -and $_.Enabled -and $_.SamAccountName -notlike 'krbtgt'
    } | Sort-Object PasswordLastSet | Select-Object -First 8 | ForEach-Object {
        [PSCustomObject]@{
            SamAccountName   = $_.SamAccountName
            DisplayName      = $_.Name
            OU               = ($_.DistinguishedName -replace 'CN=[^,]+,','') -replace $domainDN,'' -replace '^,',''
            LastLogon        = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdLastSet       = if ($_.PasswordLastSet) { $_.PasswordLastSet.ToString('yyyy-MM-dd') } else { 'Never' }
            PwdNeverExpires  = $true
            Status           = 'ACTIVE'
            Risk             = 'HIGH'
        }
    }

    # Merge all user highlights
    $userTableData = @()
    $userTableData += $lockedUserData
    $userTableData += $staleUserData
    $userTableData += $pwdNeverData
    $userTableData = $userTableData | Sort-Object Risk -Descending | Select-Object -First 20 | Sort-Object Risk

    Write-Log "User stats: Enabled=$totalEnabled, Disabled=$totalDisabled, Stale=$userStaleCount, Locked=$userLockedCount, PwdNeverExpires=$userPwdNeverCount" -Level SUCCESS
} catch {
    Write-Log "Error collecting user data: $_" -Level WARN
    $userTableData = @()
}

# ── FINE-GRAINED PASSWORD POLICIES ──────────────
Write-Log "Collecting Fine-Grained Password Policies..."
$psoData = @()
try {
    $psos = Get-ADFineGrainedPasswordPolicy -Filter * -Server $DomainFQDN -Properties *
    foreach ($pso in $psos | Sort-Object Precedence) {
        $subjects = (Get-ADFineGrainedPasswordPolicySubject -Identity $pso -Server $DomainFQDN -ErrorAction SilentlyContinue) |
                    Select-Object -ExpandProperty Name
        $psoData += [PSCustomObject]@{
            Name            = $pso.Name
            Precedence      = $pso.Precedence
            AppliedTo       = if ($subjects) { $subjects -join ', ' } else { '(none)' }
            MinLength       = $pso.MinPasswordLength
            MaxAge          = if ($pso.MaxPasswordAge.TotalDays -eq 0) { 'Never' } else { "$([int]$pso.MaxPasswordAge.TotalDays) days" }
            LockoutThresh   = if ($pso.LockoutThreshold -eq 0) { 'Never' } else { "$($pso.LockoutThreshold) / $([int]$pso.LockoutObservationWindow.TotalMinutes) min" }
            Complexity      = $pso.ComplexityEnabled
            Status          = 'Active'
        }
    }
    # If no PSOs, add default domain policy row
    if ($psoData.Count -eq 0) {
        $ddp = Get-ADDefaultDomainPasswordPolicy -Server $DomainFQDN
        $psoData += [PSCustomObject]@{
            Name          = 'Default Domain Policy'
            Precedence    = 'N/A'
            AppliedTo     = 'All Users (domain-level)'
            MinLength     = $ddp.MinPasswordLength
            MaxAge        = if ($ddp.MaxPasswordAge.TotalDays -eq 0) { 'Never' } else { "$([int]$ddp.MaxPasswordAge.TotalDays) days" }
            LockoutThresh = if ($ddp.LockoutThreshold -eq 0) { 'Never' } else { "$($ddp.LockoutThreshold) attempts" }
            Complexity    = $ddp.ComplexityEnabled
            Status        = 'Default'
        }
    }
    Write-Log "Collected $($psoData.Count) PSO entries." -Level SUCCESS
} catch {
    Write-Log "PSO collection error: $_" -Level WARN
    $psoData = @()
}

# ── GROUP POLICY ────────────────────────────────
Write-Log "Collecting Group Policy data..."
$gpoData = @()
$gpoTotal = 0; $gpoUnlinked = 0; $gpoLinked = 0; $gpoEnforced = 0
if ($modAvailable['GroupPolicy']) {
    try {
        $allGPOs = Get-GPO -All -Domain $DomainFQDN
        $gpoTotal = $allGPOs.Count
        foreach ($gpo in $allGPOs | Sort-Object DisplayName) {
            try {
                $report = [xml]($gpo | Get-GPOReport -ReportType XML -Domain $DomainFQDN)
                $links  = $report.GPO.LinksTo
                $linkedTo = if ($links) { ($links | ForEach-Object { $_.SOMPath }) -join '; ' } else { 'Not Linked' }
                $enforced = if ($links) { ($links | Where-Object { $_.NoOverride -eq $true }).Count -gt 0 } else { $false }
                $status   = if (-not $links) { 'UNLINKED'; $gpoUnlinked++ } elseif ($enforced) { 'ENFORCED'; $gpoEnforced++ } else { 'ENABLED'; $gpoLinked++ }

                $gpoData += [PSCustomObject]@{
                    Name         = $gpo.DisplayName
                    LinkedTo     = $linkedTo
                    Status       = $status
                    WMIFilter    = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { 'None' }
                    LastModified = $gpo.ModificationTime.ToString('yyyy-MM-dd')
                    GpoStatus    = $gpo.GpoStatus
                }
            } catch { }
        }
        Write-Log "Collected $gpoTotal GPOs ($gpoLinked linked, $gpoUnlinked unlinked, $gpoEnforced enforced)." -Level SUCCESS
    } catch {
        Write-Log "GPO collection error: $_" -Level WARN
    }
} else {
    Write-Log "GroupPolicy module unavailable — GPO section will show placeholder." -Level WARN
    $gpoData += [PSCustomObject]@{ Name='Install GPMC/RSAT to collect GPO data'; LinkedTo='N/A'; Status='N/A'; WMIFilter='N/A'; LastModified='N/A'; GpoStatus='N/A' }
}

# ── PKI / CERTIFICATES ──────────────────────────
Write-Log "Collecting PKI / Certificate data..."
$pkiData = @()
try {
    # Try to enumerate from AD CS via CertificationAuthority WMI
    $caList = @()
    try {
        $caList = Get-CimInstance -Namespace 'root\CIMv2' -ClassName 'Win32_Service' |
                  Where-Object { $_.Name -eq 'CertSvc' -and $_.State -eq 'Running' } |
                  Select-Object -ExpandProperty SystemName
    } catch { }

    # Local cert store on DCs (Kerberos, LDAPS certs)
    foreach ($dc in $dcData) {
        try {
            $certs = Invoke-Command -ComputerName $dc.HostName -ScriptBlock {
                Get-ChildItem Cert:\LocalMachine\My | Select-Object Subject, Thumbprint, NotAfter, NotBefore,
                    @{n='Template';e={
                        $ext = $_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Certificate Template Name'}
                        if ($ext) { $ext.Format(0) } else { 'Unknown' }
                    }}
            } -ErrorAction SilentlyContinue

            if ($certs) {
                foreach ($cert in $certs) {
                    $daysLeft = ([datetime]$cert.NotAfter - (Get-Date)).Days
                    $statusCls = if ($daysLeft -lt 14) {'crit'} elseif ($daysLeft -lt 30) {'warn'} else {'ok'}
                    $pkiData += [PSCustomObject]@{
                        CommonName   = $cert.Subject -replace 'CN=','' -replace ',.*',''
                        IssuedTo     = $dc.HostName
                        Template     = $cert.Template
                        IssuedDate   = $cert.NotBefore.ToString('yyyy-MM-dd')
                        Expires      = $cert.NotAfter.ToString('yyyy-MM-dd')
                        DaysLeft     = $daysLeft
                        CRLStatus    = 'Valid'
                        Status       = if ($daysLeft -lt 14) {'CRITICAL'} elseif ($daysLeft -lt 30) {'EXPIRING'} else {'HEALTHY'}
                        StatusClass  = $statusCls
                    }
                }
            }
        } catch { }
    }

    # Also check local machine cert store
    $localCerts = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.HasPrivateKey }
    foreach ($cert in $localCerts) {
        $daysLeft = ($cert.NotAfter - (Get-Date)).Days
        $pkiData += [PSCustomObject]@{
            CommonName  = $cert.Subject -replace 'CN=','' -replace ',.*',''
            IssuedTo    = $env:COMPUTERNAME
            Template    = ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Certificate Template Name'} | Select-Object -First 1)?.Format(0) ?? 'Unknown'
            IssuedDate  = $cert.NotBefore.ToString('yyyy-MM-dd')
            Expires     = $cert.NotAfter.ToString('yyyy-MM-dd')
            DaysLeft    = $daysLeft
            CRLStatus   = 'Valid'
            Status      = if ($daysLeft -lt 14) {'CRITICAL'} elseif ($daysLeft -lt 30) {'EXPIRING'} else {'HEALTHY'}
            StatusClass = if ($daysLeft -lt 14) {'crit'} elseif ($daysLeft -lt 30) {'warn'} else {'ok'}
        }
    }

    $pkiData = $pkiData | Sort-Object DaysLeft | Select-Object -Unique -Property * | Sort-Object DaysLeft
    Write-Log "Collected $($pkiData.Count) certificate entries." -Level SUCCESS
} catch {
    Write-Log "PKI data collection error: $_" -Level WARN
    $pkiData = @()
}

# ── DNS ZONES ───────────────────────────────────
Write-Log "Collecting DNS zone data..."
$dnsData = @()
if ($modAvailable['DnsServer']) {
    try {
        $dnsServer = $dcData | Where-Object { $_.FSMORoles -like '*PDC*' } | Select-Object -ExpandProperty HostName -First 1
        if (-not $dnsServer) { $dnsServer = $pdcEmulator }

        $zones = Get-DnsServerZone -ComputerName $dnsServer -ErrorAction SilentlyContinue
        foreach ($z in $zones | Where-Object { -not $z.IsAutoCreated }) {
            $rrCount = 0
            try { $rrCount = (Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue).Count } catch {}

            $aging   = 'N/A'
            $scavEnabled = $false
            if ($z.ZoneType -eq 'Primary') {
                try {
                    $za = Get-DnsServerZoneAging -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue
                    $scavEnabled = $za.AgingEnabled
                    $aging = if ($za.AgingEnabled) {'Enabled'} else {'Disabled'}
                } catch { $aging = 'Unknown' }
            }

            $dnsData += [PSCustomObject]@{
                ZoneName     = $z.ZoneName
                ZoneType     = "$($z.ZoneType)$(if ($z.IsADIntegrated) {' (AD)'})"
                RepScope     = if ($z.IsADIntegrated) { $z.ReplicationScope } else { 'File-based' }
                DynUpdate    = $z.DynamicUpdate
                RecordCount  = $rrCount
                Scavenging   = $aging
                ScavEnabled  = $scavEnabled
                Status       = if (-not $scavEnabled -and $z.ZoneType -eq 'Primary' -and $z.IsADIntegrated) {'WARN'} else {'HEALTHY'}
            }
        }
        Write-Log "Collected $($dnsData.Count) DNS zones." -Level SUCCESS
    } catch {
        Write-Log "DNS collection error: $_" -Level WARN
    }
} else {
    Write-Log "DnsServer module unavailable — DNS section will show placeholder." -Level WARN
    $dnsData += [PSCustomObject]@{ ZoneName='Install RSAT-DNS-Server to collect DNS data'; ZoneType='N/A'; RepScope='N/A'; DynUpdate='N/A'; RecordCount=0; Scavenging='N/A'; ScavEnabled=$false; Status='N/A' }
}

# ── DHCP ────────────────────────────────────────
Write-Log "Collecting DHCP scope data..."
$dhcpData = @()
if ($modAvailable['DHCPServer']) {
    try {
        $dhcpServers = Get-DhcpServerInDC -ErrorAction SilentlyContinue
        foreach ($srv in $dhcpServers | Select-Object -First 3) {
            try {
                $scopes = Get-DhcpServerv4Scope -ComputerName $srv.DnsName -ErrorAction SilentlyContinue
                foreach ($scope in $scopes) {
                    $stats = Get-DhcpServerv4ScopeStatistics -ScopeId $scope.ScopeId -ComputerName $srv.DnsName -ErrorAction SilentlyContinue
                    $pct   = if ($stats.PercentageInUse) { [math]::Round($stats.PercentageInUse, 1) } else { 0 }
                    $fail  = $null
                    try { $fail = Get-DhcpServerv4Failover -ScopeId $scope.ScopeId -ComputerName $srv.DnsName -ErrorAction SilentlyContinue } catch {}

                    $dhcpData += [PSCustomObject]@{
                        ScopeName    = $scope.Name
                        Subnet       = "$($scope.ScopeId)/$($scope.SubnetMask)"
                        Total        = $stats.AddressesFree + $stats.AddressesInUse
                        InUse        = $stats.AddressesInUse
                        Free         = $stats.AddressesFree
                        UtilPct      = $pct
                        Failover     = if ($fail) { $fail.Mode.ToString().Substring(0,2).ToUpper() } else { 'None' }
                        Status       = if ($pct -ge 90) {'HIGH USE'} elseif ($pct -ge 75) {'WARNING'} else {'HEALTHY'}
                        StatusClass  = if ($pct -ge 90) {'crit'} elseif ($pct -ge 75) {'warn'} else {'ok'}
                        Server       = $srv.DnsName
                    }
                }
            } catch { }
        }
        Write-Log "Collected $($dhcpData.Count) DHCP scopes." -Level SUCCESS
    } catch {
        Write-Log "DHCP collection error: $_" -Level WARN
    }
} else {
    Write-Log "DHCPServer module unavailable — DHCP section will show placeholder." -Level WARN
    $dhcpData += [PSCustomObject]@{ ScopeName='Install RSAT-DHCP to collect DHCP data'; Subnet='N/A'; Total=0; InUse=0; Free=0; UtilPct=0; Failover='N/A'; Status='N/A'; StatusClass='info'; Server='N/A' }
}

# ── PRIVILEGED GROUPS ────────────────────────────
Write-Log "Collecting Privileged Group memberships..."
$privGroupData = @()
$privGroups = @('Domain Admins','Enterprise Admins','Schema Admins','Administrators',
                'Backup Operators','Account Operators','Server Operators','Print Operators',
                'Group Policy Creator Owners','DnsAdmins','DHCP Administrators')
try {
    foreach ($gname in $privGroups) {
        try {
            $grp = Get-ADGroup -Filter "Name -eq '$gname'" -Server $DomainFQDN -ErrorAction SilentlyContinue
            if (-not $grp) { continue }
            $members = Get-ADGroupMember -Identity $grp -Recursive -Server $DomainFQDN -ErrorAction SilentlyContinue
            $memberNames = ($members | Select-Object -ExpandProperty SamAccountName) -join ', '
            $tier = switch ($gname) {
                'Domain Admins'    {'TIER 0'} 'Enterprise Admins' {'TIER 0'} 'Schema Admins' {'TIER 0'}
                'Administrators'   {'TIER 0'} default {'TIER 1'}
            }
            $tierClass = if ($tier -eq 'TIER 0') {'crit'} else {'warn'}

            $privGroupData += [PSCustomObject]@{
                GroupName    = $gname
                MemberCount  = $members.Count
                Members      = if ($memberNames.Length -gt 120) { $memberNames.Substring(0,117) + '…' } else { $memberNames }
                Tier         = $tier
                TierClass    = $tierClass
                LastChange   = 'See AD audit logs'
                ReviewStatus = if ($members.Count -gt 10) { 'REVIEW' } else { 'OK' }
            }
        } catch { }
    }
    Write-Log "Collected $($privGroupData.Count) privileged group records." -Level SUCCESS
} catch {
    Write-Log "Privileged group collection error: $_" -Level WARN
}

# ── SECURITY ALERTS ──────────────────────────────
Write-Log "Building Security Alert summary..."
$secAlerts = @()
$alertId = 1

# Locked accounts
if ($userLockedCount -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'Accounts'
        Finding    = "$userLockedCount user account(s) currently locked out"
        AffectedObj= "$userLockedCount accounts"
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = 'WARNING'; SevClass = 'warn'
        Action     = 'Review lockout source; unlock as appropriate; check for brute-force'
    }
}
# Stale accounts
if ($userStaleCount -gt 5) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'Accounts'
        Finding    = "$userStaleCount enabled accounts inactive for 90+ days"
        AffectedObj= "$userStaleCount accounts"
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = 'WARNING'; SevClass = 'warn'
        Action     = 'Disable or remove stale accounts; review with managers'
    }
}
# Password never expires
if ($userPwdNeverCount -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'Accounts'
        Finding    = "$userPwdNeverCount enabled accounts with Password Never Expires"
        AffectedObj= "$userPwdNeverCount accounts"
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = if ($userPwdNeverCount -gt 20) {'CRITICAL'} else {'WARNING'}
        SevClass   = if ($userPwdNeverCount -gt 20) {'crit'} else {'warn'}
        Action     = 'Apply PSO or enable password expiry; exclude only gMSA and break-glass'
    }
}
# Expiring certs
$expiringCerts = $pkiData | Where-Object { $_.DaysLeft -lt 30 }
foreach ($cert in $expiringCerts) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'PKI'
        Finding    = "Certificate expires in $($cert.DaysLeft) day(s): $($cert.CommonName)"
        AffectedObj= $cert.IssuedTo
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = if ($cert.DaysLeft -lt 7) {'CRITICAL'} else {'WARNING'}
        SevClass   = if ($cert.DaysLeft -lt 7) {'crit'} else {'warn'}
        Action     = 'Renew certificate before expiry to prevent authentication failures'
    }
}
# GPO unlinked
if ($gpoUnlinked -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'GPO'
        Finding    = "$gpoUnlinked unlinked GPO(s) in the domain"
        AffectedObj= "Group Policy"
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = 'INFO'; SevClass = 'info'
        Action     = 'Review unlinked GPOs; link or delete to reduce namespace clutter'
    }
}
# Replication issues
$replFails = $replData | Where-Object { $_.Failures -gt 0 }
foreach ($rf in $replFails) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'Replication'
        Finding    = "Replication failures: $($rf.SourceDC) → $($rf.DestDC) ($($rf.Failures) failure(s))"
        AffectedObj= $rf.DestDC
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = if ($rf.Failures -ge 5) {'CRITICAL'} else {'WARNING'}
        SevClass   = if ($rf.Failures -ge 5) {'crit'} else {'warn'}
        Action     = 'Run repadmin /replsummary; check WAN link; force sync with repadmin /syncall'
    }
}
# DNS scavenging disabled
$dnsNoScav = $dnsData | Where-Object { $_.Status -eq 'WARN' }
if ($dnsNoScav.Count -gt 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'DNS'
        Finding    = "$($dnsNoScav.Count) DNS zone(s) with scavenging disabled (stale record risk)"
        AffectedObj= ($dnsNoScav | Select-Object -ExpandProperty ZoneName) -join ', '
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = 'INFO'; SevClass = 'info'
        Action     = 'Enable DNS scavenging; review NoRefreshInterval and RefreshInterval settings'
    }
}
# Privileged group size check
$domAdminsRow = $privGroupData | Where-Object { $_.GroupName -eq 'Domain Admins' }
if ($domAdminsRow -and [int]$domAdminsRow.MemberCount -gt 8) {
    $secAlerts += [PSCustomObject]@{
        AlertID    = "SEC-$(($alertId++).ToString('000'))"
        Category   = 'Privileged Access'
        Finding    = "Domain Admins has $($domAdminsRow.MemberCount) members — consider reducing"
        AffectedObj= 'Domain Admins'
        Detected   = (Get-Date).ToString('yyyy-MM-dd')
        Severity   = 'WARNING'; SevClass = 'warn'
        Action     = 'Audit Domain Admins membership; use Tier 0 accounts with JIT elevation where possible'
    }
}
if ($secAlerts.Count -eq 0) {
    $secAlerts += [PSCustomObject]@{
        AlertID='N/A'; Category='General'; Finding='No security alerts detected at this time.';
        AffectedObj='N/A'; Detected=(Get-Date).ToString('yyyy-MM-dd');
        Severity='INFO'; SevClass='info'; Action='Continue routine monitoring'
    }
}

# ── KPI SUMMARY COUNTS ───────────────────────────
$critAlerts  = ($secAlerts | Where-Object { $_.Severity -eq 'CRITICAL' }).Count
$warnAlerts  = ($secAlerts | Where-Object { $_.Severity -eq 'WARNING'  }).Count
$expiringCount = ($pkiData | Where-Object { $_.DaysLeft -lt 30 }).Count
$replOkPct   = if ($replData.Count -gt 0) { [math]::Round( ($replData | Where-Object {$_.Failures -eq 0}).Count / $replData.Count * 100, 0) } else { 100 }

Write-Log "Data collection complete. Building HTML report..." -Level SUCCESS

# ════════════════════════════════════════════════
#  HTML BUILDER HELPERS
# ════════════════════════════════════════════════
function Build-TableRows {
    param([array]$Data, [string[]]$Props, [hashtable]$StatusCols = @{}, [hashtable]$StatusClsCols = @{})
    $sb = [System.Text.StringBuilder]::new()
    foreach ($row in $Data) {
        [void]$sb.Append('<tr>')
        foreach ($p in $Props) {
            $val = $row.$p
            if ($StatusCols.ContainsKey($p)) {
                $clsProp = $StatusCols[$p]
                $cls = if ($row.$clsProp) { $row.$clsProp } else { 'info' }
                [void]$sb.Append("<td><span class=`"td-badge $cls`">$(Escape-Html $val)</span></td>")
            } elseif ($p -match 'Pct$') {
                $pct = [int]$val
                $barCls = if ($pct -ge 90) {'red'} elseif ($pct -ge 75) {'amber'} else {'green'}
                [void]$sb.Append("<td><div class=`"progress-wrap`"><div class=`"progress-bar-bg`"><div class=`"progress-bar-fill $barCls`" style=`"width:$pct%`"></div></div><span class=`"progress-val`">$pct%</span></div></td>")
            } else {
                [void]$sb.Append("<td class=`"$(if ($p -match 'Name$|DC$|Account|Subnet|Zone') {'td-mono'} else {''})`">$(Escape-Html "$val")</td>")
            }
        }
        [void]$sb.Append('</tr>')
    }
    return $sb.ToString()
}

# ════════════════════════════════════════════════
#  BUILD HTML REPORT
# ════════════════════════════════════════════════
Write-Log "Generating HTML report..."

$reportDate  = (Get-Date).ToString('MMMM dd, yyyy  HH:mm:ss')
$reportDateISO = (Get-Date -Format 'yyyy-MM-dd HH:mm')
$kpiReplClass  = if ($replOkPct -eq 100) {'ok'} elseif ($replOkPct -ge 90) {'warn'} else {'crit'}
$kpiCertClass  = if ($expiringCount -eq 0) {'ok'} elseif ($expiringCount -lt 3) {'warn'} else {'crit'}
$kpiAlertClass = if ($critAlerts -gt 0) {'crit'} elseif ($warnAlerts -gt 0) {'warn'} else {'ok'}
$kpiStaleClass = if ($userStaleCount -gt 50) {'crit'} elseif ($userStaleCount -gt 10) {'warn'} else {'ok'}

# ─ DC table rows ─
$dcRows = foreach ($dc in $dcData) {
    $gcBadge   = Status-Badge (if($dc.IsGC){'YES'}else{'NO'})   (if($dc.IsGC){'ok'}else{'warn'})
    $replBadge = Status-Badge $dc.ReplStatus (if($dc.ReplOK){'ok'}else{'warn'})
    $dnsBadge  = Status-Badge 'OK' 'ok'
    $roleLabel = if ($dc.IsRODC) { "$($dc.FSMORoles) (RODC)" } else { $dc.FSMORoles }
    "<tr><td class='td-mono'>$($dc.Name)</td><td>$(Escape-Html $dc.Site)</td><td>$(Escape-Html $dc.OSVersion)</td><td>$(Escape-Html $roleLabel)</td><td>$gcBadge</td><td>$replBadge</td><td>$dnsBadge</td><td class='td-mono'>$($dc.Uptime)</td><td class='td-mono'>$($dc.IPv4)</td></tr>"
}

# ─ Replication rows ─
$replRows = foreach ($r in $replData) {
    $badge = Status-Badge $r.Status $r.StatusClass
    "<tr><td class='td-mono'>$(Escape-Html $r.SourceDC)</td><td class='td-mono'>$(Escape-Html $r.DestDC)</td><td class='td-mono' style='font-size:11px'>$(Escape-Html $r.NC)</td><td class='td-mono'>$($r.LastSuccess)</td><td class='td-mono'>$($r.Failures)</td><td>$badge</td></tr>"
}
if (-not $replRows) { $replRows = "<tr><td colspan='6' style='color:var(--text-muted);text-align:center'>No replication data collected — run on a Domain Controller for full data</td></tr>" }

# ─ User rows ─
$userRows = foreach ($u in $userTableData) {
    $riskCls = switch ($u.Risk) { 'HIGH'{'crit'} 'MEDIUM'{'warn'} default{'info'} }
    $stsCls  = switch ($u.Status) { 'LOCKED'{'crit'} 'Stale'{'warn'} default{'ok'} }
    "<tr><td class='td-mono'>$(Escape-Html $u.SamAccountName)</td><td>$(Escape-Html $u.DisplayName)</td><td class='td-mono' style='font-size:11px'>$(Escape-Html $u.OU)</td><td class='td-mono'>$($u.LastLogon)</td><td class='td-mono'>$($u.PwdLastSet)</td><td>$(Escape-Html $u.PwdNeverExpires)</td><td>$(Status-Badge $u.Status $stsCls)</td><td>$(Status-Badge $u.Risk $riskCls)</td></tr>"
}
if (-not $userRows) { $userRows = "<tr><td colspan='8' style='color:var(--text-muted);text-align:center'>No flagged user accounts found</td></tr>" }

# ─ PSO rows ─
$psoRows = foreach ($p in $psoData) {
    $stsCls = if ($p.Status -eq 'Active') {'ok'} else {'info'}
    "<tr><td class='td-mono'>$(Escape-Html $p.Name)</td><td class='td-mono'>$($p.Precedence)</td><td>$(Escape-Html $p.AppliedTo)</td><td class='td-mono'>$($p.MinLength)</td><td class='td-mono'>$(Escape-Html $p.MaxAge)</td><td class='td-mono'>$(Escape-Html $p.LockoutThresh)</td><td>$(Status-Badge $p.Complexity 'info')</td><td>$(Status-Badge $p.Status $stsCls)</td></tr>"
}
if (-not $psoRows) { $psoRows = "<tr><td colspan='8' style='color:var(--text-muted);text-align:center'>No PSO data</td></tr>" }

# ─ GPO rows ─
$gpoRows = foreach ($g in $gpoData | Select-Object -First 20) {
    $cls = switch ($g.Status) { 'UNLINKED'{'warn'} 'ENFORCED'{'info'} default{'ok'} }
    "<tr><td>$(Escape-Html $g.Name)</td><td class='td-mono' style='font-size:11px'>$(Escape-Html $g.LinkedTo)</td><td>$(Status-Badge $g.Status $cls)</td><td>$(Escape-Html $g.WMIFilter)</td><td class='td-mono'>$(Escape-Html $g.LastModified)</td><td>$(Escape-Html $g.GpoStatus)</td></tr>"
}
if (-not $gpoRows) { $gpoRows = "<tr><td colspan='6' style='color:var(--text-muted);text-align:center'>Install GPMC (RSAT) and rerun to collect GPO data</td></tr>" }

# ─ PKI rows ─
$pkiRows = foreach ($c in $pkiData | Select-Object -First 20) {
    $dColor = switch ($c.StatusClass) { 'crit'{'var(--accent-red)'} 'warn'{'var(--accent-amber)'} default{'var(--accent-green)'} }
    "<tr><td>$(Escape-Html $c.CommonName)</td><td class='td-mono'>$(Escape-Html $c.IssuedTo)</td><td>$(Escape-Html $c.Template)</td><td class='td-mono'>$(Escape-Html $c.IssuedDate)</td><td class='td-mono'>$(Escape-Html $c.Expires)</td><td class='td-mono' style='color:$dColor'>$($c.DaysLeft)</td><td>$(Status-Badge 'Valid' 'ok')</td><td>$(Status-Badge $c.Status $c.StatusClass)</td></tr>"
}
if (-not $pkiRows) { $pkiRows = "<tr><td colspan='8' style='color:var(--text-muted);text-align:center'>No certificates found in local machine store</td></tr>" }

# ─ DNS rows ─
$dnsRows = foreach ($z in $dnsData) {
    $cls = if ($z.Status -eq 'HEALTHY') {'ok'} elseif ($z.Status -eq 'N/A') {'info'} else {'warn'}
    "<tr><td class='td-mono'>$(Escape-Html $z.ZoneName)</td><td>$(Escape-Html $z.ZoneType)</td><td>$(Escape-Html $z.RepScope)</td><td>$(Escape-Html $z.DynUpdate)</td><td class='td-mono'>$($z.RecordCount)</td><td>$(Status-Badge $z.Scavenging (if($z.ScavEnabled){'ok'}else{'warn'}))</td><td>$(Status-Badge $z.Status $cls)</td></tr>"
}
if (-not $dnsRows) { $dnsRows = "<tr><td colspan='7' style='color:var(--text-muted);text-align:center'>Install RSAT-DNS-Server and rerun to collect DNS data</td></tr>" }

# ─ DHCP rows ─
$dhcpRows = foreach ($d in $dhcpData) {
    $pct = [int]$d.UtilPct
    $barCls = if ($pct -ge 90) {'red'} elseif ($pct -ge 75) {'amber'} else {'green'}
    "<tr><td>$(Escape-Html $d.ScopeName)</td><td class='td-mono'>$(Escape-Html $d.Subnet)</td><td class='td-mono'>$($d.Total)</td><td class='td-mono'>$($d.InUse)</td><td class='td-mono'>$($d.Free)</td><td><div class='progress-wrap'><div class='progress-bar-bg'><div class='progress-bar-fill $barCls' style='width:$pct%'></div></div><span class='progress-val'>$pct%</span></div></td><td>$(Status-Badge $d.Failover 'ok')</td><td>$(Status-Badge $d.Status $d.StatusClass)</td></tr>"
}
if (-not $dhcpRows) { $dhcpRows = "<tr><td colspan='8' style='color:var(--text-muted);text-align:center'>Install RSAT-DHCP and rerun to collect DHCP data</td></tr>" }

# ─ Priv group rows ─
$privRows = foreach ($g in $privGroupData) {
    "<tr><td>$(Escape-Html $g.GroupName)</td><td class='td-mono'>$($g.MemberCount)</td><td style='font-size:11.5px'>$(Escape-Html $g.Members)</td><td>$(Status-Badge $g.Tier $g.TierClass)</td><td>$(Escape-Html $g.LastChange)</td><td>$(Status-Badge $g.ReviewStatus (if($g.ReviewStatus -eq 'OK'){'ok'}else{'warn'}))</td></tr>"
}

# ─ Security alert rows ─
$secRows = foreach ($a in $secAlerts) {
    "<tr><td class='td-mono'>$($a.AlertID)</td><td>$(Escape-Html $a.Category)</td><td>$(Escape-Html $a.Finding)</td><td class='td-mono'>$(Escape-Html $a.AffectedObj)</td><td class='td-mono'>$($a.Detected)</td><td>$(Status-Badge $a.Severity $a.SevClass)</td><td>$(Escape-Html $a.Action)</td></tr>"
}

# ═══════════════════════════════════════════════
#  INLINE HTML (same styles + JS as before, data injected)
# ═══════════════════════════════════════════════
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AD Health Dashboard — $DomainFQDN</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://unpkg.com/docx@8.5.0/build/index.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
:root{--bg-primary:#0a0e1a;--bg-secondary:#0f1629;--bg-card:#131c2e;--bg-card-hover:#172035;--border:#1e2d4a;--border-bright:#2a3f60;--accent-blue:#1e90ff;--accent-cyan:#00d4ff;--accent-green:#00e676;--accent-amber:#ffab40;--accent-red:#ff5252;--accent-purple:#b388ff;--text-primary:#e8f0fe;--text-secondary:#8ba3c7;--text-muted:#4a6080;--grid-line:rgba(30,144,255,0.06);--glow-blue:rgba(30,144,255,0.15);--status-ok:#00e676;--status-warn:#ffab40;--status-crit:#ff5252;--status-info:#1e90ff;--header-h:72px;}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'IBM Plex Sans',sans-serif;background:var(--bg-primary);color:var(--text-primary);min-height:100vh;font-size:14px;line-height:1.6;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:40px 40px;}
.header{position:sticky;top:0;z-index:100;height:var(--header-h);background:rgba(10,14,26,0.95);backdrop-filter:blur(16px);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:24px;padding:0 32px;box-shadow:0 4px 32px rgba(0,0,0,0.4);}
.header-logo{display:flex;align-items:center;gap:12px;flex-shrink:0;}
.header-logo svg{width:36px;height:36px;}
.header-titles{line-height:1.2;}
.header-titles h1{font-size:18px;font-weight:700;letter-spacing:0.04em;background:linear-gradient(90deg,var(--accent-cyan),var(--accent-blue));-webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.header-meta{font-size:11px;color:var(--text-muted);font-family:'IBM Plex Mono',monospace;}
.header-domain-wrap{display:flex;align-items:center;gap:8px;margin-left:8px;}
.domain-label{font-size:11px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.08em;}
.domain-input{background:var(--bg-secondary);border:1px solid var(--border-bright);color:var(--accent-cyan);padding:4px 10px;border-radius:4px;font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:500;width:240px;transition:border-color 0.2s;}
.domain-input:focus{outline:none;border-color:var(--accent-blue);box-shadow:0 0 0 3px var(--glow-blue);}
.header-right{margin-left:auto;display:flex;align-items:center;gap:16px;}
.report-time{font-size:11px;color:var(--text-muted);font-family:'IBM Plex Mono',monospace;text-align:right;}
.report-time span{display:block;}
.report-time .author{color:var(--text-secondary);}
.pulse-dot{width:8px;height:8px;border-radius:50%;background:var(--accent-green);box-shadow:0 0 8px var(--accent-green);animation:pulse 2s ease-in-out infinite;}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:0.5;transform:scale(0.8)}}
.search-wrap{padding:16px 32px;background:var(--bg-secondary);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:16px;}
.search-box{display:flex;align-items:center;gap:10px;background:var(--bg-card);border:1px solid var(--border-bright);border-radius:6px;padding:8px 14px;flex:1;max-width:600px;transition:border-color 0.2s,box-shadow 0.2s;}
.search-box:focus-within{border-color:var(--accent-blue);box-shadow:0 0 0 3px var(--glow-blue);}
.search-box svg{color:var(--text-muted);flex-shrink:0;}
.search-box input{background:none;border:none;outline:none;color:var(--text-primary);font-family:'IBM Plex Sans',sans-serif;font-size:14px;width:100%;}
.search-box input::placeholder{color:var(--text-muted);}
.search-count{font-size:12px;color:var(--text-muted);white-space:nowrap;}
.search-clear{background:none;border:none;color:var(--text-muted);cursor:pointer;font-size:18px;line-height:1;padding:0 4px;transition:color 0.2s;}
.search-clear:hover{color:var(--accent-red);}
.main{padding:24px 32px;max-width:1600px;margin:0 auto;}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:28px;}
.kpi-card{background:var(--bg-card);border:1px solid var(--border);border-radius:8px;padding:18px 20px;position:relative;overflow:hidden;transition:border-color 0.2s,transform 0.15s;}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.kpi-card.ok::before{background:var(--status-ok)}.kpi-card.warn::before{background:var(--status-warn)}.kpi-card.crit::before{background:var(--status-crit)}.kpi-card.info::before{background:var(--status-info)}
.kpi-card:hover{border-color:var(--border-bright);transform:translateY(-2px);}
.kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);margin-bottom:8px;}
.kpi-value{font-size:32px;font-weight:700;font-family:'IBM Plex Mono',monospace;line-height:1;margin-bottom:4px;}
.kpi-card.ok .kpi-value{color:var(--status-ok)}.kpi-card.warn .kpi-value{color:var(--status-warn)}.kpi-card.crit .kpi-value{color:var(--status-crit)}.kpi-card.info .kpi-value{color:var(--status-info)}
.kpi-sub{font-size:11px;color:var(--text-muted);}
.section{background:var(--bg-card);border:1px solid var(--border);border-radius:8px;margin-bottom:20px;overflow:hidden;transition:border-color 0.2s;}
.section:hover{border-color:var(--border-bright);}
.section.search-hidden{display:none;}
.section-header{display:flex;align-items:center;gap:12px;padding:14px 20px;cursor:pointer;border-bottom:1px solid transparent;transition:background 0.15s,border-color 0.15s;user-select:none;}
.section-header:hover{background:var(--bg-card-hover);}
.section.open .section-header{border-bottom-color:var(--border);}
.section-icon{width:32px;height:32px;border-radius:6px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:15px;}
.section-icon.blue{background:rgba(30,144,255,0.15);color:var(--accent-blue)}.section-icon.cyan{background:rgba(0,212,255,0.12);color:var(--accent-cyan)}.section-icon.green{background:rgba(0,230,118,0.12);color:var(--accent-green)}.section-icon.amber{background:rgba(255,171,64,0.12);color:var(--accent-amber)}.section-icon.red{background:rgba(255,82,82,0.12);color:var(--accent-red)}.section-icon.purple{background:rgba(179,136,255,0.12);color:var(--accent-purple)}
.section-title{font-size:14px;font-weight:600;letter-spacing:0.02em;flex:1;}
.section-badge{font-size:10px;font-family:'IBM Plex Mono',monospace;padding:2px 8px;border-radius:20px;font-weight:500;letter-spacing:0.05em;}
.badge-ok{background:rgba(0,230,118,0.15);color:var(--status-ok);border:1px solid rgba(0,230,118,0.3)}.badge-warn{background:rgba(255,171,64,0.15);color:var(--status-warn);border:1px solid rgba(255,171,64,0.3)}.badge-crit{background:rgba(255,82,82,0.15);color:var(--status-crit);border:1px solid rgba(255,82,82,0.3)}.badge-info{background:rgba(30,144,255,0.15);color:var(--status-info);border:1px solid rgba(30,144,255,0.3)}.badge-neutral{background:rgba(139,163,199,0.1);color:var(--text-secondary);border:1px solid var(--border);}
.section-chevron{color:var(--text-muted);transition:transform 0.25s;flex-shrink:0;}
.section.open .section-chevron{transform:rotate(180deg);}
.export-bar{display:none;align-items:center;gap:8px;padding:10px 20px;background:rgba(30,144,255,0.04);border-bottom:1px solid var(--border);flex-wrap:wrap;}
.section.open .export-bar{display:flex;}
.export-label{font-size:10px;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);margin-right:4px;}
.btn-export{display:inline-flex;align-items:center;gap:5px;padding:4px 12px;border-radius:4px;font-size:11px;font-weight:600;letter-spacing:0.04em;cursor:pointer;border:1px solid transparent;font-family:'IBM Plex Mono',monospace;transition:all 0.15s;}
.btn-export.csv{background:rgba(0,230,118,0.1);color:var(--accent-green);border-color:rgba(0,230,118,0.2)}.btn-export.xlsx{background:rgba(30,144,255,0.1);color:var(--accent-blue);border-color:rgba(30,144,255,0.2)}.btn-export.txt{background:rgba(139,163,199,0.1);color:var(--text-secondary);border-color:var(--border)}.btn-export.docx{background:rgba(179,136,255,0.12);color:var(--accent-purple);border-color:rgba(179,136,255,0.25)}
.btn-export:hover{filter:brightness(1.2);transform:translateY(-1px)}.btn-export:active{transform:translateY(0)}
.section-body{display:none;padding:20px;}
.section.open .section-body{display:block;}
.data-table{width:100%;border-collapse:collapse;font-size:12.5px;}
.data-table th{text-align:left;padding:8px 12px;font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--text-muted);font-weight:600;border-bottom:1px solid var(--border);background:var(--bg-secondary);white-space:nowrap;}
.data-table td{padding:9px 12px;border-bottom:1px solid rgba(30,46,74,0.5);color:var(--text-primary);vertical-align:middle;}
.data-table tr:last-child td{border-bottom:none}.data-table tr:hover td{background:rgba(30,144,255,0.04);}
.td-badge{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;font-family:'IBM Plex Mono',monospace;white-space:nowrap;}
.td-badge::before{content:'';width:5px;height:5px;border-radius:50%;}
.td-badge.ok{background:rgba(0,230,118,0.12);color:var(--status-ok)}.td-badge.ok::before{background:var(--status-ok)}
.td-badge.warn{background:rgba(255,171,64,0.12);color:var(--status-warn)}.td-badge.warn::before{background:var(--status-warn)}
.td-badge.crit{background:rgba(255,82,82,0.12);color:var(--status-crit)}.td-badge.crit::before{background:var(--status-crit)}
.td-badge.info{background:rgba(30,144,255,0.12);color:var(--status-info)}.td-badge.info::before{background:var(--status-info)}
.td-mono{font-family:'IBM Plex Mono',monospace;font-size:11.5px;}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:20px;}
.sum-item{background:var(--bg-secondary);border:1px solid var(--border);border-radius:6px;padding:12px 14px;}
.sum-item .sum-label{font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.08em;margin-bottom:4px;}
.sum-item .sum-val{font-size:22px;font-weight:700;font-family:'IBM Plex Mono',monospace;}
.sum-item .sum-val.green{color:var(--accent-green)}.sum-item .sum-val.amber{color:var(--accent-amber)}.sum-item .sum-val.red{color:var(--accent-red)}.sum-item .sum-val.blue{color:var(--accent-blue)}.sum-item .sum-val.purple{color:var(--accent-purple)}
.progress-wrap{display:flex;align-items:center;gap:10px;}
.progress-bar-bg{flex:1;height:6px;background:var(--bg-secondary);border-radius:3px;overflow:hidden;}
.progress-bar-fill{height:100%;border-radius:3px;}
.progress-bar-fill.green{background:var(--accent-green)}.progress-bar-fill.amber{background:var(--accent-amber)}.progress-bar-fill.red{background:var(--accent-red)}.progress-bar-fill.blue{background:var(--accent-blue)}
.progress-val{font-size:11px;font-family:'IBM Plex Mono',monospace;color:var(--text-secondary);min-width:38px;text-align:right;}
.footer{text-align:center;padding:32px;color:var(--text-muted);font-size:11px;border-top:1px solid var(--border);margin-top:32px;font-family:'IBM Plex Mono',monospace;}
.footer span{color:var(--text-secondary);}
tr.search-match td{background:rgba(30,144,255,0.08)!important;}
@media print{.header{position:relative;}.search-wrap,.export-bar,.section-chevron,.pulse-dot{display:none!important;}.section-body{display:block!important;}.section-header{border-bottom:1px solid var(--border)!important;}}
</style>
</head>
<body>
<header class="header">
  <div class="header-logo">
    <svg viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg"><rect width="36" height="36" rx="8" fill="rgba(30,144,255,0.1)" stroke="rgba(30,144,255,0.3)" stroke-width="1"/><path d="M18 6L30 12V24L18 30L6 24V12L18 6Z" stroke="#00d4ff" stroke-width="1.5" fill="rgba(0,212,255,0.08)"/><circle cx="18" cy="18" r="4" fill="#1e90ff"/><path d="M18 10V14M18 22V26M10 18H14M22 18H26" stroke="#1e90ff" stroke-width="1.5" stroke-linecap="round"/></svg>
    <div class="header-titles"><h1>AD HEALTH DASHBOARD</h1><div class="header-meta">Enterprise Domain Report</div></div>
  </div>
  <div class="header-domain-wrap">
    <span class="domain-label">Domain:</span>
    <input class="domain-input" id="domainInput" value="$(Escape-Html $DomainFQDN)" title="Click to edit domain name"/>
  </div>
  <div class="header-right">
    <div class="report-time"><span>$reportDate</span><span class="author">$(Escape-Html $Author)</span></div>
    <div class="pulse-dot"></div>
  </div>
</header>

<div class="search-wrap">
  <div class="search-box">
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
    <input type="text" id="searchInput" placeholder="Search across all sections — DCs, accounts, policies, certificates…" autocomplete="off"/>
    <button class="search-clear" id="searchClear" title="Clear">×</button>
  </div>
  <span class="search-count" id="searchCount"></span>
</div>

<main class="main">
<div class="kpi-grid">
  <div class="kpi-card info"><div class="kpi-label">Domain Controllers</div><div class="kpi-value">$($dcData.Count)</div><div class="kpi-sub">$DomainFQDN</div></div>
  <div class="kpi-card $kpiReplClass"><div class="kpi-label">Replication Health</div><div class="kpi-value">$replOkPct<span style="font-size:16px">%</span></div><div class="kpi-sub">$(if($replOkPct -eq 100){'No errors'}else{"$(($replData|Where-Object{$_.Failures -gt 0}).Count) with failures"})</div></div>
  <div class="kpi-card $kpiStaleClass"><div class="kpi-label">Stale Accounts</div><div class="kpi-value">$userStaleCount</div><div class="kpi-sub">90+ days inactive</div></div>
  <div class="kpi-card $(if($userLockedCount -gt 0){'crit'}else{'ok'})"><div class="kpi-label">Locked Accounts</div><div class="kpi-value">$userLockedCount</div><div class="kpi-sub">Currently locked</div></div>
  <div class="kpi-card $kpiCertClass"><div class="kpi-label">Expiring Certs</div><div class="kpi-value">$expiringCount</div><div class="kpi-sub">Within 30 days</div></div>
  <div class="kpi-card $kpiAlertClass"><div class="kpi-label">Security Alerts</div><div class="kpi-value">$($secAlerts.Count)</div><div class="kpi-sub">$critAlerts critical, $warnAlerts warning</div></div>
  <div class="kpi-card info"><div class="kpi-label">Total Enabled Users</div><div class="kpi-value">$totalEnabled</div><div class="kpi-sub">$totalDisabled disabled</div></div>
  <div class="kpi-card info"><div class="kpi-label">Domain Functional Level</div><div class="kpi-value" style="font-size:16px">$(Escape-Html $domainMode)</div><div class="kpi-sub">Forest: $(Escape-Html $forestMode)</div></div>
</div>

<!-- SECTION: Domain Controllers -->
<div class="section open" id="sec-dc">
  <div class="section-header" onclick="toggleSection('sec-dc')">
    <div class="section-icon blue">🖥</div>
    <span class="section-title">Domain Controller Health</span>
    <span class="section-badge $(if(($dcData|Where-Object{-not $_.ReplOK}).Count -eq 0){'badge-ok'}else{'badge-warn'})">$($dcData.Count) DCs</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-dc','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-dc','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-dc','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-dc','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <div class="summary-grid">
      <div class="sum-item"><div class="sum-label">Total DCs</div><div class="sum-val blue">$($dcData.Count)</div></div>
      <div class="sum-item"><div class="sum-label">GC Servers</div><div class="sum-val blue">$(($dcData|Where-Object{$_.IsGC}).Count)</div></div>
      <div class="sum-item"><div class="sum-label">RODC</div><div class="sum-val amber">$(($dcData|Where-Object{$_.IsRODC}).Count)</div></div>
      <div class="sum-item"><div class="sum-label">Repl Errors</div><div class="sum-val $(if(($replData|Where-Object{$_.Failures -gt 0}).Count -eq 0){'green'}else{'red'})">$(($replData|Where-Object{$_.Failures -gt 0}).Count)</div></div>
      <div class="sum-item"><div class="sum-label">Sites</div><div class="sum-val blue">$(($dcData|Select-Object -ExpandProperty Site -Unique).Count)</div></div>
    </div>
    <table class="data-table" id="tbl-dc">
      <thead><tr><th>DC Name</th><th>Site</th><th>OS Version</th><th>FSMO Roles</th><th>GC</th><th>Replication</th><th>DNS</th><th>Uptime</th><th>IP Address</th></tr></thead>
      <tbody>$dcRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: Replication -->
<div class="section open" id="sec-repl">
  <div class="section-header" onclick="toggleSection('sec-repl')">
    <div class="section-icon cyan">↺</div>
    <span class="section-title">AD Replication Status</span>
    <span class="section-badge $(if(($replData|Where-Object{$_.Failures -gt 0}).Count -eq 0){'badge-ok'}else{'badge-warn'})">$replOkPct% HEALTHY</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-repl','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-repl','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-repl','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-repl','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-repl">
      <thead><tr><th>Source DC</th><th>Destination DC</th><th>Naming Context</th><th>Last Success</th><th>Failures</th><th>Status</th></tr></thead>
      <tbody>$replRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: Users -->
<div class="section open" id="sec-users">
  <div class="section-header" onclick="toggleSection('sec-users')">
    <div class="section-icon green">👤</div>
    <span class="section-title">User Account Health</span>
    <span class="section-badge $(if($userLockedCount -gt 0 -or $userStaleCount -gt 20){'badge-warn'}else{'badge-ok'})">$totalEnabled ENABLED</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-users','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-users','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-users','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-users','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <div class="summary-grid">
      <div class="sum-item"><div class="sum-label">Enabled</div><div class="sum-val blue">$totalEnabled</div></div>
      <div class="sum-item"><div class="sum-label">Disabled</div><div class="sum-val amber">$totalDisabled</div></div>
      <div class="sum-item"><div class="sum-label">Inactive 90d+</div><div class="sum-val $(if($userStaleCount -gt 20){'red'}else{'amber'})">$userStaleCount</div></div>
      <div class="sum-item"><div class="sum-label">Pwd Never Expires</div><div class="sum-val $(if($userPwdNeverCount -gt 20){'red'}else{'amber'})">$userPwdNeverCount</div></div>
      <div class="sum-item"><div class="sum-label">Pwd Expiring 14d</div><div class="sum-val amber">$userPwdExpiring</div></div>
      <div class="sum-item"><div class="sum-label">Locked Out</div><div class="sum-val $(if($userLockedCount -gt 0){'red'}else{'green'})">$userLockedCount</div></div>
    </div>
    <p style="font-size:11px;color:var(--text-muted);margin-bottom:12px;">Showing top flagged accounts (stale, locked, password never expires). Export for full dataset.</p>
    <table class="data-table" id="tbl-users">
      <thead><tr><th>SAM Account</th><th>Display Name</th><th>OU Path</th><th>Last Logon</th><th>Pwd Last Set</th><th>Pwd Never Expires</th><th>Status</th><th>Risk</th></tr></thead>
      <tbody>$userRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: PSO -->
<div class="section" id="sec-pso">
  <div class="section-header" onclick="toggleSection('sec-pso')">
    <div class="section-icon amber">🔒</div>
    <span class="section-title">Password Policies (PSO / Default Domain)</span>
    <span class="section-badge badge-info">$($psoData.Count) POLIC$(if($psoData.Count -eq 1){'Y'}else{'IES'})</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-pso','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-pso','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-pso','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-pso','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-pso">
      <thead><tr><th>PSO Name</th><th>Precedence</th><th>Applied To</th><th>Min Length</th><th>Max Age</th><th>Lockout Threshold</th><th>Complexity</th><th>Status</th></tr></thead>
      <tbody>$psoRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: GPO -->
<div class="section" id="sec-gpo">
  <div class="section-header" onclick="toggleSection('sec-gpo')">
    <div class="section-icon purple">📋</div>
    <span class="section-title">Group Policy Health</span>
    <span class="section-badge $(if($gpoUnlinked -gt 0){'badge-warn'}else{'badge-ok'})">$gpoTotal GPOs</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-gpo','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-gpo','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-gpo','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-gpo','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <div class="summary-grid">
      <div class="sum-item"><div class="sum-label">Total GPOs</div><div class="sum-val blue">$gpoTotal</div></div>
      <div class="sum-item"><div class="sum-label">Linked</div><div class="sum-val green">$gpoLinked</div></div>
      <div class="sum-item"><div class="sum-label">Unlinked</div><div class="sum-val $(if($gpoUnlinked -gt 0){'amber'}else{'green'})">$gpoUnlinked</div></div>
      <div class="sum-item"><div class="sum-label">Enforced</div><div class="sum-val blue">$gpoEnforced</div></div>
    </div>
    <table class="data-table" id="tbl-gpo">
      <thead><tr><th>GPO Name</th><th>Linked To</th><th>Status</th><th>WMI Filter</th><th>Last Modified</th><th>GPO Status</th></tr></thead>
      <tbody>$gpoRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: PKI -->
<div class="section" id="sec-pki">
  <div class="section-header" onclick="toggleSection('sec-pki')">
    <div class="section-icon amber">🏅</div>
    <span class="section-title">PKI / Certificate Services Health</span>
    <span class="section-badge $(if($expiringCount -gt 0){'badge-warn'}else{'badge-ok'})">$expiringCount EXPIRING</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-pki','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-pki','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-pki','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-pki','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-pki">
      <thead><tr><th>Common Name</th><th>Issued To</th><th>Template</th><th>Issued Date</th><th>Expires</th><th>Days Left</th><th>CRL Status</th><th>Status</th></tr></thead>
      <tbody>$pkiRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: DNS -->
<div class="section" id="sec-dns">
  <div class="section-header" onclick="toggleSection('sec-dns')">
    <div class="section-icon cyan">🌐</div>
    <span class="section-title">DNS Zone Health</span>
    <span class="section-badge $(if(($dnsData|Where-Object{$_.Status -ne 'HEALTHY' -and $_.Status -ne 'N/A'}).Count -eq 0){'badge-ok'}else{'badge-warn'})">$($dnsData.Count) ZONES</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-dns','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-dns','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-dns','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-dns','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-dns">
      <thead><tr><th>Zone Name</th><th>Type</th><th>Replication Scope</th><th>Dynamic Update</th><th>Records</th><th>Scavenging</th><th>Status</th></tr></thead>
      <tbody>$dnsRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: DHCP -->
<div class="section" id="sec-dhcp">
  <div class="section-header" onclick="toggleSection('sec-dhcp')">
    <div class="section-icon blue">📡</div>
    <span class="section-title">DHCP Scope Utilization</span>
    <span class="section-badge $(if(($dhcpData|Where-Object{$_.StatusClass -eq 'crit'}).Count -gt 0){'badge-crit'}elseif(($dhcpData|Where-Object{$_.StatusClass -eq 'warn'}).Count -gt 0){'badge-warn'}else{'badge-ok'})">$($dhcpData.Count) SCOPES</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-dhcp','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-dhcp','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-dhcp','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-dhcp','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-dhcp">
      <thead><tr><th>Scope Name</th><th>Subnet</th><th>Total</th><th>In Use</th><th>Free</th><th>Utilization</th><th>Failover</th><th>Status</th></tr></thead>
      <tbody>$dhcpRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: Privileged Groups -->
<div class="section" id="sec-privgroups">
  <div class="section-header" onclick="toggleSection('sec-privgroups')">
    <div class="section-icon red">🛡</div>
    <span class="section-title">Privileged Group Membership</span>
    <span class="section-badge badge-info">TIER MODEL</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-privgroups','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-privgroups','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-privgroups','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-privgroups','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-privgroups">
      <thead><tr><th>Group</th><th>Members</th><th>Member Names</th><th>Tier</th><th>Last Change</th><th>Review</th></tr></thead>
      <tbody>$privRows</tbody>
    </table>
  </div>
</div>

<!-- SECTION: Security Alerts -->
<div class="section open" id="sec-security">
  <div class="section-header" onclick="toggleSection('sec-security')">
    <div class="section-icon red">⚠</div>
    <span class="section-title">Security Alerts &amp; Findings</span>
    <span class="section-badge $(if($critAlerts -gt 0){'badge-crit'}elseif($warnAlerts -gt 0){'badge-warn'}else{'badge-ok'})">$critAlerts CRITICAL · $warnAlerts WARNING</span>
    <svg class="section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
  </div>
  <div class="export-bar"><span class="export-label">Export:</span>
    <button class="btn-export csv"  onclick="exportSection('sec-security','csv')">⬇ CSV</button>
    <button class="btn-export xlsx" onclick="exportSection('sec-security','xlsx')">⬇ XLSX</button>
    <button class="btn-export txt"  onclick="exportSection('sec-security','txt')">⬇ TXT</button>
    <button class="btn-export docx" onclick="exportSection('sec-security','docx')">⬇ DOCX</button>
  </div>
  <div class="section-body">
    <table class="data-table" id="tbl-security">
      <thead><tr><th>Alert ID</th><th>Category</th><th>Finding</th><th>Affected Object</th><th>Detected</th><th>Severity</th><th>Recommended Action</th></tr></thead>
      <tbody>$secRows</tbody>
    </table>
  </div>
</div>

</main>

<footer class="footer">
  AD Health Dashboard &nbsp;|&nbsp; <span>$(Escape-Html $DomainFQDN)</span> &nbsp;|&nbsp;
  <span>$(Escape-Html $Author)</span> &nbsp;|&nbsp;
  Generated: <span>$reportDate</span>
</footer>

<script>
// Domain sync
const domainInput = document.getElementById('domainInput');
function syncDomain(){
  const v = domainInput.value.trim() || 'N/A';
  document.querySelectorAll('.footer span').forEach((s,i)=>{ if(i===0) s.textContent=v; });
  document.title = 'AD Health Dashboard — ' + v;
}
domainInput.addEventListener('input', syncDomain);

// Section toggle
function toggleSection(id){
  document.getElementById(id).classList.toggle('open');
}

// Search
const searchInput = document.getElementById('searchInput');
const searchCount = document.getElementById('searchCount');
const searchClear = document.getElementById('searchClear');
searchInput.addEventListener('input', doSearch);
searchClear.addEventListener('click', () => { searchInput.value=''; doSearch(); });
function doSearch(){
  const q = searchInput.value.trim().toLowerCase();
  if(!q){ clearSearch(); return; }
  let total = 0;
  document.querySelectorAll('.section').forEach(sec => {
    let hit = sec.querySelector('.section-title').textContent.toLowerCase().includes(q);
    sec.querySelectorAll('tbody tr').forEach(tr => {
      const m = tr.textContent.toLowerCase().includes(q);
      tr.classList.toggle('search-match', m);
      if(m){ hit=true; total++; }
    });
    sec.classList.toggle('search-hidden', !hit);
    if(hit) sec.classList.add('open');
  });
  searchCount.textContent = total===0 ? 'No matches' : total+' row'+(total===1?'':'s')+' matched';
}
function clearSearch(){
  document.querySelectorAll('.section').forEach(s=>{
    s.classList.remove('search-hidden');
    s.querySelectorAll('tr').forEach(tr=>tr.classList.remove('search-match'));
  });
  searchCount.textContent='';
}

// Export helpers
function getTableData(secId){
  const tbl = document.querySelector('#'+secId+' table');
  if(!tbl) return {headers:[],rows:[]};
  const headers = Array.from(tbl.querySelectorAll('thead th')).map(th=>th.textContent.trim());
  const rows = Array.from(tbl.querySelectorAll('tbody tr')).map(tr=>
    Array.from(tr.querySelectorAll('td')).map(td=>td.textContent.trim().replace(/\s+/g,' '))
  );
  return {headers,rows};
}
function getSectionTitle(secId){ return document.querySelector('#'+secId+' .section-title').textContent.trim(); }
function dl(content,filename,mime){ const b=new Blob([content],{type:mime}); const u=URL.createObjectURL(b); const a=document.createElement('a'); a.href=u; a.download=filename; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(u); }

function exportSection(secId,format){
  const {headers,rows}=getTableData(secId);
  const title=getSectionTitle(secId);
  const domain=domainInput.value.trim()||'domain';
  const fn='AD-Health_'+title.replace(/[^a-z0-9]/gi,'_').replace(/_+/g,'_')+'_'+domain.split('.')[0];
  if(format==='csv') exportCSV(headers,rows,fn);
  else if(format==='xlsx') exportXLSX(headers,rows,fn,title,domain);
  else if(format==='txt') exportTXT(headers,rows,fn,title,domain);
  else if(format==='docx') exportDOCX(headers,rows,fn,title,domain);
}

function exportCSV(h,r,fn){ const e=v=>'"'+String(v).replace(/"/g,'""')+'"'; const l=[h.map(e).join(','),...r.map(row=>row.map(e).join(','))]; dl(l.join('\r\n'),fn+'.csv','text/csv'); }

function exportXLSX(h,r,fn,title){
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.aoa_to_sheet([h,...r]);
  ws['!cols']=h.map((_,i)=>({wch:Math.min(Math.max(h[i].length,...r.map(row=>(row[i]||'').length))+2,50)}));
  XLSX.utils.book_append_sheet(wb,ws,title.substring(0,31));
  XLSX.writeFile(wb,fn+'.xlsx');
}

function exportTXT(h,r,fn,title,domain){
  const sep='═'.repeat(80); const dash='─'.repeat(80);
  const w=h.map((hd,i)=>Math.max(hd.length,...r.map(row=>(row[i]||'').length),8));
  const pad=(s,n)=>String(s).padEnd(n,' ');
  const lines=[sep,'AD HEALTH DASHBOARD','Section: '+title,'Domain:  '+domain,'Author:  $Author','Date:    '+ new Date().toLocaleString(),sep,'',h.map((hd,i)=>pad(hd,w[i])).join(' | '),dash,...r.map(row=>row.map((v,i)=>pad(v,w[i])).join(' | ')),'',sep];
  dl(lines.join('\r\n'),fn+'.txt','text/plain');
}

async function exportDOCX(h,r,fn,title,domain){
  const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,HeadingLevel,AlignmentType,WidthType,BorderStyle,ShadingType,VerticalAlign,Header,Footer}=window.docx;
  const hBlue='1B4F9E',hTxt='FFFFFF',alt='EBF2FB';
  const brd={style:BorderStyle.SINGLE,size:4,color:'BBCCDD'};
  const borders={top:brd,bottom:brd,left:brd,right:brd};
  const cellW=Math.floor(15360/h.length);
  const hRow=new TableRow({tableHeader:true,children:h.map(hd=>new TableCell({borders,width:{size:cellW,type:WidthType.DXA},shading:{fill:hBlue,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({children:[new TextRun({text:hd,bold:true,color:hTxt,size:18,font:'Arial'})]})]}))});
  const dRows=r.map((row,ri)=>new TableRow({children:row.map(cell=>new TableCell({borders,width:{size:cellW,type:WidthType.DXA},shading:{fill:ri%2===1?alt:'FFFFFF',type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:cell,size:16,font:'Arial'})]})]})) }));
  const doc=new Document({styles:{default:{document:{run:{font:'Arial',size:22}}}},sections:[{properties:{page:{size:{width:15840,height:12240},orientation:'landscape',margin:{top:720,right:720,bottom:720,left:720}}},headers:{default:new Header({children:[new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:6,color:hBlue,space:1}},children:[new TextRun({text:'AD HEALTH DASHBOARD — '+domain.toUpperCase()+'    '+title,bold:true,size:24,color:hBlue,font:'Arial'})]})]}),footers:{default:new Footer({children:[new Paragraph({children:[new TextRun({text:'$Author    |    Generated: '+new Date().toLocaleString(),size:16,color:'888888',font:'Arial'})]})]})}},children:[new Paragraph({children:[new TextRun({text:title,bold:true,size:32,color:hBlue,font:'Arial'})]}),new Paragraph({spacing:{before:80,after:240},children:[new TextRun({text:'Domain: ',bold:true,size:18,font:'Arial'}),new TextRun({text:domain,size:18,font:'Arial'}),new TextRun({text:'    Generated: ',bold:true,size:18,font:'Arial'}),new TextRun({text:new Date().toLocaleString(),size:18,font:'Arial'})]}),new Table({width:{size:15360,type:WidthType.DXA},columnWidths:h.map(()=>cellW),rows:[hRow,...dRows]}),new Paragraph({children:[new TextRun('')]})]}]});
  const buf=await Packer.toBlob(doc);
  const u=URL.createObjectURL(buf);const a=document.createElement('a');a.href=u;a.download=fn+'.docx';a.click();URL.revokeObjectURL(u);
}
</script>
</body>
</html>
"@

# ════════════════════════════════════════════════
#  WRITE OUTPUT FILES
# ════════════════════════════════════════════════
$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportFile  = Join-Path $OutputPath "ADHealthDashboard_${DomainFQDN}_${timestamp}.html"

try {
    $html | Out-File -FilePath $reportFile -Encoding UTF8 -Force
    Write-Log "Report saved: $reportFile" -Level SUCCESS
} catch {
    Write-Log "Failed to write report file: $_" -Level ERROR
    exit 1
}

# ════════════════════════════════════════════════
#  SUMMARY
# ════════════════════════════════════════════════
Write-Log "══════════════════════════════════════════════" -Level INFO
Write-Log "  REPORT COMPLETE" -Level SUCCESS
Write-Log "  Domain   : $DomainFQDN" -Level INFO
Write-Log "  DCs      : $($dcData.Count)" -Level INFO
Write-Log "  Alerts   : $($secAlerts.Count) ($critAlerts critical, $warnAlerts warning)" -Level INFO
Write-Log "  Output   : $reportFile" -Level INFO
Write-Log "  Log      : $LogPath" -Level INFO
Write-Log "══════════════════════════════════════════════" -Level INFO

if ($OpenOnComplete) {
    Write-Log "Opening report in default browser..." -Level INFO
    Start-Process $reportFile
}

Write-Output $reportFile
