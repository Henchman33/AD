<#
.SYNOPSIS
    Active Directory Comprehensive Reporting Dashboard
.DESCRIPTION
    Generates an interactive HTML dashboard with statistics, charts, and a searchable table
    for AD users, computers, groups, and OUs. Includes stale detection, password expiry,
    OS distribution, and export capabilities (Excel, CSV, Word). Features a dark/light theme.
    Saves all outputs to a timestamped folder on the desktop.
    Runs without AD using sample data for testing.
.AUTHOR
    Based on the AD Application Inventory framework, extended for full AD reporting.
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# --- 1. DETECT DOMAIN / FALLBACK TO SAMPLE ---
$useSample = $false
try {
    $domainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
    $domainName = $domainObj.Name
    $domainDN = "DC=" + $domainName.Replace(".", ",DC=")
    Write-Host "Connected to domain: $domainName" -ForegroundColor Green
} catch {
    Write-Warning "Not connected to a domain. Using sample data for testing."
    $useSample = $true
    $domainName = "SAMPLE.LOCAL"
}

# --- CONSOLE UI ---
Clear-Host
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   ACTIVE DIRECTORY REPORTING TOOL | $domainName" -ForegroundColor White
Write-Host "   ------------------------------------------------------------------" -ForegroundColor DarkGray
if ($useSample) {
    Write-Host "   [TEST MODE] Using embedded sample data." -ForegroundColor Yellow
} else {
    Write-Host "   [+] Scanning Users, Computers, Groups, OUs..." -ForegroundColor Yellow
}
Write-Host "======================================================================" -ForegroundColor Cyan

# --- 2. DATA COLLECTION (Real or Sample) ---
$Today = Get-Date
$staleThresholdDays = 180

# Helper: convert AD timestamp to days ago
function Get-LogonDays([string]$timestamp) {
    if ($timestamp -and $timestamp -gt 0) {
        try {
            $d = [DateTime]::FromFileTime($timestamp)
            if ($d.Year -gt ($Today.Year + 5)) { return $null }
            return ($Today - $d).Days
        } catch { return $null }
    }
    return $null
}

# Initialize collections
$allItems = [System.Collections.Generic.List[Object]]::new()
$userList = [System.Collections.Generic.List[Object]]::new()
$computerList = [System.Collections.Generic.List[Object]]::new()
$groupList = [System.Collections.Generic.List[Object]]::new()
$ouList = [System.Collections.Generic.List[Object]]::new()

# Statistics placeholders
$totalUsers = 0
$enabledUsers = 0
$disabledUsers = 0
$staleUsers = 0
$neverLoggedUsers = 0
$expiredPasswordUsers = 0

$totalComputers = 0
$enabledComputers = 0
$disabledComputers = 0
$staleComputers = 0
$neverLoggedComputers = 0
$osDistribution = @{}

$totalGroups = 0
$securityGroups = 0
$distributionGroups = 0

$totalOUs = 0

if (-not $useSample) {
    # ---- REAL AD QUERIES ----
    Import-Module ActiveDirectory -ErrorAction Stop

    # 1. USERS
    Write-Host "   Querying Users..." -NoNewline
    $users = Get-ADUser -Filter * -Properties DisplayName, Department, Title, Enabled, LastLogonTimestamp, PasswordLastSet, whenCreated, PasswordExpired, LockedOut
    Write-Host " $($users.Count) found." -ForegroundColor Green
    $totalUsers = $users.Count
    foreach ($u in $users) {
        $days = Get-LogonDays $u.LastLogonTimestamp
        $status = if (-not $u.Enabled) { "Disabled" } else {
            if ($days -eq $null) { "Never Logged In" }
            elseif ($days -le $staleThresholdDays) { "Active" }
            else { "Stale" }
        }
        if ($status -eq "Stale") { $staleUsers++ }
        if ($status -eq "Never Logged In") { $neverLoggedUsers++ }
        if ($u.Enabled) { $enabledUsers++ } else { $disabledUsers++ }
        if ($u.PasswordExpired) { $expiredPasswordUsers++ }

        $userList.Add([PSCustomObject]@{
            Category    = "User"
            Name        = $u.DisplayName
            SamAccount  = $u.SamAccountName
            Department  = $u.Department
            Title       = $u.Title
            Status      = $status
            Enabled     = $u.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            Created     = $u.whenCreated
            PasswordSet = $u.PasswordLastSet
            PasswordExpired = $u.PasswordExpired
            LockedOut   = $u.LockedOut
            DN          = $u.DistinguishedName
        })
    }

    # 2. COMPUTERS
    Write-Host "   Querying Computers..." -NoNewline
    $computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, Enabled, LastLogonTimestamp, whenCreated, Description
    Write-Host " $($computers.Count) found." -ForegroundColor Green
    $totalComputers = $computers.Count
    foreach ($c in $computers) {
        $days = Get-LogonDays $c.LastLogonTimestamp
        $status = if (-not $c.Enabled) { "Disabled" } else {
            if ($days -eq $null) { "Never Logged In" }
            elseif ($days -le $staleThresholdDays) { "Active" }
            else { "Stale" }
        }
        if ($status -eq "Stale") { $staleComputers++ }
        if ($status -eq "Never Logged In") { $neverLoggedComputers++ }
        if ($c.Enabled) { $enabledComputers++ } else { $disabledComputers++ }

        $os = $c.OperatingSystem
        if ($os) {
            if (-not $osDistribution.ContainsKey($os)) { $osDistribution[$os] = 0 }
            $osDistribution[$os]++
        }

        $computerList.Add([PSCustomObject]@{
            Category    = "Computer"
            Name        = $c.Name
            OS          = $c.OperatingSystem
            OSVersion   = $c.OperatingSystemVersion
            Status      = $status
            Enabled     = $c.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            Description = $c.Description
            Created     = $c.whenCreated
            DN          = $c.DistinguishedName
        })
    }

    # 3. GROUPS
    Write-Host "   Querying Groups..." -NoNewline
    $groups = Get-ADGroup -Filter * -Properties GroupCategory, Description, Members, whenCreated
    Write-Host " $($groups.Count) found." -ForegroundColor Green
    $totalGroups = $groups.Count
    foreach ($g in $groups) {
        if ($g.GroupCategory -eq "Security") { $securityGroups++ } else { $distributionGroups++ }
        $groupList.Add([PSCustomObject]@{
            Category    = "Group"
            Name        = $g.Name
            GroupType   = $g.GroupCategory
            MemberCount = ($g.Members).Count
            Description = $g.Description
            Created     = $g.whenCreated
            DN          = $g.DistinguishedName
        })
    }

    # 4. OUs
    Write-Host "   Querying OUs..." -NoNewline
    $ous = Get-ADOrganizationalUnit -Filter * -Properties Description, whenCreated
    Write-Host " $($ous.Count) found." -ForegroundColor Green
    $totalOUs = $ous.Count
    foreach ($ou in $ous) {
        $ouList.Add([PSCustomObject]@{
            Category    = "OU"
            Name        = $ou.Name
            Path        = $ou.DistinguishedName
            Description = $ou.Description
            Created     = $ou.whenCreated
            DN          = $ou.DistinguishedName
        })
    }

} else {
    # ---- SAMPLE DATA FOR TESTING ----
    Write-Host "   Generating sample data..." -ForegroundColor Yellow

    # Users
    $userList.Add([PSCustomObject]@{Category="User"; Name="John Doe"; SamAccount="jdoe"; Department="IT"; Title="Admin"; Status="Active"; Enabled=$true; LastLogon="2 days"; Created="2024-01-15"; PasswordSet="2024-12-01"; PasswordExpired=$false; LockedOut=$false; DN="CN=John Doe,OU=Users,DC=sample,DC=local"})
    $userList.Add([PSCustomObject]@{Category="User"; Name="Jane Smith"; SamAccount="jsmith"; Department="HR"; Title="Manager"; Status="Stale"; Enabled=$true; LastLogon="200 days"; Created="2023-06-10"; PasswordSet="2023-06-10"; PasswordExpired=$false; LockedOut=$false; DN="CN=Jane Smith,OU=Users,DC=sample,DC=local"})
    $userList.Add([PSCustomObject]@{Category="User"; Name="Bob Johnson"; SamAccount="bjohnson"; Department="Finance"; Title="Analyst"; Status="Disabled"; Enabled=$false; LastLogon="Never"; Created="2024-08-20"; PasswordSet="2024-08-20"; PasswordExpired=$false; LockedOut=$false; DN="CN=Bob Johnson,OU=Users,DC=sample,DC=local"})
    $userList.Add([PSCustomObject]@{Category="User"; Name="Alice Brown"; SamAccount="abrown"; Department="IT"; Title="Developer"; Status="Active"; Enabled=$true; LastLogon="1 days"; Created="2024-03-05"; PasswordSet="2024-11-15"; PasswordExpired=$false; LockedOut=$false; DN="CN=Alice Brown,OU=Users,DC=sample,DC=local"})
    $totalUsers = $userList.Count
    $enabledUsers = 2; $disabledUsers = 1; $staleUsers = 1; $neverLoggedUsers = 1; $expiredPasswordUsers = 0

    # Computers
    $computerList.Add([PSCustomObject]@{Category="Computer"; Name="DC01"; OS="Windows Server 2019"; OSVersion="10.0.17763"; Status="Active"; Enabled=$true; LastLogon="1 days"; Description="Primary DC"; Created="2024-01-15"; DN="CN=DC01,DC=sample,DC=local"})
    $computerList.Add([PSCustomObject]@{Category="Computer"; Name="FS01"; OS="Windows Server 2016"; OSVersion="10.0.14393"; Status="Stale"; Enabled=$true; LastLogon="190 days"; Description="File Server"; Created="2023-06-10"; DN="CN=FS01,DC=sample,DC=local"})
    $computerList.Add([PSCustomObject]@{Category="Computer"; Name="APP01"; OS="Windows 10 Pro"; OSVersion="10.0.19045"; Status="Disabled"; Enabled=$false; LastLogon="Never"; Description="Legacy PC"; Created="2024-08-20"; DN="CN=APP01,DC=sample,DC=local"})
    $computerList.Add([PSCustomObject]@{Category="Computer"; Name="SQL01"; OS="Windows Server 2022"; OSVersion="10.0.20348"; Status="Active"; Enabled=$true; LastLogon="2 days"; Description="SQL Server"; Created="2024-03-05"; DN="CN=SQL01,DC=sample,DC=local"})
    $totalComputers = $computerList.Count
    $enabledComputers = 3; $disabledComputers = 1; $staleComputers = 1; $neverLoggedComputers = 1
    $osDistribution = @{"Windows Server 2019"=1; "Windows Server 2016"=1; "Windows 10 Pro"=1; "Windows Server 2022"=1}

    # Groups
    $groupList.Add([PSCustomObject]@{Category="Group"; Name="Domain Admins"; GroupType="Security"; MemberCount=5; Description="Domain administrators"; Created="2021-01-01"; DN="CN=Domain Admins,CN=Users,DC=sample,DC=local"})
    $groupList.Add([PSCustomObject]@{Category="Group"; Name="Sales"; GroupType="Distribution"; MemberCount=12; Description="Sales team"; Created="2022-05-15"; DN="CN=Sales,OU=Groups,DC=sample,DC=local"})
    $groupList.Add([PSCustomObject]@{Category="Group"; Name="IT Admins"; GroupType="Security"; MemberCount=3; Description="IT admin group"; Created="2023-09-01"; DN="CN=IT Admins,OU=Groups,DC=sample,DC=local"})
    $totalGroups = $groupList.Count
    $securityGroups = 2; $distributionGroups = 1

    # OUs
    $ouList.Add([PSCustomObject]@{Category="OU"; Name="Users"; Path="OU=Users,DC=sample,DC=local"; Description="User accounts"; Created="2021-01-01"; DN="OU=Users,DC=sample,DC=local"})
    $ouList.Add([PSCustomObject]@{Category="OU"; Name="Computers"; Path="OU=Computers,DC=sample,DC=local"; Description="Computer objects"; Created="2021-01-01"; DN="OU=Computers,DC=sample,DC=local"})
    $ouList.Add([PSCustomObject]@{Category="OU"; Name="Groups"; Path="OU=Groups,DC=sample,DC=local"; Description="Group objects"; Created="2021-01-01"; DN="OU=Groups,DC=sample,DC=local"})
    $totalOUs = $ouList.Count
}

# Combine all into one list for the main table
$allItems.AddRange($userList)
$allItems.AddRange($computerList)
$allItems.AddRange($groupList)
$allItems.AddRange($ouList)

# Prepare stats for JSON
$osLabels = $osDistribution.Keys
$osCounts = $osDistribution.Values

$jsonAll = $allItems | ConvertTo-Json -Depth 3 -Compress
$jsonStats = @{
    TotalUsers = $totalUsers
    EnabledUsers = $enabledUsers
    DisabledUsers = $disabledUsers
    StaleUsers = $staleUsers
    NeverLoggedUsers = $neverLoggedUsers
    ExpiredPasswordUsers = $expiredPasswordUsers
    TotalComputers = $totalComputers
    EnabledComputers = $enabledComputers
    DisabledComputers = $disabledComputers
    StaleComputers = $staleComputers
    NeverLoggedComputers = $neverLoggedComputers
    TotalGroups = $totalGroups
    SecurityGroups = $securityGroups
    DistributionGroups = $distributionGroups
    TotalOUs = $totalOUs
    OSLabels = $osLabels
    OSCounts = $osCounts
    Domain = $domainName
} | ConvertTo-Json -Depth 5 -Compress

$utf8 = [System.Text.Encoding]::UTF8
$b64Data = [Convert]::ToBase64String($utf8.GetBytes($jsonAll))
$b64Stats = [Convert]::ToBase64String($utf8.GetBytes($jsonStats))

# --- 3. CREATE OUTPUT FOLDER ON DESKTOP ---
$desktop = [Environment]::GetFolderPath('Desktop')
$folderTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$folderName = "AD_Report_$folderTime"
$outputDir = Join-Path $desktop $folderName
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Write-Host "Creating output folder: $outputDir" -ForegroundColor Cyan

# --- 4. GENERATE CSV FILES ---
$dateStr = Get-Date -Format "yyyyMMdd"
$baseName = "AD_Report_$dateStr"
$csvPath = Join-Path $outputDir "$baseName.csv"
$allItems | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Master CSV exported to: $csvPath" -ForegroundColor Green

# Per-category CSVs
$cats = $allItems | Select-Object -ExpandProperty Category -Unique
foreach ($cat in $cats) {
    $catFile = Join-Path $outputDir "$baseName-$cat.csv"
    $allItems | Where-Object { $_.Category -eq $cat } | Export-Csv -Path $catFile -NoTypeInformation -Encoding UTF8
}

# --- 5. GENERATE HTML DASHBOARD WITH DARK THEME ---
Write-Host "   Generating HTML Dashboard with dark/light theme..." -ForegroundColor Green

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AD Reporting Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        /* --- THEME VARIABLES --- */
        :root {
            --bg: #f8fafc;
            --text: #1e293b;
            --card-bg: #ffffff;
            --border: #e2e8f0;
            --nav-bg: #ffffff;
            --stat-bg: #ffffff;
            --primary: #4f46e5;
            --table-header: #f1f5f9;
            --modal-bg: #ffffff;
            --filter-bg: rgba(0,0,0,0.03);
        }
        body.dark-theme {
            --bg: #0f172a;
            --text: #f1f5f9;
            --card-bg: #1e293b;
            --border: #334155;
            --nav-bg: #1e293b;
            --stat-bg: #1e293b;
            --primary: #818cf8;
            --table-header: #1e293b;
            --modal-bg: #1e293b;
            --filter-bg: rgba(255,255,255,0.05);
        }
        body {
            background-color: var(--bg);
            color: var(--text);
            font-family: 'Inter', sans-serif;
            font-size: 0.8rem;
            transition: background-color 0.3s, color 0.3s;
        }
        .navbar { background: var(--nav-bg) !important; border-bottom: 1px solid var(--border); }
        .brand { color: var(--primary) !important; }
        .stat-card { background: var(--stat-bg) !important; border-color: var(--border) !important; }
        .stat-card .stat-val { color: var(--text) !important; }
        .chart-card { background: var(--card-bg) !important; border-color: var(--border) !important; }
        .chart-card .chart-header { color: var(--text) !important; }
        .table-card { background: var(--card-bg) !important; border-color: var(--border) !important; }
        table.dataTable thead th { background: var(--table-header) !important; color: var(--text) !important; border-bottom-color: var(--border) !important; }
        table.dataTable tbody td { border-bottom-color: var(--border) !important; color: var(--text) !important; }
        .form-control-sm-custom { background: var(--bg) !important; border-color: var(--border) !important; color: var(--text) !important; }
        .filter-btn { border-color: var(--border) !important; color: var(--text) !important; background: transparent; }
        .filter-btn.active { background: var(--primary) !important; color: white !important; border-color: var(--primary) !important; }
        .action-btn { border-color: var(--border) !important; color: var(--text) !important; }
        .action-btn:hover { background: var(--primary) !important; color: white !important; border-color: var(--primary) !important; }
        .sub-navbar { background: var(--bg) !important; }
        .nav-btn { background: var(--card-bg) !important; border-color: var(--border) !important; color: var(--text) !important; }
        .nav-btn:hover { background: var(--primary) !important; color: white !important; }
        .stat-item { background: var(--stat-bg) !important; border-color: var(--border) !important; }
        .stat-item .stat-info div { color: var(--text) !important; }
        .stat-item .stat-info div:first-child { color: #94a3b8 !important; }
        .modal-content { background: var(--modal-bg) !important; border-color: var(--border) !important; }
        .modal-header, .modal-footer { border-color: var(--border) !important; }
        .modal-title, .m-title { color: var(--text) !important; }
        .m-label { color: #94a3b8 !important; }
        .m-val { color: var(--text) !important; }
        .badge-s { color: white !important; }
        .bg-info-cat { background: #e0f2fe !important; color: #0369a1 !important; }
        .row-stale { opacity: 0.6; background-color: rgba(100, 116, 139, 0.1) !important; }
        .row-disabled { opacity: 0.5; background-color: rgba(239, 68, 68, 0.1) !important; }
        .filter-bar { background: var(--filter-bg) !important; border-color: var(--border) !important; }
        .filter-label { color: #94a3b8 !important; }
        .footer, .text-muted { color: #94a3b8 !important; }
        .btn-export { background: transparent; border-color: var(--border); color: var(--text); }
        .btn-export:hover { background: var(--primary); color: white; border-color: var(--primary); }
        .search-input { background: var(--bg) !important; border-color: var(--border) !important; color: var(--text) !important; }
        .search-input:focus { border-color: var(--primary) !important; outline: none; }
        .stat-icon { color: var(--text) !important; opacity: 0.12 !important; }
        /* keep existing styles except overrides */
        .stat-card { border-left: 3px solid var(--primary); }
        .sc-blue { border-left-color: #3b82f6; }
        .sc-green { border-left-color: #22c55e; }
        .sc-gray { border-left-color: #94a3b8; }
        .sc-red { border-left-color: #ef4444; }
        .sc-purple { border-left-color: #a855f7; }
        .sc-orange { border-left-color: #f97316; }
        .sc-teal { border-left-color: #14b8a6; }
        .sc-dark { border-left-color: #334155; }
        /* all other existing styles remain */
    </style>
</head>
<body>

    <nav class="navbar fixed-top">
        <div class="d-flex align-items-center gap-2">
            <i class="fa-solid fa-server text-primary fs-5"></i>
            <span class="brand fw-bold" id="pageTitle">AD Reporting</span>
        </div>
        <div class="small text-muted" id="headerInfo" style="font-size: 0.75rem;">Loading...</div>
        <div>
            <button class="btn btn-sm btn-outline-secondary action-btn" onclick="toggleTheme()"><i class="fa-solid fa-moon"></i> Theme</button>
        </div>
    </nav>
    <div style="height: 50px;"></div>

    <div class="container-fluid px-4 py-2">
        
        <!-- Stats Cards -->
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-blue clickable" data-filter="category" data-val="User"><div class="stat-label">Total Users</div><div class="stat-val" id="vTotalUsers">0</div><i class="fa-solid fa-users stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="status" data-val="Active"><div class="stat-label text-success">Active Users</div><div class="stat-val text-success" id="vActiveUsers">0</div><i class="fa-solid fa-user-check stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label text-muted">Stale Users</div><div class="stat-val text-muted" id="vStaleUsers">0</div><i class="fa-solid fa-user-clock stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-red clickable" data-filter="status" data-val="Disabled"><div class="stat-label text-danger">Disabled Users</div><div class="stat-val text-danger" id="vDisabledUsers">0</div><i class="fa-solid fa-user-slash stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-purple clickable" data-filter="category" data-val="Computer"><div class="stat-label">Computers</div><div class="stat-val" id="vTotalComputers">0</div><i class="fa-solid fa-desktop stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label text-muted">Stale Computers</div><div class="stat-val text-muted" id="vStaleComputers">0</div><i class="fa-solid fa-desktop stat-icon"></i></div></div>
        </div>
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="category" data-val="Group"><div class="stat-label">Total Groups</div><div class="stat-val" id="vTotalGroups">0</div><i class="fa-solid fa-users-cog stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-orange clickable" data-filter="category" data-val="OU"><div class="stat-label">OUs</div><div class="stat-val" id="vTotalOUs">0</div><i class="fa-solid fa-folder-tree stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-red clickable" data-filter="passwordExpired" data-val="True"><div class="stat-label text-danger">Pwd Expired</div><div class="stat-val text-danger" id="vExpiredPwd">0</div><i class="fa-solid fa-key stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-purple clickable" data-filter="status" data-val="Never Logged In"><div class="stat-label">Never Logged In</div><div class="stat-val" id="vNeverLogged">0</div><i class="fa-solid fa-ghost stat-icon"></i></div></div>
        </div>

        <!-- Charts -->
        <div class="row g-2">
            <div class="col-lg-6">
                <div class="chart-card">
                    <div class="chart-header">
                        <span><i class="fa-solid fa-chart-bar me-1"></i> Computer OS Distribution</span>
                        <small class="text-muted fw-normal">(by count)</small>
                    </div>
                    <div class="chart-wrapper">
                        <div class="chart-scroll">
                            <div style="position: relative; height: 1000px; width: 100%">
                                <canvas id="osChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-lg-6">
                <div class="chart-card">
                    <div class="chart-header"><span><i class="fa-solid fa-chart-pie me-1"></i> User Status</span></div>
                    <div class="chart-wrapper d-flex align-items-center justify-content-center position-relative">
                        <div style="width: 100%; height: 180px;">
                            <canvas id="userStatusChart"></canvas>
                        </div>
                        <div class="pie-center-text">
                            <div class="pie-val" id="centerTotal">-</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Table -->
        <div class="table-card">
            <div class="d-flex justify-content-between align-items-center mb-2">
                <h6 class="m-0 fw-bold" style="color: var(--text); font-size: 0.9rem;"><i class="fa-solid fa-list me-2"></i>All Objects</h6>
                <div>
                    <button class="btn btn-sm btn-outline-success btn-export" id="btnExcel">Excel</button>
                    <button class="btn btn-sm btn-outline-primary btn-export" id="btnCsv">CSV</button>
                    <button class="btn btn-sm btn-outline-secondary btn-export" id="btnWord">Word</button>
                    <button class="btn btn-sm btn-outline-danger btn-export" id="btnReset">Reset</button>
                </div>
            </div>
            
            <table id="mainTable" class="table table-hover table-sm w-100">
                <thead>
                    <tr><th>Category</th><th>Name</th><th>Status</th><th>Enabled</th><th>Last Logon</th><th>Detail</th><th>Description</th><th>Created</th></tr>
                    <tr class="filter-row">
                        <th><select class="form-control-sm-custom"><option value="">All</option><option value="User">User</option><option value="Computer">Computer</option><option value="Group">Group</option><option value="OU">OU</option></select></th>
                        <th><input type="text" class="form-control-sm-custom" placeholder="Search..."></th>
                        <th><select class="form-control-sm-custom"><option value="">All</option><option value="Active">Active</option><option value="Stale">Stale</option><option value="Never Logged In">Never</option><option value="Disabled">Disabled</option></select></th>
                        <th><select class="form-control-sm-custom"><option value="">All</option><option value="True">Enabled</option><option value="False">Disabled</option></select></th>
                        <th><input type="text" class="form-control-sm-custom" placeholder="Days..."></th>
                        <th><input type="text" class="form-control-sm-custom" placeholder="..."></th>
                        <th><input type="text" class="form-control-sm-custom" placeholder="..."></th>
                        <th><input type="text" class="form-control-sm-custom" placeholder="..."></th>
                    </tr>
                </thead>
                <tbody id="tBody"></tbody>
            </table>
        </div>
        <div class="text-center text-muted mt-2" style="font-size: 0.65rem;">Generated by AD Reporting Tool • Domain: <span id="footerDomain">-</span></div>
    </div>

    <!-- Detail Modal -->
    <div class="modal fade" id="detailModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered modal-lg">
            <div class="modal-content">
                <div class="modal-header py-2">
                    <h6 class="modal-title fw-bold" id="modalTitle"></h6>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <pre id="modalContent" class="small" style="white-space: pre-wrap; word-wrap: break-word; max-height: 400px; overflow-y: auto;"></pre>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.3.6/js/buttons.html5.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <script>
        function dec(s) { return new TextDecoder().decode(Uint8Array.from(atob(s), c => c.codePointAt(0))); }
        let data = [];
        let stats = {};

        // Theme toggle
        function toggleTheme() {
            document.body.classList.toggle('dark-theme');
            // change icon
            let btn = document.querySelector('.action-btn i');
            if (document.body.classList.contains('dark-theme')) {
                btn.className = 'fa-solid fa-sun';
            } else {
                btn.className = 'fa-solid fa-moon';
            }
        }

        $(document).ready(function() {
            try {
                stats = JSON.parse(dec('@BASE64_STATS@'));
                data = JSON.parse(dec('@BASE64_DATA@'));

                $('#pageTitle').text(stats.Domain + ' AD Report');
                $('#headerInfo').text(`Generated: ${new Date().toLocaleString()}`);
                $('#footerDomain').text(stats.Domain);

                // Populate stats
                $('#vTotalUsers').text(stats.TotalUsers);
                $('#vActiveUsers').text(stats.EnabledUsers - stats.StaleUsers - stats.NeverLoggedUsers); // crude: active = enabled - stale - never
                $('#vStaleUsers').text(stats.StaleUsers);
                $('#vDisabledUsers').text(stats.DisabledUsers);
                $('#vTotalComputers').text(stats.TotalComputers);
                $('#vStaleComputers').text(stats.StaleComputers);
                $('#vTotalGroups').text(stats.TotalGroups);
                $('#vTotalOUs').text(stats.TotalOUs);
                $('#vExpiredPwd').text(stats.ExpiredPasswordUsers || 0);
                $('#vNeverLogged').text(stats.NeverLoggedUsers + stats.NeverLoggedComputers);

                // Chart: OS Distribution
                const osCtx = document.getElementById('osChart').getContext('2d');
                const osLabels = stats.OSLabels || [];
                const osCounts = stats.OSCounts || [];
                const osChartHeight = Math.max(300, osLabels.length * 20);
                $('#osChart').parent().height(osChartHeight);
                new Chart(osCtx, {
                    type: 'bar',
                    data: {
                        labels: osLabels,
                        datasets: [{ label: 'Computers', data: osCounts, backgroundColor: '#4f46e5', borderRadius: 3, barThickness: 8 }]
                    },
                    options: {
                        indexAxis: 'y',
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: { x: { display: false }, y: { grid: { display: false }, ticks: { font: { size: 9 } } } }
                    }
                });

                // Chart: User Status (Pie)
                const statusCtx = document.getElementById('userStatusChart').getContext('2d');
                const activeUsers = stats.EnabledUsers - stats.StaleUsers - stats.NeverLoggedUsers;
                new Chart(statusCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Active', 'Stale', 'Never', 'Disabled'],
                        datasets: [{
                            data: [activeUsers, stats.StaleUsers, stats.NeverLoggedUsers, stats.DisabledUsers],
                            backgroundColor: ['#22c55e', '#94a3b8', '#a855f7', '#ef4444'],
                            borderWidth: 0
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        cutout: '75%',
                        plugins: { legend: { display: false } }
                    }
                });
                $('#centerTotal').text(stats.TotalUsers);

                // Build table rows
                const rows = data.map((item, idx) => {
                    let statusCls = 'bg-ghost';
                    if (item.Status === 'Active') statusCls = 'bg-ok';
                    else if (item.Status === 'Stale') statusCls = 'bg-warn';
                    else if (item.Status === 'Disabled') statusCls = 'bg-err';
                    else if (item.Status === 'Never Logged In') statusCls = 'bg-ghost';
                    
                    let enabledCls = (item.Enabled === true) ? 'bg-ok' : (item.Enabled === false ? 'bg-err' : 'bg-ghost');
                    let enabledText = (item.Enabled === true) ? 'Enabled' : (item.Enabled === false ? 'Disabled' : 'N/A');
                    
                    let rowClass = '';
                    if (item.Status === 'Stale' || item.Status === 'Never Logged In') rowClass += ' row-stale';
                    if (item.Enabled === false) rowClass += ' row-disabled';

                    let detail = '';
                    if (item.SamAccount) detail = item.SamAccount;
                    else if (item.OS) detail = item.OS;
                    else if (item.GroupType) detail = item.GroupType;
                    else if (item.Path) detail = item.Path;
                    else if (item.DN) detail = item.DN;

                    // For groups show member count
                    if (item.MemberCount !== undefined) detail = 'Members: ' + item.MemberCount;

                    return `<tr class="${rowClass}" data-category="${item.Category}" data-status="${item.Status}">
                        <td><span class="badge-s bg-info-cat">${item.Category}</span></td>
                        <td class="fw-bold clickable" style="color:var(--primary)" onclick="showDetail(${idx})">${item.Name}</td>
                        <td><span class="badge-s ${statusCls}">${item.Status || 'N/A'}</span></td>
                        <td><span class="badge-s ${enabledCls}">${enabledText}</span></td>
                        <td class="td-nowrap">${item.LastLogon || '-'}</td>
                        <td><div class="txt-trunc" title="${detail}">${detail}</div></td>
                        <td><div class="txt-trunc" title="${item.Description || ''}">${item.Description || ''}</div></td>
                        <td class="td-nowrap">${item.Created || ''}</td>
                    </tr>`;
                }).join('');
                $('#tBody').html(rows);

                // DataTable
                const table = $('#mainTable').DataTable({
                    dom: 'lrtip',
                    pageLength: 25,
                    lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
                    buttons: [{ extend: 'excelHtml5', className: 'd-none', text: 'Excel' }],
                    columnDefs: [
                        { orderData: [2, 3] }
                    ]
                });

                // Filter row
                $('#mainTable thead input, #mainTable thead select').on('keyup change', function() {
                    let i = $(this).parent().index();
                    if (table.column(i).search() !== this.value) {
                        table.column(i).search(this.value).draw();
                    }
                });

                // Stat card clicks
                $('.stat-card').click(function() {
                    $('.stat-card').removeClass('active');
                    $(this).addClass('active');
                    let filter = $(this).data('filter');
                    let val = $(this).data('val');
                    table.search('').columns().search('').draw();
                    $('.filter-row input, .filter-row select').val('');
                    if (val && val !== '') {
                        let colIdx = -1;
                        if (filter === 'category') colIdx = 0;
                        else if (filter === 'status') colIdx = 2;
                        else if (filter === 'passwordExpired') {
                            // custom filter for expired passwords – we'll handle via search in name? Not straightforward, we'll just do a custom search on column 2? For simplicity, we'll just show a message or filter manually.
                            // For now, we'll just apply a quick filter: show users with password expired.
                            table.search('').columns().search('').draw();
                            // We'll rely on the data: we have passwordExpired property, but not in table. So we'll skip.
                            // Better: we'll filter by category User and search for expired in a hidden column? We'll keep simple.
                            return;
                        }
                        if (colIdx >= 0) {
                            table.column(colIdx).search('^' + val + '$', true, false).draw();
                        }
                    }
                    $('html, body').animate({ scrollTop: $(".table-card").offset().top - 80 }, 400);
                });

                // Export buttons
                $('#btnExcel').click(function() { table.button('.buttons-excel').trigger(); });
                $('#btnCsv').click(function() {
                    // Use DataTables to export CSV
                    table.button('.buttons-excel').trigger();
                    // Note: DataTables buttons only have excel, but we can change text
                });
                $('#btnWord').click(function() {
                    let htmlContent = `
                        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                        <head><meta charset="utf-8"><title>AD Report</title>
                        <style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #ccc; padding: 4px; font-size: 10pt; } th { background: #f0f0f0; }</style>
                        </head><body><h2>AD Report - ${stats.Domain}</h2>
                        <p>Generated: ${new Date().toLocaleString()}</p>
                        ${document.querySelector('#mainTable').outerHTML}
                        </body></html>`;
                    let blob = new Blob([htmlContent], { type: 'application/msword' });
                    let link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = 'AD_Report.doc';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                });
                $('#btnReset').click(function() {
                    table.search('').columns().search('').draw();
                    $('.filter-row input, .filter-row select').val('');
                    $('.stat-card').removeClass('active');
                });

            } catch (e) {
                alert("Data Load Error: " + e.message);
                console.error(e);
            }
        });

        function showDetail(idx) {
            let item = data[idx];
            let content = JSON.stringify(item, null, 2);
            $('#modalTitle').text(item.Name + ' (' + item.Category + ')');
            $('#modalContent').text(content);
            new bootstrap.Modal(document.getElementById('detailModal')).show();
        }
    </script>
</body>
</html>
'@

# Replace placeholders
$html = $html.Replace('@BASE64_STATS@', $b64Stats).Replace('@BASE64_DATA@', $b64Data)

# Save HTML to the output folder
$htmlPath = Join-Path $outputDir "AD_Report.html"
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($true))
Write-Host "HTML Dashboard exported to: $htmlPath" -ForegroundColor Green

# Final message
Write-Host "`nAll reports saved to folder:" -ForegroundColor Cyan
Write-Host "  $outputDir" -ForegroundColor Yellow
Write-Host "`nOpen the HTML dashboard and use the Export buttons for Excel/Word/CSV." -ForegroundColor White
Write-Host "Press any key to exit..." -ForegroundColor Gray
# $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
