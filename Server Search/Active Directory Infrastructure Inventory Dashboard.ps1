<#
.SYNOPSIS
    Active Directory Infrastructure Inventory Dashboard
.DESCRIPTION
    Generates an HTML dashboard with statistics and details for:
    - Servers (stale/active detection)
    - Service Accounts (with SPN or naming patterns)
    - Managed Service Accounts (gMSA)
    - Application Objects (serviceConnectionPoint)
    - Security Groups (member counts)
    Exports findings to CSV, Excel (via HTML button), and Word (via HTML button).
    All output files are saved in a timestamped folder on the desktop.
    Runs without AD using embedded sample data for testing.
.AUTHOR
    Based on MYIGT AD Report by Stephen McKee, adapted by AI

    Will the script detect Windows Server 2025 and Linux servers?

The script filters computer objects using -Filter { OperatingSystem -like "*Server*" }. Microsoft’s AD population for Windows Server 2025 will almost certainly set the OperatingSystem attribute to something like "Windows Server 2025" or "Windows Server 2025 Datacenter", which contains "Server". So it will be included automatically.

Linux (or other non‑Windows) servers – Only if their AD OperatingSystem attribute contains the substring "Server".
In practice, many Linux distributions that are domain‑joined (via SSSD or similar) set OperatingSystem to "Linux", "Red Hat Enterprise Linux", "Ubuntu", etc. – these will be missed because they don’t include "Server".
How to include all server‑like computers (including Linux)

If you want to capture all computer objects and then categorise them by OS, you can either:

    Remove the OS filter and fetch all computer objects, then filter within your code.

    Use an expanded filter that matches common server OS patterns.

Option 1: Fetch all computers (simplest)

Replace the server query block (around line 63) with:
powershell

$servers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonTimestamp, Description, Enabled, whenCreated, ServicePrincipalNames, DistinguishedName
Then, in the loop, you can still compute $status and optionally add a flag like IsServer = $true if $s.OperatingSystem -match "Server|Linux|CentOS|Red Hat|Ubuntu|Debian" to differentiate.

Option 2: Expand the filter to include Linux patterns
powershell

$servers = Get-ADComputer -Filter { (OperatingSystem -like "*Server*") -or (OperatingSystem -like "*Linux*") -or (OperatingSystem -like "*Ubuntu*") -or (OperatingSystem -like "*CentOS*") -or (OperatingSystem -like "*Red Hat*") } -Properties ...

This will catch most domain‑joined Linux servers that have a recognizable OS name.
What about other OSes (e.g., FreeBSD, macOS Server)? You can similarly extend the filter with additional -or clauses. The script is flexible – just adjust the -Filter parameter to suit your environment.
If you want to be exhaustive, fetching all computers and then applying a regex in PowerShell is more maintainable and won’t miss any.
Quick modification for your script

Look for this line (around line 63 in the provided code):
powershell
$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } -Properties ...

Replace it with:
powershell
$servers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonTimestamp, Description, Enabled, whenCreated, ServicePrincipalNames, DistinguishedName
Then, inside the foreach ($s in $servers) loop, you can optionally add an IsServer property based on a broader match – but the existing $status logic will still work for all computers, marking them as Active/Stale/Disabled regardless of OS.
Summary
Windows Server 2025: ✅ detected.
Linux servers: ❌ not detected unless they include “Server” in their OS name – modify the filter as shown above to include them.
Feel free to adjust the filter to match your exact naming conventions. The rest of the dashboard (stats, charts, exports) will work seamlessly with the expanded data.
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
Write-Host "   AD INFRASTRUCTURE INVENTORY DASHBOARD | $domainName" -ForegroundColor White
Write-Host "   ------------------------------------------------------------------" -ForegroundColor DarkGray
if ($useSample) {
    Write-Host "   [TEST MODE] Using embedded sample data." -ForegroundColor Yellow
} else {
    Write-Host "   [+] Scanning Servers, Service Accounts, Applications, Groups..." -ForegroundColor Yellow
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
            if ($d.Year -gt ($Today.Year + 5)) { return $null }  # future bug
            return ($Today - $d).Days
        } catch { return $null }
    }
    return $null
}

# Initialize collections
$serverList = [System.Collections.Generic.List[Object]]::new()
$svcAccountList = [System.Collections.Generic.List[Object]]::new()
$gmsaList = [System.Collections.Generic.List[Object]]::new()
$appList = [System.Collections.Generic.List[Object]]::new()
$groupList = [System.Collections.Generic.List[Object]]::new()

if (-not $useSample) {
    # ---- REAL AD QUERIES ----
    Import-Module ActiveDirectory -ErrorAction Stop

    # 2a. Servers (computer objects with OperatingSystem containing "Server")
    Write-Host "   Querying Servers..." -NoNewline
    $servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } -Properties OperatingSystem, OperatingSystemVersion, LastLogonTimestamp, Description, Enabled, whenCreated, ServicePrincipalNames, DistinguishedName
    Write-Host " $($servers.Count) found." -ForegroundColor Green
    foreach ($s in $servers) {
        $days = Get-LogonDays $s.LastLogonTimestamp
        $status = if (-not $s.Enabled) { "Disabled" } else {
            if ($days -eq $null) { "Never Logged In" }
            elseif ($days -le $staleThresholdDays) { "Active" }
            else { "Stale" }
        }
        $serverList.Add([PSCustomObject]@{
            Category    = "Server"
            Name        = $s.Name
            OS          = $s.OperatingSystem
            OSVersion   = $s.OperatingSystemVersion
            Status      = $status
            Enabled     = $s.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            Created     = $s.whenCreated
            Description = $s.Description
            SPN         = ($s.ServicePrincipalNames -join ";")
            DN          = $s.DistinguishedName
        })
    }

    # 2b. Service Accounts: users with SPN or name patterns (svc*, service*, etc.) or in typical OUs
    Write-Host "   Querying Service Accounts..." -NoNewline
    $filter = { (ServicePrincipalName -like "*") -or (Name -like "svc*") -or (Name -like "service*") -or (Name -like "gmsa*") -or (Name -like "msa*") }
    $svcUsers = Get-ADUser -Filter $filter -Properties ServicePrincipalName, LastLogonTimestamp, PasswordLastSet, Enabled, Description, whenCreated, DistinguishedName
    Write-Host " $($svcUsers.Count) found." -ForegroundColor Green
    foreach ($u in $svcUsers) {
        $days = Get-LogonDays $u.LastLogonTimestamp
        $status = if (-not $u.Enabled) { "Disabled" } else {
            if ($days -eq $null) { "Never Logged In" }
            elseif ($days -le $staleThresholdDays) { "Active" }
            else { "Stale" }
        }
        $svcAccountList.Add([PSCustomObject]@{
            Category    = "Service Account"
            Name        = $u.Name
            SamAccount  = $u.SamAccountName
            SPN         = ($u.ServicePrincipalName -join ";")
            Status      = $status
            Enabled     = $u.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            PasswordSet = $u.PasswordLastSet
            Created     = $u.whenCreated
            Description = $u.Description
            DN          = $u.DistinguishedName
        })
    }

    # 2c. Managed Service Accounts (gMSA) - using Get-ADServiceAccount if available, else fallback
    Write-Host "   Querying Managed Service Accounts..." -NoNewline
    try {
        $gmsas = Get-ADServiceAccount -Filter * -Properties LastLogonTimestamp, Enabled, Description, whenCreated
        Write-Host " $($gmsas.Count) found." -ForegroundColor Green
        foreach ($g in $gmsas) {
            $days = Get-LogonDays $g.LastLogonTimestamp
            $status = if (-not $g.Enabled) { "Disabled" } else {
                if ($days -eq $null) { "Never Logged In" }
                elseif ($days -le $staleThresholdDays) { "Active" }
                else { "Stale" }
            }
            $gmsaList.Add([PSCustomObject]@{
                Category    = "gMSA"
                Name        = $g.Name
                SamAccount  = $g.SamAccountName
                Status      = $status
                Enabled     = $g.Enabled
                LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
                Created     = $g.whenCreated
                Description = $g.Description
                DN          = $g.DistinguishedName
            })
        }
    } catch {
        Write-Warning "Get-ADServiceAccount not available; skipping gMSA."
    }

    # 2d. Applications (serviceConnectionPoint objects)
    Write-Host "   Querying Application Objects (serviceConnectionPoint)..." -NoNewline
    $apps = Get-ADObject -Filter { ObjectClass -eq "serviceConnectionPoint" } -Properties keywords, serviceBindingInformation, DisplayName, Description, whenCreated, DistinguishedName
    Write-Host " $($apps.Count) found." -ForegroundColor Green
    foreach ($a in $apps) {
        $appList.Add([PSCustomObject]@{
            Category    = "Application"
            Name        = if ($a.DisplayName) { $a.DisplayName } else { $a.Name }
            Keywords    = ($a.Keywords -join ";")
            BindingInfo = ($a.serviceBindingInformation -join ";")
            Description = $a.Description
            Created     = $a.whenCreated
            DN          = $a.DistinguishedName
        })
    }

    # 2e. Security Groups (filter GroupCategory = Security)
    Write-Host "   Querying Security Groups..." -NoNewline
    $groups = Get-ADGroup -Filter { GroupCategory -eq "Security" } -Properties Description, Members, whenCreated, DistinguishedName
    Write-Host " $($groups.Count) found." -ForegroundColor Green
    foreach ($g in $groups) {
        $memberCount = ($g.Members).Count
        $groupList.Add([PSCustomObject]@{
            Category    = "Security Group"
            Name        = $g.Name
            MemberCount = $memberCount
            Description = $g.Description
            Created     = $g.whenCreated
            DN          = $g.DistinguishedName
        })
    }

} else {
    # ---- SAMPLE DATA FOR TESTING ----
    Write-Host "   Generating sample data..." -ForegroundColor Yellow
    # Sample Servers
    $serverList.Add([PSCustomObject]@{Category="Server"; Name="DC01"; OS="Windows Server 2019"; OSVersion="10.0.17763"; Status="Active"; Enabled=$true; LastLogon="2 days"; Created="2024-01-15"; Description="Primary DC"; SPN=""; DN="CN=DC01,DC=sample,DC=local"})
    $serverList.Add([PSCustomObject]@{Category="Server"; Name="FS01"; OS="Windows Server 2016"; OSVersion="10.0.14393"; Status="Stale"; Enabled=$true; LastLogon="200 days"; Created="2023-06-10"; Description="File Server"; SPN=""; DN="CN=FS01,DC=sample,DC=local"})
    $serverList.Add([PSCustomObject]@{Category="Server"; Name="APP01"; OS="Windows Server 2022"; OSVersion="10.0.20348"; Status="Disabled"; Enabled=$false; LastLogon="Never"; Created="2024-08-20"; Description="Legacy App"; SPN=""; DN="CN=APP01,DC=sample,DC=local"})
    $serverList.Add([PSCustomObject]@{Category="Server"; Name="SQL01"; OS="Windows Server 2019"; OSVersion="10.0.17763"; Status="Active"; Enabled=$true; LastLogon="1 days"; Created="2024-03-05"; Description="SQL Server"; SPN="MSSQLSvc/sql01.sample.local"; DN="CN=SQL01,DC=sample,DC=local"})

    # Service Accounts
    $svcAccountList.Add([PSCustomObject]@{Category="Service Account"; Name="svc_backup"; SamAccount="svc_backup"; SPN=""; Status="Active"; Enabled=$true; LastLogon="5 days"; PasswordSet="2024-09-01"; Created="2023-12-01"; Description="Backup service account"; DN="CN=svc_backup,OU=ServiceAccounts,DC=sample,DC=local"})
    $svcAccountList.Add([PSCustomObject]@{Category="Service Account"; Name="svc_sql"; SamAccount="svc_sql"; SPN="MSSQLSvc/sql01.sample.local"; Status="Stale"; Enabled=$true; LastLogon="200 days"; PasswordSet="2024-08-15"; Created="2023-06-01"; Description="SQL service account"; DN="CN=svc_sql,OU=ServiceAccounts,DC=sample,DC=local"})
    $svcAccountList.Add([PSCustomObject]@{Category="Service Account"; Name="svc_iis"; SamAccount="svc_iis"; SPN="HTTP/iis.sample.local"; Status="Never Logged In"; Enabled=$true; LastLogon="Never"; PasswordSet="2024-10-01"; Created="2024-10-01"; Description="IIS application pool"; DN="CN=svc_iis,OU=ServiceAccounts,DC=sample,DC=local"})

    # gMSA
    $gmsaList.Add([PSCustomObject]@{Category="gMSA"; Name="gmsa-sql"; SamAccount="gmsa-sql$"; Status="Active"; Enabled=$true; LastLogon="2 days"; Created="2024-07-01"; Description="gMSA for SQL"; DN="CN=gmsa-sql,CN=Managed Service Accounts,DC=sample,DC=local"})
    $gmsaList.Add([PSCustomObject]@{Category="gMSA"; Name="gmsa-web"; SamAccount="gmsa-web$"; Status="Stale"; Enabled=$true; LastLogon="190 days"; Created="2023-09-01"; Description="gMSA for IIS"; DN="CN=gmsa-web,CN=Managed Service Accounts,DC=sample,DC=local"})

    # Applications
    $appList.Add([PSCustomObject]@{Category="Application"; Name="SharePoint"; Keywords="SharePoint;2019"; BindingInfo="https://sp.sample.local"; Description="SharePoint farm"; Created="2024-02-10"; DN="CN=SP-APP,CN=Services,DC=sample,DC=local"})
    $appList.Add([PSCustomObject]@{Category="Application"; Name="Exchange"; Keywords="Exchange;2016"; BindingInfo="https://mail.sample.local"; Description="Exchange server"; Created="2023-11-05"; DN="CN=EX-APP,CN=Services,DC=sample,DC=local"})

    # Security Groups
    $groupList.Add([PSCustomObject]@{Category="Security Group"; Name="Domain Admins"; MemberCount=5; Description="Domain administrators"; Created="2021-01-01"; DN="CN=Domain Admins,CN=Users,DC=sample,DC=local"})
    $groupList.Add([PSCustomObject]@{Category="Security Group"; Name="SQL Admins"; MemberCount=3; Description="SQL Server admins"; Created="2023-05-15"; DN="CN=SQL Admins,OU=Groups,DC=sample,DC=local"})
    $groupList.Add([PSCustomObject]@{Category="Security Group"; Name="Backup Operators"; MemberCount=2; Description="Backup operators"; Created="2022-08-20"; DN="CN=Backup Operators,CN=Builtin,DC=sample,DC=local"})
}

# Combine all into one list for table
$allItems = [System.Collections.Generic.List[Object]]::new()
$allItems.AddRange($serverList)
$allItems.AddRange($svcAccountList)
$allItems.AddRange($gmsaList)
$allItems.AddRange($appList)
$allItems.AddRange($groupList)

# Compute stats
$totalServers = $serverList.Count
$totalSvcAccounts = $svcAccountList.Count
$totalGMSA = $gmsaList.Count
$totalApps = $appList.Count
$totalGroups = $groupList.Count

$staleServers = ($serverList | Where-Object { $_.Status -eq "Stale" }).Count
$activeServers = ($serverList | Where-Object { $_.Status -eq "Active" }).Count
$neverServers = ($serverList | Where-Object { $_.Status -eq "Never Logged In" }).Count
$disabledServers = ($serverList | Where-Object { -not $_.Enabled }).Count

$staleSvc = ($svcAccountList | Where-Object { $_.Status -eq "Stale" }).Count
$activeSvc = ($svcAccountList | Where-Object { $_.Status -eq "Active" }).Count
$neverSvc = ($svcAccountList | Where-Object { $_.Status -eq "Never Logged In" }).Count

# OS distribution for servers
$osDist = $serverList | Group-Object OS | Select-Object Name, Count | Sort-Object Count -Descending
$osLabels = $osDist.Name
$osCounts = $osDist.Count

# Pie data for server status
$serverStatusData = @{
    Active = $activeServers
    Stale  = $staleServers
    Never  = $neverServers
    Disabled = $disabledServers
}

# Build JSON for embedding
$jsonAll = $allItems | ConvertTo-Json -Depth 3 -Compress
$jsonStats = @{
    TotalServers = $totalServers
    ActiveServers = $activeServers
    StaleServers = $staleServers
    NeverServers = $neverServers
    DisabledServers = $disabledServers
    TotalSvcAccounts = $totalSvcAccounts
    ActiveSvc = $activeSvc
    StaleSvc = $staleSvc
    NeverSvc = $neverSvc
    TotalGMSA = $totalGMSA
    TotalApps = $totalApps
    TotalGroups = $totalGroups
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
$folderName = "Server_Search_$folderTime"
$outputDir = Join-Path $desktop $folderName
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Write-Host "Creating output folder: $outputDir" -ForegroundColor Cyan

# --- 4. GENERATE CSV FILES ---
$dateStr = Get-Date -Format "yyyyMMdd"
$baseName = "AD_Infra_Report_$dateStr"
$csvPath = Join-Path $outputDir "$baseName.csv"
$allItems | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV exported to: $csvPath" -ForegroundColor Green

# Also export separate CSVs per category (optional)
$cats = $allItems | Select-Object -ExpandProperty Category -Unique
foreach ($cat in $cats) {
    $catFile = Join-Path $outputDir "$baseName-$cat.csv"
    $allItems | Where-Object { $_.Category -eq $cat } | Export-Csv -Path $catFile -NoTypeInformation -Encoding UTF8
}

# --- 5. GENERATE HTML DASHBOARD ---
Write-Host "   Generating HTML Dashboard..." -ForegroundColor Green

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AD Infrastructure Inventory Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <style>
        :root { --bg: #f8fafc; --text: #334155; --primary: #4f46e5; }
        body { background-color: var(--bg); font-family: 'Inter', sans-serif; color: var(--text); font-size: 0.8rem; }
        
        .navbar { background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); padding: 0.4rem 1.5rem; }
        .brand { font-weight: 700; color: var(--primary); font-size: 1rem; }

        .stat-card {
            background: #fff; border-radius: 8px; border: 1px solid #e2e8f0; padding: 10px 12px;
            cursor: pointer; transition: 0.2s; position: relative; overflow: hidden; height: 100%;
            display: flex; flex-direction: column; justify-content: space-between;
        }
        .stat-card:hover { border-color: var(--primary); transform: translateY(-2px); box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .stat-card.active { background: #eef2ff; border-color: var(--primary); }
        
        .stat-label { 
            font-size: 0.6rem; font-weight: 700; text-transform: uppercase; color: #64748b; margin-bottom: 2px; letter-spacing: 0.5px;
        }
        .stat-val { 
            font-size: 1.4rem; font-weight: 700; color: #1e293b; line-height: 1;
        }
        .stat-icon { 
            position: absolute; right: 10px; top: 50%; transform: translateY(-50%); 
            font-size: 1.8rem; opacity: 0.12; color: #334155;
        }
        
        .sc-blue { border-left: 3px solid #3b82f6; } .sc-green { border-left: 3px solid #22c55e; } 
        .sc-gray { border-left: 3px solid #94a3b8; } .sc-red { border-left: 3px solid #ef4444; }
        .sc-purple { border-left: 3px solid #a855f7; } .sc-dark { border-left: 3px solid #334155; }
        .sc-orange { border-left: 3px solid #f97316; }

        .chart-card { background: #fff; border-radius: 8px; border: 1px solid #e2e8f0; padding: 10px 15px; height: 100%; }
        .chart-header { font-size: 0.8rem; font-weight: 600; margin-bottom: 5px; color: #475569; }
        .chart-wrapper { position: relative; height: 180px; } 
        .chart-scroll { overflow-y: auto; height: 100%; padding-right: 5px; }

        .table-card { background: #fff; border-radius: 8px; border: 1px solid #e2e8f0; padding: 12px; margin-top: 0.8rem; }
        
        table.dataTable { border-collapse: collapse !important; width: 100% !important; margin-top: 0 !important; }
        table.dataTable thead th { 
            background: #f1f5f9; font-weight: 600; font-size: 0.7rem; 
            padding: 6px 8px !important; border-bottom: 1px solid #e2e8f0; white-space: nowrap; 
        }
        table.dataTable tbody td { 
            padding: 4px 8px !important; vertical-align: middle; font-size: 0.75rem; border-bottom: 1px solid #f1f5f9;
        }
        
        .txt-trunc { max-width: 130px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display: block; }
        .td-nowrap { white-space: nowrap; }

        .badge-s { padding: 2px 6px; border-radius: 4px; font-size: 0.65rem; font-weight: 600; white-space: nowrap; }
        .bg-ok { background: #dcfce7; color: #166534; } .bg-warn { background: #ffedd5; color: #9a3412; } 
        .bg-err { background: #fee2e2; color: #991b1b; } .bg-ghost { background: #f1f5f9; color: #64748b; border: 1px solid #cbd5e1; }
        .bg-info-cat { background: #e0f2fe; color: #0369a1; }

        .row-stale { opacity: 0.6; background-color: #fafafa !important; } .row-stale:hover { opacity: 1; }
        .row-disabled { opacity: 0.5; background-color: #fef2f2 !important; }

        .form-control-sm-custom { height: 24px; padding: 1px 5px; font-size: 0.7rem; border: 1px solid #cbd5e1; border-radius: 4px; width: 100%; }
        .dataTables_length select { font-size: 0.75rem; padding: 1px 5px; } .dataTables_length { font-size: 0.75rem; color: #64748b; margin-bottom: 5px; }
        div.dataTables_info { font-size: 0.75rem; color: #64748b; }

        .pie-center-text { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); text-align: center; pointer-events: none; }
        .pie-val { font-size: 1.8rem; font-weight: 800; color: #334155; line-height: 1; }
        
        .modal-content { border: 1px solid #e2e8f0; }
        .group-badge { display: inline-block; background: var(--primary); color: white; padding: 2px 6px; border-radius: 4px; margin: 2px; font-size: 0.7rem; }
        .btn-export { font-size: 0.7rem; padding: 2px 8px; }
    </style>
</head>
<body>

    <nav class="navbar fixed-top">
        <div class="d-flex align-items-center gap-2">
            <i class="fa-solid fa-server text-primary fs-5"></i>
            <span class="brand" id="pageTitle">Infrastructure Dashboard</span>
        </div>
        <div class="small text-muted" id="headerInfo" style="font-size: 0.75rem;">Loading...</div>
    </nav>
    <div style="height: 50px;"></div>

    <div class="container-fluid px-4 py-2">
        
        <!-- Stats Cards -->
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-blue clickable" data-filter="category" data-val="Server"><div class="stat-label">Total Servers</div><div class="stat-val" id="vTotalServers">0</div><i class="fa-solid fa-server stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="status" data-val="Active"><div class="stat-label text-success">Active Servers</div><div class="stat-val text-success" id="vActiveServers">0</div><i class="fa-solid fa-check-circle stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label text-muted">Stale Servers</div><div class="stat-val text-muted" id="vStaleServers">0</div><i class="fa-solid fa-bed stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-red clickable" data-filter="status" data-val="Disabled"><div class="stat-label text-danger">Disabled Servers</div><div class="stat-val text-danger" id="vDisabledServers">0</div><i class="fa-solid fa-power-off stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-purple clickable" data-filter="category" data-val="Service Account"><div class="stat-label">Service Accounts</div><div class="stat-val" id="vSvcAccounts">0</div><i class="fa-solid fa-user-cog stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-orange clickable" data-filter="category" data-val="gMSA"><div class="stat-label">gMSA</div><div class="stat-val" id="vGMSA">0</div><i class="fa-solid fa-user-shield stat-icon"></i></div></div>
        </div>
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-dark clickable" data-filter="category" data-val="Application"><div class="stat-label">Applications</div><div class="stat-val" id="vApps">0</div><i class="fa-solid fa-apple-alt stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-blue clickable" data-filter="category" data-val="Security Group"><div class="stat-label">Security Groups</div><div class="stat-val" id="vGroups">0</div><i class="fa-solid fa-users-cog stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="status" data-val="Active"><div class="stat-label">Active Svc Accts</div><div class="stat-val text-success" id="vActiveSvc">0</div><i class="fa-solid fa-user-check stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label">Stale Svc Accts</div><div class="stat-val text-muted" id="vStaleSvc">0</div><i class="fa-solid fa-user-slash stat-icon"></i></div></div>
        </div>

        <!-- Charts -->
        <div class="row g-2">
            <div class="col-lg-8">
                <div class="chart-card">
                    <div class="chart-header">
                        <span><i class="fa-solid fa-chart-bar me-1"></i> Server OS Distribution</span>
                        <small class="text-muted fw-normal">(Scrollable)</small>
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
            <div class="col-lg-4">
                <div class="chart-card">
                    <div class="chart-header"><span><i class="fa-solid fa-chart-pie me-1"></i> Server Status</span></div>
                    <div class="chart-wrapper d-flex align-items-center justify-content-center position-relative">
                        <div style="width: 100%; height: 180px;">
                            <canvas id="statusChart"></canvas>
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
                    <tr><th>Category</th><th>Name</th><th>Status</th><th>Enabled</th><th>Last Logon</th><th>OS / SPN</th><th>Description</th><th>Created</th></tr>
                    <tr class="filter-row">
                        <th><select class="form-control-sm-custom"><option value="">All</option><option value="Server">Server</option><option value="Service Account">Service Account</option><option value="gMSA">gMSA</option><option value="Application">Application</option><option value="Security Group">Security Group</option></select></th>
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
        <div class="text-center text-muted mt-2" style="font-size: 0.65rem;">Generated by AD Infrastructure Inventory Tool • Domain: <span id="footerDomain">-</span></div>
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

        $(document).ready(function() {
            try {
                stats = JSON.parse(dec('@BASE64_STATS@'));
                data = JSON.parse(dec('@BASE64_DATA@'));

                $('#pageTitle').text(stats.Domain + ' Infrastructure');
                $('#headerInfo').text(`Generated: ${new Date().toLocaleString()}`);
                $('#footerDomain').text(stats.Domain);

                // Populate stats
                $('#vTotalServers').text(stats.TotalServers);
                $('#vActiveServers').text(stats.ActiveServers);
                $('#vStaleServers').text(stats.StaleServers);
                $('#vDisabledServers').text(stats.DisabledServers);
                $('#vSvcAccounts').text(stats.TotalSvcAccounts);
                $('#vGMSA').text(stats.TotalGMSA);
                $('#vApps').text(stats.TotalApps);
                $('#vGroups').text(stats.TotalGroups);
                $('#vActiveSvc').text(stats.ActiveSvc);
                $('#vStaleSvc').text(stats.StaleSvc);
                $('#centerTotal').text(stats.TotalServers);

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
                        datasets: [{ label: 'Servers', data: osCounts, backgroundColor: '#4f46e5', borderRadius: 3, barThickness: 8 }]
                    },
                    options: {
                        indexAxis: 'y',
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: { x: { display: false }, y: { grid: { display: false }, ticks: { font: { size: 9 } } } }
                    }
                });

                // Chart: Server Status (Pie)
                const statusCtx = document.getElementById('statusChart').getContext('2d');
                new Chart(statusCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['Active', 'Stale', 'Never', 'Disabled'],
                        datasets: [{
                            data: [stats.ActiveServers, stats.StaleServers, stats.NeverServers, stats.DisabledServers],
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

                // Build table rows
                const rows = data.map((item, idx) => {
                    let statusCls = 'bg-ghost';
                    if (item.Status === 'Active') statusCls = 'bg-ok';
                    else if (item.Status === 'Stale') statusCls = 'bg-warn';
                    else if (item.Status === 'Disabled') statusCls = 'bg-err';
                    else if (item.Status === 'Never Logged In') statusCls = 'bg-ghost';
                    
                    let enabledCls = (item.Enabled === true) ? 'bg-ok' : (item.Enabled === false ? 'bg-err' : 'bg-ghost');
                    let enabledText = (item.Enabled === true) ? 'Enabled' : (item.Enabled === false ? 'Disabled' : 'N/A');
                    
                    // Determine row styling for stale/disabled
                    let rowClass = '';
                    if (item.Status === 'Stale' || item.Status === 'Never Logged In') rowClass += ' row-stale';
                    if (item.Enabled === false) rowClass += ' row-disabled';

                    let osOrSpn = item.OS || item.SPN || item.SamAccount || '';
                    if (item.OSVersion) osOrSpn += ' (' + item.OSVersion + ')';
                    if (item.BindingInfo) osOrSpn += ' | ' + item.BindingInfo;
                    if (item.MemberCount !== undefined) osOrSpn = 'Members: ' + item.MemberCount;

                    return `<tr class="${rowClass}" data-category="${item.Category}" data-status="${item.Status}">
                        <td><span class="badge-s bg-info-cat">${item.Category}</span></td>
                        <td class="fw-bold clickable" style="color:#4f46e5" onclick="showDetail(${idx})">${item.Name}</td>
                        <td><span class="badge-s ${statusCls}">${item.Status || 'N/A'}</span></td>
                        <td><span class="badge-s ${enabledCls}">${enabledText}</span></td>
                        <td class="td-nowrap">${item.LastLogon || '-'}</td>
                        <td><div class="txt-trunc" title="${osOrSpn}">${osOrSpn}</div></td>
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
                        { orderData: [2, 3] } // sort by status then enabled
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
                    // Clear all filters
                    table.search('').columns().search('').draw();
                    $('.filter-row input, .filter-row select').val('');
                    if (val && val !== '') {
                        let colIdx = (filter === 'category') ? 0 : (filter === 'status' ? 2 : -1);
                        if (colIdx >= 0) {
                            table.column(colIdx).search('^' + val + '$', true, false).draw();
                        }
                    }
                    $('html, body').animate({ scrollTop: $(".table-card").offset().top - 80 }, 400);
                });

                // Export buttons
                $('#btnExcel').click(function() {
                    // Use DataTables button (Excel)
                    table.button('.buttons-excel').trigger();
                });

                $('#btnCsv').click(function() {
                    // Export to CSV using DataTables
                    let csv = table.buttons(0).text('CSV').trigger();
                    // Restore text
                    setTimeout(() => { table.buttons(0).text('CSV'); }, 100);
                });

                $('#btnWord').click(function() {
                    // Export table to Word (HTML to .doc)
                    let htmlContent = `
                        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                        <head><meta charset="utf-8"><title>AD Infrastructure Report</title>
                        <style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #ccc; padding: 4px; font-size: 10pt; } th { background: #f0f0f0; }</style>
                        </head><body><h2>AD Infrastructure Report - ${stats.Domain}</h2>
                        <p>Generated: ${new Date().toLocaleString()}</p>
                        ${document.querySelector('#mainTable').outerHTML}
                        </body></html>`;
                    let blob = new Blob([htmlContent], { type: 'application/msword' });
                    let link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = 'AD_Infra_Report.doc';
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
$htmlPath = Join-Path $outputDir "AD_Infra_Report.html"
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($true))
Write-Host "HTML Dashboard exported to: $htmlPath" -ForegroundColor Green

# Final message
Write-Host "`nAll reports saved to folder:" -ForegroundColor Cyan
Write-Host "  $outputDir" -ForegroundColor Yellow
Write-Host "`nYou can open the HTML dashboard and use the Export buttons for Excel/Word/CSV." -ForegroundColor White
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
