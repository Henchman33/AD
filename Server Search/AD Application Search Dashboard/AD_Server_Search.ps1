<#
.SYNOPSIS
    Active Directory Application Inventory Dashboard
.DESCRIPTION
    Discovers and reports on application-related objects in AD:
    - Service Connection Points (serviceConnectionPoint)
    - Application objects (class 'application')
    - Computers with Service Principal Names (SPNs)
    - Service Accounts (with SPNs or naming patterns)
    - Application-specific OUs (name contains 'app' or 'application')
    Generates an HTML dashboard with stats, charts, and exportable data.
    Saves all outputs to a timestamped folder on the desktop.
    Runs without AD using sample data for testing.
.AUTHOR
    AI-assisted, based on MYIGT AD Report framework
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
    $domainName = "MYIGT.com"
}

# --- CONSOLE UI ---
Clear-Host
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   AD APPLICATION INVENTORY DASHBOARD | $domainName" -ForegroundColor White
Write-Host "   ------------------------------------------------------------------" -ForegroundColor DarkGray
if ($useSample) {
    Write-Host "   [TEST MODE] Using embedded sample data." -ForegroundColor Yellow
} else {
    Write-Host "   [+] Scanning Application Objects, SPNs, Service Accounts..." -ForegroundColor Yellow
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
$appObjects = [System.Collections.Generic.List[Object]]::new()
$svcAccounts = [System.Collections.Generic.List[Object]]::new()
$appServers = [System.Collections.Generic.List[Object]]::new()
$appOUs = [System.Collections.Generic.List[Object]]::new()
$allSPNs = @()  # for summary

if (-not $useSample) {
    # ---- REAL AD QUERIES ----
    Import-Module ActiveDirectory -ErrorAction Stop

    # 1. Service Connection Points
    Write-Host "   Querying Service Connection Points..." -NoNewline
    $scps = Get-ADObject -Filter { ObjectClass -eq "serviceConnectionPoint" } -Properties DisplayName, Keywords, serviceBindingInformation, Description, whenCreated, DistinguishedName
    Write-Host " $($scps.Count) found." -ForegroundColor Green
    foreach ($s in $scps) {
        $appObjects.Add([PSCustomObject]@{
            Category    = "Service Connection Point"
            Name        = if ($s.DisplayName) { $s.DisplayName } else { $s.Name }
            Keywords    = ($s.Keywords -join ";")
            BindingInfo = ($s.serviceBindingInformation -join ";")
            Description = $s.Description
            Created     = $s.whenCreated
            DN          = $s.DistinguishedName
            Source      = "SCP"
        })
    }

    # 2. Application objects (class 'application')
    Write-Host "   Querying Application objects..." -NoNewline
    $apps = Get-ADObject -Filter { ObjectClass -eq "application" } -Properties DisplayName, Description, whenCreated, DistinguishedName
    Write-Host " $($apps.Count) found." -ForegroundColor Green
    foreach ($a in $apps) {
        $appObjects.Add([PSCustomObject]@{
            Category    = "Application Object"
            Name        = if ($a.DisplayName) { $a.DisplayName } else { $a.Name }
            Keywords    = ""
            BindingInfo = ""
            Description = $a.Description
            Created     = $a.whenCreated
            DN          = $a.DistinguishedName
            Source      = "AppClass"
        })
    }

    # 3. Computers with SPNs (Application Servers)
    Write-Host "   Querying Computers with SPNs..." -NoNewline
    $computers = Get-ADComputer -Filter { ServicePrincipalName -like "*" } -Properties OperatingSystem, ServicePrincipalName, Description, Enabled, LastLogonTimestamp, whenCreated, DistinguishedName
    Write-Host " $($computers.Count) found." -ForegroundColor Green
    foreach ($c in $computers) {
        $days = Get-LogonDays $c.LastLogonTimestamp
        $status = if (-not $c.Enabled) { "Disabled" } else {
            if ($days -eq $null) { "Never Logged In" }
            elseif ($days -le $staleThresholdDays) { "Active" }
            else { "Stale" }
        }
        $spnList = $c.ServicePrincipalName -join ";"
        $appServers.Add([PSCustomObject]@{
            Category    = "Application Server"
            Name        = $c.Name
            OS          = $c.OperatingSystem
            SPNs        = $spnList
            Status      = $status
            Enabled     = $c.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            Description = $c.Description
            Created     = $c.whenCreated
            DN          = $c.DistinguishedName
        })
        # Collect SPNs for summary
        $allSPNs += $c.ServicePrincipalName
    }

    # 4. Service Accounts (users with SPNs or naming patterns)
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
        $spnList = $u.ServicePrincipalName -join ";"
        $svcAccounts.Add([PSCustomObject]@{
            Category    = "Service Account"
            Name        = $u.Name
            SamAccount  = $u.SamAccountName
            SPNs        = $spnList
            Status      = $status
            Enabled     = $u.Enabled
            LastLogon   = if ($days -ne $null) { "$days days" } else { "Never" }
            PasswordSet = $u.PasswordLastSet
            Description = $u.Description
            Created     = $u.whenCreated
            DN          = $u.DistinguishedName
        })
        $allSPNs += $u.ServicePrincipalName
    }

    # 5. Application-like OUs (name contains 'app' or 'application')
    Write-Host "   Querying Application OUs..." -NoNewline
    $ouFilter = { (Name -like "*app*") -or (Name -like "*application*") }
    $ous = Get-ADOrganizationalUnit -Filter $ouFilter -Properties Description, whenCreated, DistinguishedName
    Write-Host " $($ous.Count) found." -ForegroundColor Green
    foreach ($ou in $ous) {
        $appOUs.Add([PSCustomObject]@{
            Category    = "Application OU"
            Name        = $ou.Name
            Path        = $ou.DistinguishedName
            Description = $ou.Description
            Created     = $ou.whenCreated
            DN          = $ou.DistinguishedName
        })
    }

} else {
    # ---- SAMPLE DATA FOR TESTING ----
    Write-Host "   Generating sample application data..." -ForegroundColor Yellow
    # Service Connection Points
    $appObjects.Add([PSCustomObject]@{Category="Service Connection Point"; Name="SharePoint Farm"; Keywords="SharePoint;2019"; BindingInfo="https://sp.sample.local"; Description="SharePoint farm SCP"; Created="2024-01-15"; DN="CN=SP-SCP,CN=Services,DC=sample,DC=local"; Source="SCP"})
    $appObjects.Add([PSCustomObject]@{Category="Service Connection Point"; Name="Exchange AutoDiscover"; Keywords="Exchange;2016"; BindingInfo="https://autodiscover.sample.local"; Description="Exchange Autodiscover"; Created="2023-06-10"; DN="CN=EX-SCP,CN=Services,DC=sample,DC=local"; Source="SCP"})
    $appObjects.Add([PSCustomObject]@{Category="Application Object"; Name="Custom App"; Keywords=""; BindingInfo=""; Description="Custom line-of-business app"; Created="2024-08-20"; DN="CN=CustomApp,CN=Applications,DC=sample,DC=local"; Source="AppClass"})

    # Application Servers
    $appServers.Add([PSCustomObject]@{Category="Application Server"; Name="SQL01"; OS="Windows Server 2019"; SPNs="MSSQLSvc/sql01.sample.local:1433;MSSQLSvc/sql01.sample.local"; Status="Active"; Enabled=$true; LastLogon="2 days"; Description="SQL Server"; Created="2024-03-05"; DN="CN=SQL01,DC=sample,DC=local"})
    $appServers.Add([PSCustomObject]@{Category="Application Server"; Name="WEB01"; OS="Windows Server 2022"; SPNs="HTTP/web.sample.local;HTTP/web"; Status="Stale"; Enabled=$true; LastLogon="200 days"; Description="Web Server"; Created="2023-09-01"; DN="CN=WEB01,DC=sample,DC=local"})
    $appServers.Add([PSCustomObject]@{Category="Application Server"; Name="EXCH01"; OS="Windows Server 2016"; SPNs="ExchangeAB/exch01.sample.local;SMTP/exch01.sample.local"; Status="Active"; Enabled=$true; LastLogon="5 days"; Description="Exchange Server"; Created="2022-11-15"; DN="CN=EXCH01,DC=sample,DC=local"})

    # Service Accounts
    $svcAccounts.Add([PSCustomObject]@{Category="Service Account"; Name="svc_sql"; SamAccount="svc_sql"; SPNs="MSSQLSvc/sql01.sample.local"; Status="Active"; Enabled=$true; LastLogon="3 days"; PasswordSet="2024-09-01"; Description="SQL service account"; Created="2023-12-01"; DN="CN=svc_sql,OU=ServiceAccounts,DC=sample,DC=local"})
    $svcAccounts.Add([PSCustomObject]@{Category="Service Account"; Name="svc_iis"; SamAccount="svc_iis"; SPNs="HTTP/web.sample.local"; Status="Stale"; Enabled=$true; LastLogon="250 days"; PasswordSet="2024-08-15"; Description="IIS app pool account"; Created="2023-06-01"; DN="CN=svc_iis,OU=ServiceAccounts,DC=sample,DC=local"})
    $svcAccounts.Add([PSCustomObject]@{Category="Service Account"; Name="svc_exch"; SamAccount="svc_exch"; SPNs="ExchangeAB/exch01.sample.local"; Status="Active"; Enabled=$true; LastLogon="1 days"; PasswordSet="2024-10-01"; Description="Exchange service account"; Created="2022-11-01"; DN="CN=svc_exch,OU=ServiceAccounts,DC=sample,DC=local"})

    # Application OUs
    $appOUs.Add([PSCustomObject]@{Category="Application OU"; Name="Applications"; Path="OU=Applications,DC=sample,DC=local"; Description="Container for application objects"; Created="2021-01-01"; DN="OU=Applications,DC=sample,DC=local"})
    $appOUs.Add([PSCustomObject]@{Category="Application OU"; Name="AppServers"; Path="OU=AppServers,DC=sample,DC=local"; Description="Servers hosting apps"; Created="2022-05-15"; DN="OU=AppServers,DC=sample,DC=local"})

    # Simulate SPN list
    $allSPNs = @("MSSQLSvc/sql01.sample.local", "HTTP/web.sample.local", "ExchangeAB/exch01.sample.local")
}

# Combine all into one list for the main table
$allItems = [System.Collections.Generic.List[Object]]::new()
$allItems.AddRange($appObjects)
$allItems.AddRange($appServers)
$allItems.AddRange($svcAccounts)
$allItems.AddRange($appOUs)

# Compute stats
$totalSCP = ($appObjects | Where-Object { $_.Source -eq "SCP" }).Count
$totalAppClass = ($appObjects | Where-Object { $_.Source -eq "AppClass" }).Count
$totalAppObjects = $appObjects.Count
$totalAppServers = $appServers.Count
$totalSvcAccounts = $svcAccounts.Count
$totalAppOUs = $appOUs.Count

# Stale counts for servers and service accounts
$staleServers = ($appServers | Where-Object { $_.Status -eq "Stale" }).Count
$activeServers = ($appServers | Where-Object { $_.Status -eq "Active" }).Count
$neverServers = ($appServers | Where-Object { $_.Status -eq "Never Logged In" }).Count
$disabledServers = ($appServers | Where-Object { -not $_.Enabled }).Count

$staleSvc = ($svcAccounts | Where-Object { $_.Status -eq "Stale" }).Count
$activeSvc = ($svcAccounts | Where-Object { $_.Status -eq "Active" }).Count

# Extract SPN service classes for chart
$spnTypes = @()
if ($allSPNs) {
    $spnTypes = $allSPNs | ForEach-Object { if ($_ -match "^([^/]+)") { $matches[1] } } | Group-Object | Select-Object Name, Count | Sort-Object Count -Descending
}
$spnLabels = $spnTypes.Name
$spnCounts = $spnTypes.Count

# Build JSON for embedding
$jsonAll = $allItems | ConvertTo-Json -Depth 3 -Compress
$jsonStats = @{
    TotalAppObjects = $totalAppObjects
    TotalSCP = $totalSCP
    TotalAppClass = $totalAppClass
    TotalAppServers = $totalAppServers
    ActiveServers = $activeServers
    StaleServers = $staleServers
    NeverServers = $neverServers
    DisabledServers = $disabledServers
    TotalSvcAccounts = $totalSvcAccounts
    ActiveSvc = $activeSvc
    StaleSvc = $staleSvc
    TotalAppOUs = $totalAppOUs
    SPNLabels = $spnLabels
    SPNCounts = $spnCounts
    Domain = $domainName
} | ConvertTo-Json -Depth 5 -Compress

$utf8 = [System.Text.Encoding]::UTF8
$b64Data = [Convert]::ToBase64String($utf8.GetBytes($jsonAll))
$b64Stats = [Convert]::ToBase64String($utf8.GetBytes($jsonStats))

# --- 3. CREATE OUTPUT FOLDER ON DESKTOP ---
$desktop = [Environment]::GetFolderPath('Desktop')
$folderTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$folderName = "Application_Search_$folderTime"
$outputDir = Join-Path $desktop $folderName
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Write-Host "Creating output folder: $outputDir" -ForegroundColor Cyan

# --- 4. GENERATE CSV FILES ---
$dateStr = Get-Date -Format "yyyyMMdd"
$baseName = "AD_App_Inventory_$dateStr"
$csvPath = Join-Path $outputDir "$baseName.csv"
$allItems | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV exported to: $csvPath" -ForegroundColor Green

# Per-category CSVs
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
    <title>AD Application Inventory Dashboard</title>
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
        .sc-teal { border-left: 3px solid #14b8a6; }

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
            <i class="fa-solid fa-apple-alt text-primary fs-5"></i>
            <span class="brand" id="pageTitle">Application Inventory</span>
        </div>
        <div class="small text-muted" id="headerInfo" style="font-size: 0.75rem;">Loading...</div>
    </nav>
    <div style="height: 50px;"></div>

    <div class="container-fluid px-4 py-2">
        
        <!-- Stats Cards -->
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-blue clickable" data-filter="category" data-val="Service Connection Point"><div class="stat-label">SCPs</div><div class="stat-val" id="vSCP">0</div><i class="fa-solid fa-link stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-purple clickable" data-filter="category" data-val="Application Object"><div class="stat-label">App Objects</div><div class="stat-val" id="vAppClass">0</div><i class="fa-solid fa-cube stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="category" data-val="Application Server"><div class="stat-label">App Servers</div><div class="stat-val" id="vAppServers">0</div><i class="fa-solid fa-server stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-orange clickable" data-filter="category" data-val="Service Account"><div class="stat-label">Service Accts</div><div class="stat-val" id="vSvcAccounts">0</div><i class="fa-solid fa-user-cog stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-teal clickable" data-filter="category" data-val="Application OU"><div class="stat-label">App OUs</div><div class="stat-val" id="vAppOUs">0</div><i class="fa-solid fa-folder-tree stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label text-muted">Stale Servers</div><div class="stat-val text-muted" id="vStaleServers">0</div><i class="fa-solid fa-bed stat-icon"></i></div></div>
        </div>
        <div class="row g-2 mb-2">
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-red clickable" data-filter="status" data-val="Disabled"><div class="stat-label text-danger">Disabled Servers</div><div class="stat-val text-danger" id="vDisabledServers">0</div><i class="fa-solid fa-power-off stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-green clickable" data-filter="status" data-val="Active"><div class="stat-label text-success">Active Svc Accts</div><div class="stat-val text-success" id="vActiveSvc">0</div><i class="fa-solid fa-user-check stat-icon"></i></div></div>
            <div class="col-xl-2 col-md-4 col-6"><div class="stat-card sc-gray clickable" data-filter="status" data-val="Stale"><div class="stat-label text-muted">Stale Svc Accts</div><div class="stat-val text-muted" id="vStaleSvc">0</div><i class="fa-solid fa-user-slash stat-icon"></i></div></div>
        </div>

        <!-- Charts -->
        <div class="row g-2">
            <div class="col-lg-6">
                <div class="chart-card">
                    <div class="chart-header">
                        <span><i class="fa-solid fa-chart-bar me-1"></i> SPN Service Types</span>
                        <small class="text-muted fw-normal">(from servers & service accounts)</small>
                    </div>
                    <div class="chart-wrapper">
                        <div class="chart-scroll">
                            <div style="position: relative; height: 1000px; width: 100%">
                                <canvas id="spnChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-lg-6">
                <div class="chart-card">
                    <div class="chart-header"><span><i class="fa-solid fa-chart-pie me-1"></i> Application Categories</span></div>
                    <div class="chart-wrapper d-flex align-items-center justify-content-center position-relative">
                        <div style="width: 100%; height: 180px;">
                            <canvas id="categoryChart"></canvas>
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
                        <th><select class="form-control-sm-custom"><option value="">All</option><option value="Service Connection Point">SCP</option><option value="Application Object">App Obj</option><option value="Application Server">App Server</option><option value="Service Account">Svc Acct</option><option value="Application OU">App OU</option></select></th>
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
        <div class="text-center text-muted mt-2" style="font-size: 0.65rem;">Generated by AD Application Inventory Tool • Domain: <span id="footerDomain">-</span></div>
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

                $('#pageTitle').text(stats.Domain + ' Application Inventory');
                $('#headerInfo').text(`Generated: ${new Date().toLocaleString()}`);
                $('#footerDomain').text(stats.Domain);

                // Populate stats
                $('#vSCP').text(stats.TotalSCP);
                $('#vAppClass').text(stats.TotalAppClass);
                $('#vAppServers').text(stats.TotalAppServers);
                $('#vSvcAccounts').text(stats.TotalSvcAccounts);
                $('#vAppOUs').text(stats.TotalAppOUs);
                $('#vStaleServers').text(stats.StaleServers);
                $('#vDisabledServers').text(stats.DisabledServers);
                $('#vActiveSvc').text(stats.ActiveSvc);
                $('#vStaleSvc').text(stats.StaleSvc);
                $('#centerTotal').text(stats.TotalAppObjects + stats.TotalAppServers + stats.TotalSvcAccounts + stats.TotalAppOUs);

                // Chart: SPN Service Types
                const spnCtx = document.getElementById('spnChart').getContext('2d');
                const spnLabels = stats.SPNLabels || [];
                const spnCounts = stats.SPNCounts || [];
                const spnChartHeight = Math.max(300, spnLabels.length * 20);
                $('#spnChart').parent().height(spnChartHeight);
                new Chart(spnCtx, {
                    type: 'bar',
                    data: {
                        labels: spnLabels,
                        datasets: [{ label: 'SPN Count', data: spnCounts, backgroundColor: '#4f46e5', borderRadius: 3, barThickness: 8 }]
                    },
                    options: {
                        indexAxis: 'y',
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: { x: { display: false }, y: { grid: { display: false }, ticks: { font: { size: 9 } } } }
                    }
                });

                // Chart: Category Distribution (Pie)
                const catCtx = document.getElementById('categoryChart').getContext('2d');
                new Chart(catCtx, {
                    type: 'doughnut',
                    data: {
                        labels: ['SCPs', 'App Objects', 'App Servers', 'Service Accts', 'App OUs'],
                        datasets: [{
                            data: [stats.TotalSCP, stats.TotalAppClass, stats.TotalAppServers, stats.TotalSvcAccounts, stats.TotalAppOUs],
                            backgroundColor: ['#3b82f6', '#a855f7', '#22c55e', '#f97316', '#14b8a6'],
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
                    
                    let rowClass = '';
                    if (item.Status === 'Stale' || item.Status === 'Never Logged In') rowClass += ' row-stale';
                    if (item.Enabled === false) rowClass += ' row-disabled';

                    let detail = '';
                    if (item.SPNs) detail = item.SPNs;
                    else if (item.BindingInfo) detail = item.BindingInfo;
                    else if (item.Keywords) detail = item.Keywords;
                    else if (item.Path) detail = item.Path;
                    else if (item.SamAccount) detail = item.SamAccount;

                    return `<tr class="${rowClass}" data-category="${item.Category}" data-status="${item.Status}">
                        <td><span class="badge-s bg-info-cat">${item.Category}</span></td>
                        <td class="fw-bold clickable" style="color:#4f46e5" onclick="showDetail(${idx})">${item.Name}</td>
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
                        let colIdx = (filter === 'category') ? 0 : (filter === 'status' ? 2 : -1);
                        if (colIdx >= 0) {
                            table.column(colIdx).search('^' + val + '$', true, false).draw();
                        }
                    }
                    $('html, body').animate({ scrollTop: $(".table-card").offset().top - 80 }, 400);
                });

                // Export buttons
                $('#btnExcel').click(function() { table.button('.buttons-excel').trigger(); });
                $('#btnCsv').click(function() {
                    let csv = table.buttons(0).text('CSV').trigger();
                    setTimeout(() => { table.buttons(0).text('CSV'); }, 100);
                });
                $('#btnWord').click(function() {
                    let htmlContent = `
                        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
                        <head><meta charset="utf-8"><title>AD Application Report</title>
                        <style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #ccc; padding: 4px; font-size: 10pt; } th { background: #f0f0f0; }</style>
                        </head><body><h2>AD Application Inventory - ${stats.Domain}</h2>
                        <p>Generated: ${new Date().toLocaleString()}</p>
                        ${document.querySelector('#mainTable').outerHTML}
                        </body></html>`;
                    let blob = new Blob([htmlContent], { type: 'application/msword' });
                    let link = document.createElement('a');
                    link.href = URL.createObjectURL(blob);
                    link.download = 'AD_App_Report.doc';
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
$htmlPath = Join-Path $outputDir "AD_App_Inventory.html"
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($true))
Write-Host "HTML Dashboard exported to: $htmlPath" -ForegroundColor Green

# Final message
Write-Host "`nAll reports saved to folder:" -ForegroundColor Cyan
Write-Host "  $outputDir" -ForegroundColor Yellow
Write-Host "`nOpen the HTML dashboard and use the Export buttons for Excel/Word/CSV." -ForegroundColor White
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
