# Active Directory Infrastructure Documentation Script - Enhanced Version
# Run this on a Domain Controller with appropriate permissions
# Once the Script Completes it will open the users default web brower for the report and also exports .csv's for each section within the reports findings. Enjoy!!!

# Set output directory to user's desktop
$OutputPath = [Environment]::GetFolderPath("Desktop") + "\AD_Documentation"
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.html"
$ExcelFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.xlsx"

# Import required module
Import-Module ActiveDirectory

Write-Host "Starting Active Directory Infrastructure Documentation..." -ForegroundColor Green
Write-Host "Output directory: $OutputPath" -ForegroundColor Yellow
Write-Host "Go get a cup of COFFEE!! -  Honestly by the time you get back it should be done compiling the report, and it will be opened within your default browser!!!" -ForegroundColor Green

# Initialize collections for Excel export
$ExcelData = @{
    DomainControllers = @()
    ReplicationHealth = @()
    Sites = @()
    DNSZones = @()
    DHCPScopes = @()
    Tier0Accounts = @()
    Tier1Accounts = @()
    Tier2Accounts = @()
    ServiceAccounts = @()
    ExchangeServers = @()
    GroupPolicies = @()
}

# Initialize HTML report with collapsible sections
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Active Directory Infrastructure Documentation</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #003366; border-bottom: 3px solid #003366; padding-bottom: 10px; }
        h2 { color: #003366; margin-top: 30px; border-bottom: 2px solid #003366; padding-bottom: 5px; cursor: pointer; user-select: none; }
        h2:hover { background-color: #e6f3ff; }
        h2::before { content: '▼ '; font-size: 0.8em; }
        h2.collapsed::before { content: '▶ '; }
        h3 { color: #003366; margin-top: 20px; }
        .section-content { margin-left: 20px; }
        .collapsed-content { display: none; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #003366; color: white; padding: 12px; text-align: left; font-weight: bold; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #e6f0f5; }
        .info-box { background-color: white; padding: 15px; margin: 15px 0; border-left: 4px solid #003366; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .warning { border-left-color: #ff9900; }
        .success { border-left-color: #00cc66; }
        .error { border-left-color: #cc0000; }
        .timestamp { color: #666; font-size: 0.9em; }
        .critical { color: #cc0000; font-weight: bold; }
        .healthy { color: #00cc00; font-weight: bold; }
        .tier0 { background-color: #ffcccc; }
        .tier1 { background-color: #ffffcc; }
        .tier2 { background-color: #ccffcc; }
        .toggle-all { margin: 20px 0; padding: 10px 20px; background-color: #003366; color: white; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }
        .toggle-all:hover { background-color: #002244; }
        .search-container { background-color: white; padding: 20px; margin: 20px 0; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .search-box { width: 100%; max-width: 600px; padding: 12px 20px; font-size: 16px; border: 2px solid #003366; border-radius: 4px; box-sizing: border-box; }
        .search-box:focus { outline: none; border-color: #0055aa; }
        .search-info { margin-top: 10px; color: #666; font-size: 14px; }
        .highlight { background-color: #ffff00; font-weight: bold; }
        .search-button { margin-left: 10px; padding: 12px 24px; background-color: #003366; color: white; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }
        .search-button:hover { background-color: #002244; }
        .clear-button { margin-left: 10px; padding: 12px 24px; background-color: #666; color: white; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }
        .clear-button:hover { background-color: #444; }
    </style>
    <script>
        function toggleSection(element) {
            const content = element.nextElementSibling;
            const isCollapsed = content.classList.contains('collapsed-content');
            
            if (isCollapsed) {
                content.classList.remove('collapsed-content');
                element.classList.remove('collapsed');
            } else {
                content.classList.add('collapsed-content');
                element.classList.add('collapsed');
            }
        }
        
        function toggleAll() {
            const headers = document.querySelectorAll('h2');
            const firstHeader = headers[0];
            const firstContent = firstHeader.nextElementSibling;
            const shouldExpand = firstContent.classList.contains('collapsed-content');
            
            headers.forEach(header => {
                const content = header.nextElementSibling;
                if (shouldExpand) {
                    content.classList.remove('collapsed-content');
                    header.classList.remove('collapsed');
                } else {
                    content.classList.add('collapsed-content');
                    header.classList.add('collapsed');
                }
            });
        }
        
        window.onload = function() {
            document.querySelectorAll('h2').forEach(header => {
                header.addEventListener('click', function() { toggleSection(this); });
            });
        };
        
        function searchReport() {
            clearSearch();
            const searchTerm = document.getElementById('searchBox').value.toLowerCase().trim();
            
            if (searchTerm === '') {
                document.getElementById('searchInfo').textContent = 'Please enter a search term.';
                return;
            }
            
            let matchCount = 0;
            const sections = document.querySelectorAll('.section-content');
            
            // Expand all sections for searching
            document.querySelectorAll('h2').forEach(header => {
                const content = header.nextElementSibling;
                content.classList.remove('collapsed-content');
                header.classList.remove('collapsed');
            });
            
            // Search through all text content
            sections.forEach(section => {
                const walker = document.createTreeWalker(
                    section,
                    NodeFilter.SHOW_TEXT,
                    null,
                    false
                );
                
                let node;
                while (node = walker.nextNode()) {
                    const text = node.nodeValue;
                    if (text.toLowerCase().includes(searchTerm)) {
                        const span = document.createElement('span');
                        const regex = new RegExp('(' + searchTerm + ')', 'gi');
                        span.innerHTML = text.replace(regex, '<span class="highlight">$1</span>');
                        node.parentNode.replaceChild(span, node);
                        matchCount++;
                    }
                }
            });
            
            if (matchCount > 0) {
                document.getElementById('searchInfo').textContent = 'Found ' + matchCount + ' match(es). Results are highlighted in yellow.';
                // Scroll to first match
                const firstMatch = document.querySelector('.highlight');
                if (firstMatch) {
                    firstMatch.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            } else {
                document.getElementById('searchInfo').textContent = 'No matches found for "' + searchTerm + '".';
            }
        }
        
        function clearSearch() {
            // Remove all highlights
            document.querySelectorAll('.highlight').forEach(highlight => {
                const parent = highlight.parentNode;
                parent.replaceWith(parent.textContent);
            });
            document.getElementById('searchInfo').textContent = '';
        }
        
        function handleSearchKeyPress(event) {
            if (event.key === 'Enter') {
                searchReport();
            }
        }
    </script>
</head>
<body>
    <h1>Active Directory Infrastructure Documentation</h1>
    <p class="timestamp"><strong>Author: Stephen McKee - IT Operations - IGTPLC</strong></p>
    <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    <p class="timestamp">Output Location: $OutputPath</p>
    <button class="toggle-all" onclick="toggleAll()">Expand/Collapse All Sections</button>
    
    <div class="search-container">
        <h3 style="margin-top: 0;">Search Report</h3>
        <input type="text" id="searchBox" class="search-box" placeholder="Enter search term (server name, IP address, account name, etc.)..." onkeypress="handleSearchKeyPress(event)">
        <button class="search-button" onclick="searchReport()">Search</button>
        <button class="clear-button" onclick="clearSearch(); document.getElementById('searchBox').value = '';">Clear</button>
        <div id="searchInfo" class="search-info"></div>
    </div>
"@

# 1. FOREST INFORMATION
Write-Host "Gathering Forest Information..." -ForegroundColor Cyan
$Forest = Get-ADForest
$HTML += @"
    <h2>1. Forest Overview</h2>
    <div class="section-content">
    <div class="info-box success">
        <table>
            <tr><th>Property</th><th>Value</th></tr>
            <tr><td>Forest Name</td><td>$($Forest.Name)</td></tr>
            <tr><td>Forest Functional Level</td><td>$($Forest.ForestMode)</td></tr>
            <tr><td>Schema Master</td><td>$($Forest.SchemaMaster)</td></tr>
            <tr><td>Domain Naming Master</td><td>$($Forest.DomainNamingMaster)</td></tr>
            <tr><td>Root Domain</td><td>$($Forest.RootDomain)</td></tr>
            <tr><td>Total Domains</td><td>$($Forest.Domains.Count)</td></tr>
        </table>
    </div>
    <h3>Domains in Forest</h3>
    <ul>
"@
foreach ($domain in $Forest.Domains) {
    $HTML += "        <li>$domain</li>`n"
}
$HTML += "    </ul>`n</div>`n"

# 2. DOMAIN CONTROLLERS
Write-Host "Gathering Domain Controller Information..." -ForegroundColor Cyan
$DCs = Get-ADDomainController -Filter *
$HTML += @"
    <h2>2. Domain Controllers (Total: $($DCs.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Hostname</th>
            <th>Site</th>
            <th>IP Address</th>
            <th>OS Version</th>
            <th>Global Catalog</th>
            <th>FSMO Roles</th>
            <th>DNS Service</th>
        </tr>
"@

foreach ($DC in $DCs) {
    $FSMORoles = @()
    if ($DC.OperationMasterRoles) {
        $FSMORoles = $DC.OperationMasterRoles -join ", "
    } else {
        $FSMORoles = "None"
    }
    
    $IsGC = if ($DC.IsGlobalCatalog) { "Yes" } else { "No" }
    $IsGCHTML = if ($DC.IsGlobalCatalog) { "<span class='healthy'>Yes</span>" } else { "No" }
    
    # Check DNS service
    $DNSStatus = "Unknown"
    try {
        $DNSService = Get-Service -ComputerName $DC.HostName -Name DNS -ErrorAction SilentlyContinue
        if ($DNSService) {
            $DNSStatus = $DNSService.Status
        }
    } catch {
        $DNSStatus = "Unable to query"
    }
    
    $DNSStatusHTML = if ($DNSStatus -eq "Running") { "<span class='healthy'>Running</span>" } else { "<span class='critical'>$DNSStatus</span>" }
    
    # Add to Excel data
    $ExcelData.DomainControllers += [PSCustomObject]@{
        Hostname = $DC.HostName
        Site = $DC.Site
        IPAddress = $DC.IPv4Address
        OSVersion = $DC.OperatingSystem
        GlobalCatalog = $IsGC
        FSMORoles = $FSMORoles
        DNSService = $DNSStatus
    }
    
    $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($DC.Site)</td>
            <td>$($DC.IPv4Address)</td>
            <td>$($DC.OperatingSystem)</td>
            <td>$IsGCHTML</td>
            <td>$FSMORoles</td>
            <td>$DNSStatusHTML</td>
        </tr>
"@
}
$HTML += "    </table>`n</div>`n"

# 3. REPLICATION HEALTH
Write-Host "Checking Replication Health..." -ForegroundColor Cyan
$HTML += @"
    <h2>3. Active Directory Replication Health</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Source DC</th>
            <th>Destination DC</th>
            <th>Last Replication</th>
            <th>Status</th>
            <th>Failures</th>
        </tr>
"@

foreach ($DC in $DCs) {
    try {
        $ReplPartners = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
        foreach ($Partner in $ReplPartners) {
            $LastRepl = $Partner.LastReplicationSuccess
            $TimeSince = (Get-Date) - $LastRepl
            $Status = if ($TimeSince.TotalHours -lt 24) { "Healthy" } else { "Warning" }
            $StatusHTML = if ($TimeSince.TotalHours -lt 24) { "<span class='healthy'>Healthy</span>" } else { "<span class='critical'>Warning</span>" }
            $Failures = $Partner.ConsecutiveReplicationFailures
            $FailuresHTML = if ($Partner.ConsecutiveReplicationFailures -eq 0) { "<span class='healthy'>0</span>" } else { "<span class='critical'>$($Partner.ConsecutiveReplicationFailures)</span>" }
            
            # Add to Excel data
            $ExcelData.ReplicationHealth += [PSCustomObject]@{
                SourceDC = $Partner.Partner
                DestinationDC = $DC.HostName
                LastReplication = $LastRepl.ToString("yyyy-MM-dd HH:mm:ss")
                Status = $Status
                Failures = $Failures
            }
            
            $HTML += @"
        <tr>
            <td>$($Partner.Partner)</td>
            <td>$($DC.HostName)</td>
            <td>$($LastRepl.ToString("yyyy-MM-dd HH:mm:ss"))</td>
            <td>$StatusHTML</td>
            <td>$FailuresHTML</td>
        </tr>
"@
        }
    } catch {
        $HTML += @"
        <tr>
            <td colspan="5"><span class='critical'>Unable to query replication data for $($DC.HostName)</span></td>
        </tr>
"@
    }
}
$HTML += "    </table>`n</div>`n"

# 4. SITES AND SUBNETS
Write-Host "Gathering Sites and Subnets..." -ForegroundColor Cyan
$Sites = Get-ADReplicationSite -Filter *
$HTML += @"
    <h2>4. Active Directory Sites (Total: $($Sites.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Site Name</th>
            <th>Description</th>
            <th>Subnets</th>
            <th>Domain Controllers</th>
        </tr>
"@

foreach ($Site in $Sites) {
    $Subnets = Get-ADReplicationSubnet -Filter "Site -eq '$($Site.DistinguishedName)'" | Select-Object -ExpandProperty Name
    $SubnetList = if ($Subnets) { ($Subnets -join ", ") } else { "None configured" }
    $SubnetListHTML = if ($Subnets) { ($Subnets -join "<br>") } else { "None configured" }
    
    $SiteDCs = $DCs | Where-Object { $_.Site -eq $Site.Name } | Select-Object -ExpandProperty HostName
    $DCList = if ($SiteDCs) { ($SiteDCs -join ", ") } else { "None" }
    $DCListHTML = if ($SiteDCs) { ($SiteDCs -join "<br>") } else { "None" }
    
    # Add to Excel data
    $ExcelData.Sites += [PSCustomObject]@{
        SiteName = $Site.Name
        Description = $Site.Description
        Subnets = $SubnetList
        DomainControllers = $DCList
    }
    
    $HTML += @"
        <tr>
            <td>$($Site.Name)</td>
            <td>$($Site.Description)</td>
            <td>$SubnetListHTML</td>
            <td>$DCListHTML</td>
        </tr>
"@
}
$HTML += "    </table>`n</div>`n"

# 5. DNS SERVERS
Write-Host "Gathering DNS Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>5. DNS Servers on Domain Controllers</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Server</th>
            <th>DNS Zones</th>
            <th>Zone Type</th>
            <th>Dynamic Updates</th>
        </tr>
"@

foreach ($DC in $DCs) {
    try {
        $DNSZones = Get-DnsServerZone -ComputerName $DC.HostName -ErrorAction SilentlyContinue
        if ($DNSZones) {
            foreach ($Zone in $DNSZones | Where-Object { $_.ZoneType -ne "Cache" }) {
                # Add to Excel data
                $ExcelData.DNSZones += [PSCustomObject]@{
                    Server = $DC.HostName
                    ZoneName = $Zone.ZoneName
                    ZoneType = $Zone.ZoneType
                    DynamicUpdate = $Zone.DynamicUpdate
                }
                
                $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($Zone.ZoneName)</td>
            <td>$($Zone.ZoneType)</td>
            <td>$($Zone.DynamicUpdate)</td>
        </tr>
"@
            }
        }
    } catch {
        $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td colspan="3"><span class='critical'>Unable to query DNS zones</span></td>
        </tr>
"@
    }
}
$HTML += "    </table>`n</div>`n"

# 6. DHCP SERVERS
Write-Host "Gathering DHCP Server Information..." -ForegroundColor Cyan
try {
    $DHCPServers = Get-DhcpServerInDC -ErrorAction Stop
    $HTML += @"
    <h2>6. DHCP Servers (Total: $($DHCPServers.Count))</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>Server Name</th>
            <th>IP Address</th>
            <th>Scopes</th>
            <th>Scope Range</th>
            <th>Scope State</th>
        </tr>
"@
    
    foreach ($DHCPServer in $DHCPServers) {
        try {
            $Scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer.DnsName -ErrorAction SilentlyContinue
            if ($Scopes) {
                foreach ($Scope in $Scopes) {
                    $ScopeState = $Scope.State
                    $ScopeStateHTML = if ($Scope.State -eq "Active") { "<span class='healthy'>Active</span>" } else { "<span class='critical'>$($Scope.State)</span>" }
                    
                    # Add to Excel data
                    $ExcelData.DHCPScopes += [PSCustomObject]@{
                        ServerName = $DHCPServer.DnsName
                        IPAddress = $DHCPServer.IPAddress
                        ScopeName = $Scope.Name
                        StartRange = $Scope.StartRange
                        EndRange = $Scope.EndRange
                        State = $ScopeState
                    }
                    
                    $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td>$($Scope.Name)</td>
            <td>$($Scope.StartRange) - $($Scope.EndRange)</td>
            <td>$ScopeStateHTML</td>
        </tr>
"@
                }
            } else {
                $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td colspan="3">No scopes configured or unable to query</td>
        </tr>
"@
            }
        } catch {
            $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td colspan="3"><span class='critical'>Unable to query scopes</span></td>
        </tr>
"@
        }
    }
    $HTML += "    </table>`n</div>`n"
} catch {
    $HTML += @"
    <h2>6. DHCP Servers</h2>
    <div class="section-content">
    <div class="info-box warning">
        <p><span class='critical'>Unable to retrieve DHCP servers from Active Directory.</span></p>
        <p>This may be because no DHCP servers are authorized or you lack permissions.</p>
    </div>
    </div>
"@
}

# 7. PRIVILEGED ACCOUNTS - TIER 0
Write-Host "Gathering Tier 0 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>7. Tier 0 Accounts (Highest Privilege - Domain/Enterprise/Schema Admins)</h2>
    <div class="section-content">
    <div class="info-box">
        <p><strong>Note:</strong> This section includes "Tier 0 Operators" group and standard administrative groups. Tier 0 Operators may contain Domain Admins and other highly privileged accounts.</p>
    </div>
"@

$Tier0Groups = @(
    "Tier 0 Operators",
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators"
)

# Also search for groups in Admin OU
try {
    $AdminOU = Get-ADOrganizationalUnit -Filter "Name -eq 'Admin'" -ErrorAction SilentlyContinue
    if ($AdminOU) {
        Write-Host "  Found Admin OU, searching for additional Tier 0 groups..." -ForegroundColor Yellow
        $Tier0GroupsInOU = Get-ADGroup -Filter * -SearchBase $AdminOU.DistinguishedName -SearchScope Subtree | 
            Where-Object { $_.Name -like "*Tier*0*" -or $_.Name -like "*T0*" }
        foreach ($Group in $Tier0GroupsInOU) {
            if ($Tier0Groups -notcontains $Group.Name) {
                $Tier0Groups += $Group.Name
                Write-Host "    Added group: $($Group.Name)" -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host "  Note: Could not search Admin OU for additional Tier 0 groups" -ForegroundColor Yellow
}

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

$Tier0Count = 0
foreach ($GroupName in $Tier0Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            Write-Host "  Processing group: $GroupName" -ForegroundColor Cyan
            $Members = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction SilentlyContinue
            foreach ($Member in $Members) {
                try {
                    if ($Member.objectClass -eq "user") {
                        $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                        $EnabledStatus = $User.Enabled
                        $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                        $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                        $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                        
                        $Tier0Count++
                        
                        # Add to Excel data
                        $ExcelData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $User.Name
                            SamAccountName = $User.SamAccountName
                            AccountType = "User"
                            Enabled = $EnabledStatus
                            LastLogon = $LastLogon
                            PasswordLastSet = $PwdLastSet
                        }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($User.Name) ($($User.SamAccountName))</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                    } else {
                        $ExcelData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $Member.Name
                            SamAccountName = $Member.SamAccountName
                            AccountType = $Member.objectClass
                            Enabled = "N/A"
                            LastLogon = "N/A"
                            PasswordLastSet = "N/A"
                        }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                    }
                } catch {
                    $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                }
            }
        }
    } catch {
        $HTML += "<tr class='tier0'><td>$GroupName</td><td colspan='5'><span class='critical'>Unable to query group</span></td></tr>"
    }
}

if ($Tier0Count -eq 0) {
    $HTML += "<tr class='tier0'><td colspan='6'>No Tier 0 accounts found. Please verify 'Tier 0 Operators' group exists in your Admin OU.</td></tr>"
}

$HTML += "</table></div>"

Write-Host "  Found $Tier0Count Tier 0 user accounts" -ForegroundColor Green

# 8. TIER 1 ACCOUNTS
Write-Host "Gathering Tier 1 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>8. Tier 1 Accounts (Server/Infrastructure Management)</h2>
    <div class="section-content">
    <div class="info-box">
        <p><strong>Note:</strong> This section includes "Tier 1 Operators" group and standard server management groups.</p>
    </div>
"@

$Tier1Groups = @(
    "Tier 1 Operators",
    "Server Operators",
    "Backup Operators",
    "Account Operators",
    "Print Operators"
)

# Also search for groups in Admin OU
try {
    $AdminOU = Get-ADOrganizationalUnit -Filter "Name -eq 'Admin'" -ErrorAction SilentlyContinue
    if ($AdminOU) {
        Write-Host "  Searching Admin OU for additional Tier 1 groups..." -ForegroundColor Yellow
        $Tier1GroupsInOU = Get-ADGroup -Filter * -SearchBase $AdminOU.DistinguishedName -SearchScope Subtree | 
            Where-Object { $_.Name -like "*Tier*1*" -or $_.Name -like "*T1*" }
        foreach ($Group in $Tier1GroupsInOU) {
            if ($Tier1Groups -notcontains $Group.Name) {
                $Tier1Groups += $Group.Name
                Write-Host "    Added group: $($Group.Name)" -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host "  Note: Could not search Admin OU for additional Tier 1 groups" -ForegroundColor Yellow
}

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

$Tier1Count = 0
foreach ($GroupName in $Tier1Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            Write-Host "  Processing group: $GroupName" -ForegroundColor Cyan
            $Members = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = $User.Enabled
                            $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            $Tier1Count++
                            
                            # Add to Excel data
                            $ExcelData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                SamAccountName = $User.SamAccountName
                                AccountType = "User"
                                Enabled = $EnabledStatus
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($User.Name) ($($User.SamAccountName))</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            $ExcelData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
                                SamAccountName = $Member.SamAccountName
                                AccountType = $Member.objectClass
                                Enabled = "N/A"
                                LastLogon = "N/A"
                                PasswordLastSet = "N/A"
                            }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                        }
                    } catch {
                        $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                    }
                }
            } else {
                $HTML += "<tr class='tier1'><td>$GroupName</td><td colspan='5'>No members</td></tr>"
            }
        }
    } catch {
        $HTML += "<tr class='tier1'><td>$GroupName</td><td colspan='5'><span class='critical'>Unable to query group</span></td></tr>"
    }
}

if ($Tier1Count -eq 0) {
    $HTML += "<tr class='tier1'><td colspan='6'>No Tier 1 accounts found. Please verify 'Tier 1 Operators' group exists in your Admin OU.</td></tr>"
}

$HTML += "</table></div>"

Write-Host "  Found $Tier1Count Tier 1 user accounts" -ForegroundColor Green

# 9. TIER 2 ACCOUNTS
Write-Host "Gathering Tier 2 Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>9. Tier 2 Accounts (User/Workstation Management)</h2>
    <div class="section-content">
    <div class="info-box">
        <p><strong>Note:</strong> This section includes "Tier 2 Operators" group and standard help desk/desktop support groups.</p>
    </div>
"@

$Tier2Groups = @(
    "Tier 2 Operators",
    "Help Desk",
    "Helpdesk Operators",
    "Desktop Support",
    "Remote Desktop Users"
)

# Also search for groups in Admin OU
try {
    $AdminOU = Get-ADOrganizationalUnit -Filter "Name -eq 'Admin'" -ErrorAction SilentlyContinue
    if ($AdminOU) {
        Write-Host "  Searching Admin OU for additional Tier 2 groups..." -ForegroundColor Yellow
        $Tier2GroupsInOU = Get-ADGroup -Filter * -SearchBase $AdminOU.DistinguishedName -SearchScope Subtree | 
            Where-Object { $_.Name -like "*Tier*2*" -or $_.Name -like "*T2*" }
        foreach ($Group in $Tier2GroupsInOU) {
            if ($Tier2Groups -notcontains $Group.Name) {
                $Tier2Groups += $Group.Name
                Write-Host "    Added group: $($Group.Name)" -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host "  Note: Could not search Admin OU for additional Tier 2 groups" -ForegroundColor Yellow
}

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

$Tier2Count = 0
foreach ($GroupName in $Tier2Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            Write-Host "  Processing group: $GroupName" -ForegroundColor Cyan
            $Members = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = $User.Enabled
                            $EnabledStatusHTML = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            $Tier2Count++
                            
                            # Add to Excel data
                            $ExcelData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                SamAccountName = $User.SamAccountName
                                AccountType = "User"
                                Enabled = $EnabledStatus
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($User.Name) ($($User.SamAccountName))</td><td>User</td><td>$EnabledStatusHTML</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            $ExcelData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
                                SamAccountName = $Member.SamAccountName
                                AccountType = $Member.objectClass
                                Enabled = "N/A"
                                LastLogon = "N/A"
                                PasswordLastSet = "N/A"
                            }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($Member.Name)</td><td>$($Member.objectClass)</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>"
                        }
                    } catch {
                        $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($Member.Name)</td><td>Unknown</td><td colspan='3'>Error querying</td></tr>"
                    }
                }
            } else {
                $HTML += "<tr class='tier2'><td>$GroupName</td><td colspan='5'>No members</td></tr>"
            }
        }
    } catch {
        continue
    }
}

if ($Tier2Count -eq 0) {
    $HTML += "<tr class='tier2'><td colspan='6'>No Tier 2 accounts found. Please verify 'Tier 2 Operators' group exists in your Admin OU.</td></tr>"
}

$HTML += "</table></div>"

Write-Host "  Found $Tier2Count Tier 2 user accounts" -ForegroundColor Green

# 10. SERVICE ACCOUNTS
Write-Host "Gathering Service Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>10. Service Accounts</h2>
    <div class="section-content">
"@

$ServiceAccounts = Get-ADUser -Filter * -Properties ServicePrincipalName, Description, PasswordLastSet, LastLogonDate, Enabled | 
    Where-Object { 
        ($_.ServicePrincipalName -ne $null) -or 
        ($_.SamAccountName -like "svc_*") -or 
        ($_.SamAccountName -like "svc-*") -or
        ($_.SamAccountName -like "*service*") -or
        ($_.Description -like "*service account*")
    }

$HTML += @"
    <table>
        <tr>
            <th>Account Name</th>
            <th>Description</th>
            <th>Enabled</th>
            <th>Password Last Set</th>
            <th>Last Logon</th>
            <th>SPN Count</th>
        </tr>
"@

foreach ($SvcAcct in $ServiceAccounts) {
    $EnabledStatus = $SvcAcct.Enabled
    $EnabledStatusHTML = if ($SvcAcct.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
    $PwdLastSet = if ($SvcAcct.PasswordLastSet) { $SvcAcct.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
    $LastLogon = if ($SvcAcct.LastLogonDate) { $SvcAcct.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
    $SPNCount = if ($SvcAcct.ServicePrincipalName) { $SvcAcct.ServicePrincipalName.Count } else { 0 }
    
    # Add to Excel data
    $ExcelData.ServiceAccounts += [PSCustomObject]@{
        AccountName = $SvcAcct.SamAccountName
        Description = $SvcAcct.Description
        Enabled = $EnabledStatus
        PasswordLastSet = $PwdLastSet
        LastLogon = $LastLogon
        SPNCount = $SPNCount
    }
    
    $HTML += @"
        <tr>
            <td>$($SvcAcct.SamAccountName)</td>
            <td>$($SvcAcct.Description)</td>
            <td>$EnabledStatusHTML</td>
            <td>$PwdLastSet</td>
            <td>$LastLogon</td>
            <td>$SPNCount</td>
        </tr>
"@
}

$HTML += "</table></div>"

# 11. EXCHANGE SERVERS
Write-Host "Gathering Exchange Server Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>11. Exchange Servers</h2>
    <div class="section-content">
"@

try {
    $ConfigNC = (Get-ADRootDSE).configurationNamingContext
    $ExchangeServers = Get-ADObject -Filter {objectClass -eq "msExchExchangeServer"} -SearchBase $ConfigNC -Properties Name, msExchServerSite, serialNumber, versionNumber, msExchCurrentServerRoles, networkAddress, whenCreated
    
    if ($ExchangeServers) {
        $HTML += @"
        <table>
            <tr>
                <th>Server Name</th>
                <th>Site</th>
                <th>Roles</th>
                <th>Version</th>
                <th>FQDN</th>
                <th>Created</th>
            </tr>
"@
        
        foreach ($ExchServer in $ExchangeServers) {
            $Roles = switch ($ExchServer.msExchCurrentServerRoles) {
                2 { "Mailbox" }
                4 { "Client Access" }
                16 { "Unified Messaging" }
                32 { "Hub Transport" }
                64 { "Edge Transport" }
                54 { "Mailbox, Client Access, Hub Transport" }
                default { $ExchServer.msExchCurrentServerRoles }
            }
            
            $SiteName = if ($ExchServer.msExchServerSite) {
                ($ExchServer.msExchServerSite -split ",")[0] -replace "CN=", ""
            } else {
                "Unknown"
            }
            
            $FQDN = "N/A"
            if ($ExchServer.networkAddress) {
                $FQDN = ($ExchServer.networkAddress | Where-Object { $_ -like "ncacn_ip_tcp:*" }) -replace "ncacn_ip_tcp:", ""
            }
            
            $Version = "Unknown"
            if ($ExchServer.serialNumber) {
                $VersionNumber = $ExchServer.serialNumber
                if ($VersionNumber -like "Version 15.2*") { $Version = "Exchange 2019" }
                elseif ($VersionNumber -like "Version 15.1*") { $Version = "Exchange 2016" }
                elseif ($VersionNumber -like "Version 15.0*") { $Version = "Exchange 2013" }
                elseif ($VersionNumber -like "Version 14.*") { $Version = "Exchange 2010" }
                else { $Version = $VersionNumber }
            }
            
            # Add to Excel data
            $ExcelData.ExchangeServers += [PSCustomObject]@{
                ServerName = $ExchServer.Name
                Site = $SiteName
                Roles = $Roles
                Version = $Version
                FQDN = $FQDN
                Created = $ExchServer.whenCreated.ToString("yyyy-MM-dd")
            }
            
            $HTML += @"
            <tr>
                <td>$($ExchServer.Name)</td>
                <td>$SiteName</td>
                <td>$Roles</td>
                <td>$Version</td>
                <td>$FQDN</td>
                <td>$($ExchServer.whenCreated.ToString("yyyy-MM-dd"))</td>
            </tr>
"@
        }
        
        $HTML += "</table>"
    } else {
        $HTML += @"
        <div class="info-box">
            <p>No Exchange servers found in Active Directory.</p>
        </div>
"@
    }
} catch {
    $HTML += @"
    <div class="info-box warning">
        <p><span class='critical'>Unable to query Exchange servers from Active Directory.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 12. GROUP POLICIES
Write-Host "Gathering Group Policy Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>12. Group Policy Objects</h2>
    <div class="section-content">
"@

try {
    $GPOs = Get-GPO -All
    $HTML += @"
    <table>
        <tr>
            <th>GPO Name</th>
            <th>Status</th>
            <th>Created</th>
            <th>Modified</th>
            <th>User Version</th>
            <th>Computer Version</th>
        </tr>
"@
    
    foreach ($GPO in $GPOs) {
        $Status = $GPO.GpoStatus
        $StatusHTML = if ($Status -eq "AllSettingsEnabled") { 
            "<span class='healthy'>Enabled</span>" 
        } elseif ($Status -like "*Disabled") { 
            "<span class='critical'>$Status</span>" 
        } else { 
            $Status 
        }
        
        # Add to Excel data
        $ExcelData.GroupPolicies += [PSCustomObject]@{
            GPOName = $GPO.DisplayName
            Status = $Status
            Created = $GPO.CreationTime.ToString("yyyy-MM-dd HH:mm")
            Modified = $GPO.ModificationTime.ToString("yyyy-MM-dd HH:mm")
            UserVersion = $GPO.User.DSVersion
            ComputerVersion = $GPO.Computer.DSVersion
        }
        
        $HTML += @"
        <tr>
            <td>$($GPO.DisplayName)</td>
            <td>$StatusHTML</td>
            <td>$($GPO.CreationTime.ToString("yyyy-MM-dd HH:mm"))</td>
            <td>$($GPO.ModificationTime.ToString("yyyy-MM-dd HH:mm"))</td>
            <td>$($GPO.User.DSVersion)</td>
            <td>$($GPO.Computer.DSVersion)</td>
        </tr>
"@
    }
    
    $HTML += "</table>"
} catch {
    $HTML += @"
    <div class="info-box warning">
        <p><span class='critical'>Unable to retrieve Group Policy Objects.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 13. FSMO ROLES SUMMARY
Write-Host "Generating FSMO Roles Summary..." -ForegroundColor Cyan
$Domain = Get-ADDomain
$HTML += @"
    <h2>13. FSMO Roles Summary</h2>
    <div class="section-content">
    <div class="info-box">
        <h3>Forest-Wide Roles</h3>
        <table>
            <tr><th>Role</th><th>Holder</th></tr>
            <tr><td>Schema Master</td><td>$($Forest.SchemaMaster)</td></tr>
            <tr><td>Domain Naming Master</td><td>$($Forest.DomainNamingMaster)</td></tr>
        </table>
        
        <h3>Domain-Wide Roles</h3>
        <table>
            <tr><th>Role</th><th>Holder</th></tr>
            <tr><td>PDC Emulator</td><td>$($Domain.PDCEmulator)</td></tr>
            <tr><td>RID Master</td><td>$($Domain.RIDMaster)</td></tr>
            <tr><td>Infrastructure Master</td><td>$($Domain.InfrastructureMaster)</td></tr>
        </table>
    </div>
    </div>
"@

# Close HTML
$HTML += @"
    <div class="info-box" style="margin-top: 40px;">
        <p><strong>Documentation Complete</strong></p>
        <p>This report provides a comprehensive overview of your Active Directory infrastructure.</p>
        <p><strong>Files Generated:</strong></p>
        <ul>
            <li>HTML Report: $ReportFile</li>
            <li>Excel Report: $ExcelFile</li>
        </ul>
    </div>
</body>
</html>
"@

# Save HTML report
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Host "HTML Documentation Complete!" -ForegroundColor Green
Write-Host "Report saved to: $ReportFile" -ForegroundColor Yellow

# Export to Excel using ImportExcel module or COM object
Write-Host "Exporting data to Excel..." -ForegroundColor Cyan

# Try using ImportExcel module first (if available)
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    
    # Export each dataset to a separate worksheet
    $ExcelData.DomainControllers | Export-Excel -Path $ExcelFile -WorksheetName "Domain Controllers" -AutoSize -TableName "DomainControllers" -TableStyle Medium2
    $ExcelData.ReplicationHealth | Export-Excel -Path $ExcelFile -WorksheetName "Replication Health" -AutoSize -TableName "ReplicationHealth" -TableStyle Medium2
    $ExcelData.Sites | Export-Excel -Path $ExcelFile -WorksheetName "Sites and Subnets" -AutoSize -TableName "Sites" -TableStyle Medium2
    $ExcelData.DNSZones | Export-Excel -Path $ExcelFile -WorksheetName "DNS Zones" -AutoSize -TableName "DNSZones" -TableStyle Medium2
    $ExcelData.DHCPScopes | Export-Excel -Path $ExcelFile -WorksheetName "DHCP Scopes" -AutoSize -TableName "DHCPScopes" -TableStyle Medium2
    $ExcelData.Tier0Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 0 Accounts" -AutoSize -TableName "Tier0" -TableStyle Medium2
    $ExcelData.Tier1Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 1 Accounts" -AutoSize -TableName "Tier1" -TableStyle Medium2
    $ExcelData.Tier2Accounts | Export-Excel -Path $ExcelFile -WorksheetName "Tier 2 Accounts" -AutoSize -TableName "Tier2" -TableStyle Medium2
    $ExcelData.ServiceAccounts | Export-Excel -Path $ExcelFile -WorksheetName "Service Accounts" -AutoSize -TableName "ServiceAccounts" -TableStyle Medium2
    $ExcelData.ExchangeServers | Export-Excel -Path $ExcelFile -WorksheetName "Exchange Servers" -AutoSize -TableName "ExchangeServers" -TableStyle Medium2
    $ExcelData.GroupPolicies | Export-Excel -Path $ExcelFile -WorksheetName "Group Policies" -AutoSize -TableName "GroupPolicies" -TableStyle Medium2
    
    Write-Host "Excel file created successfully using ImportExcel module!" -ForegroundColor Green
    Write-Host "Excel file saved to: $ExcelFile" -ForegroundColor Yellow
} else {
    # Fallback to CSV exports
    Write-Host "ImportExcel module not found. Exporting to CSV files instead..." -ForegroundColor Yellow
    Write-Host "To create Excel files, install the module with: Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    
    $ExcelData.DomainControllers | Export-Csv "$OutputPath\DomainControllers_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ReplicationHealth | Export-Csv "$OutputPath\ReplicationHealth_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Sites | Export-Csv "$OutputPath\Sites_$Timestamp.csv" -NoTypeInformation
    $ExcelData.DNSZones | Export-Csv "$OutputPath\DNSZones_$Timestamp.csv" -NoTypeInformation
    $ExcelData.DHCPScopes | Export-Csv "$OutputPath\DHCPScopes_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier0Accounts | Export-Csv "$OutputPath\Tier0Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier1Accounts | Export-Csv "$OutputPath\Tier1Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.Tier2Accounts | Export-Csv "$OutputPath\Tier2Accounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ServiceAccounts | Export-Csv "$OutputPath\ServiceAccounts_$Timestamp.csv" -NoTypeInformation
    $ExcelData.ExchangeServers | Export-Csv "$OutputPath\ExchangeServers_$Timestamp.csv" -NoTypeInformation
    $ExcelData.GroupPolicies | Export-Csv "$OutputPath\GroupPolicies_$Timestamp.csv" -NoTypeInformation
    
    Write-Host "CSV files created successfully!" -ForegroundColor Green
    Write-Host "CSV files saved to: $OutputPath" -ForegroundColor Yellow
}

Write-Host "Opening HTML report in default browser..." -ForegroundColor Cyan
Start-Process $ReportFile

Write-Host "========================================" -ForegroundColor Green
Write-Host "BOO-YAA!! !Documentation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "All files saved to: $OutputPath" -ForegroundColor Yellow
Write-Host "This Report should IMPRESS YYOUR BOSS!!! ENJOY!!!: $OutputPath" -ForegroundColor Yellow
