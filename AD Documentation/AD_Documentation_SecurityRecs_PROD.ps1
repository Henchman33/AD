# Active Directory Infrastructure Documentation Script - Enhanced Version
# Author : Steve McKee - Server Administrator 2
# Run this on a Domain Controller with appropriate permissions
# Added Security Recommendations at the end of the HTML report

# Set output directory to user's desktop
$OutputPath = [Environment]::GetFolderPath("Desktop") + "\AD_Documentation"
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.html"

# Import required module
Import-Module ActiveDirectory

Write-Host "Starting Active Directory Infrastructure Documentation..." -ForegroundColor Green
Write-Host "Output directory: $OutputPath" -ForegroundColor Yellow

# Initialize HTML report with collapsible sections and search functionality
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Active Directory Infrastructure Documentation</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f8f9fa; color: #333; }
        h1 { color: #1a365d; border-bottom: 3px solid #2d3748; padding-bottom: 10px; }
        h2 { color: #2d3748; margin-top: 30px; border-bottom: 2px solid #4a5568; padding-bottom: 5px; cursor: pointer; user-select: none; background-color: #edf2f7; padding: 10px; border-radius: 4px; }
        h2:hover { background-color: #e2e8f0; }
        h2::before { content: '▼ '; font-size: 0.8em; color: #2b6cb0; }
        h2.collapsed::before { content: '▶ '; }
        h3 { color: #2c5282; margin-top: 20px; border-left: 4px solid #2b6cb0; padding-left: 10px; }
        .section-content { margin-left: 15px; margin-top: 10px; }
        .collapsed-content { display: none; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-radius: 4px; overflow: hidden; }
        th { background-color: #2c5282; color: white; padding: 12px; text-align: left; font-weight: bold; border-bottom: 2px solid #2d3748; }
        td { padding: 10px; border-bottom: 1px solid #e2e8f0; }
        tr:hover { background-color: #f7fafc; }
        .info-box { background-color: white; padding: 15px; margin: 15px 0; border-left: 4px solid #2c5282; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-radius: 4px; }
        .warning { border-left-color: #dd6b20; background-color: #fffaf0; }
        .success { border-left-color: #38a169; background-color: #f0fff4; }
        .error { border-left-color: #e53e3e; background-color: #fff5f5; }
        .timestamp { color: #718096; font-size: 0.9em; }
        .critical { color: #e53e3e; font-weight: bold; }
        .healthy { color: #38a169; font-weight: bold; }
        .tier0 { background-color: #fed7d7; }
        .tier1 { background-color: #feebc8; }
        .tier2 { background-color: #c6f6d5; }
        .toggle-all { margin: 20px 0; padding: 10px 20px; background-color: #2c5282; color: white; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; transition: background-color 0.2s; }
        .toggle-all:hover { background-color: #2a4365; }
        .search-container { margin: 20px 0; padding: 15px; background-color: white; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .search-box { width: 100%; padding: 10px; border: 2px solid #e2e8f0; border-radius: 4px; font-size: 14px; transition: border-color 0.2s; }
        .search-box:focus { outline: none; border-color: #2b6cb0; }
        .search-stats { margin-top: 10px; color: #718096; font-size: 0.9em; }
        .highlight { background-color: #ffeb3b; padding: 2px; border-radius: 2px; }
        .no-results { color: #e53e3e; padding: 20px; text-align: center; background-color: #fff5f5; border-radius: 4px; }
        .result-count { background-color: #2c5282; color: white; padding: 2px 6px; border-radius: 10px; font-size: 0.8em; margin-left: 5px; }
        .section-header { display: flex; justify-content: space-between; align-items: center; }
        .copy-button { background-color: #2d3748; color: white; border: none; padding: 5px 10px; border-radius: 3px; cursor: pointer; font-size: 0.8em; margin-left: 10px; }
        .copy-button:hover { background-color: #4a5568; }
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
        
        function performSearch() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const sections = document.querySelectorAll('.section-content');
            let totalMatches = 0;
            let sectionsWithMatches = 0;
            
            // Remove existing highlights
            document.querySelectorAll('.highlight').forEach(el => {
                const parent = el.parentNode;
                parent.replaceChild(document.createTextNode(el.textContent), el);
                parent.normalize();
            });
            
            if (searchTerm.length < 2) {
                document.getElementById('searchStats').innerHTML = 'Enter at least 2 characters to search';
                // Show all sections
                sections.forEach(section => {
                    section.style.display = 'block';
                });
                document.querySelectorAll('h2').forEach(header => {
                    header.style.display = 'flex';
                });
                // Remove result counts
                document.querySelectorAll('.result-count').forEach(el => el.remove());
                return;
            }
            
            sections.forEach(section => {
                let sectionMatches = 0;
                const textNodes = [];
                
                function findTextNodes(node) {
                    if (node.nodeType === Node.TEXT_NODE) {
                        textNodes.push(node);
                    } else {
                        node.childNodes.forEach(findTextNodes);
                    }
                }
                
                findTextNodes(section);
                
                textNodes.forEach(textNode => {
                    const text = textNode.textContent;
                    const lowerText = text.toLowerCase();
                    if (lowerText.includes(searchTerm)) {
                        const escapedSearchTerm = searchTerm.replace(/[.*+?^`{ }()|[\]\\]/g, '\\$&');
                        const regex = new RegExp(escapedSearchTerm, 'gi');
                        const newText = text.replace(regex, match => {
                            sectionMatches++;
                            totalMatches++;
                            return '<span class="highlight">' + match + '</span>';
                        });
                        
                        const newElement = document.createElement('span');
                        newElement.innerHTML = newText;
                        textNode.parentNode.replaceChild(newElement, textNode);
                    }
                });
                
                // Show/hide sections based on matches
                const header = section.previousElementSibling;
                if (sectionMatches > 0) {
                    section.style.display = 'block';
                    header.style.display = 'flex';
                    sectionsWithMatches++;
                    
                    // Update section header with match count
                    let countSpan = header.querySelector('.result-count');
                    if (!countSpan) {
                        countSpan = document.createElement('span');
                        countSpan.className = 'result-count';
                        const headerText = header.querySelector('.section-header-text');
                        if (headerText) {
                            headerText.appendChild(countSpan);
                        }
                    }
                    countSpan.textContent = sectionMatches;
                } else {
                    section.style.display = 'none';
                    header.style.display = 'none';
                    // Remove result count if no matches
                    const countSpan = header.querySelector('.result-count');
                    if (countSpan) {
                        countSpan.remove();
                    }
                }
            });
            
            // Update search statistics
            const statsElement = document.getElementById('searchStats');
            if (totalMatches > 0) {
                statsElement.innerHTML = 'Found <strong>' + totalMatches + '</strong> matches in <strong>' + sectionsWithMatches + '</strong> sections';
                statsElement.style.color = '#38a169';
            } else {
                statsElement.innerHTML = '<div class="no-results">No results found for "<strong>' + searchTerm + '</strong>"</div>';
            }
        }
        
        function clearSearch() {
            document.getElementById('searchInput').value = '';
            performSearch();
        }
        
        function copyToClipboard(text) {
            navigator.clipboard.writeText(text).then(function() {
                alert('Copied to clipboard: ' + text);
            }, function(err) {
                console.error('Could not copy text: ', err);
            });
        }
        
        function exportToCSV() {
            const tables = document.querySelectorAll('table');
            let csvContent = "Active Directory Infrastructure Report\\n";
            csvContent += "Generated: " + new Date().toLocaleString() + "\\n\\n";
            
            tables.forEach((table, index) => {
                const sectionContent = table.closest('.section-content');
                if (sectionContent) {
                    const sectionTitle = sectionContent.previousElementSibling.textContent;
                    csvContent += sectionTitle + "\\n";
                    
                    const rows = table.querySelectorAll('tr');
                    rows.forEach(row => {
                        const cells = row.querySelectorAll('th, td');
                        const rowData = Array.from(cells).map(cell => {
                            // Remove HTML tags and trim
                            let text = cell.textContent.replace(/<[^>]*>/g, '').trim();
                            // Escape quotes and wrap in quotes if contains comma
                            if (text.includes(',') || text.includes('"')) {
                                text = '"' + text.replace(/"/g, '""') + '"';
                            }
                            return text;
                        }).join(',');
                        csvContent += rowData + "\\n";
                    });
                    csvContent += "\\n";
                }
            });
            
            const blob = new Blob([csvContent], { type: 'text/csv' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'AD_Infrastructure_Report.csv';
            a.click();
            window.URL.revokeObjectURL(url);
        }
        
        window.onload = function() {
            document.querySelectorAll('h2').forEach(header => {
                header.addEventListener('click', function() { toggleSection(this); });
                
                // Add section header wrapper for better styling
                const headerText = header.innerHTML;
                header.innerHTML = '<div class="section-header"><span class="section-header-text">' + headerText + '</span></div>';
            });
            
            // Add search functionality
            const searchInput = document.getElementById('searchInput');
            searchInput.addEventListener('input', performSearch);
            searchInput.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    performSearch();
                }
            });
        };
    </script>
</head>
<body>
    <h1>Active Directory Infrastructure Documentation</h1>
    <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    <p class="timestamp">Output Location: $OutputPath</p>
    
    <div class="search-container">
        <input type="text" id="searchInput" class="search-box" placeholder="Search across all sections (minimum 2 characters)...">
        <div class="search-stats" id="searchStats">Enter search terms to find specific information</div>
    </div>
    
    <div style="display: flex; gap: 10px; margin-bottom: 20px;">
        <button class="toggle-all" onclick="toggleAll()">Expand/Collapse All Sections</button>
        <button class="toggle-all" onclick="clearSearch()" style="background-color: #718096;">Clear Search</button>
        <button class="toggle-all" onclick="exportToCSV()" style="background-color: #38a169;">Export to CSV</button>
    </div>
"@

# [REST OF THE SCRIPT SECTIONS REMAIN THE SAME AS BEFORE...]
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
    
    $IsGC = if ($DC.IsGlobalCatalog) { "<span class='healthy'>Yes</span>" } else { "No" }
    
    # Check DNS service
    $DNSStatus = "Unknown"
    try {
        $DNSService = Get-Service -ComputerName $DC.HostName -Name DNS -ErrorAction SilentlyContinue
        if ($DNSService) {
            $DNSStatus = if ($DNSService.Status -eq "Running") { "<span class='healthy'>Running</span>" } else { "<span class='critical'>$($DNSService.Status)</span>" }
        }
    } catch {
        $DNSStatus = "Unable to query"
    }
    
    $HTML += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($DC.Site)</td>
            <td>$($DC.IPv4Address)</td>
            <td>$($DC.OperatingSystem)</td>
            <td>$IsGC</td>
            <td>$FSMORoles</td>
            <td>$DNSStatus</td>
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
            $Status = if ($TimeSince.TotalHours -lt 24) { "<span class='healthy'>Healthy</span>" } else { "<span class='critical'>Warning</span>" }
            $Failures = if ($Partner.ConsecutiveReplicationFailures -eq 0) { "<span class='healthy'>0</span>" } else { "<span class='critical'>$($Partner.ConsecutiveReplicationFailures)</span>" }
            
            $HTML += @"
        <tr>
            <td>$($Partner.Partner)</td>
            <td>$($DC.HostName)</td>
            <td>$($LastRepl.ToString("yyyy-MM-dd HH:mm:ss"))</td>
            <td>$Status</td>
            <td>$Failures</td>
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
    $SubnetList = if ($Subnets) { ($Subnets -join "<br>") } else { "None configured" }
    
    $SiteDCs = $DCs | Where-Object { $_.Site -eq $Site.Name } | Select-Object -ExpandProperty HostName
    $DCList = if ($SiteDCs) { ($SiteDCs -join "<br>") } else { "None" }
    
    $HTML += @"
        <tr>
            <td>$($Site.Name)</td>
            <td>$($Site.Description)</td>
            <td>$SubnetList</td>
            <td>$DCList</td>
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
                    $ScopeState = if ($Scope.State -eq "Active") { "<span class='healthy'>Active</span>" } else { "<span class='critical'>$($Scope.State)</span>" }
                    $HTML += @"
        <tr>
            <td>$($DHCPServer.DnsName)</td>
            <td>$($DHCPServer.IPAddress)</td>
            <td>$($Scope.Name)</td>
            <td>$($Scope.StartRange) - $($Scope.EndRange)</td>
            <td>$ScopeState</td>
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

# 7. PRIVILEGED ACCOUNTS - TIER 0 (Domain Admins, Enterprise Admins, Schema Admins)
Write-Host "Gathering Tier 0 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>7. Tier 0 Accounts (Highest Privilege - Domain/Enterprise/Schema Admins)</h2>
    <div class="section-content">
"@

$Tier0Groups = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

foreach ($GroupName in $Tier0Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            foreach ($Member in $Members) {
                try {
                    if ($Member.objectClass -eq "user") {
                        $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                        $EnabledStatus = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                        $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                        $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                    } else {
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
$HTML += "</table></div>"

# 8. TIER 1 ACCOUNTS (Server Admins, Backup Operators)
Write-Host "Gathering Tier 1 Privileged Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>8. Tier 1 Accounts (Server/Infrastructure Management)</h2>
    <div class="section-content">
"@

$Tier1Groups = @(
    "Server Operators",
    "Backup Operators",
    "Account Operators",
    "Print Operators"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

foreach ($GroupName in $Tier1Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
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
$HTML += "</table></div>"

# 9. TIER 2 ACCOUNTS (Helpdesk, User Management)
Write-Host "Gathering Tier 2 Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>9. Tier 2 Accounts (User/Workstation Management)</h2>
    <div class="section-content">
    <div class="info-box">
        <p><strong>Note:</strong> Tier 2 typically includes help desk and user support groups. Common groups are listed below. Customize the script to add organization-specific Tier 2 groups.</p>
    </div>
"@

$Tier2Groups = @(
    "Help Desk",
    "Helpdesk Operators",
    "Desktop Support",
    "Remote Desktop Users"
)

$HTML += "<table><tr><th>Group</th><th>Member Name</th><th>Account Type</th><th>Enabled</th><th>Last Logon</th><th>Password Last Set</th></tr>"

$Tier2Found = $false
foreach ($GroupName in $Tier2Groups) {
    try {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -ErrorAction SilentlyContinue
        if ($Group) {
            $Tier2Found = $true
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            if ($Members) {
                foreach ($Member in $Members) {
                    try {
                        if ($Member.objectClass -eq "user") {
                            $User = Get-ADUser -Identity $Member.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet -ErrorAction SilentlyContinue
                            $EnabledStatus = if ($User.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
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

if (-not $Tier2Found) {
    $HTML += "<tr class='tier2'><td colspan='6'>No standard Tier 2 groups found. Please customize the script with your organization's Tier 2 group names.</td></tr>"
}

$HTML += "</table></div>"

# 10. SERVICE ACCOUNTS
Write-Host "Gathering Service Accounts..." -ForegroundColor Cyan
$HTML += @"
    <h2>10. Service Accounts</h2>
    <div class="section-content">
"@

# Service accounts are typically identified by naming convention or specific attributes
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
    $EnabledStatus = if ($SvcAcct.Enabled) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
    $PwdLastSet = if ($SvcAcct.PasswordLastSet) { $SvcAcct.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
    $LastLogon = if ($SvcAcct.LastLogonDate) { $SvcAcct.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
    $SPNCount = if ($SvcAcct.ServicePrincipalName) { $SvcAcct.ServicePrincipalName.Count } else { 0 }
    
    $HTML += @"
        <tr>
            <td>$($SvcAcct.SamAccountName)</td>
            <td>$($SvcAcct.Description)</td>
            <td>$EnabledStatus</td>
            <td>$PwdLastSet</td>
            <td>$LastLogon</td>
            <td>$SPNCount</td>
        </tr>
"@
}

$HTML += "</table></div>"

# 11. EXCHANGE SERVERS (Without Exchange Management Shell)
Write-Host "Gathering Exchange Server Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>11. Exchange Servers</h2>
    <div class="section-content">
"@

try {
    # Query Exchange servers from AD Configuration partition
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
            # Decode Exchange roles
            $Roles = switch ($ExchServer.msExchCurrentServerRoles) {
                2 { "Mailbox" }
                4 { "Client Access" }
                16 { "Unified Messaging" }
                32 { "Hub Transport" }
                64 { "Edge Transport" }
                54 { "Mailbox, Client Access, Hub Transport" }
                default { $ExchServer.msExchCurrentServerRoles }
            }
            
            # Get site name
            $SiteName = if ($ExchServer.msExchServerSite) {
                ($ExchServer.msExchServerSite -split ",")[0] -replace "CN=", ""
            } else {
                "Unknown"
            }
            
            # Get FQDN from network address
            $FQDN = "N/A"
            if ($ExchServer.networkAddress) {
                $FQDN = ($ExchServer.networkAddress | Where-Object { $_ -like "ncacn_ip_tcp:*" }) -replace "ncacn_ip_tcp:", ""
            }
            
            # Decode version
            $Version = "Unknown"
            if ($ExchServer.serialNumber) {
                $VersionNumber = $ExchServer.serialNumber
                if ($VersionNumber -like "Version 15.2*") { $Version = "Exchange 2019" }
                elseif ($VersionNumber -like "Version 15.1*") { $Version = "Exchange 2016" }
                elseif ($VersionNumber -like "Version 15.0*") { $Version = "Exchange 2013" }
                elseif ($VersionNumber -like "Version 14.*") { $Version = "Exchange 2010" }
                else { $Version = $VersionNumber }
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
        
        # Add summary information
        $HTML += @"
        <div class="info-box">
            <h3>Exchange Server Summary</h3>
            <p><strong>Total Exchange Servers:</strong> $($ExchangeServers.Count)</p>
            <p><strong>Versions Found:</strong> $(($ExchangeServers | Group-Object {$_.serialNumber -replace 'Version (\d+\.\d+).*','$1'} | ForEach-Object {$_.Name}) -join ', ')</p>
        </div>
"@
    } else {
        $HTML += @"
        <div class="info-box warning">
            <p>No Exchange servers found in Active Directory.</p>
            <p>This may be because:</p>
            <ul>
                <li>No Exchange servers are installed in the environment</li>
                <li>Exchange servers are not properly registered in Active Directory</li>
                <li>You lack permissions to query the Configuration partition</li>
            </ul>
        </div>
"@
    }
} catch {
    $HTML += @"
    <div class="info-box error">
        <p><span class='critical'>Error querying Exchange servers:</span> $($_.Exception.Message)</p>
        <p>This typically indicates permission issues or problems accessing the Configuration partition.</p>
    </div>
"@
}

$HTML += "</div>"

# 12. SUMMARY AND RECOMMENDATIONS
Write-Host "Generating Summary and Recommendations..." -ForegroundColor Cyan
$HTML += @"
    <h2>12. Summary and Recommendations</h2>
    <div class="section-content">
    <div class="info-box">
        <h3>Environment Summary</h3>
        <ul>
            <li><strong>Forest:</strong> $($Forest.Name)</li>
            <li><strong>Domain Controllers:</strong> $($DCs.Count)</li>
            <li><strong>Sites:</strong> $($Sites.Count)</li>
            <li><strong>Forest Functional Level:</strong> $($Forest.ForestMode)</li>
            <li><strong>Report Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</li>
        </ul>
    </div>
    
    <div class="info-box warning">
        <h3>Security Recommendations</h3>
        <ul>
            <li>Regularly review Tier 0 group memberships</li>
            <li>Ensure service accounts have descriptive descriptions</li>
            <li>Monitor for stale accounts (no recent logons)</li>
            <li>Verify replication health across all domain controllers</li>
            <li>Review DNS and DHCP server configurations regularly</li>
        </ul>
    </div>
    
    <div class="info-box success">
        <h3>Maintenance Recommendations</h3>
        <ul>
            <li>Regularly test backup and recovery procedures</li>
            <li>Monitor domain controller disk space and performance</li>
            <li>Keep domain controllers updated with latest security patches</li>
            <li>Review and update site/subnet configurations as needed</li>
            <li>Document any custom configurations not captured in this report</li>
        </ul>
    </div>
    </div>
"@

# Close the HTML document
$HTML += @"
</body>
</html>
"@

# Write the HTML report to file
Write-Host "Writing report to: $ReportFile" -ForegroundColor Green
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Host "Active Directory Infrastructure Documentation completed successfully!" -ForegroundColor Green
Write-Host "Report saved to: $ReportFile" -ForegroundColor Yellow

# Open the report in default browser
try {
    Start-Process $ReportFile
    Write-Host "Opening report in default browser..." -ForegroundColor Green
} catch {
    Write-Host "Report generated but could not open automatically. Please open manually: $ReportFile" -ForegroundColor Yellow
}
