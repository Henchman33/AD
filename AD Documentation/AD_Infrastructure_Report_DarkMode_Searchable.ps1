# Active Directory Infrastructure Documentation Script - Enhanced Version
# Run this on a Domain Controller with appropriate permissions

# Set output directory to user's desktop
$OutputPath = [Environment]::GetFolderPath("Desktop") + "\AD_Documentation"
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.html"
$CSVPath = "$OutputPath\AD_Infrastructure_Data_$Timestamp"
$ExcelFile = "$OutputPath\AD_Infrastructure_Report_$Timestamp.xlsx"

# Import required module
Import-Module ActiveDirectory

Write-Host "Starting Active Directory Infrastructure Documentation..." -ForegroundColor Green
Write-Host "Output directory: $OutputPath" -ForegroundColor Yellow
Write-Host "GO GET A CUP OF COFFEE!!! LARGER SITES MAY TAKE LONGER TO RUN!!!" -ForegroundColor Cyan

# Initialize data collection arrays for CSV/Excel export
$AllData = @{
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
    OrganizationalUnits = @()
}

# Initialize HTML report with collapsible sections
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Active Directory Infrastructure Documentation</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #1e1e1e; color: #e0e0e0; }
        h1 { color: #4fc3f7; border-bottom: 3px solid #4fc3f7; padding-bottom: 10px; }
        h2 { color: #81c784; margin-top: 30px; border-bottom: 2px solid #81c784; padding-bottom: 5px; cursor: pointer; user-select: none; }
        h2:hover { background-color: #2d2d2d; }
        h2::before { content: '▼ '; font-size: 0.8em; }
        h2.collapsed::before { content: '▶ '; }
        h3 { color: #64b5f6; margin-top: 20px; }
        .section-content { margin-left: 20px; }
        .collapsed-content { display: none; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: #2d2d2d; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
        th { background-color: #37474f; color: #81c784; padding: 12px; text-align: left; font-weight: bold; border-bottom: 2px solid #4fc3f7; }
        td { padding: 10px; border-bottom: 1px solid #424242; color: #e0e0e0; }
        tr:hover { background-color: #363636; }
        .info-box { background-color: #2d2d2d; padding: 15px; margin: 15px 0; border-left: 4px solid #4fc3f7; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
        .warning { border-left-color: #ffb74d; }
        .success { border-left-color: #81c784; }
        .error { border-left-color: #e57373; }
        .timestamp { color: #9e9e9e; font-size: 0.9em; }
        .author { color: #b0bec5; font-size: 1.1em; font-weight: bold; margin-bottom: 5px; }
        .critical { color: #ef5350; font-weight: bold; }
        .healthy { color: #66bb6a; font-weight: bold; }
        .tier0 { background-color: #4a2c2c; }
        .tier1 { background-color: #4a4a2c; }
        .tier2 { background-color: #2c4a2c; }
        .toggle-all { margin: 20px 0; padding: 10px 20px; background-color: #455a64; color: #e0e0e0; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }
        .toggle-all:hover { background-color: #546e7a; }
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
    </script>
</head>
<body>
    <h1>Active Directory Infrastructure Documentation</h1>
    <p class="author">Author: Stephen McKee - IGTPLC Systems Admin 2</p>
    <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
    <p class="timestamp">Output Location: $OutputPath</p>
    <button class="toggle-all" onclick="toggleAll()">Expand/Collapse All Sections</button>
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
    
    $IsGC = if ($DC.IsGlobalCatalog) { "<span class='healthy'>Yes</span>" } else { "No" }
    $IsGCPlain = if ($DC.IsGlobalCatalog) { "Yes" } else { "No" }
    
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
    
    $DNSStatusPlain = $DNSStatus -replace "<span class='healthy'>|<span class='critical'>|</span>", ""
    
    # Add to data collection
    $AllData.DomainControllers += [PSCustomObject]@{
        Hostname = $DC.HostName
        Site = $DC.Site
        IPAddress = $DC.IPv4Address
        OSVersion = $DC.OperatingSystem
        GlobalCatalog = $IsGCPlain
        FSMORoles = $FSMORoles
        DNSService = $DNSStatusPlain
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
            $StatusPlain = if ($TimeSince.TotalHours -lt 24) { "Healthy" } else { "Warning" }
            $Failures = if ($Partner.ConsecutiveReplicationFailures -eq 0) { "<span class='healthy'>0</span>" } else { "<span class='critical'>$($Partner.ConsecutiveReplicationFailures)</span>" }
            $FailuresPlain = $Partner.ConsecutiveReplicationFailures
            
            # Add to data collection
            $AllData.ReplicationHealth += [PSCustomObject]@{
                SourceDC = $Partner.Partner
                DestinationDC = $DC.HostName
                LastReplication = $LastRepl.ToString("yyyy-MM-dd HH:mm:ss")
                Status = $StatusPlain
                Failures = $FailuresPlain
            }
            
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
    $SubnetListPlain = if ($Subnets) { ($Subnets -join "; ") } else { "None configured" }
    
    $SiteDCs = $DCs | Where-Object { $_.Site -eq $Site.Name } | Select-Object -ExpandProperty HostName
    $DCList = if ($SiteDCs) { ($SiteDCs -join "<br>") } else { "None" }
    $DCListPlain = if ($SiteDCs) { ($SiteDCs -join "; ") } else { "None" }
    
    # Add to data collection
    $AllData.Sites += [PSCustomObject]@{
        SiteName = $Site.Name
        Description = $Site.Description
        Subnets = $SubnetListPlain
        DomainControllers = $DCListPlain
    }
    
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
                # Add to data collection
                $AllData.DNSZones += [PSCustomObject]@{
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
                    $ScopeState = if ($Scope.State -eq "Active") { "<span class='healthy'>Active</span>" } else { "<span class='critical'>$($Scope.State)</span>" }
                    $ScopeStatePlain = $Scope.State
                    
                    # Add to data collection
                    $AllData.DHCPScopes += [PSCustomObject]@{
                        ServerName = $DHCPServer.DnsName
                        IPAddress = $DHCPServer.IPAddress
                        ScopeName = $Scope.Name
                        StartRange = $Scope.StartRange
                        EndRange = $Scope.EndRange
                        State = $ScopeStatePlain
                    }
                    
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
                        $EnabledPlain = if ($User.Enabled) { "Yes" } else { "No" }
                        $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                        $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                        
                        # Add to data collection
                        $AllData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $User.Name
                            AccountType = "User"
                            Enabled = $EnabledPlain
                            LastLogon = $LastLogon
                            PasswordLastSet = $PwdLastSet
                        }
                        
                        $HTML += "<tr class='tier0'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                    } else {
                        # Add to data collection
                        $AllData.Tier0Accounts += [PSCustomObject]@{
                            Group = $GroupName
                            MemberName = $Member.Name
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
                            $EnabledPlain = if ($User.Enabled) { "Yes" } else { "No" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            # Add to data collection
                            $AllData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                AccountType = "User"
                                Enabled = $EnabledPlain
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier1'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            # Add to data collection
                            $AllData.Tier1Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
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
                            $EnabledPlain = if ($User.Enabled) { "Yes" } else { "No" }
                            $LastLogon = if ($User.LastLogonDate) { $User.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                            $PwdLastSet = if ($User.PasswordLastSet) { $User.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
                            
                            # Add to data collection
                            $AllData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $User.Name
                                AccountType = "User"
                                Enabled = $EnabledPlain
                                LastLogon = $LastLogon
                                PasswordLastSet = $PwdLastSet
                            }
                            
                            $HTML += "<tr class='tier2'><td>$GroupName</td><td>$($User.Name)</td><td>User</td><td>$EnabledStatus</td><td>$LastLogon</td><td>$PwdLastSet</td></tr>"
                        } else {
                            # Add to data collection
                            $AllData.Tier2Accounts += [PSCustomObject]@{
                                Group = $GroupName
                                MemberName = $Member.Name
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
    $EnabledPlain = if ($SvcAcct.Enabled) { "Yes" } else { "No" }
    $PwdLastSet = if ($SvcAcct.PasswordLastSet) { $SvcAcct.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Never" }
    $LastLogon = if ($SvcAcct.LastLogonDate) { $SvcAcct.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
    $SPNCount = if ($SvcAcct.ServicePrincipalName) { $SvcAcct.ServicePrincipalName.Count } else { 0 }
    
    # Add to data collection
    $AllData.ServiceAccounts += [PSCustomObject]@{
        AccountName = $SvcAcct.SamAccountName
        Description = $SvcAcct.Description
        Enabled = $EnabledPlain
        PasswordLastSet = $PwdLastSet
        LastLogon = $LastLogon
        SPNCount = $SPNCount
    }
    
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
            
            # Add to data collection
            $AllData.ExchangeServers += [PSCustomObject]@{
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
        <div class="info-box warning">
            <p>No Exchange servers found in Active Directory.</p>
        </div>
"@
    }
} catch {
    $HTML += @"
    <div class="info-box error">
        <p><span class='critical'>Unable to query Exchange servers from Active Directory.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 12. GROUP POLICY OBJECTS
Write-Host "Gathering Group Policy Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>12. Group Policy Objects</h2>
    <div class="section-content">
"@

try {
    $GPOs = Get-GPO -All | Sort-Object DisplayName
    
    $HTML += @"
    <table>
        <tr>
            <th>GPO Name</th>
            <th>Status</th>
            <th>Created</th>
            <th>Modified</th>
            <th>Linked To</th>
        </tr>
"@
    
    foreach ($GPO in $GPOs) {
        $GpoStatus = if ($GPO.GpoStatus -eq "AllSettingsEnabled") { "<span class='healthy'>Enabled</span>" } else { $GPO.GpoStatus }
        $GpoStatusPlain = $GPO.GpoStatus
        
        # Get GPO links
        $Links = @()
        try {
            $Report = [xml](Get-GPOReport -Guid $GPO.Id -ReportType Xml)
            $LinksTo = $Report.GPO.LinksTo
            if ($LinksTo) {
                foreach ($Link in $LinksTo) {
                    $Links += $Link.SOMPath
                }
            }
        } catch {
            $Links += "Unable to query"
        }
        
        $LinksList = if ($Links.Count -gt 0) { ($Links -join "<br>") } else { "Not linked" }
        $LinksListPlain = if ($Links.Count -gt 0) { ($Links -join "; ") } else { "Not linked" }
        
        # Add to data collection
        $AllData.GroupPolicies += [PSCustomObject]@{
            GPOName = $GPO.DisplayName
            Status = $GpoStatusPlain
            Created = $GPO.CreationTime.ToString("yyyy-MM-dd")
            Modified = $GPO.ModificationTime.ToString("yyyy-MM-dd")
            LinkedTo = $LinksListPlain
        }
        
        $HTML += @"
        <tr>
            <td>$($GPO.DisplayName)</td>
            <td>$GpoStatus</td>
            <td>$($GPO.CreationTime.ToString("yyyy-MM-dd"))</td>
            <td>$($GPO.ModificationTime.ToString("yyyy-MM-dd"))</td>
            <td>$LinksList</td>
        </tr>
"@
    }
    
    $HTML += "</table>"
} catch {
    $HTML += @"
    <div class="info-box error">
        <p><span class='critical'>Unable to retrieve Group Policy Objects.</span></p>
        <p>Error: $($_.Exception.Message)</p>
    </div>
"@
}

$HTML += "</div>"

# 13. ORGANIZATIONAL UNITS
Write-Host "Gathering Organizational Unit Information..." -ForegroundColor Cyan
$HTML += @"
    <h2>13. Organizational Units</h2>
    <div class="section-content">
    <table>
        <tr>
            <th>OU Name</th>
            <th>Distinguished Name</th>
            <th>Description</th>
            <th>Protected from Deletion</th>
        </tr>
"@

$OUs = Get-ADOrganizationalUnit -Filter * -Properties Description, ProtectedFromAccidentalDeletion | Sort-Object DistinguishedName

foreach ($OU in $OUs) {
    $Protected = if ($OU.ProtectedFromAccidentalDeletion) { "<span class='healthy'>Yes</span>" } else { "<span class='critical'>No</span>" }
    $ProtectedPlain = if ($OU.ProtectedFromAccidentalDeletion) { "Yes" } else { "No" }
    
    # Add to data collection
    $AllData.OrganizationalUnits += [PSCustomObject]@{
        OUName = $OU.Name
        DistinguishedName = $OU.DistinguishedName
        Description = $OU.Description
        ProtectedFromDeletion = $ProtectedPlain
    }
    
    $HTML += @"
        <tr>
            <td>$($OU.Name)</td>
            <td>$($OU.DistinguishedName)</td>
            <td>$($OU.Description)</td>
            <td>$Protected</td>
        </tr>
"@
}

$HTML += "</table></div>"

# Close HTML
$HTML += @"
    <div style="margin-top: 50px; padding: 20px; background-color: #2d2d2d; border-left: 4px solid #4fc3f7;">
        <p><strong>Documentation Complete</strong></p>
        <p>Report generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        <p>All files saved to: $OutputPath</p>
    </div>
</body>
</html>
"@

# Save HTML report
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8

Write-Host "`nExporting data to CSV files..." -ForegroundColor Cyan

# Export each dataset to CSV
if ($AllData.DomainControllers.Count -gt 0) {
    $AllData.DomainControllers | Export-Csv -Path "$CSVPath`_DomainControllers.csv" -NoTypeInformation
    Write-Host "  - Domain Controllers exported" -ForegroundColor Green
}

if ($AllData.ReplicationHealth.Count -gt 0) {
    $AllData.ReplicationHealth | Export-Csv -Path "$CSVPath`_ReplicationHealth.csv" -NoTypeInformation
    Write-Host "  - Replication Health exported" -ForegroundColor Green
}

if ($AllData.Sites.Count -gt 0) {
    $AllData.Sites | Export-Csv -Path "$CSVPath`_Sites.csv" -NoTypeInformation
    Write-Host "  - Sites exported" -ForegroundColor Green
}

if ($AllData.DNSZones.Count -gt 0) {
    $AllData.DNSZones | Export-Csv -Path "$CSVPath`_DNSZones.csv" -NoTypeInformation
    Write-Host "  - DNS Zones exported" -ForegroundColor Green
}

if ($AllData.DHCPScopes.Count -gt 0) {
    $AllData.DHCPScopes | Export-Csv -Path "$CSVPath`_DHCPScopes.csv" -NoTypeInformation
    Write-Host "  - DHCP Scopes exported" -ForegroundColor Green
}

if ($AllData.Tier0Accounts.Count -gt 0) {
    $AllData.Tier0Accounts | Export-Csv -Path "$CSVPath`_Tier0Accounts.csv" -NoTypeInformation
    Write-Host "  - Tier 0 Accounts exported" -ForegroundColor Green
}

if ($AllData.Tier1Accounts.Count -gt 0) {
    $AllData.Tier1Accounts | Export-Csv -Path "$CSVPath`_Tier1Accounts.csv" -NoTypeInformation
    Write-Host "  - Tier 1 Accounts exported" -ForegroundColor Green
}

if ($AllData.Tier2Accounts.Count -gt 0) {
    $AllData.Tier2Accounts | Export-Csv -Path "$CSVPath`_Tier2Accounts.csv" -NoTypeInformation
    Write-Host "  - Tier 2 Accounts exported" -ForegroundColor Green
}

if ($AllData.ServiceAccounts.Count -gt 0) {
    $AllData.ServiceAccounts | Export-Csv -Path "$CSVPath`_ServiceAccounts.csv" -NoTypeInformation
    Write-Host "  - Service Accounts exported" -ForegroundColor Green
}

if ($AllData.ExchangeServers.Count -gt 0) {
    $AllData.ExchangeServers | Export-Csv -Path "$CSVPath`_ExchangeServers.csv" -NoTypeInformation
    Write-Host "  - Exchange Servers exported" -ForegroundColor Green
}

if ($AllData.GroupPolicies.Count -gt 0) {
    $AllData.GroupPolicies | Export-Csv -Path "$CSVPath`_GroupPolicies.csv" -NoTypeInformation
    Write-Host "  - Group Policies exported" -ForegroundColor Green
}

if ($AllData.OrganizationalUnits.Count -gt 0) {
    $AllData.OrganizationalUnits | Export-Csv -Path "$CSVPath`_OrganizationalUnits.csv" -NoTypeInformation
    Write-Host "  - Organizational Units exported" -ForegroundColor Green
}

Write-Host "`nExporting data to Excel..." -ForegroundColor Cyan

# Export to Excel using COM object
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Add()
    
    # Remove default sheets except one
    while ($Workbook.Worksheets.Count -gt 1) {
        $Workbook.Worksheets.Item($Workbook.Worksheets.Count).Delete()
    }
    
    $SheetIndex = 1
    
    # Function to add worksheet with data
    function Add-WorksheetWithData {
        param (
            [string]$SheetName,
            [array]$Data
        )
        
        if ($Data.Count -eq 0) { return }
        
        if ($SheetIndex -gt 1) {
            $Sheet = $Workbook.Worksheets.Add([System.Reflection.Missing]::Value, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
        } else {
            $Sheet = $Workbook.Worksheets.Item(1)
        }
        
        $Sheet.Name = $SheetName
        
        # Add headers
        $Properties = $Data[0].PSObject.Properties.Name
        $Col = 1
        foreach ($Prop in $Properties) {
            $Sheet.Cells.Item(1, $Col) = $Prop
            $Sheet.Cells.Item(1, $Col).Font.Bold = $true
            $Sheet.Cells.Item(1, $Col).Interior.Color = 15123099  # Light blue
            $Col++
        }
        
        # Add data
        $Row = 2
        foreach ($Item in $Data) {
            $Col = 1
            foreach ($Prop in $Properties) {
                $Sheet.Cells.Item($Row, $Col) = $Item.$Prop
                $Col++
            }
            $Row++
        }
        
        # Auto-fit columns
        $Sheet.UsedRange.Columns.AutoFit() | Out-Null
        
        $Script:SheetIndex++
    }
    
    # Add all sheets
    Add-WorksheetWithData -SheetName "DomainControllers" -Data $AllData.DomainControllers
    Add-WorksheetWithData -SheetName "ReplicationHealth" -Data $AllData.ReplicationHealth
    Add-WorksheetWithData -SheetName "Sites" -Data $AllData.Sites
    Add-WorksheetWithData -SheetName "DNSZones" -Data $AllData.DNSZones
    Add-WorksheetWithData -SheetName "DHCPScopes" -Data $AllData.DHCPScopes
    Add-WorksheetWithData -SheetName "Tier0Accounts" -Data $AllData.Tier0Accounts
    Add-WorksheetWithData -SheetName "Tier1Accounts" -Data $AllData.Tier1Accounts
    Add-WorksheetWithData -SheetName "Tier2Accounts" -Data $AllData.Tier2Accounts
    Add-WorksheetWithData -SheetName "ServiceAccounts" -Data $AllData.ServiceAccounts
    Add-WorksheetWithData -SheetName "ExchangeServers" -Data $AllData.ExchangeServers
    Add-WorksheetWithData -SheetName "GroupPolicies" -Data $AllData.GroupPolicies
    Add-WorksheetWithData -SheetName "OUs" -Data $AllData.OrganizationalUnits
    
    # Save and close
    $Workbook.SaveAs($ExcelFile)
    $Workbook.Close()
    $Excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "  - Excel file created successfully" -ForegroundColor Green
} catch {
    Write-Host "  - Warning: Could not create Excel file. Excel may not be installed." -ForegroundColor Yellow
    Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "Documentation complete!" -ForegroundColor Green
Write-Host "HTML Report saved to: $ReportFile" -ForegroundColor Yellow
Write-Host "CSV files saved to: $CSVPath`_*.csv" -ForegroundColor Yellow
if (Test-Path $ExcelFile) {
    Write-Host "Excel file saved to: $ExcelFile" -ForegroundColor Yellow
}
Write-Host "Opening report in default browser..." -ForegroundColor Cyan

# Open the report in default browser
Start-Process $ReportFile

Write-Host "Script execution completed successfully!" -ForegroundColor Green
