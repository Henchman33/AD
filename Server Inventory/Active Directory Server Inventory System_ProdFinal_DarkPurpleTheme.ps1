<#
.SYNOPSIS
    Active Directory Server Inventory Report Generator
.DESCRIPTION
    Generates comprehensive inventory reports of all servers in Active Directory domain
    Exports to CSV, Excel (XLS), and HTML formats with search functionality
.AUTHOR
    Stephen McKee - IGTPLC
.NOTES
    Run this script on a PDC Domain Controller with appropriate permissions
    Requires Active Directory PowerShell Module
#>

#Requires -Modules ActiveDirectory
#Requires -RunAsAdministrator

# Script Information
$Author = "Stephen McKee - IGTPLC"
$ReportTitle = "Active Directory Server Inventory Report"

# Generate timestamp for folder and files
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$ReportFolderName = "AD Server Report $TimeStamp"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ReportFolder = Join-Path -Path $DesktopPath -ChildPath $ReportFolderName

# Create report folder
Write-Host "Creating report folder: $ReportFolder" -ForegroundColor Cyan
New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null

# Define file paths
$CSVPath = Join-Path -Path $ReportFolder -ChildPath "ServerInventory_$TimeStamp.csv"
$XLSPath = Join-Path -Path $ReportFolder -ChildPath "ServerInventory_$TimeStamp.xls"
$HTMLPath = Join-Path -Path $ReportFolder -ChildPath "ServerInventory_$TimeStamp.html"
$LogPath = Join-Path -Path $ReportFolder -ChildPath "Report_Log_$TimeStamp.txt"

# Function to write log
function Write-Log {
    param([string]$Message)
    $LogTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$LogTime - $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    Write-Host $LogMessage -ForegroundColor Green
}

Write-Log "=========================================="
Write-Log "Active Directory Server Inventory Report"
Write-Log "Author: $Author"
Write-Log "=========================================="
Write-Log "Report generation started"

# Get domain information
$Domain = Get-ADDomain
$DomainName = $Domain.DNSRoot
Write-Log "Domain: $DomainName"

# Query all server computers from Active Directory
Write-Log "Querying Active Directory for server objects..."
$Servers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -Properties * | Sort-Object Name

Write-Log "Found $($Servers.Count) server(s) in Active Directory"

# Initialize array for server data
$ServerInventory = @()
$Counter = 0

Write-Log "This will take some time to run, go get a COFFEE!!!..."
Write-Log "Gathering detailed information for each server..."

foreach ($Server in $Servers) {
    $Counter++
    Write-Progress -Activity "Processing Servers" -Status "Processing $($Server.Name) ($Counter of $($Servers.Count))" -PercentComplete (($Counter / $Servers.Count) * 100)
    
    Write-Host "Processing: $($Server.Name)" -ForegroundColor Yellow
    
    # Get IPv4 Address
    $IPv4Address = "N/A"
    if ($Server.IPv4Address) {
        $IPv4Address = $Server.IPv4Address
    } else {
        # Try to resolve DNS
        try {
            $DNS = [System.Net.Dns]::GetHostEntry($Server.Name)
            $IPv4 = $DNS.AddressList | Where-Object {$_.AddressFamily -eq 'InterNetwork'} | Select-Object -First 1
            if ($IPv4) {
                $IPv4Address = $IPv4.IPAddressToString
            }
        } catch {
            $IPv4Address = "Unable to resolve"
        }
    }
    
    # Get Operating System
    $OSDistribution = if ($Server.OperatingSystem) { $Server.OperatingSystem } else { "N/A" }
    $OSType = $OSDistribution
    
    # Get Last Logged On User
    $LastLogonUser = "N/A"
    $LastLogonDate = "N/A"
    
    if ($Server.LastLogonDate) {
        $LastLogonDate = $Server.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    # Try to get last logged on user from various properties
    if ($Server.LastLogon) {
        try {
            $LastLogonDateTime = [DateTime]::FromFileTime($Server.LastLogon)
            if ($LastLogonDateTime.Year -gt 1601) {
                $LastLogonDate = $LastLogonDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            }
        } catch {}
    }
    
    # Get last logged on user (this requires additional query)
    try {
        $LastUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Server.Name -ErrorAction SilentlyContinue
        if ($LastUser.UserName) {
            $LastLogonUser = $LastUser.UserName
        }
    } catch {
        # If WMI fails, leave as N/A
    }
    
    # Get Description
    $Description = if ($Server.Description) { $Server.Description } else { "N/A" }
    
    # Get Server Roles (if available)
    $ServerRoles = "N/A"
    try {
        if (Test-Connection -ComputerName $Server.Name -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            $Roles = Get-WmiObject -Class Win32_ServerFeature -ComputerName $Server.Name -ErrorAction SilentlyContinue
            if ($Roles) {
                $RoleNames = ($Roles | Select-Object -ExpandProperty Name) -join "; "
                if ($RoleNames) {
                    $ServerRoles = $RoleNames
                }
            }
        }
    } catch {
        $ServerRoles = "Unable to query"
    }
    
    # Get Distinguished Name
    $DistinguishedName = if ($Server.DistinguishedName) { $Server.DistinguishedName } else { "N/A" }
    
    # Extract OU from Distinguished Name
    $OU = "N/A"
    if ($Server.DistinguishedName) {
        $DNParts = $Server.DistinguishedName -split ','
        $OUParts = $DNParts | Where-Object { $_ -like "OU=*" }
        if ($OUParts) {
            $OU = ($OUParts -join ',')
        } else {
            $OU = "Default Computers Container"
        }
    }
    
    # Create custom object with all information
    $ServerObject = [PSCustomObject]@{
        'Server Name' = $Server.Name
        'IPv4 Address' = $IPv4Address
        'OS Distribution' = $OSDistribution
        'OS Type' = $OSType
        'Domain' = $DomainName
        'Last Logged In User' = $LastLogonUser
        'Last Logon Date/Time' = $LastLogonDate
        'Description' = $Description
        'Server Roles' = $ServerRoles
        'Distinguished Name' = $DistinguishedName
        'Organizational Unit' = $OU
        'Enabled' = $Server.Enabled
        'Created' = if ($Server.Created) { $Server.Created.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
        'Modified' = if ($Server.Modified) { $Server.Modified.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
    }
    
    $ServerInventory += $ServerObject
}

Write-Progress -Activity "Processing Servers" -Completed

Write-Log "Data collection completed"

# Export to CSV
Write-Log "Exporting to CSV format..."
$ServerInventory | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV export completed: $CSVPath"

# Export to Excel (XLS) format - using CSV with .xls extension
Write-Log "Exporting to Excel format..."
$ServerInventory | Export-Csv -Path $XLSPath -NoTypeInformation -Encoding UTF8 -Delimiter "`t"
Write-Log "Excel export completed: $XLSPath"

# Generate HTML Report with search functionality
Write-Log "Generating HTML report..."

$HTMLHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportTitle</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 32px;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .header .meta {
            font-size: 14px;
            opacity: 0.9;
            margin-top: 10px;
        }
        
        .controls {
            padding: 20px 30px;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }
        
        .search-box {
            flex: 1;
            min-width: 250px;
        }
        
        .search-box input {
            width: 100%;
            padding: 12px 20px;
            border: 2px solid #dee2e6;
            border-radius: 5px;
            font-size: 14px;
            transition: all 0.3s;
        }
        
        .search-box input:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        .stats {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        
        .stat-box {
            padding: 10px 20px;
            background: white;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        
        .stat-box .label {
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
        }
        
        .stat-box .value {
            font-size: 24px;
            font-weight: bold;
            color: #667eea;
        }
        
        .table-container {
            padding: 30px;
            overflow-x: auto;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            font-size: 13px;
        }
        
        thead {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        th {
            padding: 15px 10px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
            cursor: pointer;
            user-select: none;
            position: relative;
        }
        
        th:hover {
            background: rgba(255,255,255,0.1);
        }
        
        th::after {
            content: '⇅';
            position: absolute;
            right: 10px;
            opacity: 0.5;
        }
        
        td {
            padding: 12px 10px;
            border-bottom: 1px solid #f0f0f0;
        }
        
        tbody tr {
            transition: background-color 0.2s;
        }
        
        tbody tr:hover {
            background: #f8f9fa;
        }
        
        tbody tr:nth-child(even) {
            background: #fafbfc;
        }
        
        tbody tr:nth-child(even):hover {
            background: #f0f2f5;
        }
        
        .no-results {
            text-align: center;
            padding: 40px;
            color: #999;
            font-size: 16px;
            display: none;
        }
        
        .footer {
            padding: 20px 30px;
            background: #f8f9fa;
            border-top: 1px solid #dee2e6;
            text-align: center;
            color: #666;
            font-size: 12px;
        }
        
        .enabled {
            color: #28a745;
            font-weight: bold;
        }
        
        .disabled {
            color: #dc3545;
            font-weight: bold;
        }
        
        @media (max-width: 768px) {
            .controls {
                flex-direction: column;
            }
            
            .search-box {
                width: 100%;
            }
            
            table {
                font-size: 11px;
            }
            
            th, td {
                padding: 8px 5px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>$ReportTitle</h1>
            <div class="meta">
                <strong>Author:</strong> $Author<br>
                <strong>Domain:</strong> $DomainName<br>
                <strong>Generated:</strong> $(Get-Date -Format "MMMM dd, yyyy - HH:mm:ss")<br>
                <strong>Server Count:</strong> $($ServerInventory.Count)
            </div>
        </div>
        
        <div class="controls">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="🔍 Search servers by name, IP, OS, description, or OU..." onkeyup="searchTable()">
            </div>
            <div class="stats">
                <div class="stat-box">
                    <div class="label">Total Servers</div>
                    <div class="value" id="totalCount">$($ServerInventory.Count)</div>
                </div>
                <div class="stat-box">
                    <div class="label">Visible</div>
                    <div class="value" id="visibleCount">$($ServerInventory.Count)</div>
                </div>
            </div>
        </div>
        
        <div class="table-container">
            <table id="serverTable">
                <thead>
                    <tr>
                        <th onclick="sortTable(0)">Server Name</th>
                        <th onclick="sortTable(1)">IPv4 Address</th>
                        <th onclick="sortTable(2)">OS Distribution</th>
                        <th onclick="sortTable(3)">OS Type</th>
                        <th onclick="sortTable(4)">Domain</th>
                        <th onclick="sortTable(5)">Last Logged In User</th>
                        <th onclick="sortTable(6)">Last Logon Date/Time</th>
                        <th onclick="sortTable(7)">Description</th>
                        <th onclick="sortTable(8)">Server Roles</th>
                        <th onclick="sortTable(9)">Distinguished Name</th>
                        <th onclick="sortTable(10)">Organizational Unit</th>
                        <th onclick="sortTable(11)">Enabled</th>
                    </tr>
                </thead>
                <tbody>
"@

$HTMLBody = ""
foreach ($Server in $ServerInventory) {
    $EnabledClass = if ($Server.Enabled -eq $true) { "enabled" } else { "disabled" }
    $EnabledText = if ($Server.Enabled -eq $true) { "Yes" } else { "No" }
    
    $HTMLBody += @"
                    <tr>
                        <td><strong>$($Server.'Server Name')</strong></td>
                        <td>$($Server.'IPv4 Address')</td>
                        <td>$($Server.'OS Distribution')</td>
                        <td>$($Server.'OS Type')</td>
                        <td>$($Server.'Domain')</td>
                        <td>$($Server.'Last Logged In User')</td>
                        <td>$($Server.'Last Logon Date/Time')</td>
                        <td>$($Server.'Description')</td>
                        <td>$($Server.'Server Roles')</td>
                        <td style="font-size: 11px;">$($Server.'Distinguished Name')</td>
                        <td>$($Server.'Organizational Unit')</td>
                        <td class="$EnabledClass">$EnabledText</td>
                    </tr>
"@
}

$HTMLFooter = @"
                </tbody>
            </table>
            <div class="no-results" id="noResults">
                No servers found matching your search criteria.
            </div>
        </div>
        
        <div class="footer">
            Report generated by PowerShell on $(Get-Date -Format "MMMM dd, yyyy") at $(Get-Date -Format "HH:mm:ss")<br>
            © $Author | Active Directory Server Inventory System
        </div>
    </div>
    
    <script>
        function searchTable() {
            const input = document.getElementById('searchInput');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('serverTable');
            const tbody = table.getElementsByTagName('tbody')[0];
            const rows = tbody.getElementsByTagName('tr');
            const noResults = document.getElementById('noResults');
            let visibleCount = 0;
            
            for (let i = 0; i < rows.length; i++) {
                const cells = rows[i].getElementsByTagName('td');
                let found = false;
                
                for (let j = 0; j < cells.length; j++) {
                    const cell = cells[j];
                    if (cell) {
                        const textValue = cell.textContent || cell.innerText;
                        if (textValue.toUpperCase().indexOf(filter) > -1) {
                            found = true;
                            break;
                        }
                    }
                }
                
                if (found) {
                    rows[i].style.display = '';
                    visibleCount++;
                } else {
                    rows[i].style.display = 'none';
                }
            }
            
            document.getElementById('visibleCount').textContent = visibleCount;
            
            if (visibleCount === 0) {
                noResults.style.display = 'block';
                table.style.display = 'none';
            } else {
                noResults.style.display = 'none';
                table.style.display = 'table';
            }
        }
        
        function sortTable(columnIndex) {
            const table = document.getElementById('serverTable');
            const tbody = table.getElementsByTagName('tbody')[0];
            const rows = Array.from(tbody.getElementsByTagName('tr'));
            const isAscending = tbody.dataset.sortOrder === 'asc';
            
            rows.sort((a, b) => {
                const aValue = a.getElementsByTagName('td')[columnIndex].textContent.trim();
                const bValue = b.getElementsByTagName('td')[columnIndex].textContent.trim();
                
                if (isAscending) {
                    return aValue.localeCompare(bValue, undefined, { numeric: true });
                } else {
                    return bValue.localeCompare(aValue, undefined, { numeric: true });
                }
            });
            
            tbody.dataset.sortOrder = isAscending ? 'desc' : 'asc';
            
            rows.forEach(row => tbody.appendChild(row));
        }
    </script>
</body>
</html>
"@

$HTMLContent = $HTMLHeader + $HTMLBody + $HTMLFooter
$HTMLContent | Out-File -FilePath $HTMLPath -Encoding UTF8

Write-Log "HTML report completed: $HTMLPath"

# Summary
Write-Log "=========================================="
Write-Log "Report Generation Summary"
Write-Log "=========================================="
Write-Log "Total Servers Processed: $($ServerInventory.Count)"
Write-Log "Report Folder: $ReportFolder"
Write-Log "Files Generated:"
Write-Log "  - CSV: ServerInventory_$TimeStamp.csv"
Write-Log "  - Excel: ServerInventory_$TimeStamp.xls"
Write-Log "  - HTML: ServerInventory_$TimeStamp.html"
Write-Log "  - Log: Report_Log_$TimeStamp.txt"
Write-Log "=========================================="
Write-Log "Report generation completed successfully!"
Write-Log "=========================================="

# Open the report folder
Write-Host "Opening report folder..." -ForegroundColor Green
Start-Process explorer.exe -ArgumentList $ReportFolder

# Display completion message
Write-Host "=========================================="
Write-Host "REPORT GENERATION COMPLETED" -ForegroundColor Green
Write-Host "=========================================="
Write-Host "Report Location: $ReportFolder" -ForegroundColor Yellow
Write-Host "Total Servers: $($ServerInventory.Count)" -ForegroundColor Cyan
Write-Host "Files Generated:" -ForegroundColor White
Write-Host "  ✓ CSV Export" -ForegroundColor Green
Write-Host "  ✓ Excel Export" -ForegroundColor Green
Write-Host "  ✓ HTML Report (with search)" -ForegroundColor Green
Write-Host "  ✓ Log File" -ForegroundColor Green
Write-Host "=========================================="
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
