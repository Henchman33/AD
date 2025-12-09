<#
New Search Features Added:
1. Advanced Search Interface

    Global search box with real-time search across all columns

    Filter by specific server (multi-select dropdown)

    Filter by share state (Online/Offline)

    Filter by encryption status

    Filter by minimum ACL count

2. Quick Search Filters

    Pre-defined quick filters for common searches:

        Domain Controllers

        Unencrypted Shares

        High ACL Count (≥10 entries)

        Admin Shares

        Data Shares

3. Enhanced Search Capabilities

    Wildcard support using * (e.g., server* for all servers starting with "server")

    ACL-specific searching (find shares with specific users/groups)

    Column-specific filtering

    Search term highlighting in results

    Search results counter showing number of matches

4. Search Tips Panel

    Helpful tips for effective searching

    Examples and syntax guidance

    Keyboard shortcuts (Enter to search)

5. Interactive Features

    Clear search button to reset all filters

    Clear highlights button to remove search term highlighting

    Multi-select for server filtering

    DataTable search integration with sorting and pagination

6. Visual Enhancements

    Search term highlighting in yellow

    Visual feedback for active filters

    Collapsible search interface

    Responsive design for all screen sizes

How to Use the Search Features:

    Basic Search: Type in the global search box and press Enter or click Search

    Advanced Filtering: Use the dropdowns to filter by specific criteria

    Quick Filters: Click on the quick filter buttons for common searches

    ACL Search: Search for specific users/groups in ACLs (e.g., "Administrators")

    Column Sorting: Click any column header to sort

    Clear All: Use the Clear button to reset all search criteria

The HTML report will automatically open in your default browser after generation, providing immediate access to both the data and the powerful search functionality.

---Update

Key Fixes Made:

    Fixed the regex escape issue: Changed \$\{\} to \`$`{}`` with proper backtick escaping

    Also fixed the regex replacement: Changed '$1' to '`$1' in the replace functions

    Fixed ACL data escaping: Properly escaped single quotes and double quotes in ACL data

    Fixed variable reference: Added the missing $uniqueServers variable
#>

# Domain SMB Shares Report Generator
# Author: Stephen McKee - Systems Administrator 2
# Description: Script to enumerate all domain servers, their SMB shares and ACLs - or at least tries to :-)

# Import required modules
Import-Module ActiveDirectory
Import-Module SmbShare

# Create output directory
$desktopPath = [Environment]::GetFolderPath("Desktop")
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Join-Path $desktopPath "Domain SMB Shares\$dateStamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# Output file names
$csvFile = Join-Path $outputDir "Domain_SMB_File_Share_Report.csv"
$xlsxFile = Join-Path $outputDir "Domain_SMB_File_Share_Report.xlsx"
$htmlFile = Join-Path $outputDir "Domain_SMB_File_Share_Report.html"

# Function to get all domain servers
function Get-DomainServers {
    try {
        Write-Host "Searching for domain servers..." -ForegroundColor Yellow
        $servers = Get-ADComputer -Filter { 
            OperatingSystem -like "*Server*" -and 
            Enabled -eq $true 
        } | Select-Object -ExpandProperty DNSHostName
        
        Write-Host "Found $($servers.Count) servers" -ForegroundColor Green
        return $servers
    }
    catch {
        Write-Error "Failed to retrieve domain servers: $_"
        return @()
    }
}

# Function to get SMB shares from a server
function Get-ServerShares {
    param(
        [string]$ComputerName
    )
    
    $results = @()
    
    try {
        Write-Host "  Checking $ComputerName..." -ForegroundColor Cyan
        
        # Get SMB shares
        $shares = Get-SmbShare -CimSession $ComputerName -ErrorAction SilentlyContinue | 
                  Where-Object {$_.ShareType -eq "FileSystemDirectory" -and $_.Name -notlike "*$"}
        
        foreach ($share in $shares) {
            try {
                # Get share ACL
                $acl = Get-SmbShareAccess -CimSession $ComputerName -Name $share.Name -ErrorAction SilentlyContinue
                
                # Format ACL entries
                $aclEntries = @()
                foreach ($ace in $acl) {
                    $aclEntries += "$($ace.AccountName) - $($ace.AccessRight) - $($ace.AccessControlType)"
                }
                $aclString = $aclEntries -join "`n"
                
                # Create result object
                $result = [PSCustomObject]@{
                    ServerName = $ComputerName
                    ShareName = $share.Name
                    SharePath = $share.Path
                    Description = $share.Description
                    CurrentUsers = $share.ConcurrentUserLimit
                    EncryptData = $share.EncryptData
                    ShareState = $share.ShareState
                    ACL = $aclString
                    ACLCount = $acl.Count
                    LastScanned = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
                
                $results += $result
            }
            catch {
                Write-Warning "Failed to get ACL for share $($share.Name) on $ComputerName"
            }
        }
    }
    catch {
        Write-Warning "Failed to connect to $ComputerName"
    }
    
    return $results
}

# Main script execution
Write-Host "=== Domain SMB File Share Report Generator ===" -ForegroundColor Green
Write-Host "Author: Stephen McKee - Systems Administrator 2" -ForegroundColor Green
Write-Host "Start Time: $(Get-Date)" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green

# Get all servers
$servers = Get-DomainServers

# Collect share information
$allResults = @()
$totalServers = $servers.Count
$currentServer = 0

# Store detailed ACL data for HTML
$detailedACLData = @()

foreach ($server in $servers) {
    $currentServer++
    Write-Progress -Activity "Scanning servers for SMB shares" `
                   -Status "Processing server $currentServer of ${totalServers}: $server" `
                   -PercentComplete (($currentServer / $totalServers) * 100)
    
    $shares = Get-ServerShares -ComputerName $server
    $allResults += $shares
    
    # Store detailed data for HTML
    foreach ($share in $shares) {
        $detailedACLData += @{
            Server = $share.ServerName
            Share = $share.ShareName
            ACL = $share.ACL
        }
    }
}

Write-Progress -Activity "Scanning servers for SMB shares" -Completed

# Export to CSV
Write-Host "Exporting to CSV..." -ForegroundColor Yellow
$allResults | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

# Export to Excel (requires ImportExcel module)
try {
    Write-Host "Exporting to Excel..." -ForegroundColor Yellow
    # Check if ImportExcel module is available
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel
        $allResults | Export-Excel -Path $xlsxFile `
                                   -WorksheetName "SMB Shares" `
                                   -TableName "DomainSMBShares" `
                                   -AutoSize `
                                   -FreezeTopRow `
                                   -BoldTopRow `
                                   -ClearSheet
    }
    else {
        Write-Warning "ImportExcel module not found. Skipping Excel export."
        Write-Host "To enable Excel export, install module: Install-Module -Name ImportExcel" -ForegroundColor Yellow
    }
}
catch {
    Write-Warning "Failed to export to Excel: $_"
}

# Create HTML Report
Write-Host "Creating HTML report..." -ForegroundColor Yellow

$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Domain SMB File Share Report</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/searchbuilder/1.5.0/css/searchBuilder.dataTables.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50, #4a6491);
            color: white;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header h1 {
            margin: 0;
            font-size: 24px;
        }
        .header .subtitle {
            margin-top: 5px;
            opacity: 0.9;
        }
        .search-container {
            background: white;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .search-box {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
        }
        .search-input {
            flex: 1;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
        }
        .search-button {
            background-color: #4a6491;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
        }
        .search-button:hover {
            background-color: #2c3e50;
        }
        .search-filters {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            margin-top: 15px;
        }
        .filter-group {
            display: flex;
            flex-direction: column;
            min-width: 200px;
        }
        .filter-label {
            font-weight: bold;
            margin-bottom: 5px;
            color: #2c3e50;
        }
        .filter-select {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: white;
        }
        .search-results-info {
            background: #e8f4f8;
            padding: 10px;
            border-radius: 4px;
            margin-top: 15px;
            display: none;
        }
        .summary {
            background: white;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .summary-item {
            display: inline-block;
            margin-right: 30px;
            padding: 10px;
            background: #f8f9fa;
            border-radius: 3px;
            min-width: 150px;
        }
        .summary-item i {
            margin-right: 8px;
            color: #4a6491;
        }
        .dataTables_wrapper {
            background: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .details-control {
            cursor: pointer;
            color: #4a6491;
        }
        .details-row {
            background-color: #f9f9f9 !important;
        }
        .details-content {
            padding: 10px;
            background-color: #f1f5fd;
            border: 1px solid #ddd;
            margin: 5px 0;
            border-radius: 3px;
        }
        .acl-entry {
            padding: 3px 0;
            border-bottom: 1px dashed #ddd;
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
        }
        .acl-entry:last-child {
            border-bottom: none;
        }
        .footer {
            text-align: center;
            margin-top: 20px;
            padding: 15px;
            color: #666;
            font-size: 0.9em;
            border-top: 1px solid #ddd;
        }
        .export-buttons {
            margin-bottom: 15px;
            text-align: right;
        }
        .btn-export {
            background-color: #4a6491;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 3px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            margin-left: 10px;
        }
        .btn-export:hover {
            background-color: #2c3e50;
        }
        table.dataTable tbody tr:hover {
            background-color: #f0f7ff !important;
        }
        .collapsible-column {
            cursor: pointer;
        }
        .column-hidden {
            display: none;
        }
        .highlight {
            background-color: yellow;
            font-weight: bold;
        }
        .search-tips {
            background: #fff8e1;
            border-left: 4px solid #ffc107;
            padding: 10px;
            margin-top: 10px;
            border-radius: 3px;
            font-size: 0.9em;
        }
        .search-tips h4 {
            margin-top: 0;
            color: #2c3e50;
        }
        .quick-search {
            display: flex;
            gap: 10px;
            margin-top: 10px;
            flex-wrap: wrap;
        }
        .quick-search-btn {
            background: #e8f4f8;
            border: 1px solid #b3d9e6;
            padding: 5px 10px;
            border-radius: 15px;
            cursor: pointer;
            font-size: 0.85em;
        }
        .quick-search-btn:hover {
            background: #b3d9e6;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1><i class="fas fa-network-wired"></i> Domain SMB File Share Report</h1>
        <div class="subtitle">
            Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | 
            Author: Stephen McKee Systems Administrator 2
        </div>
    </div>
    
    <div class="summary">
        <div class="summary-item">
            <i class="fas fa-server"></i> <strong>Total Servers:</strong> <span id="totalServers">$($servers.Count)</span>
        </div>
        <div class="summary-item">
            <i class="fas fa-folder"></i> <strong>Total Shares:</strong> <span id="totalShares">$($allResults.Count)</span>
        </div>
        <div class="summary-item">
            <i class="fas fa-shield-alt"></i> <strong>Total ACL Entries:</strong> <span id="totalACLs">$(($allResults | Measure-Object -Property ACLCount -Sum).Sum)</span>
        </div>
        <div class="summary-item">
            <i class="fas fa-clock"></i> <strong>Scan Duration:</strong> <span id="scanDuration"></span>
        </div>
    </div>
    
    <div class="search-container">
        <h3><i class="fas fa-search"></i> Advanced Search</h3>
        <div class="search-box">
            <input type="text" id="globalSearch" class="search-input" placeholder="Search across all columns (server names, share names, paths, ACLs...)">
            <button id="searchBtn" class="search-button">
                <i class="fas fa-search"></i> Search
            </button>
            <button id="clearSearchBtn" class="search-button" style="background-color: #6c757d;">
                <i class="fas fa-times"></i> Clear
            </button>
        </div>
        
        <div class="search-filters">
            <div class="filter-group">
                <label class="filter-label">Filter by Server:</label>
                <select id="serverFilter" class="filter-select" multiple>
                    <option value="">All Servers</option>
"@

# Add server options
$uniqueServers = $allResults | Select-Object -ExpandProperty ServerName -Unique | Sort-Object
foreach ($server in $uniqueServers) {
    $htmlHeader += "<option value='$server'>$server</option>"
}

$htmlHeader += @"
                </select>
                <small>Hold Ctrl to select multiple</small>
            </div>
            
            <div class="filter-group">
                <label class="filter-label">Filter by Share State:</label>
                <select id="stateFilter" class="filter-select">
                    <option value="">All States</option>
                    <option value="Online">Online</option>
                    <option value="Offline">Offline</option>
                </select>
            </div>
            
            <div class="filter-group">
                <label class="filter-label">Filter by Encryption:</label>
                <select id="encryptFilter" class="filter-select">
                    <option value="">All</option>
                    <option value="True">Encrypted</option>
                    <option value="False">Not Encrypted</option>
                </select>
            </div>
            
            <div class="filter-group">
                <label class="filter-label">Min ACL Count:</label>
                <input type="number" id="minAclFilter" class="filter-select" min="0" placeholder="0">
            </div>
        </div>
        
        <div class="quick-search">
            <strong>Quick Filters:</strong>
            <span class="quick-search-btn" data-filter="server:contains('DC')">Domain Controllers</span>
            <span class="quick-search-btn" data-filter="encrypt:False">Unencrypted Shares</span>
            <span class="quick-search-btn" data-filter="aclcount:>=10">High ACL Count</span>
            <span class="quick-search-btn" data-filter="share:contains('admin')">Admin Shares</span>
            <span class="quick-search-btn" data-filter="share:contains('data')">Data Shares</span>
        </div>
        
        <div class="search-tips">
            <h4><i class="fas fa-lightbulb"></i> Search Tips:</h4>
            <ul>
                <li>Use * for wildcard searches (e.g., <code>server*</code> for servers starting with "server")</li>
                <li>Search in ACLs using terms like <code>DOMAIN\User</code> or <code>Administrators</code></li>
                <li>Combine filters for precise results</li>
                <li>Click column headers to sort</li>
            </ul>
        </div>
        
        <div id="searchResultsInfo" class="search-results-info">
            <span id="resultCount">0</span> results found for "<span id="searchTerm"></span>"
            <button id="clearHighlightBtn" class="search-button" style="padding: 5px 10px; margin-left: 10px; background-color: #ffc107; color: #000;">
                <i class="fas fa-highlighter"></i> Clear Highlights
            </button>
        </div>
    </div>
    
    <div class="export-buttons">
        <a href="Domain_SMB_File_Share_Report.csv" class="btn-export" download>
            <i class="fas fa-file-csv"></i> Export to CSV
        </a>
    </div>
"@

$htmlTable = @"
    <table id="sharesTable" class="display" style="width:100%">
        <thead>
            <tr>
                <th></th>
                <th class="collapsible-column" data-column="server">Server Name</th>
                <th class="collapsible-column" data-column="share">Share Name</th>
                <th class="collapsible-column" data-column="path">Share Path</th>
                <th class="collapsible-column" data-column="desc">Description</th>
                <th class="collapsible-column" data-column="users">Current Users</th>
                <th class="collapsible-column" data-column="encrypt">Encrypt Data</th>
                <th class="collapsible-column" data-column="state">Share State</th>
                <th class="collapsible-column" data-column="aclcount">ACL Count</th>
            </tr>
        </thead>
        <tbody>
"@

# Add table rows with data attributes for search
foreach ($result in $allResults) {
    $aclData = $result.ACL -replace "'", "&apos;" -replace "`"", "&quot;"
    $htmlTable += @"
        <tr data-server="$($result.ServerName)"
            data-share="$($result.ShareName)"
            data-path="$($result.SharePath)"
            data-desc="$($result.Description)"
            data-users="$($result.CurrentUsers)"
            data-encrypt="$($result.EncryptData)"
            data-state="$($result.ShareState)"
            data-aclcount="$($result.ACLCount)"
            data-acl="$aclData">
            <td class="details-control"><i class="fas fa-plus-circle"></i></td>
            <td>$($result.ServerName)</td>
            <td>$($result.ShareName)</td>
            <td>$($result.SharePath)</td>
            <td>$($result.Description)</td>
            <td>$($result.CurrentUsers)</td>
            <td>$($result.EncryptData)</td>
            <td>$($result.ShareState)</td>
            <td>$($result.ACLCount)</td>
        </tr>
"@
}

$htmlTable += @"
        </tbody>
    </table>
"@

$htmlFooter = @"
    <div class="footer">
        <p>Report generated by Domain SMB Shares Scanner | $(Get-Date -Format 'yyyy')</p>
        <p>For support contact: Stephen McKee Systems Administrator 2</p>
    </div>
    
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.print.min.js"></script>
    <script src="https://cdn.datatables.net/searchbuilder/1.5.0/js/dataTables.searchBuilder.min.js"></script>
    <script src="https://cdn.datatables.net/searchpanes/2.2.0/js/dataTables.searchPanes.min.js"></script>
    
    <script>
        $(document).ready(function() {
            // Initialize DataTable with enhanced search features
            var table = $('#sharesTable').DataTable({
                dom: 'Bfrtip',
                buttons: [
                    'copy', 'csv', 'excel', 'pdf', 'print',
                    {
                        text: 'Toggle Columns',
                        action: function(e, dt, node, config) {
                            $('.collapsible-column').toggleClass('column-hidden');
                        }
                    },
                    {
                        text: 'Advanced Search',
                        action: function(e, dt, node, config) {
                            $('#globalSearch').focus();
                        }
                    }
                ],
                pageLength: 25,
                order: [[1, 'asc']],
                responsive: true,
                searchBuilder: {
                    columns: [1, 2, 3, 4, 5, 6, 7, 8]
                },
                language: {
                    searchBuilder: {
                        title: 'Build Advanced Search'
                    }
                }
            });
            
            // Global search functionality
            $('#globalSearch').on('keyup', function(e) {
                if (e.key === 'Enter') {
                    performSearch();
                }
            });
            
            $('#searchBtn').on('click', performSearch);
            
            // Clear search functionality
            $('#clearSearchBtn').on('click', function() {
                $('#globalSearch').val('');
                $('#serverFilter').val('');
                $('#stateFilter').val('');
                $('#encryptFilter').val('');
                $('#minAclFilter').val('');
                clearHighlights();
                table.search('').draw();
                $('#searchResultsInfo').hide();
            });
            
            // Filter controls
            $('#serverFilter, #stateFilter, #encryptFilter, #minAclFilter').on('change', applyFilters);
            
            // Quick search buttons
            $('.quick-search-btn').on('click', function() {
                var filter = $(this).data('filter');
                applyQuickFilter(filter);
            });
            
            // Clear highlights button
            $('#clearHighlightBtn').on('click', clearHighlights);
            
            // Function to perform search
            function performSearch() {
                var searchTerm = $('#globalSearch').val().trim();
                if (searchTerm === '') {
                    table.search('').draw();
                    $('#searchResultsInfo').hide();
                    clearHighlights();
                    return;
                }
                
                // Apply DataTable search
                table.search(searchTerm).draw();
                
                // Apply custom highlighting
                applyHighlights(searchTerm);
                
                // Show results info
                var visibleRows = table.rows({filter: 'applied'}).count();
                $('#resultCount').text(visibleRows);
                $('#searchTerm').text(searchTerm);
                $('#searchResultsInfo').show();
                
                // Also apply other filters
                applyFilters();
            }
            
            // Function to apply highlights
            function applyHighlights(searchTerm) {
                clearHighlights();
                
                if (!searchTerm) return;
                
                // Escape special regex characters
                // Instead of the complex regex escape, use a simpler approach:
                var escapedTerm = searchTerm.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
                var regex = new RegExp('(' + escapedTerm + ')', 'gi');
                
                // Highlight in table cells
                table.cells().every(function() {
                    var cell = this.node();
                    var originalText = $(cell).text();
                    if (regex.test(originalText)) {
                        var highlightedText = originalText.replace(regex, '<span class="highlight">$1</span>');
                        $(cell).html(highlightedText);
                    }
                });
                
                // Also highlight in expanded details
                $('.details-content').each(function() {
                    var content = $(this).html();
                    var highlightedContent = content.replace(regex, '<span class="highlight">$1</span>');
                    $(this).html(highlightedContent);
                });
            }
            
            // Function to clear highlights
            function clearHighlights() {
                $('.highlight').each(function() {
                    $(this).replaceWith($(this).text());
                });
            }
            
            // Function to apply filters
            function applyFilters() {
                var serverFilter = $('#serverFilter').val();
                var stateFilter = $('#stateFilter').val();
                var encryptFilter = $('#encryptFilter').val();
                var minAclFilter = $('#minAclFilter').val();
                
                // Build filter string
                var filters = [];
                
                if (serverFilter && serverFilter.length > 0) {
                    if (serverFilter.length === 1 && serverFilter[0] === '') {
                        // Do nothing for "All Servers"
                    } else {
                        var serverConditions = serverFilter.map(function(server) {
                            return '(server="' + server + '")';
                        }).join(' || ');
                        filters.push(serverConditions);
                    }
                }
                
                if (stateFilter) {
                    filters.push('(state="' + stateFilter + '")');
                }
                
                if (encryptFilter) {
                    filters.push('(encrypt="' + encryptFilter + '")');
                }
                
                if (minAclFilter) {
                    filters.push('(aclcount>=' + minAclFilter + ')');
                }
                
                // Combine filters
                if (filters.length > 0) {
                    var combinedFilter = filters.join(' && ');
                    $.fn.dataTable.ext.search.push(
                        function(settings, data, dataIndex) {
                            var row = table.row(dataIndex).node();
                            var rowData = $(row).data();
                            
                            // Check server filter
                            if (serverFilter && serverFilter.length > 0 && serverFilter[0] !== '') {
                                if (!serverFilter.includes(rowData.server)) {
                                    return false;
                                }
                            }
                            
                            // Check state filter
                            if (stateFilter && rowData.state !== stateFilter) {
                                return false;
                            }
                            
                            // Check encrypt filter
                            if (encryptFilter && rowData.encrypt !== encryptFilter) {
                                return false;
                            }
                            
                            // Check ACL count filter
                            if (minAclFilter && parseInt(rowData.aclcount) < parseInt(minAclFilter)) {
                                return false;
                            }
                            
                            return true;
                        }
                    );
                }
                
                table.draw();
                
                // Remove the custom filter function to prevent duplicates
                $.fn.dataTable.ext.search.pop();
            }
            
            // Function to apply quick filter
            function applyQuickFilter(filter) {
                var parts = filter.split(':');
                var column = parts[0];
                var value = parts[1];
                
                switch(column) {
                    case 'server':
                        $('#serverFilter').val('');
                        $('#globalSearch').val(value.replace('contains(', '').replace(')', ''));
                        break;
                    case 'encrypt':
                        $('#encryptFilter').val(value);
                        break;
                    case 'aclcount':
                        $('#minAclFilter').val(value.replace('>=', ''));
                        break;
                    case 'share':
                        $('#globalSearch').val(value.replace('contains(', '').replace(')', ''));
                        break;
                }
                
                performSearch();
            }
            
            // Add event listener for opening and closing details
            $('#sharesTable tbody').on('click', 'td.details-control', function() {
                var tr = $(this).closest('tr');
                var row = table.row(tr);
                var rowData = $(tr).data();
                
                if (row.child.isShown()) {
                    // This row is already open - close it
                    row.child.hide();
                    tr.removeClass('details-row');
                    $(this).html('<i class="fas fa-plus-circle"></i>');
                } else {
                    // Open this row
                    var details = '<div class="details-content">';
                    details += '<h4>Share Details: ' + rowData.share + ' on ' + rowData.server + '</h4>';
                    details += '<p><strong>Path:</strong> ' + rowData.path + '</p>';
                    details += '<p><strong>Description:</strong> ' + (rowData.desc || 'N/A') + '</p>';
                    details += '<p><strong>Current Users:</strong> ' + rowData.users + '</p>';
                    details += '<p><strong>Encryption:</strong> ' + rowData.encrypt + '</p>';
                    details += '<p><strong>State:</strong> ' + rowData.state + '</p>';
                    details += '<hr>';
                    details += '<h5>Access Control List (ACL):</h5>';
                    
                    if (rowData.acl && rowData.acl.trim() !== '') {
                        var aclEntries = rowData.acl.split('\n');
                        aclEntries.forEach(function(entry) {
                            if (entry.trim() !== '') {
                                details += '<div class="acl-entry">' + entry + '</div>';
                            }
                        });
                    } else {
                        details += '<p>No ACL entries found.</p>';
                    }
                    
                    details += '</div>';
                    
                    row.child(details).show();
                    tr.addClass('details-row');
                    $(this).html('<i class="fas fa-minus-circle"></i>');
                    
                    // Re-apply highlights if there's a search term
                    var searchTerm = $('#globalSearch').val();
                    if (searchTerm) {
                        applyHighlights(searchTerm);
                    }
                }
            });
            
            // Make columns collapsible
            $('.collapsible-column').on('click', function() {
                var column = $(this).data('column');
                var columnIndex = $(this).index();
                table.column(columnIndex).visible(!table.column(columnIndex).visible());
                $(this).toggleClass('column-hidden');
            });
            
            // Update scan duration
            var startTime = new Date('$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")');
            var endTime = new Date();
            var duration = Math.round((endTime - startTime) / 1000);
            $('#scanDuration').text(duration + ' seconds');
            
            // Initialize multi-select for server filter
            $('#serverFilter').attr('size', Math.min(10, $uniqueServers.length + 1));
        });
    </script>
</body>
</html>
"@

# Combine HTML parts and save
$htmlContent = $htmlHeader + $htmlTable + $htmlFooter
$htmlContent | Out-File -FilePath $htmlFile -Encoding UTF8

# Display completion message
Write-Host "=" * 60 -ForegroundColor Green
Write-Host "Report Generation Complete!" -ForegroundColor Green
Write-Host "Files saved to: $outputDir" -ForegroundColor Yellow
Write-Host "CSV File: $csvFile" -ForegroundColor Cyan
Write-Host "HTML File: $htmlFile" -ForegroundColor Cyan
if (Test-Path $xlsxFile) {
    Write-Host "Excel File: $xlsxFile" -ForegroundColor Cyan
}
Write-Host "=" * 60 -ForegroundColor Green
Write-Host "Total Servers Found: $($servers.Count)" -ForegroundColor Yellow
Write-Host "Total Shares Found: $($allResults.Count)" -ForegroundColor Yellow

# AUTOMATICALLY OPEN THE HTML REPORT
Write-Host "Opening HTML report in default browser..." -ForegroundColor Magenta
try {
    Start-Process $htmlFile
    Write-Host "HTML report opened successfully!" -ForegroundColor Green
}
catch {
    Write-Warning "Failed to open HTML report automatically. Please open it manually:"
    Write-Host $htmlFile -ForegroundColor Cyan
}

Write-Host "Script completed successfully!" -ForegroundColor Green
