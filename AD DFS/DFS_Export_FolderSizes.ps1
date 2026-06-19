<#
.SYNOPSIS
    DFS Namespace Audit Script with Folder Sizes and Owner Information
.DESCRIPTION
    Scans all DFS namespaces from the specified DFS server, exports folder targets,
    source servers, share paths, security groups, folder sizes, and owner to CSV, Excel, and HTML.
.NOTES
    Run from any server with DFS Management tools installed
    Requires: ActiveDirectory module, ImportExcel module (for XLSX)
.PARAMETER NamespaceServer
    The DFS namespace server to query (default: local computer)
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

# ===== CONFIGURATION =====
# Specify your DFS namespace server here
$NamespaceServer = "DFS01.xxx.com"  # Change this to your DFS server

# Set to $false to skip folder size calculation (saves time on large shares)
$IncludeFolderSize = $true
# =========================

# Create export folder on desktop
$exportBase = Join-Path $env:USERPROFILE "Desktop\DFS_Export"
if (-not (Test-Path $exportBase)) {
    New-Item -ItemType Directory -Path $exportBase -Force | Out-Null
}

# Timestamp for file naming
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $exportBase "DFS_Inventory_$timestamp.csv"
$xlsxPath = Join-Path $exportBase "DFS_Inventory_$timestamp.xlsx"
$htmlPath = Join-Path $exportBase "DFS_Inventory_$timestamp.html"

# Initialize results array
$results = @()

# Function to get NTFS security groups for a given folder path
function Get-FolderSecurityGroups {
    param([string]$FolderPath)
    
    $groups = @()
    try {
        if (Test-Path $FolderPath) {
            $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
            foreach ($access in $acl.Access) {
                if ($access.FileSystemRights -match "Read|Modify|FullControl|Write") {
                    $identity = $access.IdentityReference.Value
                    if ($identity -notmatch "BUILTIN\\|NT AUTHORITY\\") {
                        $groups += $identity
                    }
                }
            }
        } else {
            $groups = @("ERROR: Path not accessible - $FolderPath")
        }
    } catch {
        $groups = @("ERROR: Unable to read permissions - $_")
    }
    
    if ($groups.Count -eq 0) { $groups = @("No security groups found") }
    return ($groups -join "; ")
}

# Function to get folder owner
function Get-FolderOwner {
    param([string]$FolderPath)
    
    try {
        if (Test-Path $FolderPath) {
            $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
            $owner = $acl.Owner
            if ($owner) {
                return $owner
            } else {
                return "No owner found"
            }
        } else {
            return "ERROR: Path not accessible"
        }
    } catch {
        return "ERROR: $_"
    }
}

# Function to get folder size (recursive) with error handling
function Get-FolderSize {
    param(
        [string]$FolderPath,
        [switch]$HumanReadable
    )
    
    $sizeBytes = 0
    $sizeHuman = "0 B"
    
    try {
        if (Test-Path $FolderPath) {
            $files = Get-ChildItem -Path $FolderPath -Recurse -File -ErrorAction SilentlyContinue
            if ($files) {
                $sizeBytes = ($files | Measure-Object -Property Length -Sum).Sum
            }
            if ($sizeBytes -ge 1TB) {
                $sizeHuman = "{0:N2} TB" -f ($sizeBytes / 1TB)
            } elseif ($sizeBytes -ge 1GB) {
                $sizeHuman = "{0:N2} GB" -f ($sizeBytes / 1GB)
            } elseif ($sizeBytes -ge 1MB) {
                $sizeHuman = "{0:N2} MB" -f ($sizeBytes / 1MB)
            } elseif ($sizeBytes -ge 1KB) {
                $sizeHuman = "{0:N2} KB" -f ($sizeBytes / 1KB)
            } else {
                $sizeHuman = "$sizeBytes B"
            }
        } else {
            $sizeBytes = -1
            $sizeHuman = "ERROR: Path not accessible"
        }
    } catch {
        $sizeBytes = -1
        $sizeHuman = "ERROR: $_"
    }
    
    if ($HumanReadable) {
        return $sizeHuman
    } else {
        return $sizeBytes
    }
}

# Get all DFS namespaces from the specified server
Write-Host "Retrieving DFS namespaces from server: $NamespaceServer" -ForegroundColor Cyan
Write-Host "--------------------------------------------------------" -ForegroundColor Cyan

try {
    $namespaces = @(Get-DfsnRoot -ComputerName $NamespaceServer -ErrorAction Stop)
    if ($namespaces.Count -eq 0) {
        Write-Host "No DFS namespaces found on server: $NamespaceServer" -ForegroundColor Red
        exit 1
    }
    Write-Host "Found $($namespaces.Count) namespace(s)" -ForegroundColor Green
} catch {
    Write-Host "Error getting DFS namespaces from $NamespaceServer : $_" -ForegroundColor Red
    Write-Host "`nPossible causes:" -ForegroundColor Yellow
    Write-Host "  - The server name is incorrect or unreachable" -ForegroundColor Yellow
    Write-Host "  - You don't have permissions to query DFS namespaces" -ForegroundColor Yellow
    Write-Host "  - The DFS Management tools are not installed" -ForegroundColor Yellow
    Write-Host "  - The server is not a DFS namespace server" -ForegroundColor Yellow
    exit 1
}

$failedNamespaces = @()
$failedFolders = @()
$totalProcessed = 0
$namespaceCount = $namespaces.Count

foreach ($ns in $namespaces) {
    Write-Host "`nProcessing namespace: $($ns.Path)" -ForegroundColor Green
    
    try {
        $folders = @()
        try {
            $folders = @(Get-DfsnFolder -Path "$($ns.Path)\*" -ComputerName $NamespaceServer -ErrorAction SilentlyContinue)
            Write-Host "  Found $($folders.Count) folder(s)" -ForegroundColor Gray
        } catch {
            Write-Host "  ⚠️ Warning: Could not enumerate folders for $($ns.Path): $_" -ForegroundColor Yellow
        }
        
        # Root targets
        $rootTargets = @()
        try {
            $rootTargets = @(Get-DfsnRootTarget -Path $ns.Path -ComputerName $NamespaceServer -ErrorAction SilentlyContinue)
        } catch {
            Write-Host "  ⚠️ Warning: Could not get root targets for $($ns.Path): $_" -ForegroundColor Yellow
        }
        
        if ($rootTargets.Count -gt 0) {
            Write-Host "  Processing $($rootTargets.Count) root target(s)" -ForegroundColor Gray
            foreach ($rootTarget in $rootTargets) {
                try {
                    Write-Host "    - Processing: $($rootTarget.TargetPath)" -ForegroundColor DarkGray
                    $securityGroups = Get-FolderSecurityGroups -FolderPath $rootTarget.TargetPath
                    $owner = Get-FolderOwner -FolderPath $rootTarget.TargetPath
                    
                    $folderSizeBytes = 0
                    $folderSizeHuman = "N/A"
                    if ($IncludeFolderSize) {
                        $folderSizeBytes = Get-FolderSize -FolderPath $rootTarget.TargetPath
                        $folderSizeHuman = Get-FolderSize -FolderPath $rootTarget.TargetPath -HumanReadable
                    }
                    
                    $results += [PSCustomObject]@{
                        Namespace         = $ns.Path
                        FolderPath        = $ns.Path
                        FolderTarget      = $rootTarget.TargetPath
                        SourceServer      = ($rootTarget.TargetPath -split '\\')[2]
                        SourceShare       = $rootTarget.TargetPath
                        SecurityGroups    = $securityGroups
                        Owner             = $owner
                        FolderSizeBytes   = $folderSizeBytes
                        FolderSizeHuman   = $folderSizeHuman
                    }
                    $totalProcessed++
                } catch {
                    Write-Host "    ⚠️ Warning: Could not process root target $($rootTarget.TargetPath): $_" -ForegroundColor Yellow
                    $results += [PSCustomObject]@{
                        Namespace         = $ns.Path
                        FolderPath        = $ns.Path
                        FolderTarget      = $rootTarget.TargetPath
                        SourceServer      = ($rootTarget.TargetPath -split '\\')[2]
                        SourceShare       = $rootTarget.TargetPath
                        SecurityGroups    = "ERROR: Could not retrieve security groups"
                        Owner             = "ERROR: Could not retrieve owner"
                        FolderSizeBytes   = -1
                        FolderSizeHuman   = "ERROR"
                    }
                }
            }
        }
        
        # Subfolders
        if ($folders.Count -gt 0) {
            Write-Host "  Processing subfolders..." -ForegroundColor Gray
            foreach ($folder in $folders) {
                try {
                    $folderTargets = @(Get-DfsnFolderTarget -Path $folder.Path -ComputerName $NamespaceServer -ErrorAction SilentlyContinue)
                    
                    if ($folderTargets.Count -gt 0) {
                        foreach ($target in $folderTargets) {
                            try {
                                Write-Host "    - Processing: $($target.TargetPath)" -ForegroundColor DarkGray
                                $securityGroups = Get-FolderSecurityGroups -FolderPath $target.TargetPath
                                $owner = Get-FolderOwner -FolderPath $target.TargetPath
                                
                                $folderSizeBytes = 0
                                $folderSizeHuman = "N/A"
                                if ($IncludeFolderSize) {
                                    $folderSizeBytes = Get-FolderSize -FolderPath $target.TargetPath
                                    $folderSizeHuman = Get-FolderSize -FolderPath $target.TargetPath -HumanReadable
                                }
                                
                                $results += [PSCustomObject]@{
                                    Namespace         = $ns.Path
                                    FolderPath        = $folder.Path
                                    FolderTarget      = $target.TargetPath
                                    SourceServer      = ($target.TargetPath -split '\\')[2]
                                    SourceShare       = $target.TargetPath
                                    SecurityGroups    = $securityGroups
                                    Owner             = $owner
                                    FolderSizeBytes   = $folderSizeBytes
                                    FolderSizeHuman   = $folderSizeHuman
                                }
                                $totalProcessed++
                            } catch {
                                Write-Host "    ⚠️ Warning: Could not process target $($target.TargetPath): $_" -ForegroundColor Yellow
                                $results += [PSCustomObject]@{
                                    Namespace         = $ns.Path
                                    FolderPath        = $folder.Path
                                    FolderTarget      = $target.TargetPath
                                    SourceServer      = ($target.TargetPath -split '\\')[2]
                                    SourceShare       = $target.TargetPath
                                    SecurityGroups    = "ERROR: Could not retrieve security groups"
                                    Owner             = "ERROR: Could not retrieve owner"
                                    FolderSizeBytes   = -1
                                    FolderSizeHuman   = "ERROR"
                                }
                            }
                        }
                    } else {
                        Write-Host "    - No targets found for: $($folder.Path)" -ForegroundColor DarkGray
                    }
                } catch {
                    Write-Host "  ⚠️ Warning: Could not process folder $($folder.Path): $_" -ForegroundColor Yellow
                    $failedFolders += $folder.Path
                }
            }
        }
        
    } catch {
        Write-Host "  ❌ Error processing namespace $($ns.Path): $_" -ForegroundColor Red
        $failedNamespaces += $ns.Path
        continue
    }
}

# Export to CSV
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`n✅ CSV exported to $csvPath" -ForegroundColor Yellow
} else {
    Write-Host "`n⚠️ No data collected, CSV not created" -ForegroundColor Red
}

# Export to Excel
if ($results.Count -gt 0) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module for XLSX export..." -ForegroundColor Yellow
        Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    }
    try {
        Import-Module ImportExcel -ErrorAction Stop
        $results | Export-Excel -Path $xlsxPath -AutoSize -AutoFilter -FreezeTopRow -WorksheetName "DFS_Inventory" -TableName "DFSTable" -TableStyle Medium9
        Write-Host "✅ Excel exported to $xlsxPath" -ForegroundColor Yellow
    } catch {
        Write-Host "⚠️ Could not export to Excel: $_" -ForegroundColor Red
    }
}

# Generate HTML report if we have data
if ($results.Count -gt 0) {
    $htmlTemplate = @'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>DFS Namespace Inventory</title>
    <style>
        body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f4f4f4; }
        h1 { color: #2c3e50; }
        .info { background: #d9edf7; padding: 10px; border-radius: 5px; margin-bottom: 20px; }
        .warning { background: #fcf8e3; padding: 10px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #f0ad4e; }
        .search-box { margin: 20px 0; }
        .search-box input { padding: 8px; width: 300px; font-size: 16px; border-radius: 4px; border: 1px solid #ccc; }
        .search-box button { padding: 8px 15px; margin-left: 10px; background: #2980b9; color: white; border: none; border-radius: 4px; cursor: pointer; }
        .search-box button:hover { background: #1f6392; }
        .namespace { background: #fff; margin-bottom: 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .namespace-header { background: #2980b9; color: white; padding: 12px; cursor: pointer; border-radius: 6px 6px 0 0; font-size: 18px; font-weight: bold; }
        .namespace-header:hover { background: #1f6392; }
        .namespace-content { padding: 10px; display: none; }
        .ns-table { width: 100%; border-collapse: collapse; }
        .ns-table th, .ns-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .ns-table th { background: #3498db; color: white; }
        .ns-table tr:nth-child(even) { background: #f9f9f9; }
        .search-highlight { background-color: #ffeb3b; font-weight: bold; }
        .error-cell { background-color: #ffe6e6; color: #cc0000; }
        .footer { text-align: center; margin-top: 30px; font-size: 12px; color: #777; }
        .size-column { text-align: right; }
    </style>
</head>
<body>
    <h1>DFS Namespace Audit Report</h1>
    <div class="info">
        <strong>Generated:</strong> TIMESTAMP<br>
        <strong>DFS Server:</strong> NAMESPACESERVER<br>
        <strong>Total Namespaces:</strong> NAMESPACECOUNT<br>
        <strong>Total Folder Targets:</strong> TARGETCOUNT<br>
        <strong>Successfully Processed:</strong> PROCESSEDCOUNT<br>
        <strong>Folder Size Included:</strong> SIZEFLAG
    </div>
    WARNING_SECTION
    <div class="search-box">
        <input type="text" id="searchInput" onkeyup="searchTable()" placeholder="Search folders, targets, servers, groups...">
        <button onclick="expandAll()">Expand All</button>
        <button onclick="collapseAll()">Collapse All</button>
    </div>
    <div id="namespace-container">
        NAMESPACE_CONTENT
    </div>
    <div class="footer">
        DFS Inventory | Generated from NAMESPACESERVER
    </div>

    <script>
        function toggleSection(id) {
            var content = document.getElementById("content-" + id);
            if (content.style.display === "none" || content.style.display === "") {
                content.style.display = "block";
            } else {
                content.style.display = "none";
            }
        }

        function searchTable() {
            var input = document.getElementById("searchInput").value.toLowerCase();
            var allRows = document.querySelectorAll(".ns-table tbody tr");
            
            allRows.forEach(function(row) {
                var text = row.innerText.toLowerCase();
                
                if (text.indexOf(input) > -1) {
                    row.style.display = "";
                    if (input !== "") {
                        var cells = row.querySelectorAll("td");
                        cells.forEach(function(cell) {
                            var cellText = cell.innerText;
                            var regex = new RegExp("(" + input.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi");
                            cell.innerHTML = cellText.replace(regex, "<span class=\"search-highlight\">$1</span>");
                        });
                    }
                } else {
                    row.style.display = "none";
                }
            });
        }

        function expandAll() {
            var contents = document.querySelectorAll(".namespace-content");
            contents.forEach(function(content) {
                content.style.display = "block";
            });
        }

        function collapseAll() {
            var contents = document.querySelectorAll(".namespace-content");
            contents.forEach(function(content) {
                content.style.display = "none";
            });
        }
    </script>
</body>
</html>
'@

    # Build warning section
    $warningSection = ""
    if ($failedNamespaces.Count -gt 0 -or $failedFolders.Count -gt 0) {
        $warningSection = '<div class="warning"><strong>⚠️ Warnings:</strong><br>'
        if ($failedNamespaces.Count -gt 0) {
            $warningSection += '<strong>Failed Namespaces:</strong><br>'
            foreach ($failedNs in $failedNamespaces) {
                $warningSection += "• $failedNs<br>"
            }
        }
        if ($failedFolders.Count -gt 0) {
            $warningSection += '<strong>Failed Folders (partial data):</strong><br>'
            foreach ($failedFolder in $failedFolders) {
                $warningSection += "• $failedFolder<br>"
            }
        }
        $warningSection += '</div>'
    }

    # Build namespace content
    $namespaceContent = ""
    $nsId = 0
    $grouped = @($results | Group-Object Namespace)

    foreach ($group in $grouped) {
        $nsId++
        $nsName = $group.Name
        $namespaceContent += @"
    <div class="namespace">
        <div class="namespace-header" onclick="toggleSection($nsId)">
            📁 Namespace: $nsName
        </div>
        <div id="content-$nsId" class="namespace-content">
            <table class="ns-table">
                <thead>
                    <tr>
                        <th>Folder Path</th>
                        <th>Folder Target</th>
                        <th>Source Server</th>
                        <th>Source Share</th>
                        <th>Security Groups</th>
                        <th>Owner</th>
                        <th>Folder Size</th>
                    </tr>
                </thead>
                <tbody>
"@
        $groupItems = @($group.Group)
        
        foreach ($item in $groupItems) {
            $folderPath = $item.FolderPath -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $folderTarget = $item.FolderTarget -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $sourceServer = $item.SourceServer -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $sourceShare = $item.SourceShare -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $secGroups = $item.SecurityGroups -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $owner = $item.Owner -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            $sizeHuman = $item.FolderSizeHuman -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
            
            # Add error class if any field contains ERROR
            $rowClass = ""
            if ($secGroups -match "ERROR" -or $owner -match "ERROR" -or $sizeHuman -match "ERROR") {
                $rowClass = ' class="error-cell"'
            }
            
            $namespaceContent += @"
                    <tr$rowClass>
                        <td>$folderPath</td>
                        <td>$folderTarget</td>
                        <td>$sourceServer</td>
                        <td>$sourceShare</td>
                        <td>$secGroups</td>
                        <td>$owner</td>
                        <td class="size-column">$sizeHuman</td>
                    </tr>
"@
        }
        $namespaceContent += @"
                </tbody>
            </table>
        </div>
    </div>
"@
    }

    # Replace placeholders in the template
    $htmlContent = $htmlTemplate -replace "TIMESTAMP", (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $htmlContent = $htmlContent -replace "NAMESPACESERVER", $NamespaceServer
    $htmlContent = $htmlContent -replace "COMPUTERNAME", $env:COMPUTERNAME
    $htmlContent = $htmlContent -replace "NAMESPACECOUNT", $namespaceCount
    $htmlContent = $htmlContent -replace "TARGETCOUNT", $results.Count
    $htmlContent = $htmlContent -replace "PROCESSEDCOUNT", $totalProcessed
    $htmlContent = $htmlContent -replace "SIZEFLAG", $(if ($IncludeFolderSize) { "Yes" } else { "No (skipped)" })
    $htmlContent = $htmlContent -replace "WARNING_SECTION", $warningSection
    $htmlContent = $htmlContent -replace "NAMESPACE_CONTENT", $namespaceContent

    # Save the HTML file
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "✅ HTML report exported to $htmlPath" -ForegroundColor Yellow
    
    try {
        Start-Process $htmlPath
        Write-Host "📊 HTML report opened in your default browser" -ForegroundColor Cyan
    } catch {
        Write-Host "⚠️ Could not automatically open HTML report: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n⚠️ No data was collected. HTML report not created." -ForegroundColor Red
}

Write-Host "`n✅ Export complete!" -ForegroundColor Green
Write-Host "📁 Files saved to: $exportBase" -ForegroundColor Cyan

# Display summary
Write-Host "`n📊 Summary:" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "DFS Server:        $NamespaceServer" -ForegroundColor White
Write-Host "Total namespaces:  $namespaceCount" -ForegroundColor White
Write-Host "Total targets:      $($results.Count)" -ForegroundColor White
Write-Host "Processed:          $totalProcessed" -ForegroundColor Green
Write-Host "Failed namespaces:  $($failedNamespaces.Count)" -ForegroundColor $(if ($failedNamespaces.Count -gt 0) { "Red" } else { "Green" })
Write-Host "Failed folders:     $($failedFolders.Count)" -ForegroundColor $(if ($failedFolders.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "Folder size:        $(if ($IncludeFolderSize) { "Included" } else { "Skipped" })" -ForegroundColor Cyan
Write-Host "Owner info:         Included" -ForegroundColor Cyan

if ($failedNamespaces.Count -gt 0) {
    Write-Host "`n⚠️ Failed namespaces:" -ForegroundColor Yellow
    foreach ($failedNs in $failedNamespaces) {
        Write-Host "  - $failedNs" -ForegroundColor Yellow
    }
}

if ($failedFolders.Count -gt 0 -and $failedFolders.Count -le 10) {
    Write-Host "`n⚠️ Failed folders:" -ForegroundColor Yellow
    foreach ($failedFolder in $failedFolders) {
        Write-Host "  - $failedFolder" -ForegroundColor Yellow
    }
} elseif ($failedFolders.Count -gt 10) {
    Write-Host "`n⚠️ $($failedFolders.Count) folders failed (check HTML report for details)" -ForegroundColor Yellow
}

Write-Host "`n📁 Export files location: $exportBase" -ForegroundColor Cyan
