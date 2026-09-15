# Requires the ActiveDirectory module (RSAT)
Import-Module ActiveDirectory -ErrorAction Stop

# Get the current user's Desktop
$DesktopPath = [Environment]::GetFolderPath('Desktop')
$TimeStamp   = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$BaseName    = "MYIGT - PrintServerReport_$TimeStamp"

$CsvPath  = Join-Path $DesktopPath "$BaseName.csv"
$XlsxPath = Join-Path $DesktopPath "$BaseName.xlsx"
$HtmlPath = Join-Path $DesktopPath "$BaseName.html"

# 1. Get unique server names from all print queue objects
$ServerNames = Get-ADObject -Filter 'objectCategory -eq "printQueue"' `
    -Property ServerName |
    Select-Object -ExpandProperty ServerName |
    Where-Object { $_ } |
    Sort-Object -Unique

if (-not $ServerNames) {
    Write-Warning "No print queue servers found."
    return
}

# 2. Build the server report
$ServerReport = foreach ($server in $ServerNames) {
    $Location     = ''
    $Description  = ''
    $IPAddress    = 'N/A'
    $OnlineStatus = 'Unknown'

    # Get Location and Description from the AD computer object
    try {
        $adComp = Get-ADComputer -Identity $server -Properties Location, Description -ErrorAction Stop
        $Location    = $adComp.Location
        $Description = $adComp.Description
    } catch {
        # Leave blank if computer object not found or inaccessible
    }

    # Resolve IPv4 address
    try {
        $dnsResult = Resolve-DnsName -Name $server -Type A -ErrorAction Stop
        $ip = ($dnsResult | Where-Object { $_.IPAddress } | Select-Object -First 1).IPAddress
        if ($ip) { $IPAddress = $ip }
    } catch {
        # IP remains 'N/A'
    }

    # Check online status via ICMP ping
    try {
        $ping = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction Stop
        $OnlineStatus = if ($ping) { 'Online' } else { 'Offline' }
    } catch {
        $OnlineStatus = 'Offline'
    }

    [PSCustomObject]@{
        ServerName   = $server
        Location     = $Location
        Description  = $Description
        IPAddress    = $IPAddress
        OnlineStatus = $OnlineStatus
    }
}

# --- Export to CSV ---
$ServerReport | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# --- Export to HTML ---
$HtmlHead = @"
<style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
    h1 { color: #333; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
    th { background-color: #f2f2f2; }
    tr:nth-child(even) { background-color: #fafafa; }
</style>
"@

$ServerReport |
    ConvertTo-Html -Title 'MYIGT - Print Server Report' -Head $HtmlHead `
        -PreContent "<h1>MYIGT - Print Server Report</h1><p>Generated: $(Get-Date)</p>" |
    Out-File -FilePath $HtmlPath -Encoding UTF8

# --- Export to XLSX ---
$XlsxCreated = $false

if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel -ErrorAction Stop

    $ServerReport |
        Export-Excel -Path $XlsxPath -WorksheetName 'PrintServers' `
            -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

    $XlsxCreated = $true
}
else {
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        $Workbook  = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Worksheets.Item(1)
        $Worksheet.Name = 'PrintServers'

        # Headers
        $headers = @('ServerName','Location','Description','IPAddress','OnlineStatus')
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $Worksheet.Cells.Item(1, $col + 1) = $headers[$col]
        }

        # Data
        $row = 2
        foreach ($item in $ServerReport) {
            $Worksheet.Cells.Item($row, 1) = $item.ServerName
            $Worksheet.Cells.Item($row, 2) = $item.Location
            $Worksheet.Cells.Item($row, 3) = $item.Description
            $Worksheet.Cells.Item($row, 4) = $item.IPAddress
            $Worksheet.Cells.Item($row, 5) = $item.OnlineStatus
            $row++
        }

        # Formatting
        $Worksheet.Rows.Item(1).Font.Bold = $true
        $Worksheet.Columns.Item('A:E').AutoFit() | Out-Null
        $Worksheet.Application.ActiveWindow.SplitRow = 1
        $Worksheet.Application.ActiveWindow.FreezePanes = $true

        $Workbook.SaveAs($XlsxPath, 51) # 51 = xlOpenXMLWorkbook
        $Workbook.Close($false)
        $Excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null

        $XlsxCreated = $true
    }
    catch {
        Write-Warning "XLSX export failed. Install ImportExcel with: Install-Module ImportExcel -Scope CurrentUser. Error: $($_.Exception.Message)"
    }
    finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Write-Host "Reports exported to: $DesktopPath"
Write-Host "CSV : $CsvPath"
Write-Host "HTML: $HtmlPath"

if ($XlsxCreated) {
    Write-Host "XLSX: $XlsxPath"
}
else {
    Write-Host "XLSX: not created"
}
