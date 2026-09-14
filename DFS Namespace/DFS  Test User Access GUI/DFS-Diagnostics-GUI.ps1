#requires -Version 5.1
<#
    *** Original PowerShell Code by Radu Vuia!!! Thanks Radu! ***
    GUI Wrapper created by Stephen McKee
    DFS Access Diagnostics - GUI wrapper
    Wraps the DFS/SMB diagnostic logic in a WinForms front end.
    Run with:  powershell.exe -STA -File .\DFS-Diagnostics-GUI.ps1

    DFS Diagnostics GUI Walkthrough
    The window gives you a clean way to launch the diagnostic, monitor progress, and review outcomes without touching the command line.
    UNC path input – Pre-filled with your original DFS path, but you can change it to any \server\share\folder target.
    Report output folder – Choose where the timestamped report and CSV summary are saved; use Browse to pick a location.
    Write access toggle – Tick the checkbox to test creating and deleting a small temp file in the share, which is otherwise skipped.
    Run button – Disables the controls, starts the background runspace, and streams each check result into the output box with color coding (green PASS, red FAIL, orange WARN, blue STEP).
    Action buttons – Open the last report folder, copy the output text, or clear the output area. A status bar at the bottom shows elapsed time and final failure/warning counts.
    Optimization Tip: You can adjust the default UNC path in the $tbPath.Text line and the output folder in the $tbOut.Text line near the top of the GUI section to match your environment. The background polling interval is set by $timer.Interval = 150 (milliseconds).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =====================================================================
#  WORKER SCRIPTBLOCK  (runs in a background runspace)
# =====================================================================
$workerScript = {
    param($Path, $OutputFolder, $TestWriteAccess, $Sync)

    $ErrorActionPreference = 'Continue'
    $ProgressPreference    = 'SilentlyContinue'

    function Send-Line {
        param([string]$Text, [string]$Status = 'INFO')
        $Sync.Queue.Enqueue([pscustomobject]@{ Status = $Status; Text = $Text })
    }

    try {
        $timeStamp    = Get-Date -Format 'yyyyMMdd-HHmmss'
        $reportFolder = Join-Path $OutputFolder "DFS-$timeStamp"
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
        $reportFile = Join-Path $reportFolder 'Diagnostic-Report.txt'
        $resultFile = Join-Path $reportFolder 'Summary.csv'
        $results    = New-Object System.Collections.Generic.List[object]

        $Sync.ReportFolder = $reportFolder

        function Add-Result {
            param([string]$Category, [string]$Test, [string]$Target,
                  [ValidateSet('PASS','FAIL','WARN','INFO')][string]$Status,
                  [string]$Details)
            $item = [pscustomobject]@{
                Time     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Category = $Category
                Test     = $Test
                Target   = $Target
                Status   = $Status
                Details  = ($Details -replace "`r?`n", ' | ')
            }
            $results.Add($item)
            Send-Line ("[{0}] {1} - {2}: {3}" -f $Status, $Category, $Test, $Details) $Status
        }

        function Invoke-CapturedCommand {
            param([string]$Name, [scriptblock]$Command)
            "`r`n===== $Name =====" | Out-File $reportFile -Append -Encoding utf8
            try {
                (& $Command 2>&1 | Out-String -Width 4096).TrimEnd() |
                    Out-File $reportFile -Append -Encoding utf8
            } catch {
                $_ | Out-String | Out-File $reportFile -Append -Encoding utf8
            }
        }

        function Test-TcpPort {
            param([string]$ComputerName, [int]$Port, [string]$Purpose)
            try {
                $test = Test-NetConnection -ComputerName $ComputerName -Port $Port `
                            -InformationLevel Detailed -WarningAction SilentlyContinue
                if ($test.TcpTestSucceeded) {
                    Add-Result 'Network' "TCP $Port ($Purpose)" $ComputerName 'PASS' 'Port accessible'
                } else {
                    Add-Result 'Network' "TCP $Port ($Purpose)" $ComputerName 'FAIL' 'Port blocked, service unavailable, or route problem'
                }
                $test | Format-List * | Out-String -Width 4096 | Out-File $reportFile -Append -Encoding utf8
            } catch {
                Add-Result 'Network' "TCP $Port ($Purpose)" $ComputerName 'FAIL' $_.Exception.Message
            }
        }

        # -------------------------------------------------------------
        $trimmed         = $Path.TrimEnd('\')
        $parts           = $trimmed -split '\\' | Where-Object { $_ }
        $namespaceServer = $parts[0]
        $shareName       = $parts[1]

        "DFS/SMB diagnostic report`r`nPath: $Path`r`nStarted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" |
            Set-Content $reportFile -Encoding utf8

        Add-Result 'Context' 'Computer' $env:COMPUTERNAME 'INFO' "Domain=$env:USERDOMAIN User=$env:USERDOMAIN\$env:USERNAME"
        Add-Result 'Context' 'Target path' $Path 'INFO' "NamespaceHost=$namespaceServer Share=$shareName"

        Send-Line 'Collecting identity, IP configuration and Kerberos tickets ...' 'STEP'
        Invoke-CapturedCommand 'WHOAMI' { whoami /all }
        Invoke-CapturedCommand 'IP configuration' { ipconfig /all }
        Invoke-CapturedCommand 'Kerberos tickets (before access)' { klist }

        Send-Line "Resolving $namespaceServer ..." 'STEP'
        try {
            $dns       = Resolve-DnsName -Name $namespaceServer -ErrorAction Stop
            $addresses = @($dns | Where-Object IPAddress | Select-Object -ExpandProperty IPAddress)
            Add-Result 'DNS' 'Resolve namespace host' $namespaceServer 'PASS' ($addresses -join ', ')
            $dns | Format-Table -AutoSize | Out-String -Width 4096 | Out-File $reportFile -Append -Encoding utf8
        } catch {
            Add-Result 'DNS' 'Resolve namespace host' $namespaceServer 'FAIL' $_.Exception.Message
        }

        Send-Line "Tracing route to $namespaceServer (this can take a moment) ..." 'STEP'
        Invoke-CapturedCommand "Route to $namespaceServer" { tracert.exe -d -w 1000 $namespaceServer }

        Send-Line "Testing TCP ports on $namespaceServer ..." 'STEP'
        Test-TcpPort $namespaceServer 445 'SMB'
        Test-TcpPort $namespaceServer 139 'NetBIOS/SMB legacy'
        Test-TcpPort $namespaceServer 135 'RPC Endpoint Mapper'

        Send-Line "Testing access to $Path ..." 'STEP'
        try {
            $null = Get-Item -LiteralPath $Path -ErrorAction Stop
            Add-Result 'Access' 'Path exists / traverse' $Path 'PASS' 'The current user can reach the path'
        } catch {
            Add-Result 'Access' 'Path exists / traverse' $Path 'FAIL' `
                ("{0} (HRESULT 0x{1:X8})" -f $_.Exception.Message, $_.Exception.HResult)
        }

        try {
            $items = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop |
                        Select-Object -First 5 Name, Mode, Length, LastWriteTime)
            Add-Result 'Access' 'List folder' $Path 'PASS' ("Listing succeeded; sampled {0} item(s)" -f $items.Count)
            $items | Format-Table -AutoSize | Out-String -Width 4096 | Out-File $reportFile -Append -Encoding utf8
        } catch {
            Add-Result 'Access' 'List folder' $Path 'FAIL' $_.Exception.Message
        }

        Send-Line 'Collecting DFS referrals, SMB connections and ACLs ...' 'STEP'
        Invoke-CapturedCommand 'DFS client cache / referrals' { dfsutil.exe /pktinfo }
        Invoke-CapturedCommand 'SMB connections' { Get-SmbConnection | Sort-Object ServerName, ShareName | Format-Table -AutoSize }
        Invoke-CapturedCommand 'Kerberos tickets (after access)' { klist }
        Invoke-CapturedCommand 'Share-visible ACL (icacls)' { icacls.exe $Path }

        try {
            $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
            Add-Result 'Permissions' 'Read ACL' $Path 'PASS' "Owner=$($acl.Owner); Access rules=$($acl.Access.Count)"
            $acl.Access | Select-Object IdentityReference, FileSystemRights, AccessControlType, IsInherited, InheritanceFlags, PropagationFlags |
                Format-Table -AutoSize | Out-String -Width 4096 | Out-File $reportFile -Append -Encoding utf8
        } catch {
            Add-Result 'Permissions' 'Read ACL' $Path 'WARN' ("Could not read ACL: {0}" -f $_.Exception.Message)
        }

        if ($TestWriteAccess) {
            $probe = Join-Path $Path ('.DFS_WriteTest_{0}_{1}.tmp' -f $env:COMPUTERNAME, [guid]::NewGuid().ToString('N'))
            try {
                [System.IO.File]::WriteAllText($probe, 'Temporary access test')
                Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
                Add-Result 'Access' 'Create/delete file' $Path 'PASS' 'Write and delete succeeded'
            } catch {
                Add-Result 'Access' 'Create/delete file' $Path 'FAIL' $_.Exception.Message
                if (Test-Path -LiteralPath $probe) {
                    Add-Result 'Access' 'Cleanup test file' $probe 'WARN' 'Temporary file could not be removed; remove it manually'
                }
            }
        } else {
            Add-Result 'Access' 'Write test' $Path 'INFO' 'Skipped. Tick "Test write access" to test create/delete rights'
        }

        # ---- Follow the actual DFS targets seen in the SMB session ----
        $targetServers = @()
        try {
            $targetServers = @(Get-SmbConnection -ErrorAction Stop |
                Where-Object { $_.ServerName -and $_.ServerName -ne $namespaceServer } |
                Select-Object -ExpandProperty ServerName -Unique)
        } catch {}

        foreach ($server in $targetServers) {
            Send-Line "Testing DFS target $server ..." 'STEP'
            try {
                $ips = @(Resolve-DnsName $server -ErrorAction Stop | Where-Object IPAddress | Select-Object -ExpandProperty IPAddress)
                Add-Result 'DFS target' 'DNS' $server 'PASS' ($ips -join ', ')
            } catch {
                Add-Result 'DFS target' 'DNS' $server 'FAIL' $_.Exception.Message
            }
            Test-TcpPort $server 445 'SMB target'
        }

        Send-Line 'Reading recent SMB client event logs ...' 'STEP'
        Invoke-CapturedCommand 'Relevant Windows SMB client events' {
            Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SMBClient/Connectivity'; StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue |
                Select-Object -First 50 TimeCreated, Id, LevelDisplayName, Message | Format-List
        }
        Invoke-CapturedCommand 'Relevant Windows SMB security events' {
            Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SMBClient/Security'; StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue |
                Select-Object -First 50 TimeCreated, Id, LevelDisplayName, Message | Format-List
        }

        $results | Export-Csv -Path $resultFile -NoTypeInformation -Encoding UTF8
        "`r`n===== SUMMARY =====" | Out-File $reportFile -Append -Encoding utf8
        $results | Format-Table -AutoSize | Out-String -Width 4096 | Out-File $reportFile -Append -Encoding utf8

        $Sync.FailCount = @($results | Where-Object Status -eq 'FAIL').Count
        $Sync.WarnCount = @($results | Where-Object Status -eq 'WARN').Count
    }
    catch {
        Send-Line ("Unexpected error: {0}" -f $_.Exception.Message) 'FAIL'
        $Sync.FailCount = -1
    }
    finally {
        $Sync.Done = $true
    }
}

# =====================================================================
#  GUI
# =====================================================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'DFS Access Diagnostics'
$form.ClientSize      = New-Object System.Drawing.Size(844, 660)
$form.StartPosition   = 'CenterScreen'
$form.MinimumSize     = New-Object System.Drawing.Size(780, 620)
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Path -------------------------------------------------------------
$lblPath           = New-Object System.Windows.Forms.Label
$lblPath.Text      = 'DFS path (UNC):'
$lblPath.Location  = New-Object System.Drawing.Point(12, 12)
$lblPath.Size      = New-Object System.Drawing.Size(400, 18)

$tbPath            = New-Object System.Windows.Forms.TextBox
$tbPath.Location   = New-Object System.Drawing.Point(12, 32)
$tbPath.Size       = New-Object System.Drawing.Size(820, 25)
$tbPath.Anchor     = 'Top,Left,Right'
$tbPath.Text       = '\\RNOP-DFSR01.myigt.com\data02\JVRoyalDept'

# --- Output folder ----------------------------------------------------
$lblOut            = New-Object System.Windows.Forms.Label
$lblOut.Text       = 'Report output folder:'
$lblOut.Location   = New-Object System.Drawing.Point(12, 68)
$lblOut.Size       = New-Object System.Drawing.Size(400, 18)

$tbOut             = New-Object System.Windows.Forms.TextBox
$tbOut.Location    = New-Object System.Drawing.Point(12, 88)
$tbOut.Size        = New-Object System.Drawing.Size(716, 25)
$tbOut.Anchor      = 'Top,Left,Right'
$tbOut.Text        = Join-Path $env:USERPROFILE 'Desktop\DFS-Diagnostics'

$btnBrowse         = New-Object System.Windows.Forms.Button
$btnBrowse.Text    = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(736, 87)
$btnBrowse.Size    = New-Object System.Drawing.Size(96, 26)
$btnBrowse.Anchor  = 'Top,Right'

# --- Options ----------------------------------------------------------
$chkWrite          = New-Object System.Windows.Forms.CheckBox
$chkWrite.Text     = 'Test write access (creates and deletes a small temp file in the share)'
$chkWrite.Location = New-Object System.Drawing.Point(12, 122)
$chkWrite.Size     = New-Object System.Drawing.Size(600, 22)

# --- Buttons ----------------------------------------------------------
$btnRun            = New-Object System.Windows.Forms.Button
$btnRun.Text       = 'Run Diagnostics'
$btnRun.Location   = New-Object System.Drawing.Point(12, 152)
$btnRun.Size       = New-Object System.Drawing.Size(150, 30)
$btnRun.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnRun.ForeColor  = [System.Drawing.Color]::White
$btnRun.FlatStyle  = 'Flat'

$btnOpen           = New-Object System.Windows.Forms.Button
$btnOpen.Text      = 'Open Report Folder'
$btnOpen.Location  = New-Object System.Drawing.Point(170, 152)
$btnOpen.Size      = New-Object System.Drawing.Size(150, 30)
$btnOpen.Enabled   = $false

$btnCopy           = New-Object System.Windows.Forms.Button
$btnCopy.Text      = 'Copy Output'
$btnCopy.Location  = New-Object System.Drawing.Point(328, 152)
$btnCopy.Size      = New-Object System.Drawing.Size(110, 30)
$btnCopy.Enabled   = $false

$btnClear          = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(446, 152)
$btnClear.Size     = New-Object System.Drawing.Size(80, 30)

# --- Output box -------------------------------------------------------
$rtb               = New-Object System.Windows.Forms.RichTextBox
$rtb.Location      = New-Object System.Drawing.Point(12, 192)
$rtb.Size          = New-Object System.Drawing.Size(820, 398)
$rtb.Anchor        = 'Top,Bottom,Left,Right'
$rtb.ReadOnly      = $true
$rtb.BackColor     = [System.Drawing.Color]::FromArgb(250, 250, 250)
$rtb.Font          = New-Object System.Drawing.Font('Consolas', 9)
$rtb.WordWrap      = $true
$rtb.DetectUrls    = $false
$rtb.HideSelection = $false

# --- Status + progress ------------------------------------------------
$lblStatus         = New-Object System.Windows.Forms.Label
$lblStatus.Text    = 'Ready.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 600)
$lblStatus.Size    = New-Object System.Drawing.Size(820, 20)
$lblStatus.Anchor  = 'Bottom,Left,Right'

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(12, 624)
$progress.Size     = New-Object System.Drawing.Size(820, 14)
$progress.Anchor   = 'Bottom,Left,Right'
$progress.Style    = 'Blocks'
$progress.Value    = 0

$form.Controls.AddRange(@(
    $lblPath, $tbPath,
    $lblOut, $tbOut, $btnBrowse,
    $chkWrite,
    $btnRun, $btnOpen, $btnCopy, $btnClear,
    $rtb, $lblStatus, $progress
))

# =====================================================================
#  HELPERS
# =====================================================================
function Add-OutputLine {
    param([string]$Text, [string]$Status = 'INFO')

    $color = switch ($Status) {
        'PASS'  { [System.Drawing.Color]::FromArgb(0, 128, 0) }
        'FAIL'  { [System.Drawing.Color]::FromArgb(200, 0, 0) }
        'WARN'  { [System.Drawing.Color]::FromArgb(200, 120, 0) }
        'STEP'  { [System.Drawing.Color]::FromArgb(0, 80, 200) }
        default { [System.Drawing.Color]::FromArgb(40, 40, 40) }
    }

    $rtb.SelectionStart  = $rtb.TextLength
    $rtb.SelectionLength = 0
    $rtb.SelectionColor  = $color
    $rtb.AppendText($Text + "`r`n")
    $rtb.SelectionColor  = $rtb.ForeColor
    $rtb.ScrollToCaret()
}

# =====================================================================
#  STATE
# =====================================================================
$script:sync            = $null
$script:runspace        = $null
$script:ps              = $null
$script:handle          = $null
$script:lastReportFolder = $null
$script:startTime       = $null

# =====================================================================
#  TIMER - drains the worker queue into the GUI
# =====================================================================
$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = 150

$timer.Add_Tick({
    if ($null -eq $script:sync) { return }

    $item = $null
    while ($script:sync.Queue.TryDequeue([ref]$item)) {
        Add-OutputLine $item.Text $item.Status
        $item = $null
    }

    if ($script:startTime) {
        $elapsed = (Get-Date) - $script:startTime
        $lblStatus.Text = ("Running ... elapsed {0:mm\:ss}" -f $elapsed)
    }

    if ($script:sync.Done -and $script:sync.Queue.IsEmpty) {
        $timer.Stop()

        $progress.Style = 'Blocks'
        $progress.MarqueeAnimationSpeed = 0
        $progress.Value = 0

        $btnRun.Enabled   = $true
        $btnOpen.Enabled  = $true
        $btnCopy.Enabled  = $true
        $tbPath.Enabled   = $true
        $tbOut.Enabled    = $true
        $btnBrowse.Enabled = $true
        $chkWrite.Enabled = $true

        if ($script:sync.ReportFolder) { $script:lastReportFolder = $script:sync.ReportFolder }

        $fails = $script:sync.FailCount
        $warns = $script:sync.WarnCount

        if ($fails -gt 0) {
            $lblStatus.Text = "Completed with $fails failure(s) and $warns warning(s)."
            Add-OutputLine '' 'INFO'
            Add-OutputLine "===== DONE: $fails FAIL / $warns WARN =====" 'FAIL'
        } elseif ($fails -lt 0) {
            $lblStatus.Text = 'Completed with an unexpected error. See output.'
            Add-OutputLine '' 'INFO'
            Add-OutputLine '===== DONE: script error =====" ' 'FAIL'
        } else {
            $lblStatus.Text = "Completed successfully. $warns warning(s)."
            Add-OutputLine '' 'INFO'
            Add-OutputLine "===== DONE: no failures / $warns WARN =====" 'PASS'
        }

        if ($script:lastReportFolder) {
            Add-OutputLine "Report folder: $script:lastReportFolder" 'INFO'
        }

        if ($script:ps)       { $script:ps.Dispose();       $script:ps = $null }
        if ($script:runspace) { $script:runspace.Dispose(); $script:runspace = $null }
        $script:handle = $null
    }
})

# =====================================================================
#  EVENT HANDLERS
# =====================================================================
$btnBrowse.Add_Click({
    $dlg             = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select where diagnostic reports are saved'
    $dlg.ShowNewFolderButton = $true
    if (Test-Path -LiteralPath $tbOut.Text) { $dlg.SelectedPath = $tbOut.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $tbOut.Text = $dlg.SelectedPath
    }
})

$btnClear.Add_Click({
    $rtb.Clear()
    $lblStatus.Text = 'Ready.'
})

$btnCopy.Add_Click({
    if ($rtb.TextLength -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($rtb.Text)
        $lblStatus.Text = 'Output copied to clipboard.'
    }
})

$btnOpen.Add_Click({
    $target = $script:lastReportFolder
    if ($target -and (Test-Path -LiteralPath $target)) {
        Start-Process explorer.exe $target
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            'No report folder is available yet. Run the diagnostics first.',
            'Open Report Folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
})

$btnRun.Add_Click({
    $path = $tbPath.Text.Trim()

    if ($path -notmatch '^\\\\[^\\]+\\[^\\]+') {
        [System.Windows.Forms.MessageBox]::Show(
            "Enter a UNC path in the form \\server\share\folder.",
            'Invalid path',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $out = $tbOut.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($out)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter an output folder for the reports.',
            'Missing output folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $out)) {
            New-Item -Path $out -ItemType Directory -Force | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not create the output folder:`r`n$($_.Exception.Message)",
            'Output folder error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # ---- reset UI ----
    $rtb.Clear()
    $btnRun.Enabled    = $false
    $btnOpen.Enabled   = $false
    $btnCopy.Enabled   = $false
    $tbPath.Enabled    = $false
    $tbOut.Enabled     = $false
    $btnBrowse.Enabled = $false
    $chkWrite.Enabled  = $false

    $progress.Style = 'Marquee'
    $progress.MarqueeAnimationSpeed = 30
    $script:startTime = Get-Date

    Add-OutputLine "DFS access diagnostics started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 'STEP'
    Add-OutputLine "Path   : $path" 'INFO'
    Add-OutputLine "Output : $out" 'INFO'
    Add-OutputLine "Write test : $($chkWrite.Checked)" 'INFO'
    Add-OutputLine ('-' * 78) 'INFO'

    # ---- shared state ----
    $script:sync = [hashtable]::Synchronized(@{
        Queue        = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
        Done         = $false
        ReportFolder = $null
        FailCount    = 0
        WarnCount    = 0
    })

    # ---- background runspace ----
    $script:runspace               = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = 'MTA'
    $script:runspace.ThreadOptions  = 'ReuseThread'
    $script:runspace.Open()

    $script:ps           = [powershell]::Create()
    $script:ps.Runspace  = $script:runspace
    $null = $script:ps.AddScript($workerScript.ToString())
    $null = $script:ps.AddArgument($path)
    $null = $script:ps.AddArgument($out)
    $null = $script:ps.AddArgument([bool]$chkWrite.Checked)
    $null = $script:ps.AddArgument($script:sync)

    $script:handle = $script:ps.BeginInvoke()

    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    if ($script:ps)       { try { $script:ps.Dispose() }       catch {} }
    if ($script:runspace) { try { $script:runspace.Dispose() } catch {} }
})

# =====================================================================
#  SHOW
# =====================================================================
[void]$form.ShowDialog()
$form.Dispose()
