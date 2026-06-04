# powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\path\Enterprise_Server_Deployment_Manager_V4.ps1"
<#
.SYNOPSIS
    Enterprise Server Deployment Manager v4 - IGT PLC
.DESCRIPTION
    Multi-server remote management tool: connectivity tests, WinRM checks, PowerShell
    execution, system inventory, service management, hotfix reporting, event log review,
    and scheduled task listing. All operations log to C:\ProgramData\IGTDeploymentManager\Logs.
.NOTES
    Author  : Stephen McKee - Server Administrator - IGT PLC
    Version : 4.0
    Requires: PowerShell 5.1  |  WinRM enabled on targets for remoting functions
#>

#Requires -Version 5.1

Set-StrictMode -Off   # WinForms event closures reference outer-scope variables
$ErrorActionPreference = 'Continue'

# ══════════════════════════════════════════════════════════════════════════════
#  ASSEMBLIES
# ══════════════════════════════════════════════════════════════════════════════
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ══════════════════════════════════════════════════════════════════════════════
$LogFolder = "C:\ProgramData\IGTDeploymentManager\Logs"
if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
$LogFile   = Join-Path $LogFolder ("Deployment_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$ts [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    # Append to status box if it already exists
    if ($script:txtStatus -and -not $script:txtStatus.IsDisposed) {
        $script:txtStatus.SelectionStart  = $script:txtStatus.TextLength
        $script:txtStatus.SelectionLength = 0
        $script:txtStatus.SelectionColor  = switch ($Level) {
            'SUCCESS' { [System.Drawing.Color]::FromArgb(166,227,161) }
            'WARN'    { [System.Drawing.Color]::FromArgb(249,226,175) }
            'ERROR'   { [System.Drawing.Color]::FromArgb(243,139,168) }
            default   { [System.Drawing.Color]::FromArgb(180,192,214) }
        }
        $script:txtStatus.AppendText("$entry`r`n")
        $script:txtStatus.ScrollToCaret()
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  COLOUR PALETTE  (Catppuccin Mocha-inspired dark)
# ══════════════════════════════════════════════════════════════════════════════
$clr = @{
    Base      = [System.Drawing.Color]::FromArgb(30,  30,  46 )   # #1E1E2E
    Surface0  = [System.Drawing.Color]::FromArgb(42,  42,  62 )   # #2A2A3E
    Surface1  = [System.Drawing.Color]::FromArgb(49,  49,  73 )   # #313149
    Surface2  = [System.Drawing.Color]::FromArgb(54,  58,  79 )   # #363A4F
    Overlay   = [System.Drawing.Color]::FromArgb(108, 112, 134)   # #6C7086
    Text      = [System.Drawing.Color]::FromArgb(205, 214, 244)   # #CDD6F4
    Subtext   = [System.Drawing.Color]::FromArgb(147, 153, 178)   # #9399B2
    Blue      = [System.Drawing.Color]::FromArgb(137, 180, 250)   # #89B4FA
    Green     = [System.Drawing.Color]::FromArgb(166, 227, 161)   # #A6E3A1
    Yellow    = [System.Drawing.Color]::FromArgb(249, 226, 175)   # #F9E2AF
    Red       = [System.Drawing.Color]::FromArgb(243, 139, 168)   # #F38BA8
    Peach     = [System.Drawing.Color]::FromArgb(250, 179, 135)   # #FAB387
    Purple    = [System.Drawing.Color]::FromArgb(203, 166, 247)   # #CBA6F7
    Teal      = [System.Drawing.Color]::FromArgb(148, 226, 213)   # #94E2D5
    Border    = [System.Drawing.Color]::FromArgb(69,  71,  90 )   # #45475A
    RowOnline = [System.Drawing.Color]::FromArgb(38,  56,  44 )   # dark green tint
    RowOffline= [System.Drawing.Color]::FromArgb(56,  36,  44 )   # dark red tint
    RowWarn   = [System.Drawing.Color]::FromArgb(56,  50,  30 )   # dark yellow tint
}

# Fonts
$fntDefault = New-Object System.Drawing.Font("Segoe UI", 9)
$fntBold    = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fntMono    = New-Object System.Drawing.Font("Consolas",  9)
$fntMonoSm  = New-Object System.Drawing.Font("Consolas",  8)
$fntTitle   = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$fntSmall   = New-Object System.Drawing.Font("Segoe UI",  8)

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: style a GroupBox
# ══════════════════════════════════════════════════════════════════════════════
function New-StyledGroup {
    param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text      = $Text
    $g.Location  = New-Object System.Drawing.Point($X, $Y)
    $g.Size      = New-Object System.Drawing.Size($W, $H)
    $g.ForeColor = $clr.Blue
    $g.BackColor = $clr.Surface0
    $g.Font      = $fntBold
    return $g
}

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: styled Button
# ══════════════════════════════════════════════════════════════════════════════
function New-StyledButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y, [int]$W = 190, [int]$H = 32,
        [System.Drawing.Color]$FG,
        [System.Drawing.Color]$BG
    )
    if (-not $PSBoundParameters.ContainsKey('FG')) { $FG = $clr.Text }
    if (-not $PSBoundParameters.ContainsKey('BG')) { $BG = $clr.Surface1 }
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X, $Y)
    $b.Size      = New-Object System.Drawing.Size($W, $H)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderColor      = $clr.Border
    $b.FlatAppearance.BorderSize       = 1
    $b.FlatAppearance.MouseOverBackColor = $clr.Surface2
    $b.BackColor = $BG
    $b.ForeColor = $FG
    $b.Font      = $fntDefault
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: styled Label
# ══════════════════════════════════════════════════════════════════════════════
function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 100, [int]$H = 20,
          [System.Drawing.Color]$FG)
    if (-not $PSBoundParameters.ContainsKey('FG')) { $FG = $clr.Subtext }
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.Location  = New-Object System.Drawing.Point($X, $Y)
    $l.Size      = New-Object System.Drawing.Size($W, $H)
    $l.ForeColor = $FG
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.Font      = $fntDefault
    return $l
}

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: styled TextBox
# ══════════════════════════════════════════════════════════════════════════════
function New-StyledTextBox {
    param([int]$X, [int]$Y, [int]$W, [int]$H,
          [bool]$Multi = $false, [bool]$Mono = $false, [bool]$Password = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = New-Object System.Drawing.Point($X, $Y)
    $t.Size      = New-Object System.Drawing.Size($W, $H)
    $t.BackColor = $clr.Surface1
    $t.ForeColor = $clr.Text
    $t.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $t.Font      = if ($Mono) { $fntMono } else { $fntDefault }
    if ($Multi)    { $t.Multiline = $true; $t.ScrollBars = "Vertical" }
    if ($Password) { $t.UseSystemPasswordChar = $true }
    return $t
}

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER: style DataGridView
# ══════════════════════════════════════════════════════════════════════════════
function Set-GridStyle {
    param([System.Windows.Forms.DataGridView]$Grid)
    $Grid.BackgroundColor            = $clr.Base
    $Grid.GridColor                  = $clr.Border
    $Grid.BorderStyle                = [System.Windows.Forms.BorderStyle]::None
    $Grid.DefaultCellStyle.BackColor = $clr.Surface0
    $Grid.DefaultCellStyle.ForeColor = $clr.Text
    $Grid.DefaultCellStyle.Font      = $fntDefault
    $Grid.DefaultCellStyle.SelectionBackColor = $clr.Surface2
    $Grid.DefaultCellStyle.SelectionForeColor = $clr.Text
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $clr.Base
    $Grid.AlternatingRowsDefaultCellStyle.ForeColor = $clr.Text
    $Grid.ColumnHeadersDefaultCellStyle.BackColor   = $clr.Surface1
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor   = $clr.Blue
    $Grid.ColumnHeadersDefaultCellStyle.Font        = $fntBold
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $clr.Surface1
    $Grid.ColumnHeadersBorderStyle  = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $Grid.RowHeadersVisible         = $false
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.ReadOnly                  = $true
    $Grid.AllowUserToAddRows        = $false
    $Grid.AllowUserToDeleteRows     = $false
    $Grid.AutoSizeRowsMode          = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
}

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN FORM
# ══════════════════════════════════════════════════════════════════════════════
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Enterprise Server Deployment Manager  v4  |  IGT PLC"
$form.Size            = New-Object System.Drawing.Size(1500, 960)
$form.StartPosition   = "CenterScreen"
$form.BackColor       = $clr.Base
$form.ForeColor       = $clr.Text
$form.Font            = $fntDefault
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize     = New-Object System.Drawing.Size(1200, 800)

# ── Title banner ───────────────────────────────────────────────────────────────
$pnlBanner = New-Object System.Windows.Forms.Panel
$pnlBanner.Location  = New-Object System.Drawing.Point(0, 0)
$pnlBanner.Size      = New-Object System.Drawing.Size(1500, 48)
$pnlBanner.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 37)
$pnlBanner.Anchor    = "Top,Left,Right"
$form.Controls.Add($pnlBanner)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "  ■  Enterprise Server Deployment Manager"
$lblTitle.Location  = New-Object System.Drawing.Point(6, 8)
$lblTitle.Size      = New-Object System.Drawing.Size(700, 30)
$lblTitle.ForeColor = $clr.Blue
$lblTitle.Font      = $fntTitle
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$pnlBanner.Controls.Add($lblTitle)

$lblVersion = New-Object System.Windows.Forms.Label
$lblVersion.Text      = "v4.0  |  IGT PLC  |  Stephen McKee"
$lblVersion.Location  = New-Object System.Drawing.Point(1150, 14)
$lblVersion.Size      = New-Object System.Drawing.Size(320, 20)
$lblVersion.ForeColor = $clr.Overlay
$lblVersion.Font      = $fntSmall
$lblVersion.BackColor = [System.Drawing.Color]::Transparent
$lblVersion.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$pnlBanner.Controls.Add($lblVersion)

# ── Separator under banner ─────────────────────────────────────────────────────
$sepBanner = New-Object System.Windows.Forms.Panel
$sepBanner.Location  = New-Object System.Drawing.Point(0, 48)
$sepBanner.Size      = New-Object System.Drawing.Size(1500, 1)
$sepBanner.BackColor = $clr.Border
$form.Controls.Add($sepBanner)

# ══════════════════════════════════════════════════════════════════════════════
#  LEFT COLUMN  (Credentials + Server List)  X=10, Y=58
# ══════════════════════════════════════════════════════════════════════════════

# ── Credentials GroupBox ───────────────────────────────────────────────────────
$grpCreds = New-StyledGroup "  Credentials" 10 58 410 195
$form.Controls.Add($grpCreds)

$grpCreds.Controls.Add((New-Label "Domain"   10 32 100 20))
$cmbDomain = New-Object System.Windows.Forms.ComboBox
$cmbDomain.Location  = New-Object System.Drawing.Point(115, 28)
$cmbDomain.Size      = New-Object System.Drawing.Size(270, 22)
$cmbDomain.BackColor = $clr.Surface1
$cmbDomain.ForeColor = $clr.Text
$cmbDomain.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cmbDomain.Font      = $fntDefault
$cmbDomain.Items.AddRange(@("MYIGT.COM","AD.IGT.COM","IGTSAP.AD.IGT.COM","IS.AD.IGT.COM","ADE.EC.IGT.COM","CUSTOM"))
$cmbDomain.SelectedIndex = 0
$grpCreds.Controls.Add($cmbDomain)

$grpCreds.Controls.Add((New-Label "Username" 10 68 100 20))
$script:txtUser = New-StyledTextBox 115 64 270 22
$grpCreds.Controls.Add($script:txtUser)

$grpCreds.Controls.Add((New-Label "Password" 10 103 100 20))
$script:txtPassword = New-StyledTextBox 115 99 270 22 -Password $true
$grpCreds.Controls.Add($script:txtPassword)

$chkCurrent = New-Object System.Windows.Forms.CheckBox
$chkCurrent.Text      = "Use Current Windows Credentials"
$chkCurrent.Location  = New-Object System.Drawing.Point(10, 134)
$chkCurrent.Size      = New-Object System.Drawing.Size(260, 22)
$chkCurrent.ForeColor = $clr.Subtext
$chkCurrent.BackColor = [System.Drawing.Color]::Transparent
$chkCurrent.Font      = $fntDefault
$grpCreds.Controls.Add($chkCurrent)

# Test creds button
$btnTestCreds = New-StyledButton "  Verify Credentials" 10 162 190 28 -FG $clr.Teal
$grpCreds.Controls.Add($btnTestCreds)

# ── Server List GroupBox ───────────────────────────────────────────────────────
$grpServers = New-StyledGroup "  Server List" 10 264 410 310
$form.Controls.Add($grpServers)

$script:txtServers = New-StyledTextBox 10 24 388 210 -Multi $true
$script:txtServers.Font  = $fntMono
$script:txtServers.ScrollBars = "Vertical"
$grpServers.Controls.Add($script:txtServers)

# Server list buttons
$btnImport = New-StyledButton "  Import from File" 10 243 188 26 -FG $clr.Blue
$btnClearSrv = New-StyledButton "  Clear List"    205 243 188 26 -FG $clr.Yellow
$grpServers.Controls.Add($btnImport)
$grpServers.Controls.Add($btnClearSrv)

$lblSrvHint = New-Label "One hostname per line  |  supports comma-separated" 10 276 388 18
$lblSrvHint.Font      = $fntSmall
$lblSrvHint.ForeColor = $clr.Overlay
$grpServers.Controls.Add($lblSrvHint)

# ── Quick Actions GroupBox ─────────────────────────────────────────────────────
$grpQuick = New-StyledGroup "  Quick Actions" 10 585 410 320
$form.Controls.Add($grpQuick)

$quickButtons = @(
    @{ Text = "  Test Connectivity (Ping)";    FG = $clr.Green  }
    @{ Text = "  Test WinRM Connectivity";     FG = $clr.Teal   }
    @{ Text = "  Get System Information";      FG = $clr.Blue   }
    @{ Text = "  Get Installed Hotfixes";      FG = $clr.Purple }
    @{ Text = "  List Services";               FG = $clr.Blue   }
    @{ Text = "  List Scheduled Tasks";        FG = $clr.Purple }
    @{ Text = "  Get Event Log Errors (24h)";  FG = $clr.Red    }
    @{ Text = "  Get Disk Space";              FG = $clr.Teal   }
)

$script:quickBtns = @()
$qY = 26
foreach ($qb in $quickButtons) {
    $b = New-StyledButton $qb.Text 10 $qY 388 28 -FG $qb.FG
    $b.Tag = $qb.Text.Trim()
    $grpQuick.Controls.Add($b)
    $script:quickBtns += $b
    $qY += 34
}

# ══════════════════════════════════════════════════════════════════════════════
#  CENTRE COLUMN  (Snippet Library + PS Script Console)  X=430
# ══════════════════════════════════════════════════════════════════════════════

# ── Snippet Library ────────────────────────────────────────────────────────────
$grpSnippet = New-StyledGroup "  Script Snippet Library" 430 58 820 70
$form.Controls.Add($grpSnippet)

$grpSnippet.Controls.Add((New-Label "Category" 10 30 80 22))
$cmbSnippetCat = New-Object System.Windows.Forms.ComboBox
$cmbSnippetCat.Location  = New-Object System.Drawing.Point(96, 26)
$cmbSnippetCat.Size      = New-Object System.Drawing.Size(160, 22)
$cmbSnippetCat.BackColor = $clr.Surface1
$cmbSnippetCat.ForeColor = $clr.Text
$cmbSnippetCat.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cmbSnippetCat.DropDownStyle = "DropDownList"
$grpSnippet.Controls.Add($cmbSnippetCat)

$grpSnippet.Controls.Add((New-Label "Snippet" 270 30 60 22))
$cmbSnippet = New-Object System.Windows.Forms.ComboBox
$cmbSnippet.Location  = New-Object System.Drawing.Point(336, 26)
$cmbSnippet.Size      = New-Object System.Drawing.Size(360, 22)
$cmbSnippet.BackColor = $clr.Surface1
$cmbSnippet.ForeColor = $clr.Text
$cmbSnippet.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cmbSnippet.DropDownStyle = "DropDownList"
$grpSnippet.Controls.Add($cmbSnippet)

$btnLoadSnippet = New-StyledButton "  Load Snippet" 706 24 106 28 -FG $clr.Blue
$grpSnippet.Controls.Add($btnLoadSnippet)

# Snippet data
$script:Snippets = @{
    "General" = @{
        "Get Services (All)"           = "Get-Service | Select-Object Name,Status,StartType | Sort-Object Name"
        "Get Services (Running)"       = "Get-Service | Where-Object { `$_.Status -eq 'Running' } | Select-Object Name,Status,StartType"
        "Get Installed Hotfixes"       = "Get-HotFix | Select-Object HotFixID,Description,InstalledOn | Sort-Object InstalledOn -Descending"
        "Get Local Admins"             = "Get-LocalGroupMember -Group 'Administrators' | Select-Object Name,ObjectClass,PrincipalSource"
        "Get Uptime"                   = "`$os = Get-CimInstance Win32_OperatingSystem; `$uptime = (Get-Date) - `$os.LastBootUpTime; Write-Output `"Uptime: `$(`$uptime.Days)d `$(`$uptime.Hours)h `$(`$uptime.Minutes)m  |  Last Boot: `$(`$os.LastBootUpTime)`""
        "Get OS Info"                  = "Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber,OSArchitecture,TotalVisibleMemorySize,FreePhysicalMemory"
        "Get CPU Info"                 = "Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed"
        "Get Network Adapters"         = "Get-NetAdapter | Where-Object { `$_.Status -eq 'Up' } | Select-Object Name,InterfaceDescription,MacAddress,LinkSpeed"
        "Get IP Configuration"         = "Get-NetIPAddress | Where-Object { `$_.AddressFamily -eq 'IPv4' -and `$_.PrefixOrigin -ne 'WellKnown' } | Select-Object InterfaceAlias,IPAddress,PrefixLength"
    }
    "Disk" = @{
        "Get Disk Space"               = "Get-PSDrive -PSProvider FileSystem | Select-Object Name,@{n='Used(GB)';e={[Math]::Round(`$_.Used/1GB,2)}},@{n='Free(GB)';e={[Math]::Round(`$_.Free/1GB,2)}},@{n='Total(GB)';e={[Math]::Round((`$_.Used+`$_.Free)/1GB,2)}} | Format-Table -AutoSize"
        "Get Disk Info (WMI)"          = "Get-CimInstance Win32_DiskDrive | Select-Object Model,@{n='Size(GB)';e={[Math]::Round(`$_.Size/1GB,0)}},SerialNumber,Status"
        "Get Volume Info"              = "Get-Volume | Where-Object { `$_.DriveLetter } | Select-Object DriveLetter,FileSystemLabel,FileSystem,@{n='Size(GB)';e={[Math]::Round(`$_.Size/1GB,2)}},@{n='Free(GB)';e={[Math]::Round(`$_.SizeRemaining/1GB,2)}},HealthStatus"
        "Find Large Files (>500MB)"    = "Get-ChildItem -Path C:\ -Recurse -ErrorAction SilentlyContinue | Where-Object { !`$_.PSIsContainer -and `$_.Length -gt 500MB } | Select-Object FullName,@{n='Size(MB)';e={[Math]::Round(`$_.Length/1MB,1)}} | Sort-Object 'Size(MB)' -Descending | Select -First 20"
    }
    "Services" = @{
        "List All Services"            = "Get-Service | Select-Object Name,DisplayName,Status,StartType | Sort-Object Name | Format-Table -AutoSize"
        "List Stopped Auto Services"   = "Get-Service | Where-Object { `$_.StartType -eq 'Automatic' -and `$_.Status -eq 'Stopped' } | Select-Object Name,DisplayName,Status"
        "Start a Service"              = "Start-Service -Name 'ServiceName'  # Replace ServiceName"
        "Stop a Service"               = "Stop-Service  -Name 'ServiceName'  # Replace ServiceName"
        "Restart a Service"            = "Restart-Service -Name 'ServiceName' -Force  # Replace ServiceName"
        "Set Service StartType Auto"   = "Set-Service -Name 'ServiceName' -StartupType Automatic"
    }
    "Events" = @{
        "System Errors (24h)"          = "`$Start = (Get-Date).AddHours(-24); Get-EventLog -LogName System -EntryType Error -After `$Start | Select-Object TimeGenerated,Source,EventID,Message | Sort-Object TimeGenerated -Descending | Select -First 50"
        "Application Errors (24h)"     = "`$Start = (Get-Date).AddHours(-24); Get-EventLog -LogName Application -EntryType Error -After `$Start | Select-Object TimeGenerated,Source,EventID,Message | Sort-Object TimeGenerated -Descending | Select -First 50"
        "Security Events (24h)"        = "`$Start = (Get-Date).AddHours(-24); Get-EventLog -LogName Security -After `$Start | Select-Object TimeGenerated,Source,EventID,Message | Sort-Object TimeGenerated -Descending | Select -First 50"
        "Failed Logons (4625)"         = "Get-WinEvent -FilterHashtable @{LogName='Security';Id=4625;StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue | Select-Object TimeCreated,@{n='User';e={`$_.Properties[5].Value}},@{n='IP';e={`$_.Properties[19].Value}} | Select -First 30"
        "Clear System Event Log"       = "Clear-EventLog -LogName System"
        "Clear Application Event Log"  = "Clear-EventLog -LogName Application"
    }
    "Scheduled Tasks" = @{
        "List All Tasks"               = "Get-ScheduledTask | Select-Object TaskName,TaskPath,State | Sort-Object TaskPath,TaskName | Format-Table -AutoSize"
        "List Running Tasks"           = "Get-ScheduledTask | Where-Object { `$_.State -eq 'Running' } | Select-Object TaskName,TaskPath"
        "List Disabled Tasks"          = "Get-ScheduledTask | Where-Object { `$_.State -eq 'Disabled' } | Select-Object TaskName,TaskPath"
        "Get Task Details"             = "Get-ScheduledTask -TaskName 'TaskName' | Get-ScheduledTaskInfo  # Replace TaskName"
    }
    "Windows Update" = @{
        "Check WSUS Last Sync"         = "Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect' | Select-Object LastSuccessTime"
        "Get Pending Reboots"          = "Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'"
        "Get Windows Update Log"       = "Get-Content C:\Windows\WindowsUpdate.log -Tail 100 -ErrorAction SilentlyContinue"
        "Check SCCM Client Status"     = "Get-CimInstance -Namespace root\ccm -ClassName SMS_Client | Select-Object ClientVersion,ClientState"
        "Trigger SCCM Update Scan"     = "Invoke-CimMethod -Namespace root\ccm -ClassName SMS_Client -MethodName TriggerSchedule -Arguments @{sScheduleID='{00000000-0000-0000-0000-000000000113}'}"
    }
    "Security" = @{
        "Check Open Ports"             = "Get-NetTCPConnection -State Listen | Select-Object LocalAddress,LocalPort,State | Sort-Object LocalPort"
        "Get Firewall Profiles"        = "Get-NetFirewallProfile | Select-Object Name,Enabled,DefaultInboundAction,DefaultOutboundAction"
        "Get Logged-On Users"          = "query user /server:`$env:COMPUTERNAME"
        "Get Failed Logons (1h)"       = "Get-WinEvent -FilterHashtable @{LogName='Security';Id=4625;StartTime=(Get-Date).AddHours(-1)} -ErrorAction SilentlyContinue | Measure-Object | Select-Object Count"
        "Check AV Status"              = "Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct | Select-Object DisplayName,productState"
        "Get BitLocker Status"         = "Get-BitLockerVolume | Select-Object MountPoint,VolumeStatus,ProtectionStatus,EncryptionPercentage"
    }
}

# Populate snippet category combo
foreach ($cat in ($script:Snippets.Keys | Sort-Object)) {
    $cmbSnippetCat.Items.Add($cat) | Out-Null
}
$cmbSnippetCat.SelectedIndex = 0

# ── PS Script Console ──────────────────────────────────────────────────────────
$grpPS = New-StyledGroup "  Remote PowerShell Console" 430 138 820 390
$form.Controls.Add($grpPS)

$script:txtScript = New-StyledTextBox 10 24 798 350 -Multi $true -Mono $true
$script:txtScript.ScrollBars = "Both"
$script:txtScript.WordWrap   = $false
$script:txtScript.Text = @"
# Enter PowerShell commands to execute on all listed servers
# Use the Snippet Library above to load pre-built commands

Get-Service | Select-Object Name, Status | Sort-Object Name
"@
$grpPS.Controls.Add($script:txtScript)

# ── Execute Buttons row ────────────────────────────────────────────────────────
$grpExec = New-StyledGroup "  Execute" 430 538 820 80
$form.Controls.Add($grpExec)

$btnExecute   = New-StyledButton "  Execute PowerShell"         10 22 220 34 -FG $clr.Green
$btnExecAsync = New-StyledButton "  Execute (Background Jobs)"  240 22 220 34 -FG $clr.Teal
$btnStopJobs  = New-StyledButton "  Stop All Jobs"              470 22 150 34 -FG $clr.Red
$btnClearGrid = New-StyledButton "  Clear Results"              630 22 150 34 -FG $clr.Yellow
$btnExportCsv = New-StyledButton "  Export to CSV"              790 22 120 34 -FG $clr.Purple
$grpExec.Controls.Add($btnExecute)
$grpExec.Controls.Add($btnExecAsync)
$grpExec.Controls.Add($btnStopJobs)
$grpExec.Controls.Add($btnClearGrid)
$grpExec.Controls.Add($btnExportCsv)

# ── Progress Bar ───────────────────────────────────────────────────────────────
$pnlProgress = New-Object System.Windows.Forms.Panel
$pnlProgress.Location  = New-Object System.Drawing.Point(430, 628)
$pnlProgress.Size      = New-Object System.Drawing.Size(820, 36)
$pnlProgress.BackColor = $clr.Surface0
$form.Controls.Add($pnlProgress)

$pnlProgress.Controls.Add((New-Label "Progress:" 6 8 72 20 -FG $clr.Subtext))

$script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$script:ProgressBar.Location  = New-Object System.Drawing.Point(82, 8)
$script:ProgressBar.Size      = New-Object System.Drawing.Size(628, 20)
$script:ProgressBar.Style     = "Continuous"
$script:ProgressBar.BackColor = $clr.Surface1
$script:ProgressBar.ForeColor = $clr.Green
$pnlProgress.Controls.Add($script:ProgressBar)

$script:lblProgress = New-Label "Idle" 718 8 96 20 -FG $clr.Subtext
$pnlProgress.Controls.Add($script:lblProgress)

# ── Results DataGridView ───────────────────────────────────────────────────────
$grpGrid = New-StyledGroup "  Results" 430 674 820 230
$form.Controls.Add($grpGrid)

$script:Grid = New-Object System.Windows.Forms.DataGridView
$script:Grid.Location = New-Object System.Drawing.Point(10, 22)
$script:Grid.Size     = New-Object System.Drawing.Size(798, 198)
$script:Grid.Anchor   = "Top,Left,Right,Bottom"
Set-GridStyle $script:Grid

$script:Grid.ColumnCount = 5
$script:Grid.Columns[0].Name  = "Server";    $script:Grid.Columns[0].Width = 160
$script:Grid.Columns[1].Name  = "Status";    $script:Grid.Columns[1].Width = 90
$script:Grid.Columns[2].Name  = "Result";    $script:Grid.Columns[2].AutoSizeMode = "Fill"
$script:Grid.Columns[3].Name  = "Duration";  $script:Grid.Columns[3].Width = 80
$script:Grid.Columns[4].Name  = "Timestamp"; $script:Grid.Columns[4].Width = 145
$grpGrid.Controls.Add($script:Grid)

# ══════════════════════════════════════════════════════════════════════════════
#  RIGHT COLUMN  (Status Log + Service Manager)  X=1260
# ══════════════════════════════════════════════════════════════════════════════

# ── Status / Log ───────────────────────────────────────────────────────────────
$grpStatus = New-StyledGroup "  Activity Log" 1260 58 220 580
$form.Controls.Add($grpStatus)

$script:txtStatus = New-Object System.Windows.Forms.RichTextBox
$script:txtStatus.Location  = New-Object System.Drawing.Point(6, 24)
$script:txtStatus.Size      = New-Object System.Drawing.Size(206, 510)
$script:txtStatus.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 32)
$script:txtStatus.ForeColor = $clr.Text
$script:txtStatus.Font      = $fntMonoSm
$script:txtStatus.ReadOnly  = $true
$script:txtStatus.ScrollBars = "Vertical"
$script:txtStatus.WordWrap  = $true
$script:txtStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$grpStatus.Controls.Add($script:txtStatus)

$btnClearLog = New-StyledButton "  Clear Log" 6 540 206 26 -FG $clr.Overlay
$grpStatus.Controls.Add($btnClearLog)

# ── Service Manager ────────────────────────────────────────────────────────────
$grpSvcMgr = New-StyledGroup "  Remote Service Manager" 1260 648 220 256
$form.Controls.Add($grpSvcMgr)

$grpSvcMgr.Controls.Add((New-Label "Service Name:" 6 28 110 20))
$script:txtSvcName = New-StyledTextBox 6 48 206 22
$grpSvcMgr.Controls.Add($script:txtSvcName)

$grpSvcMgr.Controls.Add((New-Label "Action:" 6 78 80 20))

$cmbSvcAction = New-Object System.Windows.Forms.ComboBox
$cmbSvcAction.Location  = New-Object System.Drawing.Point(6, 96)
$cmbSvcAction.Size      = New-Object System.Drawing.Size(206, 22)
$cmbSvcAction.BackColor = $clr.Surface1
$cmbSvcAction.ForeColor = $clr.Text
$cmbSvcAction.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cmbSvcAction.DropDownStyle = "DropDownList"
$cmbSvcAction.Items.AddRange(@("Start","Stop","Restart","Pause","Resume","Get Status"))
$cmbSvcAction.SelectedIndex = 0
$grpSvcMgr.Controls.Add($cmbSvcAction)

$btnSvcExecute = New-StyledButton "  Run Service Action" 6 128 206 30 -FG $clr.Peach
$grpSvcMgr.Controls.Add($btnSvcExecute)

$script:lblSvcStatus = New-Label "Ready" 6 168 206 40
$script:lblSvcStatus.ForeColor  = $clr.Subtext
$script:lblSvcStatus.Font       = $fntSmall
$script:lblSvcStatus.AutoSize   = $false
$grpSvcMgr.Controls.Add($script:lblSvcStatus)

# ── Status bar at bottom ────────────────────────────────────────────────────────
$pnlStatusBar = New-Object System.Windows.Forms.Panel
$pnlStatusBar.Location  = New-Object System.Drawing.Point(0, 912)
$pnlStatusBar.Size      = New-Object System.Drawing.Size(1500, 24)
$pnlStatusBar.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 37)
$pnlStatusBar.Anchor    = "Bottom,Left,Right"
$form.Controls.Add($pnlStatusBar)

$script:lblStatusBar = New-Object System.Windows.Forms.Label
$script:lblStatusBar.Text      = "  Ready  |  Log: $LogFile"
$script:lblStatusBar.Location  = New-Object System.Drawing.Point(0, 3)
$script:lblStatusBar.Size      = New-Object System.Drawing.Size(1500, 18)
$script:lblStatusBar.ForeColor = $clr.Overlay
$script:lblStatusBar.Font      = $fntSmall
$script:lblStatusBar.BackColor = [System.Drawing.Color]::Transparent
$pnlStatusBar.Controls.Add($script:lblStatusBar)

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Get-ServerList {
    $raw = $script:txtServers.Text -split "`r`n|`n|,"
    return ($raw | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
}

function Get-CredentialObject {
    if ($chkCurrent.Checked) { return $null }
    $domain = $cmbDomain.Text
    $user   = $script:txtUser.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($user))     { throw "Username is required." }
    if ([string]::IsNullOrWhiteSpace($script:txtPassword.Text)) { throw "Password is required." }
    $fullUser = if ($domain -ne "CUSTOM") { "$domain\$user" } else { $user }
    $secure   = ConvertTo-SecureString $script:txtPassword.Text -AsPlainText -Force
    return New-Object System.Management.Automation.PSCredential($fullUser, $secure)
}

function Set-StatusBar { param([string]$Text)
    $script:lblStatusBar.Text = "  $Text  |  Log: $LogFile"
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-GridRow {
    param([string]$Server, [string]$Status, [string]$Result, [string]$Duration)
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $idx  = $script:Grid.Rows.Add($Server, $Status, $Result, $Duration, $ts)
    $row  = $script:Grid.Rows[$idx]
    switch -Wildcard ($Status) {
        "Online"  { $row.DefaultCellStyle.BackColor = $clr.RowOnline  }
        "Success" { $row.DefaultCellStyle.BackColor = $clr.RowOnline  }
        "Offline" { $row.DefaultCellStyle.BackColor = $clr.RowOffline }
        "Failed"  { $row.DefaultCellStyle.BackColor = $clr.RowOffline }
        "Error"   { $row.DefaultCellStyle.BackColor = $clr.RowOffline }
        "Warn*"   { $row.DefaultCellStyle.BackColor = $clr.RowWarn    }
        "Partial" { $row.DefaultCellStyle.BackColor = $clr.RowWarn    }
    }
    $script:Grid.FirstDisplayedScrollingRowIndex = $idx
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-RemoteScript {
    param([string]$Server, [scriptblock]$ScriptBlock,
          [System.Management.Automation.PSCredential]$Credential)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $params = @{
            ComputerName = $Server
            ScriptBlock  = $ScriptBlock
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $params.Credential = $Credential }
        $result = Invoke-Command @params
        $sw.Stop()
        $out = ($result | Out-String).Trim()
        if ($out.Length -gt 300) { $out = $out.Substring(0,300) + "  [truncated]" }
        return @{ Status = "Success"; Result = $out; Duration = "$($sw.Elapsed.TotalSeconds.ToString('0.0'))s" }
    }
    catch {
        $sw.Stop()
        return @{ Status = "Failed"; Result = $_.Exception.Message; Duration = "$($sw.Elapsed.TotalSeconds.ToString('0.0'))s" }
    }
}

function Start-BulkOperation {
    param([string]$OperationName, [scriptblock]$PerServer)
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No servers in the list.", "No Servers", 'OK', 'Warning') | Out-Null
        return
    }
    $script:Grid.Rows.Clear()
    $script:ProgressBar.Maximum = $servers.Count
    $script:ProgressBar.Value   = 0
    Write-Log "[$OperationName] Starting on $($servers.Count) server(s)"
    Set-StatusBar "$OperationName  –  0 / $($servers.Count)"
    $i = 0
    foreach ($srv in $servers) {
        $i++
        $script:lblProgress.Text = "$i / $($servers.Count)"
        Set-StatusBar "$OperationName  –  $i / $($servers.Count)  ($srv)"
        & $PerServer $srv
        $script:ProgressBar.Value = $i
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Log "[$OperationName] Completed"
    Set-StatusBar "$OperationName  –  Done  ($($servers.Count) server(s))"
    $script:lblProgress.Text = "Done"
}

# ══════════════════════════════════════════════════════════════════════════════
#  SNIPPET LIBRARY EVENTS
# ══════════════════════════════════════════════════════════════════════════════
$cmbSnippetCat.Add_SelectedIndexChanged({
    $cat = $cmbSnippetCat.SelectedItem.ToString()
    $cmbSnippet.Items.Clear()
    foreach ($key in ($script:Snippets[$cat].Keys | Sort-Object)) {
        $cmbSnippet.Items.Add($key) | Out-Null
    }
    if ($cmbSnippet.Items.Count -gt 0) { $cmbSnippet.SelectedIndex = 0 }
})

$btnLoadSnippet.Add_Click({
    $cat  = $cmbSnippetCat.SelectedItem
    $name = $cmbSnippet.SelectedItem
    if (-not $cat -or -not $name) { return }
    $code = $script:Snippets[$cat.ToString()][$name.ToString()]
    if ($code) {
        $script:txtScript.Text = $code
        Write-Log "Loaded snippet: $cat > $name"
    }
})

# Trigger initial category population
$cmbSnippetCat.SelectedIndex = 0

# ══════════════════════════════════════════════════════════════════════════════
#  BUTTON EVENTS
# ══════════════════════════════════════════════════════════════════════════════

# ── Import server list from file ────────────────────────────────────────────────
$btnImport.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = "Import Server List"
    $ofd.Filter = "Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Get-Content $ofd.FileName | ForEach-Object { $_.Trim() } |
                     Where-Object { $_ -ne "" -and $_ -notmatch "^#" }
            # Handle CSV with a header column called "Server" or "ComputerName"
            if ($ofd.FileName -like "*.csv") {
                $csv = Import-Csv $ofd.FileName -ErrorAction SilentlyContinue
                $col = ($csv[0].PSObject.Properties.Name | Where-Object { $_ -match "Server|Computer|Name" } | Select-Object -First 1)
                if ($col) { $lines = $csv.$col | Where-Object { $_ -ne "" } }
            }
            $script:txtServers.Text = $lines -join "`r`n"
            Write-Log "Imported $($lines.Count) servers from $($ofd.FileName)"
        }
        catch { Write-Log "Import failed: $($_.Exception.Message)" 'ERROR' }
    }
})

# ── Clear server list ──────────────────────────────────────────────────────────
$btnClearSrv.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Clear the server list?","Confirm",'YesNo','Question') -eq 'Yes') {
        $script:txtServers.Clear()
    }
})

# ── Clear results grid ─────────────────────────────────────────────────────────
$btnClearGrid.Add_Click({
    $script:Grid.Rows.Clear()
    $script:ProgressBar.Value  = 0
    $script:lblProgress.Text   = "Idle"
    Write-Log "Results cleared"
})

# ── Clear log ──────────────────────────────────────────────────────────────────
$btnClearLog.Add_Click({
    $script:txtStatus.Clear()
})

# ── Export results to CSV ──────────────────────────────────────────────────────
$btnExportCsv.Add_Click({
    if ($script:Grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export", 'OK', 'Warning') | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title      = "Export Results"
    $sfd.Filter     = "CSV files (*.csv)|*.csv"
    $sfd.FileName   = "DeploymentResults_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    $sfd.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $rows = foreach ($row in $script:Grid.Rows) {
                [PSCustomObject]@{
                    Server    = $row.Cells[0].Value
                    Status    = $row.Cells[1].Value
                    Result    = $row.Cells[2].Value
                    Duration  = $row.Cells[3].Value
                    Timestamp = $row.Cells[4].Value
                }
            }
            $rows | Export-Csv $sfd.FileName -NoTypeInformation -Force
            Write-Log "Exported $($rows.Count) rows to $($sfd.FileName)" 'SUCCESS'
            Set-StatusBar "Exported to $($sfd.FileName)"
        }
        catch { Write-Log "Export failed: $($_.Exception.Message)" 'ERROR' }
    }
})

# ── Right-click grid context menu ──────────────────────────────────────────────
$ctxGrid = New-Object System.Windows.Forms.ContextMenuStrip
$ctxGrid.BackColor = $clr.Surface1
$ctxGrid.ForeColor = $clr.Text

$menuCopyRow = $ctxGrid.Items.Add("Copy Row to Clipboard")
$menuCopyRow.add_Click({
    if ($script:Grid.SelectedRows.Count -gt 0) {
        $row  = $script:Grid.SelectedRows[0]
        $text = "$($row.Cells[0].Value)`t$($row.Cells[1].Value)`t$($row.Cells[2].Value)`t$($row.Cells[3].Value)`t$($row.Cells[4].Value)"
        [System.Windows.Forms.Clipboard]::SetText($text)
        Write-Log "Row copied to clipboard"
    }
})

$menuCopyResult = $ctxGrid.Items.Add("Copy Result Cell")
$menuCopyResult.add_Click({
    if ($script:Grid.SelectedRows.Count -gt 0) {
        $val = $script:Grid.SelectedRows[0].Cells[2].Value
        if ($val) { [System.Windows.Forms.Clipboard]::SetText($val.ToString()) }
    }
})

$ctxGrid.Items.Add("-") | Out-Null

$menuRetry = $ctxGrid.Items.Add("Retry Selected Server")
$menuRetry.add_Click({
    if ($script:Grid.SelectedRows.Count -gt 0) {
        $srv = $script:Grid.SelectedRows[0].Cells[0].Value
        if ($srv) {
            $cred = try { Get-CredentialObject } catch { $null }
            $sb   = [ScriptBlock]::Create($script:txtScript.Text)
            $res  = Invoke-RemoteScript -Server $srv -ScriptBlock $sb -Credential $cred
            Add-GridRow $srv $res.Status $res.Result $res.Duration
            Write-Log "Retry $srv : $($res.Status)"
        }
    }
})

$script:Grid.ContextMenuStrip = $ctxGrid

# ── Verify Credentials ─────────────────────────────────────────────────────────
$btnTestCreds.Add_Click({
    if ($chkCurrent.Checked) {
        Write-Log "Using current credentials: $env:USERDOMAIN\$env:USERNAME" 'INFO'
        return
    }
    try {
        $cred = Get-CredentialObject
        Write-Log "Credential object built for: $($cred.UserName)" 'SUCCESS'
        Set-StatusBar "Credentials validated for $($cred.UserName)"
    }
    catch {
        Write-Log "Credential error: $($_.Exception.Message)" 'ERROR'
    }
})

# ── Test Connectivity (Ping) ───────────────────────────────────────────────────
$script:quickBtns[0].Add_Click({
    Start-BulkOperation "Ping Test" {
        param($srv)
        try {
            $ping = Test-Connection $srv -Count 2 -Quiet -ErrorAction Stop
            $status = if ($ping) { "Online" } else { "Offline" }
            Add-GridRow $srv $status (if ($ping) { "Ping success" } else { "No response" }) "-"
            Write-Log "$srv : $status"
        }
        catch { Add-GridRow $srv "Error" $_.Exception.Message "-"; Write-Log "$srv : $($_.Exception.Message)" 'ERROR' }
    }
})

# ── Test WinRM ─────────────────────────────────────────────────────────────────
$script:quickBtns[1].Add_Click({
    Start-BulkOperation "WinRM Test" {
        param($srv)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $cred   = try { Get-CredentialObject } catch { $null }
            $params = @{ ComputerName = $srv; ScriptBlock = { $env:COMPUTERNAME }; ErrorAction = 'Stop' }
            if ($cred) { $params.Credential = $cred }
            $name   = Invoke-Command @params
            $sw.Stop()
            Add-GridRow $srv "Online" "WinRM OK  |  Hostname: $name" "$($sw.Elapsed.TotalSeconds.ToString('0.0'))s"
            Write-Log "$srv WinRM OK ($name)" 'SUCCESS'
        }
        catch {
            $sw.Stop()
            Add-GridRow $srv "Failed" "WinRM: $($_.Exception.Message)" "$($sw.Elapsed.TotalSeconds.ToString('0.0'))s"
            Write-Log "$srv WinRM FAIL: $($_.Exception.Message)" 'ERROR'
        }
    }
})

# ── Get System Info ────────────────────────────────────────────────────────────
$script:quickBtns[2].Add_Click({
    Start-BulkOperation "System Info" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            $os  = Get-CimInstance Win32_OperatingSystem
            $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
            $up  = (Get-Date) - $os.LastBootUpTime
            "OS: {0} | Build: {1} | CPU: {2} | RAM: {3}GB Free/{4}GB Total | Uptime: {5}d {6}h" -f `
                $os.Caption, $os.BuildNumber, $cpu.Name,
                [Math]::Round($os.FreePhysicalMemory/1MB,1),
                [Math]::Round($os.TotalVisibleMemorySize/1MB,1),
                $up.Days, $up.Hours
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        Add-GridRow $srv $res.Status $res.Result $res.Duration
        Write-Log "$srv System Info: $($res.Status)"
    }
})

# ── Get Hotfixes ───────────────────────────────────────────────────────────────
$script:quickBtns[3].Add_Click({
    Start-BulkOperation "Hotfix Report" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            $hf = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 5
            ($hf | ForEach-Object { "$($_.HotFixID) – $($_.InstalledOn.ToString('yyyy-MM-dd'))" }) -join " | "
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        Add-GridRow $srv $res.Status "Last 5 patches: $($res.Result)" $res.Duration
        Write-Log "$srv Hotfixes: $($res.Status)"
    }
})

# ── List Services ──────────────────────────────────────────────────────────────
$script:quickBtns[4].Add_Click({
    Start-BulkOperation "Service List" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            $stopped = Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -eq 'Stopped' }
            $total   = (Get-Service).Count
            "Total: $total services  |  Auto/Stopped: $($stopped.Count)  |  $(($stopped | Select-Object -First 3 -ExpandProperty Name) -join ', ')$(if($stopped.Count -gt 3){' ...'})"
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        $st  = if ($res.Status -eq "Success" -and $res.Result -match "Auto/Stopped: 0") { "Success" } elseif ($res.Status -eq "Success") { "Warn" } else { "Failed" }
        Add-GridRow $srv $st $res.Result $res.Duration
        Write-Log "$srv Services: $($res.Status)"
    }
})

# ── List Scheduled Tasks ───────────────────────────────────────────────────────
$script:quickBtns[5].Add_Click({
    Start-BulkOperation "Scheduled Tasks" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            $tasks   = Get-ScheduledTask
            $running = ($tasks | Where-Object { $_.State -eq 'Running' }).Count
            $ready   = ($tasks | Where-Object { $_.State -eq 'Ready'   }).Count
            $disabled= ($tasks | Where-Object { $_.State -eq 'Disabled'}).Count
            "Total: $($tasks.Count)  |  Running: $running  |  Ready: $ready  |  Disabled: $disabled"
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        Add-GridRow $srv $res.Status $res.Result $res.Duration
        Write-Log "$srv Scheduled Tasks: $($res.Status)"
    }
})

# ── Event Log Errors (24h) ─────────────────────────────────────────────────────
$script:quickBtns[6].Add_Click({
    Start-BulkOperation "Event Log Errors" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            $start  = (Get-Date).AddHours(-24)
            $sysErr = @(Get-EventLog -LogName System      -EntryType Error -After $start -ErrorAction SilentlyContinue).Count
            $appErr = @(Get-EventLog -LogName Application -EntryType Error -After $start -ErrorAction SilentlyContinue).Count
            "System errors: $sysErr  |  Application errors: $appErr  (last 24h)"
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        $st  = if ($res.Status -ne "Success") { "Failed" } elseif ($res.Result -match "errors: 0\s+\|.*errors: 0") { "Success" } else { "Warn" }
        Add-GridRow $srv $st $res.Result $res.Duration
        Write-Log "$srv Events: $($res.Status)"
    }
})

# ── Get Disk Space ─────────────────────────────────────────────────────────────
$script:quickBtns[7].Add_Click({
    Start-BulkOperation "Disk Space" {
        param($srv)
        $cred = try { Get-CredentialObject } catch { $null }
        $sb   = {
            Get-PSDrive -PSProvider FileSystem |
            ForEach-Object {
                $total = $_.Used + $_.Free
                if ($total -gt 0) {
                    $pct = [Math]::Round($_.Used / $total * 100, 0)
                    "$($_.Name): $([Math]::Round($_.Free/1GB,1))GB free / $([Math]::Round($total/1GB,1))GB  ($pct% used)"
                }
            }
        }
        $res = Invoke-RemoteScript $srv $sb $cred
        $warnThreshold = $res.Result -match "9[0-9]% used"
        $st  = if ($res.Status -ne "Success") { "Failed" } elseif ($warnThreshold) { "Warn" } else { "Success" }
        Add-GridRow $srv $st $res.Result $res.Duration
        Write-Log "$srv Disk: $($res.Status)"
    }
})

# ── Execute PowerShell (sequential) ───────────────────────────────────────────
$btnExecute.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:txtScript.Text)) {
        [System.Windows.Forms.MessageBox]::Show("No script to execute.", "Execute", 'OK', 'Warning') | Out-Null
        return
    }
    $cred = try { Get-CredentialObject } catch {
        Write-Log "Credential error: $($_.Exception.Message)" 'ERROR'
        return
    }
    $sb = [ScriptBlock]::Create($script:txtScript.Text)
    Start-BulkOperation "Execute PS" {
        param($srv)
        Write-Log "Executing on $srv"
        $res = Invoke-RemoteScript $srv $sb $cred
        Add-GridRow $srv $res.Status $res.Result $res.Duration
        Write-Log "$srv : $($res.Status)"
    }
})

# ── Execute (Background Jobs) ──────────────────────────────────────────────────
$btnExecAsync.Add_Click({
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No servers in the list.", "Execute", 'OK', 'Warning') | Out-Null
        return
    }
    $cred   = try { Get-CredentialObject } catch { $null }
    $sb     = [ScriptBlock]::Create($script:txtScript.Text)
    $script:Grid.Rows.Clear()
    $script:ProgressBar.Maximum = $servers.Count
    $script:ProgressBar.Value   = 0

    Write-Log "Launching $($servers.Count) background jobs"
    Set-StatusBar "Launching background jobs..."

    $jobs = @{}
    foreach ($srv in $servers) {
        $params = @{ ComputerName = $srv; ScriptBlock = $sb; AsJob = $true }
        if ($cred) { $params.Credential = $cred }
        try {
            $j = Invoke-Command @params
            $jobs[$srv] = $j
            Write-Log "Job started on $srv  (JobId: $($j.Id))"
        }
        catch { Write-Log "Could not start job on $srv : $($_.Exception.Message)" 'ERROR' }
    }

    Write-Log "Waiting for $($jobs.Count) jobs..."
    Set-StatusBar "Waiting for background jobs..."
    $done = 0
    foreach ($kv in $jobs.GetEnumerator()) {
        $srv = $kv.Key
        $j   = $kv.Value
        try {
            $j | Wait-Job -Timeout 120 | Out-Null
            if ($j.State -eq 'Completed') {
                $result = Receive-Job $j | Out-String
                $out = $result.Trim()
                if ($out.Length -gt 300) { $out = $out.Substring(0,300) + "  [truncated]" }
                Add-GridRow $srv "Success" $out "async"
                Write-Log "$srv Job completed" 'SUCCESS'
            }
            else {
                Add-GridRow $srv "Partial" "Job state: $($j.State)" "async"
                Write-Log "$srv Job state: $($j.State)" 'WARN'
            }
        }
        catch { Add-GridRow $srv "Failed" $_.Exception.Message "async"; Write-Log "$srv Job error: $($_.Exception.Message)" 'ERROR' }
        finally { Remove-Job $j -Force -ErrorAction SilentlyContinue }
        $done++
        $script:ProgressBar.Value = $done
        $script:lblProgress.Text  = "$done / $($jobs.Count)"
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Log "All background jobs complete" 'SUCCESS'
    Set-StatusBar "Background jobs complete – $($jobs.Count) server(s)"
})

# ── Stop all background jobs ───────────────────────────────────────────────────
$btnStopJobs.Add_Click({
    $jobs = Get-Job -ErrorAction SilentlyContinue
    if ($jobs) {
        $jobs | Stop-Job -ErrorAction SilentlyContinue
        $jobs | Remove-Job -Force -ErrorAction SilentlyContinue
        Write-Log "Stopped and removed $($jobs.Count) background job(s)" 'WARN'
        Set-StatusBar "All jobs stopped"
    }
    else {
        Write-Log "No active jobs found" 'INFO'
    }
})

# ── Remote Service Manager ─────────────────────────────────────────────────────
$btnSvcExecute.Add_Click({
    $svcName = $script:txtSvcName.Text.Trim()
    $action  = $cmbSvcAction.SelectedItem.ToString()
    $servers = Get-ServerList

    if ([string]::IsNullOrWhiteSpace($svcName)) {
        $script:lblSvcStatus.ForeColor = $clr.Red
        $script:lblSvcStatus.Text = "Enter a service name."
        return
    }
    if ($servers.Count -eq 0) {
        $script:lblSvcStatus.ForeColor = $clr.Yellow
        $script:lblSvcStatus.Text = "No servers listed."
        return
    }

    $cred = try { Get-CredentialObject } catch { $null }
    $script:lblSvcStatus.ForeColor = $clr.Subtext
    $script:lblSvcStatus.Text = "Running '$action' on $($servers.Count) server(s)…"

    $sb = switch ($action) {
        "Start"      { [ScriptBlock]::Create("Start-Service   -Name '$svcName' -ErrorAction Stop; (Get-Service '$svcName').Status") }
        "Stop"       { [ScriptBlock]::Create("Stop-Service    -Name '$svcName' -Force -ErrorAction Stop; (Get-Service '$svcName').Status") }
        "Restart"    { [ScriptBlock]::Create("Restart-Service -Name '$svcName' -Force -ErrorAction Stop; (Get-Service '$svcName').Status") }
        "Pause"      { [ScriptBlock]::Create("Suspend-Service -Name '$svcName' -ErrorAction Stop; (Get-Service '$svcName').Status") }
        "Resume"     { [ScriptBlock]::Create("Resume-Service  -Name '$svcName' -ErrorAction Stop; (Get-Service '$svcName').Status") }
        "Get Status" { [ScriptBlock]::Create("Get-Service -Name '$svcName' | Select-Object Name,Status,StartType | Format-Table -AutoSize | Out-String") }
    }

    $ok = 0; $fail = 0
    foreach ($srv in $servers) {
        $res = Invoke-RemoteScript $srv $sb $cred
        Add-GridRow $srv $res.Status "[$svcName] {$action}: $($res.Result)" $res.Duration
        Write-Log "$srv Service '$svcName' $action : $($res.Status)"
        if ($res.Status -eq "Success") { $ok++ } else { $fail++ }
    }
    $script:lblSvcStatus.ForeColor = if ($fail -eq 0) { $clr.Green } else { $clr.Yellow }
    $script:lblSvcStatus.Text = "Done – OK: $ok  |  Failed: $fail"
})

# ══════════════════════════════════════════════════════════════════════════════
#  FORM EVENTS
# ══════════════════════════════════════════════════════════════════════════════
$form.Add_Shown({ $form.Activate() })

$form.Add_FormClosing({
    Get-Job -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue
})

# ══════════════════════════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════════════════════════

# Trigger snippet category load
$cmbSnippetCat.SelectedIndex = 0

Write-Log "Enterprise Server Deployment Manager v4 started  |  User: $env:USERDOMAIN\$env:USERNAME" 'INFO'
Set-StatusBar "Ready"

[void]$form.ShowDialog()
