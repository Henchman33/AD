#Requires -Version 5.1
<#
.SYNOPSIS
    Add Bulk AD Server Access Tool
    GUI-based tool for adding users and managed service accounts to local groups
    across multiple servers in multi-domain environments.

.DESCRIPTION
    Features:
      - Add multiple AD users to multiple servers in one operation
      - Add Managed Service Accounts (MSAs/gMSAs) to servers
      - Paste-in server lists or import from CSV
      - Select target local group (Administrators, RDP Users, etc.)
      - Credential profiles for multiple domains
      - CSV import/export of user/server lists
      - Real-time progress bar
      - Error grid view
      - Timestamped log files saved to Desktop

.NOTES
    Author  : Steve (Sysadmin II)
    Version : 2.0
    Requires: ActiveDirectory module (RSAT), WinForms, PowerShell 5.1+
    Run in  : PowerShell ISE (or standard PS console)
#>

# ─────────────────────────────────────────────
# BOOTSTRAP – ensure we are STA for WinForms
# ─────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning "Re-launching in STA mode for WinForms compatibility..."
    $args0 = '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe -ArgumentList $args0
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ─────────────────────────────────────────────
# LOGGING SETUP
# ─────────────────────────────────────────────
$LogRoot    = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Add Bulk AD Server Access Tool\Logs'
if (-not (Test-Path $LogRoot)) { New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null }
$LogFile    = Join-Path $LogRoot ("log_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$ErrorRows  = [System.Collections.Generic.List[PSObject]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    # Also echo to PS console for ISE Output pane
    switch ($Level) {
        'SUCCESS' { Write-Host $line -ForegroundColor Green  }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red    }
        default   { Write-Host $line -ForegroundColor Cyan   }
    }
}

Write-Log "Tool started. Log file: $LogFile"

# ─────────────────────────────────────────────
# CREDENTIAL PROFILES STORE  (in-memory)
# ─────────────────────────────────────────────
$CredProfiles = [System.Collections.Generic.Dictionary[string,PSCredential]]::new()

# ─────────────────────────────────────────────
# COLORS / FONTS  (dark industrial theme)
# ─────────────────────────────────────────────
$clrBg        = [System.Drawing.Color]::FromArgb(28, 28, 36)
$clrPanel     = [System.Drawing.Color]::FromArgb(38, 38, 50)
$clrAccent    = [System.Drawing.Color]::FromArgb(0, 150, 215)
$clrAccentHov = [System.Drawing.Color]::FromArgb(0, 180, 255)
$clrText      = [System.Drawing.Color]::FromArgb(220, 220, 230)
$clrMuted     = [System.Drawing.Color]::FromArgb(120, 120, 140)
$clrSuccess   = [System.Drawing.Color]::FromArgb(50, 200, 100)
$clrWarn      = [System.Drawing.Color]::FromArgb(255, 190, 0)
$clrError     = [System.Drawing.Color]::FromArgb(220, 60, 60)
$clrBorder    = [System.Drawing.Color]::FromArgb(60, 60, 80)

$fntTitle  = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$fntLabel  = New-Object System.Drawing.Font('Segoe UI', 9,  [System.Drawing.FontStyle]::Bold)
$fntNormal = New-Object System.Drawing.Font('Segoe UI', 9)
$fntMono   = New-Object System.Drawing.Font('Consolas', 9)
$fntSmall  = New-Object System.Drawing.Font('Segoe UI', 8)

# ─────────────────────────────────────────────
# HELPER: styled controls
# ─────────────────────────────────────────────
function New-StyledButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$W=130, [int]$H=28)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text; $b.Location = [System.Drawing.Point]::new($X,$Y)
    $b.Size      = [System.Drawing.Size]::new($W,$H)
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor    = $clrAccent
    $b.FlatAppearance.BorderSize     = 1
    $b.FlatAppearance.MouseOverBackColor = $clrAccentHov
    $b.BackColor = $clrPanel; $b.ForeColor = $clrText; $b.Font = $fntLabel
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function New-StyledTextBox {
    param([int]$X,[int]$Y,[int]$W,[int]$H,[bool]$Multi=$false,[bool]$Password=$false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = [System.Drawing.Point]::new($X,$Y)
    $t.Size      = [System.Drawing.Size]::new($W,$H)
    $t.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
    $t.ForeColor = $clrText; $t.Font = $fntMono
    $t.BorderStyle = 'FixedSingle'
    if ($Multi)    { $t.Multiline = $true; $t.ScrollBars = 'Vertical' }
    if ($Password) { $t.PasswordChar = [char]0x2022 }
    return $t
}

function New-StyledLabel {
    param([string]$Text,[int]$X,[int]$Y,[int]$W=200,[int]$H=18,[System.Drawing.Font]$Font=$fntLabel)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = [System.Drawing.Point]::new($X,$Y)
    $l.Size = [System.Drawing.Size]::new($W,$H)
    $l.ForeColor = $clrText; $l.BackColor = [System.Drawing.Color]::Transparent
    $l.Font = $Font
    return $l
}

function New-GroupBox {
    param([string]$Text,[int]$X,[int]$Y,[int]$W,[int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text; $g.Location = [System.Drawing.Point]::new($X,$Y)
    $g.Size = [System.Drawing.Size]::new($W,$H)
    $g.ForeColor = $clrAccent; $g.BackColor = $clrPanel; $g.Font = $fntLabel
    return $g
}

# ─────────────────────────────────────────────
# MAIN FORM
# ─────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Add Bulk AD Server Access Tool  v2.0"
$form.Size            = [System.Drawing.Size]::new(1100, 820)
$form.MinimumSize     = [System.Drawing.Size]::new(1100, 820)
$form.BackColor       = $clrBg
$form.ForeColor       = $clrText
$form.Font            = $fntNormal
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'

# ── Title bar strip ──────────────────────────
$pnlTitle = New-Object System.Windows.Forms.Panel
$pnlTitle.Dock      = 'Top'
$pnlTitle.Height    = 48
$pnlTitle.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$lblTitle = New-StyledLabel "  🖥  Add Bulk AD Server Access Tool" 0 12 600 26 $fntTitle
$lblTitle.ForeColor = $clrAccent
$lblLogPath = New-StyledLabel "" 10 34 900 14 $fntSmall
$lblLogPath.ForeColor = $clrMuted
$lblLogPath.Text = "  Log: $LogFile"
$pnlTitle.Controls.AddRange(@($lblTitle,$lblLogPath))
$form.Controls.Add($pnlTitle)

# ── TabControl ───────────────────────────────
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location  = [System.Drawing.Point]::new(8, 56)
$tabs.Size      = [System.Drawing.Size]::new(1074, 700)
$tabs.BackColor = $clrBg
$tabs.Font      = $fntLabel

function New-Tab([string]$text) {
    $tp = New-Object System.Windows.Forms.TabPage
    $tp.Text      = $text
    $tp.BackColor = $clrBg
    $tp.ForeColor = $clrText
    return $tp
}

$tabMain   = New-Tab "  Main Operation  "
$tabCreds  = New-Tab "  Credential Profiles  "
$tabMSA    = New-Tab "  Managed Service Accounts  "
$tabLog    = New-Tab "  Live Log  "
$tabErrors = New-Tab "  Error Grid  "
$tabs.TabPages.AddRange(@($tabMain,$tabCreds,$tabMSA,$tabLog,$tabErrors))
$form.Controls.Add($tabs)

# ═══════════════════════════════════════════════════════════════
#  TAB 1 – MAIN OPERATION
# ═══════════════════════════════════════════════════════════════

# ── LEFT: Servers ──────────────────────────
$gbServers = New-GroupBox "Target Servers" 6 6 340 380
$tabMain.Controls.Add($gbServers)

$lblSrvHint = New-StyledLabel "Paste one server per line (or use Import):" 8 20 310 16 $fntSmall
$lblSrvHint.ForeColor = $clrMuted
$gbServers.Controls.Add($lblSrvHint)

$txtServers = New-StyledTextBox 8 40 322 290 $true
$txtServers.Font = $fntMono
$txtServers.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$gbServers.Controls.Add($txtServers)

$btnImportSrv = New-StyledButton "📂 Import CSV" 8 338 120 26
$btnExportSrv = New-StyledButton "💾 Export CSV" 136 338 120 26
$gbServers.Controls.AddRange(@($btnImportSrv,$btnExportSrv))

# ── CENTER: Users ───────────────────────────
$gbUsers = New-GroupBox "Users to Add  (DOMAIN\username)" 355 6 340 380
$tabMain.Controls.Add($gbUsers)

$lblUsrHint = New-StyledLabel "Paste one DOMAIN\username per line (or Import):" 8 20 320 16 $fntSmall
$lblUsrHint.ForeColor = $clrMuted
$gbUsers.Controls.Add($lblUsrHint)

$txtUsers = New-StyledTextBox 8 40 322 290 $true
$txtUsers.Font = $fntMono
$txtUsers.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$gbUsers.Controls.Add($txtUsers)

$btnImportUsr = New-StyledButton "📂 Import CSV" 8 338 120 26
$btnExportUsr = New-StyledButton "💾 Export CSV" 136 338 120 26
$gbUsers.Controls.AddRange(@($btnImportUsr,$btnExportUsr))

# ── RIGHT: Options ──────────────────────────
$gbOptions = New-GroupBox "Options" 704 6 360 380
$tabMain.Controls.Add($gbOptions)

# Group dropdown
$gbOptions.Controls.Add((New-StyledLabel "Target Local Group:" 10 22 200 18))
$cmbGroup = New-Object System.Windows.Forms.ComboBox
$cmbGroup.Location = [System.Drawing.Point]::new(10,42)
$cmbGroup.Size     = [System.Drawing.Size]::new(330,24)
$cmbGroup.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
$cmbGroup.ForeColor = $clrText
$cmbGroup.Font      = $fntMono
$cmbGroup.DropDownStyle = 'DropDown'
@('Administrators','Remote Desktop Users','Remote Management Users',
  'Backup Operators','Performance Monitor Users','Network Configuration Operators',
  'Event Log Readers','Distributed COM Users') | ForEach-Object { [void]$cmbGroup.Items.Add($_) }
$cmbGroup.SelectedIndex = 0
$gbOptions.Controls.Add($cmbGroup)

$gbOptions.Controls.Add((New-StyledLabel "  ↑ Or type a custom group name" 10 68 300 14 $fntSmall))

# Credential profile
$gbOptions.Controls.Add((New-StyledLabel "Credential Profile (for target servers):" 10 94 300 18))
$cmbCredProf = New-Object System.Windows.Forms.ComboBox
$cmbCredProf.Location = [System.Drawing.Point]::new(10,114)
$cmbCredProf.Size     = [System.Drawing.Size]::new(250,24)
$cmbCredProf.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
$cmbCredProf.ForeColor = $clrText; $cmbCredProf.Font = $fntMono
$cmbCredProf.DropDownStyle = 'DropDownList'
[void]$cmbCredProf.Items.Add("[ Current Windows Session ]")
$cmbCredProf.SelectedIndex = 0
$gbOptions.Controls.Add($cmbCredProf)

$lblCredNote = New-StyledLabel "  ↑ Add profiles on the Credential Profiles tab" 10 140 330 14 $fntSmall
$lblCredNote.ForeColor = $clrMuted
$gbOptions.Controls.Add($lblCredNote)

# Checkboxes
$chkPing = New-Object System.Windows.Forms.CheckBox
$chkPing.Text = "Test connectivity (ping) before connecting"
$chkPing.Location = [System.Drawing.Point]::new(10,168)
$chkPing.Size     = [System.Drawing.Size]::new(330,20)
$chkPing.Checked  = $true
$chkPing.ForeColor = $clrText; $chkPing.BackColor = [System.Drawing.Color]::Transparent
$chkPing.Font = $fntNormal
$gbOptions.Controls.Add($chkPing)

$chkVerify = New-Object System.Windows.Forms.CheckBox
$chkVerify.Text = "Verify membership after add"
$chkVerify.Location = [System.Drawing.Point]::new(10,192)
$chkVerify.Size     = [System.Drawing.Size]::new(330,20)
$chkVerify.Checked  = $true
$chkVerify.ForeColor = $clrText; $chkVerify.BackColor = [System.Drawing.Color]::Transparent
$chkVerify.Font = $fntNormal
$gbOptions.Controls.Add($chkVerify)

$chkSkipExisting = New-Object System.Windows.Forms.CheckBox
$chkSkipExisting.Text = "Skip (don't error) if user already a member"
$chkSkipExisting.Location = [System.Drawing.Point]::new(10,216)
$chkSkipExisting.Size     = [System.Drawing.Size]::new(330,20)
$chkSkipExisting.Checked  = $true
$chkSkipExisting.ForeColor = $clrText; $chkSkipExisting.BackColor = [System.Drawing.Color]::Transparent
$chkSkipExisting.Font = $fntNormal
$gbOptions.Controls.Add($chkSkipExisting)

# Separator
$lblSep = New-Object System.Windows.Forms.Label
$lblSep.Location  = [System.Drawing.Point]::new(8, 246)
$lblSep.Size      = [System.Drawing.Size]::new(334, 1)
$lblSep.BackColor = $clrBorder
$gbOptions.Controls.Add($lblSep)

# Summary counts
$lblSummary = New-StyledLabel "Servers: 0   |   Users: 0" 10 255 330 16 $fntSmall
$lblSummary.ForeColor = $clrMuted
$gbOptions.Controls.Add($lblSummary)

# Refresh summary on text change
$updateSummary = {
    $sc = ($txtServers.Lines | Where-Object { $_.Trim() -ne '' }).Count
    $uc = ($txtUsers.Lines   | Where-Object { $_.Trim() -ne '' }).Count
    $lblSummary.Text = "Servers: $sc   |   Users: $uc   |   Operations: $($sc * $uc)"
}
$txtServers.Add_TextChanged($updateSummary)
$txtUsers.Add_TextChanged($updateSummary)

# ── PROGRESS ───────────────────────────────
$gbProgress = New-GroupBox "Progress" 6 394 1058 80
$tabMain.Controls.Add($gbProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = [System.Drawing.Point]::new(10,22)
$progressBar.Size     = [System.Drawing.Size]::new(940,22)
$progressBar.Style    = 'Continuous'
$progressBar.ForeColor = $clrAccent
$gbProgress.Controls.Add($progressBar)

$lblProgressDetail = New-StyledLabel "Ready." 10 50 1000 18 $fntSmall
$lblProgressDetail.ForeColor = $clrMuted
$gbProgress.Controls.Add($lblProgressDetail)

# ── ACTION BUTTONS ─────────────────────────
$btnRun    = New-StyledButton "▶  RUN" 6 484 160 40
$btnRun.Font     = New-Object System.Drawing.Font('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
$btnRun.FlatAppearance.BorderColor = $clrSuccess
$btnRun.ForeColor = $clrSuccess

$btnClear  = New-StyledButton "🗑  Clear All" 176 484 130 40
$btnReport = New-StyledButton "📋  Error Grid" 316 484 130 40
$btnOpenLog = New-StyledButton "📄  Open Log" 456 484 130 40
$tabMain.Controls.AddRange(@($btnRun,$btnClear,$btnReport,$btnOpenLog))

# ═══════════════════════════════════════════════════════════════
#  TAB 2 – CREDENTIAL PROFILES
# ═══════════════════════════════════════════════════════════════
$gbCredList = New-GroupBox "Saved Credential Profiles" 6 6 400 620
$tabCreds.Controls.Add($gbCredList)

$lbCredProfiles = New-Object System.Windows.Forms.ListBox
$lbCredProfiles.Location  = [System.Drawing.Point]::new(8,22)
$lbCredProfiles.Size      = [System.Drawing.Size]::new(380,490)
$lbCredProfiles.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$lbCredProfiles.ForeColor = $clrText; $lbCredProfiles.Font = $fntMono
$lbCredProfiles.BorderStyle = 'FixedSingle'
$gbCredList.Controls.Add($lbCredProfiles)

$btnRemoveCred = New-StyledButton "🗑 Remove Selected" 8 522 200 26
$gbCredList.Controls.Add($btnRemoveCred)

$gbAddCred = New-GroupBox "Add / Update Credential Profile" 420 6 640 280
$tabCreds.Controls.Add($gbAddCred)

$gbAddCred.Controls.Add((New-StyledLabel "Profile Name  (e.g. CORP, DMZ, LAB):" 10 24 400 18))
$txtCredName = New-StyledTextBox 10 44 300 22
$gbAddCred.Controls.Add($txtCredName)

$gbAddCred.Controls.Add((New-StyledLabel "Domain & Username  (DOMAIN\username):" 10 76 400 18))
$txtCredUser = New-StyledTextBox 10 96 300 22
$txtCredUser.Text = $env:USERDOMAIN + '\'
$gbAddCred.Controls.Add($txtCredUser)

$gbAddCred.Controls.Add((New-StyledLabel "Password:" 10 128 200 18))
$txtCredPass = New-StyledTextBox 10 148 300 22 $false $true
$gbAddCred.Controls.Add($txtCredPass)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Show password"
$chkShowPass.Location = [System.Drawing.Point]::new(10,176)
$chkShowPass.Size     = [System.Drawing.Size]::new(200,20)
$chkShowPass.ForeColor = $clrText; $chkShowPass.BackColor = [System.Drawing.Color]::Transparent
$chkShowPass.Font = $fntNormal
$chkShowPass.Add_CheckedChanged({
    $txtCredPass.PasswordChar = if ($chkShowPass.Checked) { [char]0 } else { [char]0x2022 }
})
$gbAddCred.Controls.Add($chkShowPass)

$btnSaveCred = New-StyledButton "💾 Save Profile" 10 206 150 28
$gbAddCred.Controls.Add($btnSaveCred)

$lblCredStatus = New-StyledLabel "" 170 210 400 18 $fntSmall
$gbAddCred.Controls.Add($lblCredStatus)

$gbCredHelp = New-GroupBox "Usage Notes" 420 298 640 200
$tabCreds.Controls.Add($gbCredHelp)
$helpText  = "• Create one profile per domain/environment (e.g. CORP, DMZ, LAB).`r`n"
$helpText += "• Credentials are held in memory only — they are NOT written to disk.`r`n"
$helpText += "• Select a profile on the Main Operation tab before running.`r`n"
$helpText += "• Use 'Current Windows Session' if your logon already has rights.`r`n"
$helpText += "• For cross-domain, create a profile for each domain's admin account.`r`n"
$helpText += "• Passwords are masked; use the 'Show password' checkbox to verify."
$rtbCredHelp = New-Object System.Windows.Forms.RichTextBox
$rtbCredHelp.Location  = [System.Drawing.Point]::new(10,22)
$rtbCredHelp.Size      = [System.Drawing.Size]::new(618,160)
$rtbCredHelp.Text      = $helpText
$rtbCredHelp.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
$rtbCredHelp.ForeColor = $clrMuted; $rtbCredHelp.Font = $fntSmall
$rtbCredHelp.ReadOnly  = $true; $rtbCredHelp.BorderStyle = 'None'
$gbCredHelp.Controls.Add($rtbCredHelp)

# ═══════════════════════════════════════════════════════════════
#  TAB 3 – MANAGED SERVICE ACCOUNTS
# ═══════════════════════════════════════════════════════════════
$gbMSAServers = New-GroupBox "Target Servers for MSA" 6 6 340 540
$tabMSA.Controls.Add($gbMSAServers)

$txtMSAServers = New-StyledTextBox 8 22 322 460 $true
$txtMSAServers.Font = $fntMono
$txtMSAServers.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$gbMSAServers.Controls.Add($txtMSAServers)

$btnImportMSASrv = New-StyledButton "📂 Import CSV" 8 492 120 26
$gbMSAServers.Controls.Add($btnImportMSASrv)

$gbMSAAccounts = New-GroupBox "MSA / gMSA Accounts  (DOMAIN\account$)" 355 6 340 540
$tabMSA.Controls.Add($gbMSAAccounts)

$lblMSAHint = New-StyledLabel "gMSA names typically end with `$`:" 8 20 310 16 $fntSmall
$lblMSAHint.ForeColor = $clrMuted
$gbMSAAccounts.Controls.Add($lblMSAHint)

$txtMSAAccounts = New-StyledTextBox 8 40 322 440 $true
$txtMSAAccounts.Font = $fntMono
$txtMSAAccounts.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$gbMSAAccounts.Controls.Add($txtMSAAccounts)

$btnImportMSAAcc = New-StyledButton "📂 Import CSV" 8 492 120 26
$gbMSAAccounts.Controls.Add($btnImportMSAAcc)

$gbMSAOptions = New-GroupBox "MSA Options" 704 6 360 280
$tabMSA.Controls.Add($gbMSAOptions)

$gbMSAOptions.Controls.Add((New-StyledLabel "Target Local Group:" 10 22 200 18))
$cmbMSAGroup = New-Object System.Windows.Forms.ComboBox
$cmbMSAGroup.Location = [System.Drawing.Point]::new(10,42)
$cmbMSAGroup.Size     = [System.Drawing.Size]::new(330,24)
$cmbMSAGroup.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
$cmbMSAGroup.ForeColor = $clrText; $cmbMSAGroup.Font = $fntMono
$cmbMSAGroup.DropDownStyle = 'DropDown'
@('Administrators','Remote Desktop Users','Remote Management Users',
  'Backup Operators','Performance Monitor Users') | ForEach-Object { [void]$cmbMSAGroup.Items.Add($_) }
$cmbMSAGroup.SelectedIndex = 0
$gbMSAOptions.Controls.Add($cmbMSAGroup)

$gbMSAOptions.Controls.Add((New-StyledLabel "Credential Profile:" 10 76 200 18))
$cmbMSACredProf = New-Object System.Windows.Forms.ComboBox
$cmbMSACredProf.Location = [System.Drawing.Point]::new(10,96)
$cmbMSACredProf.Size     = [System.Drawing.Size]::new(330,24)
$cmbMSACredProf.BackColor = [System.Drawing.Color]::FromArgb(22,22,30)
$cmbMSACredProf.ForeColor = $clrText; $cmbMSACredProf.Font = $fntMono
$cmbMSACredProf.DropDownStyle = 'DropDownList'
[void]$cmbMSACredProf.Items.Add("[ Current Windows Session ]")
$cmbMSACredProf.SelectedIndex = 0
$gbMSAOptions.Controls.Add($cmbMSACredProf)

$gbMSAOptions.Controls.Add((New-StyledLabel "NOTE: gMSA must already be installed on" 10 130 330 16 $fntSmall))
$gbMSAOptions.Controls.Add((New-StyledLabel "each target server via Install-ADServiceAccount." 10 148 330 16 $fntSmall))

$btnRunMSA = New-StyledButton "▶  ADD MSAs" 704 300 160 40
$btnRunMSA.Font = New-Object System.Drawing.Font('Segoe UI',11,[System.Drawing.FontStyle]::Bold)
$btnRunMSA.FlatAppearance.BorderColor = $clrWarn
$btnRunMSA.ForeColor = $clrWarn
$tabMSA.Controls.Add($btnRunMSA)

$msaProgressBar = New-Object System.Windows.Forms.ProgressBar
$msaProgressBar.Location = [System.Drawing.Point]::new(6,560)
$msaProgressBar.Size     = [System.Drawing.Size]::new(1050,18)
$msaProgressBar.Style    = 'Continuous'
$tabMSA.Controls.Add($msaProgressBar)

$lblMSAStatus = New-StyledLabel "Ready." 6 582 900 16 $fntSmall
$lblMSAStatus.ForeColor = $clrMuted
$tabMSA.Controls.Add($lblMSAStatus)

# ═══════════════════════════════════════════════════════════════
#  TAB 4 – LIVE LOG
# ═══════════════════════════════════════════════════════════════
$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock       = 'Fill'
$rtbLog.BackColor  = [System.Drawing.Color]::FromArgb(10,10,16)
$rtbLog.ForeColor  = $clrText
$rtbLog.Font       = $fntMono
$rtbLog.ReadOnly   = $true
$rtbLog.ScrollBars = 'Vertical'
$rtbLog.WordWrap   = $false
$tabLog.Controls.Add($rtbLog)

$pnlLogBtns = New-Object System.Windows.Forms.Panel
$pnlLogBtns.Dock   = 'Bottom'; $pnlLogBtns.Height = 36
$pnlLogBtns.BackColor = $clrPanel
$btnClearLog  = New-StyledButton "🗑 Clear View" 4 4 120 26
$btnSaveLog   = New-StyledButton "💾 Save Log" 132 4 120 26
$btnOpenLogFolder = New-StyledButton "📂 Open Log Folder" 260 4 150 26
$pnlLogBtns.Controls.AddRange(@($btnClearLog,$btnSaveLog,$btnOpenLogFolder))
$tabLog.Controls.Add($pnlLogBtns)

# ═══════════════════════════════════════════════════════════════
#  TAB 5 – ERROR GRID
# ═══════════════════════════════════════════════════════════════
$dgvErrors = New-Object System.Windows.Forms.DataGridView
$dgvErrors.Dock              = 'Fill'
$dgvErrors.BackgroundColor   = [System.Drawing.Color]::FromArgb(18,18,26)
$dgvErrors.ForeColor         = $clrText
$dgvErrors.GridColor         = $clrBorder
$dgvErrors.Font              = $fntMono
$dgvErrors.BorderStyle       = 'None'
$dgvErrors.RowHeadersVisible = $false
$dgvErrors.AllowUserToAddRows = $false
$dgvErrors.ReadOnly          = $true
$dgvErrors.SelectionMode     = 'FullRowSelect'
$dgvErrors.AutoSizeColumnsMode = 'Fill'
$dgvErrors.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(28,28,42)
$dgvErrors.ColumnHeadersDefaultCellStyle.ForeColor = $clrAccent
$dgvErrors.ColumnHeadersDefaultCellStyle.Font      = $fntLabel
$dgvErrors.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(18,18,26)
$dgvErrors.DefaultCellStyle.ForeColor = $clrText
$dgvErrors.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(40,80,120)
$dgvErrors.EnableHeadersVisualStyles = $false

# Add columns
foreach ($col in @('Timestamp','Server','Account','Group','Status','Message')) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText = $col; $c.Name = $col
    [void]$dgvErrors.Columns.Add($c)
}
$tabErrors.Controls.Add($dgvErrors)

$pnlErrBtns = New-Object System.Windows.Forms.Panel
$pnlErrBtns.Dock = 'Bottom'; $pnlErrBtns.Height = 36; $pnlErrBtns.BackColor = $clrPanel
$btnExportErrors = New-StyledButton "💾 Export Errors CSV" 4 4 170 26
$btnClearErrors  = New-StyledButton "🗑 Clear Grid" 182 4 120 26
$pnlErrBtns.Controls.AddRange(@($btnExportErrors,$btnClearErrors))
$tabErrors.Controls.Add($pnlErrBtns)

# ─────────────────────────────────────────────
# HELPER: Add line to Live Log tab
# ─────────────────────────────────────────────
function Add-LogLine {
    param([string]$Text,[System.Drawing.Color]$Color=$clrText)
    if ($rtbLog.InvokeRequired) {
        $rtbLog.Invoke([Action[string,System.Drawing.Color]]{
            param($t,$c) 
            $rtbLog.SelectionStart  = $rtbLog.TextLength
            $rtbLog.SelectionLength = 0
            $rtbLog.SelectionColor  = $c
            $rtbLog.AppendText("$t`n")
            $rtbLog.ScrollToCaret()
        }, $Text, $Color)
    } else {
        $rtbLog.SelectionStart  = $rtbLog.TextLength
        $rtbLog.SelectionLength = 0
        $rtbLog.SelectionColor  = $Color
        $rtbLog.AppendText("$Text`n")
        $rtbLog.ScrollToCaret()
    }
}

function Write-UILog {
    param([string]$Message,[string]$Level='INFO')
    Write-Log $Message $Level
    $color = switch ($Level) {
        'SUCCESS' { $clrSuccess }
        'WARNING' { $clrWarn   }
        'ERROR'   { $clrError  }
        default   { $clrMuted  }
    }
    $ts = Get-Date -Format 'HH:mm:ss'
    Add-LogLine "[$ts][$Level] $Message" $color
}

# ─────────────────────────────────────────────
# HELPER: Add row to Error Grid
# ─────────────────────────────────────────────
function Add-ErrorRow {
    param([string]$Server,[string]$Account,[string]$Group,[string]$Status,[string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    [void]$dgvErrors.Rows.Add($ts,$Server,$Account,$Group,$Status,$Message)
    # Color row
    $row = $dgvErrors.Rows[$dgvErrors.Rows.Count - 1]
    $row.DefaultCellStyle.ForeColor = if ($Status -eq 'SUCCESS') { $clrSuccess } elseif ($Status -eq 'SKIPPED') { $clrWarn } else { $clrError }
    $ErrorRows.Add([PSCustomObject]@{
        Timestamp=$ts;Server=$Server;Account=$Account;Group=$Group;Status=$Status;Message=$Message
    })
}

# ─────────────────────────────────────────────
# CORE: Add user/account to a server's local group
# ─────────────────────────────────────────────
function Add-AccountToLocalGroup {
    param(
        [string]$Server,
        [string]$Account,
        [string]$Group,
        [bool]$SkipExisting,
        [bool]$Verify,
        [System.Management.Automation.PSCredential]$Credential
    )

    $scriptBlock = {
        param($acc,$grp,$skipEx,$verify)
        $result = @{ Status=''; Message=''; Verified=$false }

        try {
            # Parse domain\user or domain\account$
            if ($acc -match '^(.+)\\(.+)$') {
                $dom  = $matches[1]
                $user = $matches[2]
            } else {
                $dom  = $env:USERDOMAIN
                $user = $acc
            }

            $group = [ADSI]"WinNT://./$grp,group"
            $members = @($group.Invoke('Members') | ForEach-Object { $_.GetType().InvokeMember('Name','GetProperty',$null,$_,$null) })

            # Normalize user for comparison (strip trailing $)
            $userComp = $user.TrimEnd('$').ToLower()
            $alreadyMember = $members | Where-Object { $_.TrimEnd('$').ToLower() -eq $userComp }

            if ($alreadyMember -and $skipEx) {
                $result.Status  = 'SKIPPED'
                $result.Message = "Already a member of $grp"
            } elseif ($alreadyMember) {
                $result.Status  = 'ERROR'
                $result.Message = "Already a member of $grp (skip-existing is OFF)"
            } else {
                $group.Add("WinNT://$dom/$user")
                $result.Status  = 'SUCCESS'
                $result.Message = "Added to $grp"
            }

            if ($verify -and $result.Status -eq 'SUCCESS') {
                $membersAfter = @($group.Invoke('Members') | ForEach-Object { $_.GetType().InvokeMember('Name','GetProperty',$null,$_,$null) })
                $confirmed = $membersAfter | Where-Object { $_.TrimEnd('$').ToLower() -eq $userComp }
                $result.Verified = [bool]$confirmed
                if (-not $result.Verified) {
                    $result.Status  = 'WARNING'
                    $result.Message = "Add succeeded but membership verification failed"
                }
            }
        } catch {
            $result.Status  = 'ERROR'
            $result.Message = $_.Exception.Message
        }
        return $result
    }

    $icmParams = @{
        ComputerName  = $Server
        ScriptBlock   = $scriptBlock
        ArgumentList  = $Account, $Group, $SkipExisting, $Verify
        ErrorAction   = 'Stop'
    }
    if ($Credential) { $icmParams.Credential = $Credential }

    try {
        $res = Invoke-Command @icmParams
        return $res
    } catch {
        return @{ Status='ERROR'; Message="Invoke-Command failed: $($_.Exception.Message)" }
    }
}

# ─────────────────────────────────────────────
# BUTTON EVENTS
# ─────────────────────────────────────────────

# ── Credential: Save ──
$btnSaveCred.Add_Click({
    $name = $txtCredName.Text.Trim()
    $user = $txtCredUser.Text.Trim()
    $pass = $txtCredPass.Text

    if (-not $name) { $lblCredStatus.Text = "⚠ Profile name required."; $lblCredStatus.ForeColor=$clrWarn; return }
    if (-not $user) { $lblCredStatus.Text = "⚠ Username required.";      $lblCredStatus.ForeColor=$clrWarn; return }
    if (-not $pass) { $lblCredStatus.Text = "⚠ Password required.";      $lblCredStatus.ForeColor=$clrWarn; return }

    $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
    $cred    = New-Object System.Management.Automation.PSCredential($user,$secPass)
    $CredProfiles[$name] = $cred

    # Refresh comboboxes
    foreach ($cb in @($cmbCredProf,$cmbMSACredProf)) {
        if (-not $cb.Items.Contains($name)) { [void]$cb.Items.Add($name) }
    }
    if (-not $lbCredProfiles.Items.Contains($name)) { [void]$lbCredProfiles.Items.Add($name) }

    $lblCredStatus.Text = "✔ Profile '$name' saved."; $lblCredStatus.ForeColor = $clrSuccess
    Write-UILog "Credential profile '$name' saved for user '$user'." 'INFO'
    $txtCredPass.Clear()
})

# ── Credential: Remove ──
$btnRemoveCred.Add_Click({
    $sel = $lbCredProfiles.SelectedItem
    if (-not $sel) { return }
    $CredProfiles.Remove($sel)
    $lbCredProfiles.Items.Remove($sel)
    foreach ($cb in @($cmbCredProf,$cmbMSACredProf)) { $cb.Items.Remove($sel) }
    Write-UILog "Credential profile '$sel' removed." 'INFO'
})

# ── Import Servers CSV ──
$btnImportSrv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title  = "Import Server List"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtServers.Lines = $lines
            Write-UILog "Imported $($lines.Count) servers from $($ofd.FileName)" 'INFO'
        } catch { Write-UILog "CSV import error: $_" 'ERROR' }
    }
})

# ── Export Servers CSV ──
$btnExportSrv.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"; $sfd.FileName = "servers.csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        $txtServers.Lines | Where-Object { $_.Trim() } | ForEach-Object { [PSCustomObject]@{Server=$_} } |
            Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Servers exported to $($sfd.FileName)" 'INFO'
    }
})

# ── Import Users CSV ──
$btnImportUsr.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title  = "Import User List"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtUsers.Lines = $lines
            Write-UILog "Imported $($lines.Count) users from $($ofd.FileName)" 'INFO'
        } catch { Write-UILog "CSV import error: $_" 'ERROR' }
    }
})

# ── Export Users CSV ──
$btnExportUsr.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"; $sfd.FileName = "users.csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        $txtUsers.Lines | Where-Object { $_.Trim() } | ForEach-Object { [PSCustomObject]@{Account=$_} } |
            Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Users exported to $($sfd.FileName)" 'INFO'
    }
})

# ── Clear All ──
$btnClear.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Clear all server and user lists?","Confirm Clear",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq 'Yes') {
        $txtServers.Clear(); $txtUsers.Clear()
        $progressBar.Value = 0; $lblProgressDetail.Text = "Cleared."
    }
})

# ── Open Log ──
$btnOpenLog.Add_Click({ Start-Process notepad.exe -ArgumentList $LogFile })

# ── Error Grid Export ──
$btnExportErrors.Add_Click({
    if ($ErrorRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No error data to export.","Export","OK","Information") | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "operation_results_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    if ($sfd.ShowDialog() -eq 'OK') {
        $ErrorRows | Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Results exported to $($sfd.FileName)" 'INFO'
    }
})

$btnClearErrors.Add_Click({ $dgvErrors.Rows.Clear(); $ErrorRows.Clear() })

# ── Log tab buttons ──
$btnClearLog.Add_Click({ $rtbLog.Clear() })
$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt"
    $sfd.FileName = "export_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    if ($sfd.ShowDialog() -eq 'OK') {
        $rtbLog.Text | Set-Content $sfd.FileName
    }
})
$btnOpenLogFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $LogRoot })

# ── Error Grid: open when report button clicked ──
$btnReport.Add_Click({ $tabs.SelectedTab = $tabErrors })

# ── MSA Import ──
$btnImportMSASrv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtMSAServers.Lines = $lines
        } catch { Write-UILog "CSV import error: $_" 'ERROR' }
    }
})
$btnImportMSAAcc.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtMSAAccounts.Lines = $lines
        } catch { Write-UILog "CSV import error: $_" 'ERROR' }
    }
})

# ─────────────────────────────────────────────
# MAIN RUN LOGIC
# ─────────────────────────────────────────────
$btnRun.Add_Click({

    $servers = $txtServers.Lines | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $users   = $txtUsers.Lines   | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $group   = $cmbGroup.Text.Trim()
    $profSel = $cmbCredProf.Text

    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one server.","Validation","OK","Warning") | Out-Null; return
    }
    if ($users.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one user.","Validation","OK","Warning") | Out-Null; return
    }
    if (-not $group) {
        [System.Windows.Forms.MessageBox]::Show("Please select a target group.","Validation","OK","Warning") | Out-Null; return
    }

    $cred = $null
    if ($profSel -ne "[ Current Windows Session ]" -and $CredProfiles.ContainsKey($profSel)) {
        $cred = $CredProfiles[$profSel]
        Write-UILog "Using credential profile: $profSel" 'INFO'
    } else {
        Write-UILog "Using current Windows session credentials." 'INFO'
    }

    $skipExisting = $chkSkipExisting.Checked
    $verify       = $chkVerify.Checked
    $doPing       = $chkPing.Checked

    $total   = $servers.Count * $users.Count
    $done    = 0
    $success = 0; $skipped = 0; $errors = 0

    $progressBar.Maximum = $total
    $progressBar.Value   = 0

    Write-UILog "=== Operation Start: $total operations ($($servers.Count) servers × $($users.Count) users) ===" 'INFO'
    $tabs.SelectedTab = $tabLog

    foreach ($server in $servers) {

        # Ping test
        if ($doPing) {
            Write-UILog "Pinging $server..." 'INFO'
            $pingOK = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $pingOK) {
                Write-UILog "UNREACHABLE: $server — skipping all users for this server." 'WARNING'
                foreach ($user in $users) {
                    Add-ErrorRow $server $user $group 'ERROR' "Server unreachable (ping failed)"
                    $errors++; $done++
                    $progressBar.Value = [Math]::Min($done,$total)
                    $lblProgressDetail.Text = "[$done/$total] $server — unreachable"
                    [System.Windows.Forms.Application]::DoEvents()
                }
                continue
            }
            Write-UILog "Reachable: $server" 'INFO'
        }

        foreach ($user in $users) {

            $lblProgressDetail.Text = "[$done/$total] Adding '$user' to '$group' on '$server'..."
            [System.Windows.Forms.Application]::DoEvents()

            Write-UILog "Processing: $user → $server [$group]" 'INFO'

            $res = Add-AccountToLocalGroup -Server $server -Account $user -Group $group `
                       -SkipExisting $skipExisting -Verify $verify -Credential $cred

            switch ($res.Status) {
                'SUCCESS' {
                    $success++
                    Write-UILog "SUCCESS: $user on $server → $($res.Message)" 'SUCCESS'
                    Add-ErrorRow $server $user $group 'SUCCESS' $res.Message
                }
                'SKIPPED' {
                    $skipped++
                    Write-UILog "SKIPPED: $user on $server → $($res.Message)" 'WARNING'
                    Add-ErrorRow $server $user $group 'SKIPPED' $res.Message
                }
                'WARNING' {
                    $errors++
                    Write-UILog "WARNING: $user on $server → $($res.Message)" 'WARNING'
                    Add-ErrorRow $server $user $group 'WARNING' $res.Message
                }
                default {
                    $errors++
                    Write-UILog "ERROR: $user on $server → $($res.Message)" 'ERROR'
                    Add-ErrorRow $server $user $group 'ERROR' $res.Message
                }
            }

            $done++
            $progressBar.Value = [Math]::Min($done,$total)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $summary = "=== Complete: $total ops | ✔ $success succeeded | ⚠ $skipped skipped | ✖ $errors errors ==="
    Write-UILog $summary 'INFO'
    $lblProgressDetail.Text = $summary

    if ($errors -gt 0) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "$errors error(s) occurred.`nSwitch to the Error Grid tab?",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -eq 'Yes') { $tabs.SelectedTab = $tabErrors }
    } else {
        [System.Windows.Forms.MessageBox]::Show($summary,"Operation Complete","OK","Information") | Out-Null
    }
})

# ─────────────────────────────────────────────
# MSA RUN LOGIC
# ─────────────────────────────────────────────
$btnRunMSA.Add_Click({

    $servers  = $txtMSAServers.Lines  | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $accounts = $txtMSAAccounts.Lines | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $group    = $cmbMSAGroup.Text.Trim()
    $profSel  = $cmbMSACredProf.Text

    if ($servers.Count -eq 0 -or $accounts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter servers and MSA accounts.","Validation","OK","Warning") | Out-Null; return
    }

    $cred = $null
    if ($profSel -ne "[ Current Windows Session ]" -and $CredProfiles.ContainsKey($profSel)) {
        $cred = $CredProfiles[$profSel]
    }

    $total = $servers.Count * $accounts.Count
    $done  = 0
    $msaProgressBar.Maximum = $total
    $msaProgressBar.Value   = 0

    Write-UILog "=== MSA Operation Start: $total operations ===" 'INFO'

    foreach ($server in $servers) {
        foreach ($acc in $accounts) {

            $lblMSAStatus.Text = "[$done/$total] Adding MSA '$acc' to '$group' on '$server'..."
            [System.Windows.Forms.Application]::DoEvents()

            $res = Add-AccountToLocalGroup -Server $server -Account $acc -Group $group `
                       -SkipExisting $true -Verify $true -Credential $cred

            $statusLabel = $res.Status
            Write-UILog "MSA [$statusLabel] $acc → $server [$group]: $($res.Message)" $statusLabel
            Add-ErrorRow $server $acc $group $statusLabel $res.Message

            $done++
            $msaProgressBar.Value = [Math]::Min($done,$total)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $lblMSAStatus.Text = "MSA operation complete. $done operations processed."
    Write-UILog "=== MSA Operation Complete ===" 'INFO'
    [System.Windows.Forms.MessageBox]::Show("MSA operation complete.`nCheck the Error Grid / Live Log for details.","Done","OK","Information") | Out-Null
})

# ─────────────────────────────────────────────
# FORM LOAD
# ─────────────────────────────────────────────
$form.Add_Shown({
    Write-UILog "Tool ready. Log directory: $LogRoot" 'INFO'
    Write-UILog "Add credential profiles on the 'Credential Profiles' tab before running." 'INFO'
})

# ─────────────────────────────────────────────
# LAUNCH
# ─────────────────────────────────────────────
[void]$form.ShowDialog()
$form.Dispose()
