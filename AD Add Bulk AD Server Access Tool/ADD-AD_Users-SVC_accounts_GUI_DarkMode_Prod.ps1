#Requires -Version 5.1
<#
.SYNOPSIS
    Add Bulk AD Server Access Tool v2.0
    GUI-based tool for adding users and MSAs to local groups across multiple servers.

.NOTES
    Run in PowerShell ISE or standard PowerShell console.
    Requires WinRM on target servers.
    Log files saved to: Desktop\Add Bulk AD Server Access Tool\Logs\
#>

# -----------------------------------------------
# BOOTSTRAP - ensure STA mode for WinForms
# -----------------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $args0 = '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', $MyInvocation.MyCommand.Definition
    Start-Process powershell.exe -ArgumentList $args0
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# -----------------------------------------------
# LOGGING SETUP
# -----------------------------------------------
$LogRoot = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Add Bulk AD Server Access Tool\Logs'
if (-not (Test-Path $LogRoot)) { New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null }
$LogFile  = Join-Path $LogRoot ("log_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$ErrorRows = [System.Collections.Generic.List[PSObject]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    switch ($Level) {
        'SUCCESS' { Write-Host $line -ForegroundColor Green  }
        'WARNING' { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red    }
        default   { Write-Host $line -ForegroundColor Cyan   }
    }
}

Write-Log "Tool started. Log: $LogFile"

# -----------------------------------------------
# CREDENTIAL PROFILES (in-memory only)
# -----------------------------------------------
$CredProfiles = [System.Collections.Generic.Dictionary[string,PSCredential]]::new()

# -----------------------------------------------
# COLORS - Dark Grey Theme
# -----------------------------------------------
$clrBg        = [System.Drawing.Color]::FromArgb(58, 58, 62)
$clrPanel     = [System.Drawing.Color]::FromArgb(70, 70, 75)
$clrAccent    = [System.Drawing.Color]::FromArgb(0, 150, 215)
$clrAccentHov = [System.Drawing.Color]::FromArgb(0, 180, 255)
$clrText      = [System.Drawing.Color]::FromArgb(225, 225, 230)
$clrMuted     = [System.Drawing.Color]::FromArgb(160, 160, 170)
$clrSuccess   = [System.Drawing.Color]::FromArgb(50, 200, 100)
$clrWarn      = [System.Drawing.Color]::FromArgb(255, 190, 0)
$clrError     = [System.Drawing.Color]::FromArgb(220, 60, 60)
$clrBorder    = [System.Drawing.Color]::FromArgb(90, 90, 100)
$clrInputBg   = [System.Drawing.Color]::FromArgb(52, 52, 56)
$clrInputBg2  = [System.Drawing.Color]::FromArgb(62, 62, 66)
$clrTitleBar  = [System.Drawing.Color]::FromArgb(45, 45, 50)
$clrGridHdr   = [System.Drawing.Color]::FromArgb(55, 55, 60)

# -----------------------------------------------
# FONTS
# -----------------------------------------------
$fntTitle  = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$fntLabel  = New-Object System.Drawing.Font('Segoe UI', 9,  [System.Drawing.FontStyle]::Bold)
$fntNormal = New-Object System.Drawing.Font('Segoe UI', 9)
$fntMono   = New-Object System.Drawing.Font('Consolas', 9)
$fntSmall  = New-Object System.Drawing.Font('Segoe UI', 8)

# -----------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------
function New-StyledButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$W=130, [int]$H=28)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = [System.Drawing.Point]::new($X, $Y)
    $b.Size      = [System.Drawing.Size]::new($W, $H)
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor        = $clrAccent
    $b.FlatAppearance.BorderSize         = 1
    $b.FlatAppearance.MouseOverBackColor = $clrAccentHov
    $b.BackColor = $clrPanel
    $b.ForeColor = $clrText
    $b.Font      = $fntLabel
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function New-StyledTextBox {
    param([int]$X, [int]$Y, [int]$W, [int]$H,
          [bool]$Multi=$false, [bool]$Password=$false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = [System.Drawing.Point]::new($X, $Y)
    $t.Size        = [System.Drawing.Size]::new($W, $H)
    $t.BackColor   = $clrInputBg
    $t.ForeColor   = $clrText
    $t.Font        = $fntMono
    $t.BorderStyle = 'FixedSingle'
    if ($Multi)    { $t.Multiline = $true; $t.ScrollBars = 'Vertical' }
    if ($Password) { $t.PasswordChar = [char]0x2022 }
    return $t
}

function New-StyledLabel {
    param([string]$Text, [int]$X, [int]$Y, [int]$W=200, [int]$H=18,
          [System.Drawing.Font]$Font=$fntLabel)
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.Location  = [System.Drawing.Point]::new($X, $Y)
    $l.Size      = [System.Drawing.Size]::new($W, $H)
    $l.ForeColor = $clrText
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.Font      = $Font
    return $l
}

function New-GroupBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text      = $Text
    $g.Location  = [System.Drawing.Point]::new($X, $Y)
    $g.Size      = [System.Drawing.Size]::new($W, $H)
    $g.ForeColor = $clrAccent
    $g.BackColor = $clrPanel
    $g.Font      = $fntLabel
    return $g
}

# -----------------------------------------------
# MAIN FORM
# -----------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Add Bulk AD Server Access Tool  v2.0"
$form.Size            = [System.Drawing.Size]::new(1100, 820)
$form.MinimumSize     = [System.Drawing.Size]::new(1100, 820)
$form.BackColor       = $clrBg
$form.ForeColor       = $clrText
$form.Font            = $fntNormal
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'Sizable'

# Title bar
$pnlTitle           = New-Object System.Windows.Forms.Panel
$pnlTitle.Dock      = 'Top'
$pnlTitle.Height    = 48
$pnlTitle.BackColor = $clrTitleBar

$lblTitle           = New-StyledLabel "  [AD TOOL]  Add Bulk AD Server Access Tool" 0 10 660 26 $fntTitle
$lblTitle.ForeColor = $clrAccent

$lblLogPath           = New-StyledLabel "" 10 34 980 14 $fntSmall
$lblLogPath.ForeColor = $clrMuted
$lblLogPath.Text      = "  Log: $LogFile"

$pnlTitle.Controls.AddRange(@($lblTitle, $lblLogPath))
$form.Controls.Add($pnlTitle)

# -----------------------------------------------
# TAB CONTROL
# -----------------------------------------------
$tabs           = New-Object System.Windows.Forms.TabControl
$tabs.Location  = [System.Drawing.Point]::new(8, 56)
$tabs.Size      = [System.Drawing.Size]::new(1074, 700)
$tabs.BackColor = $clrBg
$tabs.Font      = $fntLabel

function New-Tab {
    param([string]$text)
    $tp           = New-Object System.Windows.Forms.TabPage
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
$tabs.TabPages.AddRange(@($tabMain, $tabCreds, $tabMSA, $tabLog, $tabErrors))
$form.Controls.Add($tabs)

# =======================================================
# TAB 1 - MAIN OPERATION
# =======================================================

# Left panel: Servers
$gbServers = New-GroupBox "Target Servers" 6 6 340 380
$tabMain.Controls.Add($gbServers)

$lblSrvHint           = New-StyledLabel "Paste one server per line  (or use Import):" 8 20 310 16 $fntSmall
$lblSrvHint.ForeColor = $clrMuted
$gbServers.Controls.Add($lblSrvHint)

$txtServers = New-StyledTextBox 8 40 322 290 $true
$gbServers.Controls.Add($txtServers)

$btnImportSrv = New-StyledButton "Import CSV"  8   338 120 26
$btnExportSrv = New-StyledButton "Export CSV"  136 338 120 26
$gbServers.Controls.AddRange(@($btnImportSrv, $btnExportSrv))

# Center panel: Users
$gbUsers = New-GroupBox "Users to Add  (DOMAIN\username)" 355 6 340 380
$tabMain.Controls.Add($gbUsers)

$lblUsrHint           = New-StyledLabel "Paste one DOMAIN\username per line  (or Import):" 8 20 320 16 $fntSmall
$lblUsrHint.ForeColor = $clrMuted
$gbUsers.Controls.Add($lblUsrHint)

$txtUsers = New-StyledTextBox 8 40 322 290 $true
$gbUsers.Controls.Add($txtUsers)

$btnImportUsr = New-StyledButton "Import CSV"  8   338 120 26
$btnExportUsr = New-StyledButton "Export CSV"  136 338 120 26
$gbUsers.Controls.AddRange(@($btnImportUsr, $btnExportUsr))

# Right panel: Options
$gbOptions = New-GroupBox "Options" 704 6 360 380
$tabMain.Controls.Add($gbOptions)

$gbOptions.Controls.Add((New-StyledLabel "Target Local Group:" 10 22 200 18))

$cmbGroup               = New-Object System.Windows.Forms.ComboBox
$cmbGroup.Location      = [System.Drawing.Point]::new(10, 42)
$cmbGroup.Size          = [System.Drawing.Size]::new(330, 24)
$cmbGroup.BackColor     = $clrInputBg2
$cmbGroup.ForeColor     = $clrText
$cmbGroup.Font          = $fntMono
$cmbGroup.DropDownStyle = 'DropDown'
@(
    'Administrators',
    'Remote Desktop Users',
    'Remote Management Users',
    'Backup Operators',
    'Performance Monitor Users',
    'Network Configuration Operators',
    'Event Log Readers',
    'Distributed COM Users'
) | ForEach-Object { [void]$cmbGroup.Items.Add($_) }
$cmbGroup.SelectedIndex = 0
$gbOptions.Controls.Add($cmbGroup)

$lblGrpHint           = New-StyledLabel "  (or type a custom group name above)" 10 68 300 14 $fntSmall
$lblGrpHint.ForeColor = $clrMuted
$gbOptions.Controls.Add($lblGrpHint)

$gbOptions.Controls.Add((New-StyledLabel "Credential Profile:" 10 94 300 18))

$cmbCredProf               = New-Object System.Windows.Forms.ComboBox
$cmbCredProf.Location      = [System.Drawing.Point]::new(10, 114)
$cmbCredProf.Size          = [System.Drawing.Size]::new(330, 24)
$cmbCredProf.BackColor     = $clrInputBg2
$cmbCredProf.ForeColor     = $clrText
$cmbCredProf.Font          = $fntMono
$cmbCredProf.DropDownStyle = 'DropDownList'
[void]$cmbCredProf.Items.Add("[ Current Windows Session ]")
$cmbCredProf.SelectedIndex = 0
$gbOptions.Controls.Add($cmbCredProf)

$lblCredNote           = New-StyledLabel "  (add profiles on Credential Profiles tab)" 10 140 330 14 $fntSmall
$lblCredNote.ForeColor = $clrMuted
$gbOptions.Controls.Add($lblCredNote)

$chkPing           = New-Object System.Windows.Forms.CheckBox
$chkPing.Text      = "Test connectivity (ping) before connecting"
$chkPing.Location  = [System.Drawing.Point]::new(10, 168)
$chkPing.Size      = [System.Drawing.Size]::new(330, 20)
$chkPing.Checked   = $true
$chkPing.ForeColor = $clrText
$chkPing.BackColor = [System.Drawing.Color]::Transparent
$chkPing.Font      = $fntNormal
$gbOptions.Controls.Add($chkPing)

$chkVerify           = New-Object System.Windows.Forms.CheckBox
$chkVerify.Text      = "Verify membership after add"
$chkVerify.Location  = [System.Drawing.Point]::new(10, 192)
$chkVerify.Size      = [System.Drawing.Size]::new(330, 20)
$chkVerify.Checked   = $true
$chkVerify.ForeColor = $clrText
$chkVerify.BackColor = [System.Drawing.Color]::Transparent
$chkVerify.Font      = $fntNormal
$gbOptions.Controls.Add($chkVerify)

$chkSkipExisting           = New-Object System.Windows.Forms.CheckBox
$chkSkipExisting.Text      = "Skip (don't error) if user already a member"
$chkSkipExisting.Location  = [System.Drawing.Point]::new(10, 216)
$chkSkipExisting.Size      = [System.Drawing.Size]::new(330, 20)
$chkSkipExisting.Checked   = $true
$chkSkipExisting.ForeColor = $clrText
$chkSkipExisting.BackColor = [System.Drawing.Color]::Transparent
$chkSkipExisting.Font      = $fntNormal
$gbOptions.Controls.Add($chkSkipExisting)

$lblSepLine           = New-Object System.Windows.Forms.Label
$lblSepLine.Location  = [System.Drawing.Point]::new(8, 246)
$lblSepLine.Size      = [System.Drawing.Size]::new(334, 1)
$lblSepLine.BackColor = $clrBorder
$gbOptions.Controls.Add($lblSepLine)

$lblSummary           = New-StyledLabel "Servers: 0   |   Users: 0   |   Operations: 0" 10 255 330 16 $fntSmall
$lblSummary.ForeColor = $clrMuted
$gbOptions.Controls.Add($lblSummary)

$updateSummary = {
    $sc = ($txtServers.Lines | Where-Object { $_.Trim() -ne '' }).Count
    $uc = ($txtUsers.Lines   | Where-Object { $_.Trim() -ne '' }).Count
    $lblSummary.Text = "Servers: $sc   |   Users: $uc   |   Operations: $($sc * $uc)"
}
$txtServers.Add_TextChanged($updateSummary)
$txtUsers.Add_TextChanged($updateSummary)

# Progress bar
$gbProgress = New-GroupBox "Progress" 6 394 1058 80
$tabMain.Controls.Add($gbProgress)

$progressBar           = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location  = [System.Drawing.Point]::new(10, 22)
$progressBar.Size      = [System.Drawing.Size]::new(940, 22)
$progressBar.Style     = 'Continuous'
$progressBar.ForeColor = $clrAccent
$gbProgress.Controls.Add($progressBar)

$lblProgressDetail           = New-StyledLabel "Ready." 10 50 1000 18 $fntSmall
$lblProgressDetail.ForeColor = $clrMuted
$gbProgress.Controls.Add($lblProgressDetail)

# Action buttons
$btnRun                                 = New-StyledButton "[ RUN ]"    6   484 160 40
$btnRun.Font                            = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnRun.FlatAppearance.BorderColor      = $clrSuccess
$btnRun.ForeColor                       = $clrSuccess

$btnClear   = New-StyledButton "Clear All"   176 484 130 40
$btnReport  = New-StyledButton "Error Grid"  316 484 130 40
$btnOpenLog = New-StyledButton "Open Log"    456 484 130 40
$tabMain.Controls.AddRange(@($btnRun, $btnClear, $btnReport, $btnOpenLog))

# =======================================================
# TAB 2 - CREDENTIAL PROFILES
# =======================================================
$gbCredList = New-GroupBox "Saved Credential Profiles" 6 6 400 620
$tabCreds.Controls.Add($gbCredList)

$lbCredProfiles             = New-Object System.Windows.Forms.ListBox
$lbCredProfiles.Location    = [System.Drawing.Point]::new(8, 22)
$lbCredProfiles.Size        = [System.Drawing.Size]::new(380, 490)
$lbCredProfiles.BackColor   = $clrInputBg
$lbCredProfiles.ForeColor   = $clrText
$lbCredProfiles.Font        = $fntMono
$lbCredProfiles.BorderStyle = 'FixedSingle'
$gbCredList.Controls.Add($lbCredProfiles)

$btnRemoveCred = New-StyledButton "Remove Selected" 8 522 180 26
$gbCredList.Controls.Add($btnRemoveCred)

$gbAddCred = New-GroupBox "Add / Update Credential Profile" 420 6 640 280
$tabCreds.Controls.Add($gbAddCred)

$gbAddCred.Controls.Add((New-StyledLabel "Profile Name  (e.g. CORP, DMZ, LAB):" 10 24 400 18))
$txtCredName = New-StyledTextBox 10 44 300 22
$gbAddCred.Controls.Add($txtCredName)

$gbAddCred.Controls.Add((New-StyledLabel "Domain and Username  (DOMAIN\username):" 10 76 400 18))
$txtCredUser      = New-StyledTextBox 10 96 300 22
$txtCredUser.Text = $env:USERDOMAIN + '\'
$gbAddCred.Controls.Add($txtCredUser)

$gbAddCred.Controls.Add((New-StyledLabel "Password:" 10 128 200 18))
$txtCredPass = New-StyledTextBox 10 148 300 22 $false $true
$gbAddCred.Controls.Add($txtCredPass)

$chkShowPass           = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text      = "Show password"
$chkShowPass.Location  = [System.Drawing.Point]::new(10, 176)
$chkShowPass.Size      = [System.Drawing.Size]::new(200, 20)
$chkShowPass.ForeColor = $clrText
$chkShowPass.BackColor = [System.Drawing.Color]::Transparent
$chkShowPass.Font      = $fntNormal
$chkShowPass.Add_CheckedChanged({
    if ($chkShowPass.Checked) {
        $txtCredPass.PasswordChar = [char]0
    } else {
        $txtCredPass.PasswordChar = [char]0x2022
    }
})
$gbAddCred.Controls.Add($chkShowPass)

$btnSaveCred = New-StyledButton "Save Profile" 10 206 150 28
$gbAddCred.Controls.Add($btnSaveCred)

$lblCredStatus = New-StyledLabel "" 170 210 400 18 $fntSmall
$gbAddCred.Controls.Add($lblCredStatus)

$gbCredHelp = New-GroupBox "Usage Notes" 420 298 640 200
$tabCreds.Controls.Add($gbCredHelp)

$rtbCredHelp             = New-Object System.Windows.Forms.RichTextBox
$rtbCredHelp.Location    = [System.Drawing.Point]::new(10, 22)
$rtbCredHelp.Size        = [System.Drawing.Size]::new(618, 160)
$rtbCredHelp.BackColor   = $clrInputBg
$rtbCredHelp.ForeColor   = $clrMuted
$rtbCredHelp.Font        = $fntSmall
$rtbCredHelp.ReadOnly    = $true
$rtbCredHelp.BorderStyle = 'None'
$noteLines = @(
    "- Create one profile per domain/environment (e.g. CORP, DMZ, LAB).",
    "- Credentials are held in memory only - NOT written to disk.",
    "- Select a profile on the Main Operation tab before running.",
    "- Use 'Current Windows Session' if your logon already has rights.",
    "- For cross-domain, create a profile for each domain's admin account.",
    "- Passwords are masked; use the Show password checkbox to verify."
)
$rtbCredHelp.Text = $noteLines -join [Environment]::NewLine
$gbCredHelp.Controls.Add($rtbCredHelp)

# =======================================================
# TAB 3 - MANAGED SERVICE ACCOUNTS
# =======================================================
$gbMSAServers = New-GroupBox "Target Servers for MSA" 6 6 340 540
$tabMSA.Controls.Add($gbMSAServers)

$txtMSAServers = New-StyledTextBox 8 22 322 460 $true
$gbMSAServers.Controls.Add($txtMSAServers)

$btnImportMSASrv = New-StyledButton "Import CSV" 8 492 120 26
$gbMSAServers.Controls.Add($btnImportMSASrv)

$gbMSAAccounts = New-GroupBox "MSA / gMSA Accounts  (DOMAIN\account`$)" 355 6 340 540
$tabMSA.Controls.Add($gbMSAAccounts)

$lblMSAHint           = New-StyledLabel "gMSA names typically end with dollar sign:" 8 20 310 16 $fntSmall
$lblMSAHint.ForeColor = $clrMuted
$gbMSAAccounts.Controls.Add($lblMSAHint)

$txtMSAAccounts = New-StyledTextBox 8 40 322 440 $true
$gbMSAAccounts.Controls.Add($txtMSAAccounts)

$btnImportMSAAcc = New-StyledButton "Import CSV" 8 492 120 26
$gbMSAAccounts.Controls.Add($btnImportMSAAcc)

$gbMSAOptions = New-GroupBox "MSA Options" 704 6 360 280
$tabMSA.Controls.Add($gbMSAOptions)

$gbMSAOptions.Controls.Add((New-StyledLabel "Target Local Group:" 10 22 200 18))

$cmbMSAGroup               = New-Object System.Windows.Forms.ComboBox
$cmbMSAGroup.Location      = [System.Drawing.Point]::new(10, 42)
$cmbMSAGroup.Size          = [System.Drawing.Size]::new(330, 24)
$cmbMSAGroup.BackColor     = $clrInputBg2
$cmbMSAGroup.ForeColor     = $clrText
$cmbMSAGroup.Font          = $fntMono
$cmbMSAGroup.DropDownStyle = 'DropDown'
@(
    'Administrators',
    'Remote Desktop Users',
    'Remote Management Users',
    'Backup Operators',
    'Performance Monitor Users'
) | ForEach-Object { [void]$cmbMSAGroup.Items.Add($_) }
$cmbMSAGroup.SelectedIndex = 0
$gbMSAOptions.Controls.Add($cmbMSAGroup)

$gbMSAOptions.Controls.Add((New-StyledLabel "Credential Profile:" 10 76 200 18))

$cmbMSACredProf               = New-Object System.Windows.Forms.ComboBox
$cmbMSACredProf.Location      = [System.Drawing.Point]::new(10, 96)
$cmbMSACredProf.Size          = [System.Drawing.Size]::new(330, 24)
$cmbMSACredProf.BackColor     = $clrInputBg2
$cmbMSACredProf.ForeColor     = $clrText
$cmbMSACredProf.Font          = $fntMono
$cmbMSACredProf.DropDownStyle = 'DropDownList'
[void]$cmbMSACredProf.Items.Add("[ Current Windows Session ]")
$cmbMSACredProf.SelectedIndex = 0
$gbMSAOptions.Controls.Add($cmbMSACredProf)

$gbMSAOptions.Controls.Add((New-StyledLabel "NOTE: gMSA must already be installed on" 10 130 330 16 $fntSmall))
$gbMSAOptions.Controls.Add((New-StyledLabel "each server via Install-ADServiceAccount." 10 148 330 16 $fntSmall))

$btnRunMSA                            = New-StyledButton "[ ADD MSAs ]" 704 300 160 40
$btnRunMSA.Font                       = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnRunMSA.FlatAppearance.BorderColor = $clrWarn
$btnRunMSA.ForeColor                  = $clrWarn
$tabMSA.Controls.Add($btnRunMSA)

$msaProgressBar          = New-Object System.Windows.Forms.ProgressBar
$msaProgressBar.Location = [System.Drawing.Point]::new(6, 560)
$msaProgressBar.Size     = [System.Drawing.Size]::new(1050, 18)
$msaProgressBar.Style    = 'Continuous'
$tabMSA.Controls.Add($msaProgressBar)

$lblMSAStatus           = New-StyledLabel "Ready." 6 582 900 16 $fntSmall
$lblMSAStatus.ForeColor = $clrMuted
$tabMSA.Controls.Add($lblMSAStatus)

# =======================================================
# TAB 4 - LIVE LOG
# =======================================================
$rtbLog            = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock       = 'Fill'
$rtbLog.BackColor  = $clrInputBg
$rtbLog.ForeColor  = $clrText
$rtbLog.Font       = $fntMono
$rtbLog.ReadOnly   = $true
$rtbLog.ScrollBars = 'Vertical'
$rtbLog.WordWrap   = $false
$tabLog.Controls.Add($rtbLog)

$pnlLogBtns           = New-Object System.Windows.Forms.Panel
$pnlLogBtns.Dock      = 'Bottom'
$pnlLogBtns.Height    = 36
$pnlLogBtns.BackColor = $clrPanel

$btnClearLog      = New-StyledButton "Clear View"      4   4 120 26
$btnSaveLog       = New-StyledButton "Save Log"        132 4 120 26
$btnOpenLogFolder = New-StyledButton "Open Log Folder" 260 4 160 26
$pnlLogBtns.Controls.AddRange(@($btnClearLog, $btnSaveLog, $btnOpenLogFolder))
$tabLog.Controls.Add($pnlLogBtns)

# =======================================================
# TAB 5 - ERROR GRID
# =======================================================
$dgvErrors                                            = New-Object System.Windows.Forms.DataGridView
$dgvErrors.Dock                                       = 'Fill'
$dgvErrors.BackgroundColor                            = $clrInputBg
$dgvErrors.ForeColor                                  = $clrText
$dgvErrors.GridColor                                  = $clrBorder
$dgvErrors.Font                                       = $fntMono
$dgvErrors.BorderStyle                                = 'None'
$dgvErrors.RowHeadersVisible                          = $false
$dgvErrors.AllowUserToAddRows                         = $false
$dgvErrors.ReadOnly                                   = $true
$dgvErrors.SelectionMode                              = 'FullRowSelect'
$dgvErrors.AutoSizeColumnsMode                        = 'Fill'
$dgvErrors.EnableHeadersVisualStyles                  = $false
$dgvErrors.ColumnHeadersDefaultCellStyle.BackColor    = $clrGridHdr
$dgvErrors.ColumnHeadersDefaultCellStyle.ForeColor    = $clrAccent
$dgvErrors.ColumnHeadersDefaultCellStyle.Font         = $fntLabel
$dgvErrors.DefaultCellStyle.BackColor                 = $clrInputBg
$dgvErrors.DefaultCellStyle.ForeColor                 = $clrText
$dgvErrors.DefaultCellStyle.SelectionBackColor        = [System.Drawing.Color]::FromArgb(40, 80, 120)

foreach ($col in @('Timestamp','Server','Account','Group','Status','Message')) {
    $c             = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.HeaderText  = $col
    $c.Name        = $col
    [void]$dgvErrors.Columns.Add($c)
}
$tabErrors.Controls.Add($dgvErrors)

$pnlErrBtns           = New-Object System.Windows.Forms.Panel
$pnlErrBtns.Dock      = 'Bottom'
$pnlErrBtns.Height    = 36
$pnlErrBtns.BackColor = $clrPanel

$btnExportErrors = New-StyledButton "Export Errors CSV" 4   4 170 26
$btnClearErrors  = New-StyledButton "Clear Grid"        182 4 120 26
$pnlErrBtns.Controls.AddRange(@($btnExportErrors, $btnClearErrors))
$tabErrors.Controls.Add($pnlErrBtns)

# -----------------------------------------------
# HELPERS: Logging and Grid
# -----------------------------------------------
function Add-LogLine {
    param([string]$Text, [System.Drawing.Color]$Color)
    $rtbLog.SelectionStart  = $rtbLog.TextLength
    $rtbLog.SelectionLength = 0
    $rtbLog.SelectionColor  = $Color
    $rtbLog.AppendText("$Text`n")
    $rtbLog.ScrollToCaret()
}

function Write-UILog {
    param([string]$Message, [string]$Level='INFO')
    Write-Log $Message $Level
    $color = switch ($Level) {
        'SUCCESS' { $clrSuccess }
        'WARNING' { $clrWarn   }
        'ERROR'   { $clrError  }
        default   { $clrMuted  }
    }
    $ts = Get-Date -Format 'HH:mm:ss'
    Add-LogLine "[$ts][$Level] $Message" $color
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-ErrorRow {
    param([string]$Server, [string]$Account, [string]$Group,
          [string]$Status, [string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    [void]$dgvErrors.Rows.Add($ts, $Server, $Account, $Group, $Status, $Message)
    $row = $dgvErrors.Rows[$dgvErrors.Rows.Count - 1]
    $row.DefaultCellStyle.ForeColor = switch ($Status) {
        'SUCCESS' { $clrSuccess }
        'SKIPPED' { $clrWarn   }
        default   { $clrError  }
    }
    $ErrorRows.Add([PSCustomObject]@{
        Timestamp = $ts
        Server    = $Server
        Account   = $Account
        Group     = $Group
        Status    = $Status
        Message   = $Message
    })
}

# -----------------------------------------------
# CORE: Add account to local group on remote server
# -----------------------------------------------
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
        param($acc, $grp, $skipEx, $verify)
        $result = @{ Status = ''; Message = '' }
        try {
            if ($acc -match '^(.+)\\(.+)$') {
                $dom  = $matches[1]
                $user = $matches[2]
            } else {
                $dom  = $env:USERDOMAIN
                $user = $acc
            }

            $grpObj  = [ADSI]"WinNT://./$grp,group"
            $members = @($grpObj.Invoke('Members') | ForEach-Object {
                $_.GetType().InvokeMember('Name','GetProperty',$null,$_,$null)
            })

            $userComp      = $user.TrimEnd('$').ToLower()
            $alreadyMember = $members | Where-Object { $_.TrimEnd('$').ToLower() -eq $userComp }

            if ($alreadyMember -and $skipEx) {
                $result.Status  = 'SKIPPED'
                $result.Message = "Already a member of $grp"
            } elseif ($alreadyMember) {
                $result.Status  = 'ERROR'
                $result.Message = "Already a member of $grp (skip-existing is OFF)"
            } else {
                $grpObj.Add("WinNT://$dom/$user")
                $result.Status  = 'SUCCESS'
                $result.Message = "Added to $grp"
            }

            if ($verify -and $result.Status -eq 'SUCCESS') {
                $membersAfter = @($grpObj.Invoke('Members') | ForEach-Object {
                    $_.GetType().InvokeMember('Name','GetProperty',$null,$_,$null)
                })
                $confirmed = $membersAfter | Where-Object { $_.TrimEnd('$').ToLower() -eq $userComp }
                if (-not $confirmed) {
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
        ComputerName = $Server
        ScriptBlock  = $scriptBlock
        ArgumentList = $Account, $Group, $SkipExisting, $Verify
        ErrorAction  = 'Stop'
    }
    if ($Credential) { $icmParams.Credential = $Credential }

    try {
        return Invoke-Command @icmParams
    } catch {
        return @{ Status = 'ERROR'; Message = "Invoke-Command failed: $($_.Exception.Message)" }
    }
}

# =======================================================
# BUTTON EVENTS
# =======================================================

$btnSaveCred.Add_Click({
    $name = $txtCredName.Text.Trim()
    $user = $txtCredUser.Text.Trim()
    $pass = $txtCredPass.Text

    if (-not $name) {
        $lblCredStatus.Text      = "Profile name required."
        $lblCredStatus.ForeColor = $clrWarn
        return
    }
    if (-not $user) {
        $lblCredStatus.Text      = "Username required."
        $lblCredStatus.ForeColor = $clrWarn
        return
    }
    if (-not $pass) {
        $lblCredStatus.Text      = "Password required."
        $lblCredStatus.ForeColor = $clrWarn
        return
    }

    $secPass             = ConvertTo-SecureString $pass -AsPlainText -Force
    $cred                = New-Object System.Management.Automation.PSCredential($user, $secPass)
    $CredProfiles[$name] = $cred

    foreach ($cb in @($cmbCredProf, $cmbMSACredProf)) {
        if (-not $cb.Items.Contains($name)) { [void]$cb.Items.Add($name) }
    }
    if (-not $lbCredProfiles.Items.Contains($name)) { [void]$lbCredProfiles.Items.Add($name) }

    $lblCredStatus.Text      = "Profile '$name' saved successfully."
    $lblCredStatus.ForeColor = $clrSuccess
    Write-UILog "Credential profile '$name' saved for user '$user'." 'INFO'
    $txtCredPass.Clear()
})

$btnRemoveCred.Add_Click({
    $sel = $lbCredProfiles.SelectedItem
    if (-not $sel) { return }
    $CredProfiles.Remove($sel)
    $lbCredProfiles.Items.Remove($sel)
    foreach ($cb in @($cmbCredProf, $cmbMSACredProf)) { $cb.Items.Remove($sel) }
    Write-UILog "Credential profile '$sel' removed." 'INFO'
})

$btnImportSrv.Add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title  = "Import Server List"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines            = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtServers.Lines = $lines
            Write-UILog "Imported $($lines.Count) servers from $($ofd.FileName)" 'INFO'
        } catch {
            Write-UILog "CSV import error: $_" 'ERROR'
        }
    }
})

$btnExportSrv.Add_Click({
    $sfd          = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "servers.csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        $txtServers.Lines | Where-Object { $_.Trim() } |
            ForEach-Object { [PSCustomObject]@{ Server = $_ } } |
            Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Servers exported to $($sfd.FileName)" 'INFO'
    }
})

$btnImportUsr.Add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $ofd.Title  = "Import User List"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines          = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtUsers.Lines = $lines
            Write-UILog "Imported $($lines.Count) users from $($ofd.FileName)" 'INFO'
        } catch {
            Write-UILog "CSV import error: $_" 'ERROR'
        }
    }
})

$btnExportUsr.Add_Click({
    $sfd          = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "users.csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        $txtUsers.Lines | Where-Object { $_.Trim() } |
            ForEach-Object { [PSCustomObject]@{ Account = $_ } } |
            Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Users exported to $($sfd.FileName)" 'INFO'
    }
})

$btnClear.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Clear all server and user lists?", "Confirm Clear",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq 'Yes') {
        $txtServers.Clear()
        $txtUsers.Clear()
        $progressBar.Value      = 0
        $lblProgressDetail.Text = "Cleared."
    }
})

$btnOpenLog.Add_Click({ Start-Process notepad.exe -ArgumentList $LogFile })
$btnReport.Add_Click({ $tabs.SelectedTab = $tabErrors })

$btnExportErrors.Add_Click({
    if ($ErrorRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No data to export.", "Export", "OK", "Information") | Out-Null
        return
    }
    $sfd          = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "operation_results_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    if ($sfd.ShowDialog() -eq 'OK') {
        $ErrorRows | Export-Csv $sfd.FileName -NoTypeInformation
        Write-UILog "Results exported to $($sfd.FileName)" 'INFO'
    }
})

$btnClearErrors.Add_Click({ $dgvErrors.Rows.Clear(); $ErrorRows.Clear() })

$btnClearLog.Add_Click({ $rtbLog.Clear() })

$btnSaveLog.Add_Click({
    $sfd          = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "Log Files (*.log)|*.log|Text Files (*.txt)|*.txt"
    $sfd.FileName = "export_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    if ($sfd.ShowDialog() -eq 'OK') {
        $rtbLog.Text | Set-Content $sfd.FileName
    }
})

$btnOpenLogFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $LogRoot })

$btnImportMSASrv.Add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines               = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtMSAServers.Lines = $lines
        } catch {
            Write-UILog "CSV import error: $_" 'ERROR'
        }
    }
})

$btnImportMSAAcc.Add_Click({
    $ofd        = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq 'OK') {
        try {
            $lines                = Import-Csv $ofd.FileName | ForEach-Object { $_.PSObject.Properties.Value[0] }
            $txtMSAAccounts.Lines = $lines
        } catch {
            Write-UILog "CSV import error: $_" 'ERROR'
        }
    }
})

# =======================================================
# MAIN RUN LOGIC
# =======================================================
$btnRun.Add_Click({

    $servers = $txtServers.Lines | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $users   = $txtUsers.Lines   | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $group   = $cmbGroup.Text.Trim()
    $profSel = $cmbCredProf.Text

    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter at least one server.", "Validation", "OK", "Warning") | Out-Null
        return
    }
    if ($users.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter at least one user.", "Validation", "OK", "Warning") | Out-Null
        return
    }
    if (-not $group) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a target group.", "Validation", "OK", "Warning") | Out-Null
        return
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
    $success = 0
    $skipped = 0
    $errors  = 0

    $progressBar.Maximum = $total
    $progressBar.Value   = 0

    Write-UILog "=== Operation Start: $total operations ($($servers.Count) servers x $($users.Count) users) ===" 'INFO'
    $tabs.SelectedTab = $tabLog

    foreach ($server in $servers) {

        if ($doPing) {
            Write-UILog "Pinging $server ..." 'INFO'
            $pingOK = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $pingOK) {
                Write-UILog "UNREACHABLE: $server - skipping all users for this server." 'WARNING'
                foreach ($user in $users) {
                    Add-ErrorRow $server $user $group 'ERROR' "Server unreachable (ping failed)"
                    $errors++
                    $done++
                    $progressBar.Value      = [Math]::Min($done, $total)
                    $lblProgressDetail.Text = "[$done/$total] $server - unreachable"
                    [System.Windows.Forms.Application]::DoEvents()
                }
                continue
            }
            Write-UILog "Reachable: $server" 'INFO'
        }

        foreach ($user in $users) {

            $lblProgressDetail.Text = "[$done/$total] Adding '$user' to '$group' on '$server' ..."
            [System.Windows.Forms.Application]::DoEvents()
            Write-UILog "Processing: $user -> $server [$group]" 'INFO'

            $res = Add-AccountToLocalGroup `
                -Server       $server `
                -Account      $user `
                -Group        $group `
                -SkipExisting $skipExisting `
                -Verify       $verify `
                -Credential   $cred

            switch ($res.Status) {
                'SUCCESS' {
                    $success++
                    Write-UILog "SUCCESS: $user on $server -> $($res.Message)" 'SUCCESS'
                    Add-ErrorRow $server $user $group 'SUCCESS' $res.Message
                }
                'SKIPPED' {
                    $skipped++
                    Write-UILog "SKIPPED: $user on $server -> $($res.Message)" 'WARNING'
                    Add-ErrorRow $server $user $group 'SKIPPED' $res.Message
                }
                'WARNING' {
                    $errors++
                    Write-UILog "WARNING: $user on $server -> $($res.Message)" 'WARNING'
                    Add-ErrorRow $server $user $group 'WARNING' $res.Message
                }
                default {
                    $errors++
                    Write-UILog "ERROR: $user on $server -> $($res.Message)" 'ERROR'
                    Add-ErrorRow $server $user $group 'ERROR' $res.Message
                }
            }

            $done++
            $progressBar.Value = [Math]::Min($done, $total)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $summary = "=== Complete: $total ops | OK: $success | Skipped: $skipped | Errors: $errors ==="
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
        [System.Windows.Forms.MessageBox]::Show(
            $summary, "Operation Complete", "OK", "Information") | Out-Null
    }
})

# =======================================================
# MSA RUN LOGIC
# =======================================================
$btnRunMSA.Add_Click({

    $servers  = $txtMSAServers.Lines  | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $accounts = $txtMSAAccounts.Lines | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $group    = $cmbMSAGroup.Text.Trim()
    $profSel  = $cmbMSACredProf.Text

    if ($servers.Count -eq 0 -or $accounts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter servers and MSA accounts.", "Validation", "OK", "Warning") | Out-Null
        return
    }

    $cred = $null
    if ($profSel -ne "[ Current Windows Session ]" -and $CredProfiles.ContainsKey($profSel)) {
        $cred = $CredProfiles[$profSel]
    }

    $total                  = $servers.Count * $accounts.Count
    $done                   = 0
    $msaProgressBar.Maximum = $total
    $msaProgressBar.Value   = 0

    Write-UILog "=== MSA Operation Start: $total operations ===" 'INFO'

    foreach ($server in $servers) {
        foreach ($acc in $accounts) {

            $lblMSAStatus.Text = "[$done/$total] Adding MSA '$acc' to '$group' on '$server' ..."
            [System.Windows.Forms.Application]::DoEvents()

            $res = Add-AccountToLocalGroup `
                -Server       $server `
                -Account      $acc `
                -Group        $group `
                -SkipExisting $true `
                -Verify       $true `
                -Credential   $cred

            Write-UILog "MSA [$($res.Status)] $acc -> $server [$group]: $($res.Message)" $res.Status
            Add-ErrorRow $server $acc $group $res.Status $res.Message

            $done++
            $msaProgressBar.Value = [Math]::Min($done, $total)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $lblMSAStatus.Text = "MSA operation complete. $done operations processed."
    Write-UILog "=== MSA Operation Complete ===" 'INFO'
    [System.Windows.Forms.MessageBox]::Show(
        "MSA operation complete.`nCheck the Error Grid or Live Log for details.",
        "Done", "OK", "Information") | Out-Null
})

# =======================================================
# FORM LOAD
# =======================================================
$form.Add_Shown({
    Write-UILog "Tool ready. Log directory: $LogRoot" 'INFO'
    Write-UILog "Tip: Add credential profiles on the Credential Profiles tab before running across domains." 'INFO'
})

# =======================================================
# LAUNCH
# =======================================================
[void]$form.ShowDialog()
$form.Dispose()
