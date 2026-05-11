#Requires -Version 5.1
<#
.SYNOPSIS
    Active Directory Kerberos Troubleshooting GUI Tool
.DESCRIPTION
    A comprehensive WinForms-based GUI tool for diagnosing and troubleshooting
    Kerberos authentication issues in Active Directory environments.
    Supports multiple domains and alternate credentials.
.NOTES
    Author      : SysAdmin Tools
    Version     : 2.0
    Requires    : PowerShell 5.1+, RSAT (ActiveDirectory module), Windows OS
    Run As      : Local admin or domain admin preferred for full diagnostic access
.EXAMPLE
    .\Invoke-KerberosTroubleshooter.ps1
#>

#region --- Bootstrap ---
$ErrorActionPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Logging ---
$LogDir  = "$env:USERPROFILE\Documents\KerbTroubleshooter"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogFile = "$LogDir\KerbLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

Write-Log "KerberosTroubleshooter started by $env:USERNAME on $env:COMPUTERNAME"

#endregion

#region --- Color / Style Palette ---
$clrBg          = [System.Drawing.Color]::FromArgb(18,  18,  30)   # Deep navy background
$clrPanel       = [System.Drawing.Color]::FromArgb(28,  28,  46)   # Slightly lighter panel
$clrCard        = [System.Drawing.Color]::FromArgb(36,  36,  58)   # Card / groupbox fill
$clrAccent      = [System.Drawing.Color]::FromArgb(82, 130, 255)   # Blue accent
$clrAccentHover = [System.Drawing.Color]::FromArgb(110,155,255)
$clrSuccess     = [System.Drawing.Color]::FromArgb(80, 200, 120)
$clrWarn        = [System.Drawing.Color]::FromArgb(255,200,  60)
$clrError       = [System.Drawing.Color]::FromArgb(255,  80,  80)
$clrText        = [System.Drawing.Color]::FromArgb(220, 220, 235)
$clrTextDim     = [System.Drawing.Color]::FromArgb(140, 140, 160)
$clrBorder      = [System.Drawing.Color]::FromArgb(55,  55,  85)
$clrInput       = [System.Drawing.Color]::FromArgb(22,  22,  38)
$clrTabSel      = [System.Drawing.Color]::FromArgb(82, 130, 255)
$clrTabNorm     = [System.Drawing.Color]::FromArgb(28,  28,  46)

$fontMono  = New-Object System.Drawing.Font("Consolas", 9)
$fontUI    = New-Object System.Drawing.Font("Segoe UI",  9)
$fontUIB   = New-Object System.Drawing.Font("Segoe UI",  9, [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$fontSmall = New-Object System.Drawing.Font("Segoe UI",  8)

#endregion

#region --- Helper Functions ---

# Shared credential store
$script:SavedCred   = $null
$script:ActiveDomain = $null

function Get-ToolCredential { return $script:SavedCred }
function Get-ActiveDomain   { return $script:ActiveDomain }

function New-StyledButton {
    param(
        [string]$Text,
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(150,32)),
        [System.Drawing.Color]$BgColor = $clrAccent,
        [System.Drawing.Color]$FgColor = [System.Drawing.Color]::White
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text            = $Text
    $btn.Location        = $Location
    $btn.Size            = $Size
    $btn.FlatStyle       = 'Flat'
    $btn.BackColor       = $BgColor
    $btn.ForeColor       = $FgColor
    $btn.Font            = $fontUIB
    $btn.Cursor          = [System.Windows.Forms.Cursors]::Hand
    $btn.FlatAppearance.BorderSize  = 0
    $btn.FlatAppearance.MouseOverBackColor = $clrAccentHover
    return $btn
}

function New-StyledLabel {
    param([string]$Text, [System.Drawing.Point]$Location,
          [System.Drawing.Size]$Size, [System.Drawing.Font]$Font = $fontUI,
          [System.Drawing.Color]$ForeColor = $clrText)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Text
    $lbl.Location  = $Location
    $lbl.Size      = $Size
    $lbl.Font      = $Font
    $lbl.ForeColor = $ForeColor
    $lbl.BackColor = [System.Drawing.Color]::Transparent
    return $lbl
}

function New-StyledTextBox {
    param([System.Drawing.Point]$Location, [System.Drawing.Size]$Size,
          [string]$Text = "", [bool]$ReadOnly = $false, [bool]$Multiline = $false,
          [bool]$Password = $false)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location   = $Location
    $tb.Size       = $Size
    $tb.Text       = $Text
    $tb.ReadOnly   = $ReadOnly
    $tb.Multiline  = $Multiline
    $tb.Font       = if ($Multiline) { $fontMono } else { $fontUI }
    $tb.BackColor  = $clrInput
    $tb.ForeColor  = $clrText
    $tb.BorderStyle = 'FixedSingle'
    if ($Password) { $tb.PasswordChar = [char]0x2022 }
    if ($Multiline) { $tb.ScrollBars = 'Vertical' }
    return $tb
}

function New-GroupBox {
    param([string]$Text, [System.Drawing.Point]$Location, [System.Drawing.Size]$Size)
    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text      = $Text
    $gb.Location  = $Location
    $gb.Size      = $Size
    $gb.Font      = $fontUIB
    $gb.ForeColor = $clrAccent
    $gb.BackColor = $clrCard
    return $gb
}

function Append-Output {
    param([System.Windows.Forms.RichTextBox]$RTB, [string]$Text,
          [System.Drawing.Color]$Color = $null)
    if ($null -eq $Color) { $Color = $clrText }
    $RTB.SelectionStart  = $RTB.TextLength
    $RTB.SelectionLength = 0
    $RTB.SelectionColor  = $Color
    $RTB.AppendText($Text + "`r`n")
    $RTB.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-Result {
    param([System.Windows.Forms.RichTextBox]$RTB, [string]$Label,
          [string]$Status, [string]$Detail = "")
    $icon  = switch ($Status) { 'PASS' { "[+]" } 'FAIL' { "[X]" } 'WARN' { "[!]" } 'INFO' { "[i]" } default { "[ ]" } }
    $color = switch ($Status) { 'PASS' { $clrSuccess } 'FAIL' { $clrError } 'WARN' { $clrWarn } default { $clrText } }
    Append-Output -RTB $RTB -Text "$icon $Label" -Color $color
    if ($Detail) { Append-Output -RTB $RTB -Text "    $Detail" -Color $clrTextDim }
    Write-Log "$Status | $Label | $Detail"
}

function New-RichOutput {
    param([System.Drawing.Point]$Location, [System.Drawing.Size]$Size)
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Location    = $Location
    $rtb.Size        = $Size
    $rtb.ReadOnly    = $true
    $rtb.BackColor   = $clrBg
    $rtb.ForeColor   = $clrText
    $rtb.Font        = $fontMono
    $rtb.BorderStyle = 'None'
    $rtb.ScrollBars  = 'Vertical'
    $rtb.WordWrap    = $false
    return $rtb
}

function Get-ADCred {
    $cred = Get-ToolCredential
    if ($null -eq $cred) { return $null }
    return $cred
}

function Invoke-WithSpinner {
    param([System.Windows.Forms.Button]$Btn, [scriptblock]$Action)
    $orig = $Btn.Text
    $Btn.Text    = "Running..."
    $Btn.Enabled = $false
    [System.Windows.Forms.Application]::DoEvents()
    try { & $Action } finally {
        $Btn.Text    = $orig
        $Btn.Enabled = $true
    }
}

#endregion

#region --- Main Form ---

$form = New-Object System.Windows.Forms.Form
$form.Text            = "  Active Directory — Kerberos Troubleshooter v2.0"
$form.Size            = New-Object System.Drawing.Size(1440, 980)
$form.MinimumSize     = New-Object System.Drawing.Size(900, 640)
$form.StartPosition   = 'CenterScreen'
$form.BackColor       = $clrBg
$form.ForeColor       = $clrText
$form.Font            = $fontUI
$form.FormBorderStyle = 'Sizable'
$form.Icon            = [System.Drawing.SystemIcons]::Shield

#endregion

#region --- Header Banner ---

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock      = 'Top'
$pnlHeader.Height    = 65
$pnlHeader.BackColor = $clrPanel

$lblTitle = New-StyledLabel -Text "  🔐  Active Directory  |  Kerberos Troubleshooter" `
    -Location (New-Object System.Drawing.Point(10, 18)) `
    -Size     (New-Object System.Drawing.Size(540, 18)) `
    -Font     $fontTitle -ForeColor $clrAccent

$lblVersion = New-StyledLabel -Text "v2.0  |  Log: $LogFile" `
    -Location (New-Object System.Drawing.Point(40, 36)) `
    -Size     (New-Object System.Drawing.Size(700, 22)) `
    -Font     $fontSmall -ForeColor $clrTextDim

$pnlHeader.Controls.AddRange(@($lblTitle, $lblVersion))
$form.Controls.Add($pnlHeader)

#endregion

#region --- Status Bar ---

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $clrPanel
$statusBar.ForeColor = $clrText

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text      = "Ready — configure connection settings in the Connection tab first."
$statusLabel.ForeColor = $clrTextDim
$statusLabel.Font      = $fontSmall

$statusDomain = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusDomain.Text      = "Domain: Not Connected"
$statusDomain.ForeColor = $clrWarn
$statusDomain.Font      = $fontUIB
$statusDomain.Spring    = $true
$statusDomain.TextAlign = 'MiddleRight'

$statusBar.Items.AddRange(@($statusLabel, $statusDomain))
$form.Controls.Add($statusBar)

#endregion

#region --- Tab Control ---

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock             = 'Fill'
$tabs.Font             = $fontUIB
$tabs.DrawMode         = 'OwnerDrawFixed'
$tabs.ItemSize         = New-Object System.Drawing.Size(148, 34)
$tabs.SizeMode         = 'Fixed'
$tabs.Padding          = New-Object System.Drawing.Point(10, 6)

# Custom tab painting
$tabs.Add_DrawItem({
    param($sender, $e)
    $tab    = $sender.TabPages[$e.Index]
    $rect   = $e.Bounds
    $sel    = ($e.Index -eq $sender.SelectedIndex)
    $bgClr  = if ($sel) { $clrAccent }    else { $clrTabNorm }
    $fgClr  = if ($sel) { [System.Drawing.Color]::White } else { $clrTextDim }
    $brush  = New-Object System.Drawing.SolidBrush($bgClr)
    $e.Graphics.FillRectangle($brush, $rect)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment     = 'Center'
    $sf.LineAlignment = 'Center'
    $txtBrush = New-Object System.Drawing.SolidBrush($fgClr)
    $e.Graphics.DrawString($tab.Text, $fontUIB, $txtBrush, [System.Drawing.RectangleF]$rect, $sf)
    $brush.Dispose(); $txtBrush.Dispose()
})

$form.Controls.Add($tabs)

function New-TabPage {
    param([string]$Title)
    $tp = New-Object System.Windows.Forms.TabPage
    $tp.Text      = $Title
    $tp.BackColor = $clrBg
    $tp.ForeColor = $clrText
    $tp.Padding   = New-Object System.Windows.Forms.Padding(10)
    return $tp
}

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 1 — CONNECTION SETTINGS
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Connection

$tabConn = New-TabPage "⚙ Connection"
$tabs.TabPages.Add($tabConn)

# Domain group
$gbDomain = New-GroupBox -Text "Domain / DC Target" `
    -Location (New-Object System.Drawing.Point(10, 10)) `
    -Size     (New-Object System.Drawing.Size(490, 180))

$tabConn.Controls.Add($gbDomain)

$lblDomName  = New-StyledLabel "Target Domain (FQDN):" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(180,20))
$tbDomain    = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,26)) -Size (New-Object System.Drawing.Size(270,22)) -Text $env:USERDNSDOMAIN

$lblDCName   = New-StyledLabel "Domain Controller (optional):" (New-Object System.Drawing.Point(12,62)) (New-Object System.Drawing.Size(185,20))
$tbDC        = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,60)) -Size (New-Object System.Drawing.Size(270,22)) -Text ""

$lblSite     = New-StyledLabel "Site Name (optional):" (New-Object System.Drawing.Point(12,96)) (New-Object System.Drawing.Size(185,20))
$tbSite      = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,94)) -Size (New-Object System.Drawing.Size(270,22))

$lblDCNote   = New-StyledLabel "Leave DC blank to auto-discover via DNS." `
    -Location (New-Object System.Drawing.Point(12,130)) -Size (New-Object System.Drawing.Size(460,16)) `
    -Font $fontSmall -ForeColor $clrTextDim

$cbUseCurrent = New-Object System.Windows.Forms.CheckBox
$cbUseCurrent.Text      = "Use current domain ($env:USERDNSDOMAIN)"
$cbUseCurrent.Location  = New-Object System.Drawing.Point(12,150)
$cbUseCurrent.Size      = New-Object System.Drawing.Size(350,20)
$cbUseCurrent.ForeColor = $clrText
$cbUseCurrent.BackColor = [System.Drawing.Color]::Transparent
$cbUseCurrent.Font      = $fontUI
$cbUseCurrent.Checked   = $true

$cbUseCurrent.Add_CheckedChanged({
    if ($cbUseCurrent.Checked) {
        $tbDomain.Text    = $env:USERDNSDOMAIN
        $tbDomain.Enabled = $false
    } else {
        $tbDomain.Enabled = $true
        $tbDomain.Focus()
    }
})
$tbDomain.Enabled = $false

$gbDomain.Controls.AddRange(@($lblDomName,$tbDomain,$lblDCName,$tbDC,$lblSite,$tbSite,$lblDCNote,$cbUseCurrent))

# Credentials group
$gbCred = New-GroupBox -Text "Alternate Credentials" `
    -Location (New-Object System.Drawing.Point(10, 200)) `
    -Size     (New-Object System.Drawing.Size(490, 190))
$tabConn.Controls.Add($gbCred)

$cbAltCred = New-Object System.Windows.Forms.CheckBox
$cbAltCred.Text      = "Use alternate credentials"
$cbAltCred.Location  = New-Object System.Drawing.Point(12,24)
$cbAltCred.Size      = New-Object System.Drawing.Size(280,20)
$cbAltCred.ForeColor = $clrText
$cbAltCred.BackColor = [System.Drawing.Color]::Transparent
$cbAltCred.Font      = $fontUI

$lblCredUser  = New-StyledLabel "Username (DOMAIN\user):" (New-Object System.Drawing.Point(12,54)) (New-Object System.Drawing.Size(180,20))
$tbCredUser   = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,52)) -Size (New-Object System.Drawing.Size(270,22))
$tbCredUser.Enabled = $false

$lblCredPass  = New-StyledLabel "Password:" (New-Object System.Drawing.Point(12,88)) (New-Object System.Drawing.Size(180,20))
$tbCredPass   = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,86)) -Size (New-Object System.Drawing.Size(270,22)) -Password $true
$tbCredPass.Enabled = $false

$btnSecureCred = New-StyledButton -Text "Use Get-Credential" `
    -Location (New-Object System.Drawing.Point(12,118)) `
    -Size (New-Object System.Drawing.Size(160,28)) `
    -BgColor ([System.Drawing.Color]::FromArgb(55,80,130))
$btnSecureCred.Enabled = $false

$lblCredStatus = New-StyledLabel "Status: Using current session credentials." `
    -Location (New-Object System.Drawing.Point(12,156)) -Size (New-Object System.Drawing.Size(460,18)) `
    -Font $fontSmall -ForeColor $clrTextDim

$cbAltCred.Add_CheckedChanged({
    $e = $cbAltCred.Checked
    $tbCredUser.Enabled   = $e
    $tbCredPass.Enabled   = $e
    $btnSecureCred.Enabled = $e
    if (-not $e) {
        $script:SavedCred     = $null
        $lblCredStatus.Text   = "Status: Using current session credentials."
        $lblCredStatus.ForeColor = $clrTextDim
    }
})

$btnSecureCred.Add_Click({
    try {
        $cred = Get-Credential -Message "Enter credentials for AD access" -UserName $tbCredUser.Text
        if ($cred) {
            $script:SavedCred   = $cred
            $tbCredUser.Text    = $cred.UserName
            $tbCredPass.Text    = "••••••••"
            $lblCredStatus.Text = "Status: Credential captured for $($cred.UserName)"
            $lblCredStatus.ForeColor = $clrSuccess
        }
    } catch { }
})

$gbCred.Controls.AddRange(@($cbAltCred,$lblCredUser,$tbCredUser,$lblCredPass,$tbCredPass,$btnSecureCred,$lblCredStatus))

# Connection button + info
$btnConnect = New-StyledButton -Text "▶  Apply & Connect" `
    -Location (New-Object System.Drawing.Point(10, 400)) `
    -Size     (New-Object System.Drawing.Size(200, 36))

$lblConnResult = New-StyledLabel "" `
    -Location (New-Object System.Drawing.Point(220, 408)) `
    -Size     (New-Object System.Drawing.Size(350, 22)) `
    -Font     $fontUIB

$tabConn.Controls.AddRange(@($btnConnect, $lblConnResult))

# Connection info panel (right side)
$gbConnInfo = New-GroupBox -Text "Current Environment Info" `
    -Location (New-Object System.Drawing.Point(515, 10)) `
    -Size     (New-Object System.Drawing.Size(490, 430))
$tabConn.Controls.Add($gbConnInfo)

$rtbConnInfo = New-RichOutput (New-Object System.Drawing.Point(8,20)) (New-Object System.Drawing.Size(470,398))
$gbConnInfo.Controls.Add($rtbConnInfo)

# Populate environment info on load
$btnConnect.Add_Click({
    Invoke-WithSpinner -Btn $btnConnect -Action {
        $rtbConnInfo.Clear()

        # Build credential for AD
        if ($cbAltCred.Checked -and $tbCredUser.Text -and $tbCredPass.Text -ne "••••••••") {
            $secPw = ConvertTo-SecureString $tbCredPass.Text -AsPlainText -Force
            $script:SavedCred = New-Object System.Management.Automation.PSCredential($tbCredUser.Text, $secPw)
        }

        $domain = $tbDomain.Text.Trim()
        $dc     = $tbDC.Text.Trim()
        if (-not $domain) {
            $lblConnResult.Text      = "✗ Domain cannot be empty."
            $lblConnResult.ForeColor = $clrError
            return
        }

        $script:ActiveDomain = $domain

        # Local machine info
        Append-Output $rtbConnInfo "═══ Local Machine ═══════════════════" $clrAccent
        Append-Output $rtbConnInfo "  Computer  : $env:COMPUTERNAME"
        Append-Output $rtbConnInfo "  User      : $env:USERDOMAIN\$env:USERNAME"
        Append-Output $rtbConnInfo "  OS        : $((Get-CimInstance Win32_OperatingSystem).Caption)"
        Append-Output $rtbConnInfo "  Domain    : $env:USERDNSDOMAIN"

        # DNS resolution test
        Append-Output $rtbConnInfo "`n═══ DNS Resolution ══════════════════" $clrAccent
        try {
            $dns = [System.Net.Dns]::GetHostAddresses($domain)
            Append-Output $rtbConnInfo "  [$domain]" $clrSuccess
            foreach ($ip in $dns) { Append-Output $rtbConnInfo "    → $($ip.IPAddressToString)" }
        } catch {
            Append-Output $rtbConnInfo "  [X] DNS resolution FAILED for $domain" $clrError
        }

        # AD Module + DC locator
        Append-Output $rtbConnInfo "`n═══ AD Module & DC Discovery ════════" $clrAccent
        $adMod = Get-Module -ListAvailable -Name ActiveDirectory
        if ($adMod) {
            Append-Output $rtbConnInfo "  [+] ActiveDirectory module found (v$($adMod.Version))" $clrSuccess
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue
            try {
                $adParams = @{ Identity = $domain }
                if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
                if ($dc) { $adParams.Server = $dc }
                $adDomain = Get-ADDomain @adParams -ErrorAction Stop
                Append-Output $rtbConnInfo "  [+] Connected to domain: $($adDomain.DNSRoot)" $clrSuccess
                Append-Output $rtbConnInfo "       PDC Emulator : $($adDomain.PDCEmulator)"
                Append-Output $rtbConnInfo "       Forest       : $($adDomain.Forest)"
                Append-Output $rtbConnInfo "       Func. Level  : $($adDomain.DomainMode)"
                $statusDomain.Text      = "Domain: $($adDomain.DNSRoot)  |  PDC: $($adDomain.PDCEmulator)"
                $statusDomain.ForeColor = $clrSuccess
                $lblConnResult.Text      = "✔ Connected to $($adDomain.DNSRoot)"
                $lblConnResult.ForeColor = $clrSuccess
            } catch {
                Append-Output $rtbConnInfo "  [X] AD connection failed: $($_.Exception.Message)" $clrError
                $lblConnResult.Text      = "✗ AD connection failed."
                $lblConnResult.ForeColor = $clrError
            }
        } else {
            Append-Output $rtbConnInfo "  [!] ActiveDirectory module NOT found." $clrWarn
            Append-Output $rtbConnInfo "      Install RSAT: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory*'" $clrTextDim
        }

        # Time check (Kerberos critical)
        Append-Output $rtbConnInfo "`n═══ System Time ═════════════════════" $clrAccent
        $localTime = Get-Date
        try {
            $dcTarget = if ($dc) { $dc } else { $domain }
            $w32tm = w32tm /query /status 2>&1
            Append-Output $rtbConnInfo "  Local Time  : $($localTime.ToString('yyyy-MM-dd HH:mm:ss'))"
            $stratum = ($w32tm | Select-String "Stratum" | Select-Object -First 1).ToString().Trim()
            $source  = ($w32tm | Select-String "Source"  | Select-Object -First 1).ToString().Trim()
            Append-Output $rtbConnInfo "  $stratum"
            Append-Output $rtbConnInfo "  $source"
        } catch {
            Append-Output $rtbConnInfo "  [!] Could not query w32tm." $clrWarn
        }

        Write-Log "Connection applied to domain: $domain"
    }
})

# Auto-populate on tab load
$tabConn.Add_Enter({
    if ($rtbConnInfo.TextLength -eq 0) {
        Append-Output $rtbConnInfo "Click '▶ Apply & Connect' to begin." $clrTextDim
    }
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 2 — KERBEROS TICKETS (KLIST)
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Tickets

$tabTickets = New-TabPage "🎫 Kerberos Tickets"
$tabs.TabPages.Add($tabTickets)

$gbKlistCtrl = New-GroupBox -Text "Ticket Controls" `
    -Location (New-Object System.Drawing.Point(10, 10)) `
    -Size     (New-Object System.Drawing.Size(280, 230))
$tabTickets.Controls.Add($gbKlistCtrl)

$btnKlist    = New-StyledButton "List All Tickets"    (New-Object System.Drawing.Point(12,28))  (New-Object System.Drawing.Size(250,30))
$btnKlistTgt = New-StyledButton "TGT Only"            (New-Object System.Drawing.Point(12,68))  (New-Object System.Drawing.Size(250,30))
$btnPurge    = New-StyledButton "Purge Tickets ⚠"    (New-Object System.Drawing.Point(12,108)) (New-Object System.Drawing.Size(250,30)) -BgColor ([System.Drawing.Color]::FromArgb(130,55,55))
$btnKInit    = New-StyledButton "New Ticket (kinit)"  (New-Object System.Drawing.Point(12,148)) (New-Object System.Drawing.Size(250,30))
$btnCopyTkt  = New-StyledButton "Copy Output"         (New-Object System.Drawing.Point(12,190)) (New-Object System.Drawing.Size(250,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,55))
$gbKlistCtrl.Controls.AddRange(@($btnKlist,$btnKlistTgt,$btnPurge,$btnKInit,$btnCopyTkt))

$gbTicketInfo = New-GroupBox -Text "Ticket Analysis" `
    -Location (New-Object System.Drawing.Point(300, 10)) `
    -Size     (New-Object System.Drawing.Size(720, 230))
$tabTickets.Controls.Add($gbTicketInfo)

$rtbTickets = New-RichOutput (New-Object System.Drawing.Point(8,20)) (New-Object System.Drawing.Size(700,625))
$rtbTickets.WordWrap = $false
$tabTickets.Controls.Add($rtbTickets)
$rtbTickets.Location = New-Object System.Drawing.Point(10,250)
$rtbTickets.Size     = New-Object System.Drawing.Size(1010,390)

# Summary boxes inside gbTicketInfo
$lblTktCount  = New-StyledLabel "Tickets Found: —"   (New-Object System.Drawing.Point(10,30))  (New-Object System.Drawing.Size(180,20)) $fontUIB $clrText
$lblTgtExp    = New-StyledLabel "TGT Expires: —"     (New-Object System.Drawing.Point(10,58))  (New-Object System.Drawing.Size(340,20)) $fontUI  $clrText
$lblEncType   = New-StyledLabel "Enc Types: —"       (New-Object System.Drawing.Point(10,82))  (New-Object System.Drawing.Size(700,20)) $fontUI  $clrText
$lblDES       = New-StyledLabel ""                   (New-Object System.Drawing.Point(10,106)) (New-Object System.Drawing.Size(700,20)) $fontUIB $clrWarn
$gbTicketInfo.Controls.AddRange(@($lblTktCount,$lblTgtExp,$lblEncType,$lblDES))

function Parse-KlistOutput {
    param([string[]]$Lines, [System.Windows.Forms.RichTextBox]$RTB)
    $RTB.Clear()
    $ticketCount = 0
    $tgtExpiry   = ""
    $encTypes    = @()
    $hasDES      = $false

    foreach ($line in $Lines) {
        if ($line -match "^#\d+") {
            $ticketCount++
            Append-Output $RTB "" 
            Append-Output $RTB $line $clrAccent
        } elseif ($line -match "Server:") {
            Append-Output $RTB $line $clrText
        } elseif ($line -match "KerbTicket Encryption Type:") {
            $enc = ($line -replace ".*KerbTicket Encryption Type:\s*","").Trim()
            $encTypes += $enc
            $c = if ($enc -match "AES256") { $clrSuccess } elseif ($enc -match "DES|RC4") { $clrWarn } else { $clrText }
            Append-Output $RTB $line $c
            if ($enc -match "DES") { $hasDES = $true }
        } elseif ($line -match "End Time:") {
            if ($line -match "krbtgt" -or $tgtExpiry -eq "") { $tgtExpiry = $line }
            Append-Output $RTB $line $clrTextDim
        } elseif ($line -match "Ticket Flags") {
            Append-Output $RTB $line ([System.Drawing.Color]::FromArgb(180,180,100))
        } else {
            Append-Output $RTB $line $clrTextDim
        }
    }

    $lblTktCount.Text = "Tickets Found: $ticketCount"
    $lblTgtExp.Text   = "TGT Expiry: $($tgtExpiry.Trim())"
    $lblEncType.Text  = "Enc Types: $($encTypes | Sort-Object -Unique | Join-String -Separator ', ')"
    if ($hasDES) { $lblDES.Text = "⚠ DES tickets detected — insecure encryption!" }
    else          { $lblDES.Text = "" }
}

$btnKlist.Add_Click({
    Invoke-WithSpinner -Btn $btnKlist -Action {
        $out = klist 2>&1
        Parse-KlistOutput -Lines $out -RTB $rtbTickets
        Write-Log "klist executed"
    }
})

$btnKlistTgt.Add_Click({
    Invoke-WithSpinner -Btn $btnKlistTgt -Action {
        $out = klist tgt 2>&1
        $rtbTickets.Clear()
        foreach ($l in $out) { Append-Output $rtbTickets $l $clrText }
        Write-Log "klist tgt executed"
    }
})

$btnPurge.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show(
        "This will purge ALL Kerberos tickets for the current session.`nYou may need to re-authenticate to resources.`n`nProceed?",
        "Purge Kerberos Tickets", 'YesNo', 'Warning')
    if ($r -eq 'Yes') {
        $out = klist purge 2>&1
        $rtbTickets.Clear()
        Append-Output $rtbTickets "Tickets purged at $(Get-Date -Format 'HH:mm:ss')" $clrWarn
        foreach ($l in $out) { Append-Output $rtbTickets $l $clrText }
        Write-Log "klist purge executed"
    }
})

$btnKInit.Add_Click({
    $domain = Get-ActiveDomain
    if (-not $domain) { $domain = $env:USERDNSDOMAIN }
    $out = klist get "krbtgt/$domain" 2>&1
    $rtbTickets.Clear()
    foreach ($l in $out) { Append-Output $rtbTickets $l $clrText }
    Write-Log "klist get krbtgt executed"
})

$btnCopyTkt.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($rtbTickets.Text)
    $btnCopyTkt.Text = "✔ Copied!"
    Start-Sleep -Milliseconds 1500
    $btnCopyTkt.Text = "Copy Output"
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 3 — PORT & CONNECTIVITY
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Connectivity

$tabConn2 = New-TabPage "🔌 Connectivity"
$tabs.TabPages.Add($tabConn2)

$gbPortCtrl = New-GroupBox -Text "Port / DC Connectivity Test" `
    -Location (New-Object System.Drawing.Point(10, 10)) -Size (New-Object System.Drawing.Size(490,200))
$tabConn2.Controls.Add($gbPortCtrl)

$lblPortTarget = New-StyledLabel "Target (DC or Domain):" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(180,20))
$tbPortTarget  = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,26)) -Size (New-Object System.Drawing.Size(270,22))

$lblPortSel    = New-StyledLabel "Port Set:" (New-Object System.Drawing.Point(12,60)) (New-Object System.Drawing.Size(180,20))
$cbPortSet     = New-Object System.Windows.Forms.ComboBox
$cbPortSet.Location = New-Object System.Drawing.Point(200,58)
$cbPortSet.Size     = New-Object System.Drawing.Size(270,22)
$cbPortSet.Font     = $fontUI
$cbPortSet.BackColor = $clrInput
$cbPortSet.ForeColor = $clrText
$cbPortSet.FlatStyle = 'Flat'
$cbPortSet.DropDownStyle = 'DropDownList'
@("Kerberos Only (88, 464)","Full AD Ports","LDAP/LDAPS","Kerberos + LDAP","All") | ForEach-Object { $cbPortSet.Items.Add($_) | Out-Null }
$cbPortSet.SelectedIndex = 0
$gbPortCtrl.Controls.AddRange(@($lblPortTarget,$tbPortTarget,$lblPortSel,$cbPortSet))

$btnTestPorts  = New-StyledButton "Test Ports" (New-Object System.Drawing.Point(12,94))  (New-Object System.Drawing.Size(150,30))
$btnTestDCList = New-StyledButton "Discover DCs" (New-Object System.Drawing.Point(170,94)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))
$btnTestNLTest = New-StyledButton "nltest /dsgetdc" (New-Object System.Drawing.Point(328,94)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))

$gbPortCtrl.Controls.AddRange(@($btnTestPorts,$btnTestDCList,$btnTestNLTest))

$lblPortNote = New-StyledLabel "Kerberos requires TCP/UDP 88 (KDC), 464 (kpasswd), 389/636 (LDAP), 3268/3269 (GC)" `
    -Location (New-Object System.Drawing.Point(12,136)) -Size (New-Object System.Drawing.Size(465,18)) -Font $fontSmall -ForeColor $clrTextDim
$gbPortCtrl.Controls.Add($lblPortNote)

# Port Legend
$gbLegend = New-GroupBox -Text "Kerberos Port Reference" `
    -Location (New-Object System.Drawing.Point(515, 10)) -Size (New-Object System.Drawing.Size(490,200))
$tabConn2.Controls.Add($gbLegend)

$portRef = @(
    "Port 88   TCP/UDP  Kerberos KDC (authentication tickets)"
    "Port 464  TCP/UDP  Kerberos kpasswd (password changes)"
    "Port 389  TCP/UDP  LDAP (directory queries)"
    "Port 636  TCP      LDAPS (secure LDAP)"
    "Port 3268 TCP      Global Catalog"
    "Port 3269 TCP      Global Catalog SSL"
    "Port 445  TCP      SMB / DCE-RPC"
    "Port 135  TCP      RPC Endpoint Mapper"
    "Port 53   TCP/UDP  DNS (critical for Kerberos)"
    "Port 49152-65535 TCP  Dynamic RPC"
)
$rtbLegend = New-RichOutput (New-Object System.Drawing.Point(8,20)) (New-Object System.Drawing.Size(472,170))
$gbLegend.Controls.Add($rtbLegend)
foreach ($l in $portRef) { Append-Output $rtbLegend $l $clrTextDim }

# Main output
$rtbPorts = New-RichOutput (New-Object System.Drawing.Point(10,220)) (New-Object System.Drawing.Size(1010,420))
$tabConn2.Controls.Add($rtbPorts)

$portSets = @{
    "Kerberos Only (88, 464)"  = @(88,464)
    "Full AD Ports"            = @(88,464,389,636,3268,3269,445,135,53,9389)
    "LDAP/LDAPS"               = @(389,636,3268,3269)
    "Kerberos + LDAP"          = @(88,464,389,636)
    "All"                      = @(53,88,135,389,445,464,636,3268,3269,9389)
}

$btnTestPorts.Add_Click({
    Invoke-WithSpinner -Btn $btnTestPorts -Action {
        $target = $tbPortTarget.Text.Trim()
        if (-not $target) { $target = Get-ActiveDomain; if (-not $target) { $target = $env:USERDNSDOMAIN } }
        $ports = $portSets[$cbPortSet.Text]
        $rtbPorts.Clear()
        Append-Output $rtbPorts "Port Connectivity Test → $target" $clrAccent
        Append-Output $rtbPorts "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" $clrTextDim
        Append-Output $rtbPorts ""

        $pass = 0; $fail = 0
        foreach ($port in $ports) {
            try {
                $result = Test-NetConnection -ComputerName $target -Port $port -WarningAction SilentlyContinue -ErrorAction Stop
                $portName = switch ($port) {
                    88 {"Kerberos KDC"} 464 {"kpasswd"} 389 {"LDAP"} 636 {"LDAPS"}
                    3268 {"Global Catalog"} 3269 {"GC SSL"} 445 {"SMB"} 135 {"RPC Mapper"}
                    53  {"DNS"} 9389 {"AD Web Svc"} default {"Unknown"}
                }
                if ($result.TcpTestSucceeded) {
                    Write-Result $rtbPorts "Port $port ($portName)" 'PASS' "Latency: $($result.PingReplyDetails.RoundtripTime)ms"
                    $pass++
                } else {
                    Write-Result $rtbPorts "Port $port ($portName)" 'FAIL' "TCP connection refused or timed out"
                    $fail++
                }
            } catch {
                Write-Result $rtbPorts "Port $port" 'FAIL' $_.Exception.Message
                $fail++
            }
        }
        Append-Output $rtbPorts ""
        Append-Output $rtbPorts "Results: $pass PASS  |  $fail FAIL" $(if ($fail -eq 0) { $clrSuccess } else { $clrWarn })
        Write-Log "Port test complete: $pass pass, $fail fail against $target"
    }
})

$btnTestDCList.Add_Click({
    Invoke-WithSpinner -Btn $btnTestDCList -Action {
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $rtbPorts.Clear()
        Append-Output $rtbPorts "Discovering Domain Controllers for: $domain" $clrAccent

        try {
            $adParams = @{}
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
            $dcs = Get-ADDomainController -Filter * -Server $domain @adParams | Select-Object Name, Site, IPv4Address, IsGlobalCatalog, OperationMasterRoles
            foreach ($dc in $dcs) {
                $roles = if ($dc.OperationMasterRoles) { $dc.OperationMasterRoles -join ", " } else { "None" }
                Append-Output $rtbPorts ""
                Append-Output $rtbPorts "  [$($dc.Name)]" $clrAccent
                Append-Output $rtbPorts "    IPv4     : $($dc.IPv4Address)"
                Append-Output $rtbPorts "    Site     : $($dc.Site)"
                Append-Output $rtbPorts "    GC       : $($dc.IsGlobalCatalog)"
                Append-Output $rtbPorts "    FSMO     : $roles" $clrTextDim
            }
        } catch {
            Append-Output $rtbPorts "ActiveDirectory module unavailable, falling back to DNS..." $clrWarn
            $nlout = nltest /dclist:$domain 2>&1
            foreach ($l in $nlout) { Append-Output $rtbPorts $l $clrText }
        }
    }
})

$btnTestNLTest.Add_Click({
    Invoke-WithSpinner -Btn $btnTestNLTest -Action {
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $rtbPorts.Clear()
        Append-Output $rtbPorts "nltest /dsgetdc:$domain" $clrAccent
        $out = nltest /dsgetdc:$domain 2>&1
        foreach ($l in $out) {
            $c = if ($l -match "ERROR|FAIL") { $clrError } elseif ($l -match "\\\\") { $clrSuccess } else { $clrText }
            Append-Output $rtbPorts $l $c
        }
        Append-Output $rtbPorts ""
        Append-Output $rtbPorts "nltest /sc_query:$domain" $clrAccent
        $out2 = nltest /sc_query:$domain 2>&1
        foreach ($l in $out2) {
            $c = if ($l -match "ERROR|FAIL") { $clrError } elseif ($l -match "SUCCESS|CONNECTED") { $clrSuccess } else { $clrText }
            Append-Output $rtbPorts $l $c
        }
        Write-Log "nltest /dsgetdc executed for $domain"
    }
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 4 — TIME SYNC (Critical for Kerberos)
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Time Sync

$tabTime = New-TabPage "⏰ Time Sync"
$tabs.TabPages.Add($tabTime)

$gbTimeCtrl = New-GroupBox -Text "Time Synchronization Diagnostics" `
    -Location (New-Object System.Drawing.Point(10,10)) -Size (New-Object System.Drawing.Size(490,160))
$tabTime.Controls.Add($gbTimeCtrl)

$btnW32Status  = New-StyledButton "w32tm Status"     (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(150,30))
$btnW32Peers   = New-StyledButton "NTP Peers"        (New-Object System.Drawing.Point(170,28)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))
$btnW32Resync  = New-StyledButton "Force Resync"     (New-Object System.Drawing.Point(328,28)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(130,55,55))
$btnTimeCheck  = New-StyledButton "Kerberos Skew Check" (New-Object System.Drawing.Point(12,68)) (New-Object System.Drawing.Size(200,30))
$btnDCTimeComp = New-StyledButton "Compare DC Times"    (New-Object System.Drawing.Point(220,68)) (New-Object System.Drawing.Size(200,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))

$lblTimeNote = New-StyledLabel "Kerberos allows maximum 5 minutes clock skew. Exceeding this causes KRB_AP_ERR_SKEW errors." `
    -Location (New-Object System.Drawing.Point(12,110)) -Size (New-Object System.Drawing.Size(460,18)) -Font $fontSmall -ForeColor $clrWarn

$gbTimeCtrl.Controls.AddRange(@($btnW32Status,$btnW32Peers,$btnW32Resync,$btnTimeCheck,$btnDCTimeComp,$lblTimeNote))

$rtbTime = New-RichOutput (New-Object System.Drawing.Point(10,180)) (New-Object System.Drawing.Size(1010,460))
$tabTime.Controls.Add($rtbTime)

# Skew summary banner
$gbSkew = New-GroupBox -Text "Clock Skew Status" `
    -Location (New-Object System.Drawing.Point(515,10)) -Size (New-Object System.Drawing.Size(490,160))
$tabTime.Controls.Add($gbSkew)

$lblSkewVal  = New-StyledLabel "Skew: Not calculated" (New-Object System.Drawing.Point(12,30)) (New-Object System.Drawing.Size(460,30)) $fontTitle $clrText
$lblSkewStat = New-StyledLabel "Run 'Kerberos Skew Check' to measure." (New-Object System.Drawing.Point(12,70)) (New-Object System.Drawing.Size(460,20)) $fontUI $clrTextDim
$lblSkewFix  = New-StyledLabel "" (New-Object System.Drawing.Point(12,96)) (New-Object System.Drawing.Size(460,50)) $fontSmall $clrWarn
$gbSkew.Controls.AddRange(@($lblSkewVal,$lblSkewStat,$lblSkewFix))

$btnW32Status.Add_Click({
    Invoke-WithSpinner -Btn $btnW32Status -Action {
        $rtbTime.Clear()
        Append-Output $rtbTime "w32tm /query /status" $clrAccent
        $out = w32tm /query /status 2>&1
        foreach ($l in $out) {
            $c = if ($l -match "Error|FAIL") { $clrError } elseif ($l -match "Source|Stratum") { $clrSuccess } else { $clrText }
            Append-Output $rtbTime $l $c
        }
        Write-Log "w32tm /query /status executed"
    }
})

$btnW32Peers.Add_Click({
    Invoke-WithSpinner -Btn $btnW32Peers -Action {
        $rtbTime.Clear()
        Append-Output $rtbTime "w32tm /query /peers" $clrAccent
        $out = w32tm /query /peers 2>&1
        foreach ($l in $out) { Append-Output $rtbTime $l $clrText }
        Append-Output $rtbTime "" 
        Append-Output $rtbTime "w32tm /query /configuration" $clrAccent
        $out2 = w32tm /query /configuration 2>&1
        foreach ($l in $out2) { Append-Output $rtbTime $l $clrText }
        Write-Log "w32tm peers/config queried"
    }
})

$btnW32Resync.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show("Force w32tm resync with /force flag?`nThis may briefly disrupt time services.","Force Resync",'YesNo','Warning')
    if ($r -eq 'Yes') {
        Invoke-WithSpinner -Btn $btnW32Resync -Action {
            $rtbTime.Clear()
            Append-Output $rtbTime "w32tm /resync /force" $clrAccent
            $out = w32tm /resync /force 2>&1
            foreach ($l in $out) {
                $c = if ($l -match "success") { $clrSuccess } elseif ($l -match "error|fail") { $clrError } else { $clrText }
                Append-Output $rtbTime $l $c
            }
            Write-Log "w32tm /resync /force executed"
        }
    }
})

$btnTimeCheck.Add_Click({
    Invoke-WithSpinner -Btn $btnTimeCheck -Action {
        $rtbTime.Clear()
        $domain  = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        Append-Output $rtbTime "Kerberos Clock Skew Analysis → $domain" $clrAccent
        Append-Output $rtbTime "Local Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')" $clrText

        try {
            $dcTarget = (nltest /dsgetdc:$domain 2>&1 | Select-String '\\\\') -replace '.*\\\\([^\s]+).*','$1'
            if ($dcTarget) {
                $dcTime = ([System.Net.Sockets.TcpClient]::new()) | ForEach-Object { $null }
                # Use .NET to get DC time via WMI
                $wmiArgs = @{ ComputerName = $dcTarget; Class = "Win32_UTCTime"; ErrorAction = "Stop" }
                if ($script:SavedCred) { $wmiArgs.Credential = $script:SavedCred }
                $utcObj = Get-CimInstance @wmiArgs
                $dcDate = [datetime]::new($utcObj.Year,$utcObj.Month,$utcObj.Day,$utcObj.Hour,$utcObj.Minute,$utcObj.Second,[DateTimeKind]::Utc)
                $localUtc = (Get-Date).ToUniversalTime()
                $skew = [Math]::Abs(($localUtc - $dcDate).TotalSeconds)
                $skewMin = [Math]::Round($skew / 60, 2)

                Append-Output $rtbTime "DC ($dcTarget) UTC : $($dcDate.ToString('yyyy-MM-dd HH:mm:ss'))" $clrText
                Append-Output $rtbTime "Local UTC          : $($localUtc.ToString('yyyy-MM-dd HH:mm:ss'))" $clrText
                Append-Output $rtbTime "Skew               : $skew seconds ($skewMin minutes)" $(if ($skew -lt 60) { $clrSuccess } elseif ($skew -lt 300) { $clrWarn } else { $clrError })

                $lblSkewVal.Text  = "Skew: $skew sec  ($skewMin min)"
                if ($skew -lt 60) {
                    $lblSkewVal.ForeColor  = $clrSuccess
                    $lblSkewStat.Text      = "✔ Excellent — well within 5-minute Kerberos tolerance."
                    $lblSkewStat.ForeColor = $clrSuccess
                    $lblSkewFix.Text = ""
                } elseif ($skew -lt 300) {
                    $lblSkewVal.ForeColor  = $clrWarn
                    $lblSkewStat.Text      = "⚠ Warning — approaching the 5-minute limit."
                    $lblSkewStat.ForeColor = $clrWarn
                    $lblSkewFix.Text = "Run: w32tm /resync /force  to correct."
                } else {
                    $lblSkewVal.ForeColor  = $clrError
                    $lblSkewStat.Text      = "✗ CRITICAL — skew exceeds 5 minutes! Kerberos will FAIL."
                    $lblSkewStat.ForeColor = $clrError
                    $lblSkewFix.Text = "Fix: net stop w32tm; w32tm /unregister; w32tm /register; net start w32tm; w32tm /resync /force"
                }
                Write-Result $rtbTime "Clock Skew" $(if ($skew -lt 300) { 'PASS' } else { 'FAIL' }) "$skew seconds ($skewMin minutes)"
            } else {
                Append-Output $rtbTime "[!] Could not resolve DC via nltest. Check Connection tab." $clrWarn
            }
        } catch {
            Append-Output $rtbTime "[X] Skew check failed: $($_.Exception.Message)" $clrError
            Append-Output $rtbTime "    Try: w32tm /stripchart /computer:<DC> /dataonly /samples:3" $clrTextDim
        }
        Write-Log "Kerberos skew check completed for $domain"
    }
})

$btnDCTimeComp.Add_Click({
    Invoke-WithSpinner -Btn $btnDCTimeComp -Action {
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $rtbTime.Clear()
        Append-Output $rtbTime "DC Time Comparison for $domain" $clrAccent
        try {
            $adParams = @{ Filter = "*"; Server = $domain }
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
            $dcs = Get-ADDomainController @adParams
            $localUtc = (Get-Date).ToUniversalTime()
            foreach ($dc in $dcs) {
                try {
                    $wmi = Get-CimInstance Win32_UTCTime -ComputerName $dc.HostName -ErrorAction Stop
                    $dcDt = [datetime]::new($wmi.Year,$wmi.Month,$wmi.Day,$wmi.Hour,$wmi.Minute,$wmi.Second,'Utc')
                    $sk = [Math]::Round([Math]::Abs(($localUtc - $dcDt).TotalSeconds),1)
                    $stat = if ($sk -lt 60) { 'PASS' } elseif ($sk -lt 300) { 'WARN' } else { 'FAIL' }
                    Write-Result $rtbTime "$($dc.Name)" $stat "Skew: $sk sec | DC Time: $($dcDt.ToString('HH:mm:ss')) UTC"
                } catch {
                    Write-Result $rtbTime "$($dc.Name)" 'WARN' "Could not reach: $($_.Exception.Message)"
                }
            }
        } catch {
            Append-Output $rtbTime "[X] AD module required for DC comparison." $clrError
        }
        Write-Log "DC time comparison completed"
    }
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 5 — SPN DIAGNOSTICS
# ════════════════════════════════════════════════════════════════════════════
#region TAB: SPN

$tabSPN = New-TabPage "🔎 SPN Check"
$tabs.TabPages.Add($tabSPN)

$gbSPNCtrl = New-GroupBox -Text "Service Principal Name Lookup" `
    -Location (New-Object System.Drawing.Point(10,10)) -Size (New-Object System.Drawing.Size(490,240))
$tabSPN.Controls.Add($gbSPNCtrl)

$lblSPNTarget = New-StyledLabel "Account / SPN / Service:" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(180,20))
$tbSPNTarget  = New-StyledTextBox -Location (New-Object System.Drawing.Point(200,26)) -Size (New-Object System.Drawing.Size(270,22))

$lblSPNType   = New-StyledLabel "Search Type:" (New-Object System.Drawing.Point(12,60)) (New-Object System.Drawing.Size(180,20))
$cbSPNType    = New-Object System.Windows.Forms.ComboBox
$cbSPNType.Location = New-Object System.Drawing.Point(200,58)
$cbSPNType.Size     = New-Object System.Drawing.Size(270,22)
$cbSPNType.Font     = $fontUI; $cbSPNType.BackColor = $clrInput; $cbSPNType.ForeColor = $clrText
$cbSPNType.FlatStyle = 'Flat'; $cbSPNType.DropDownStyle = 'DropDownList'
@("By Account (sAMAccountName)","By SPN String","By Service Class (e.g. HTTP)","Duplicate SPNs","All SPNs for Domain") | ForEach-Object { $cbSPNType.Items.Add($_) | Out-Null }
$cbSPNType.SelectedIndex = 0
$gbSPNCtrl.Controls.AddRange(@($lblSPNTarget,$tbSPNTarget,$lblSPNType,$cbSPNType))

$btnSPNSearch  = New-StyledButton "Search SPNs"   (New-Object System.Drawing.Point(12,94))  (New-Object System.Drawing.Size(150,30))
$btnSPNDupe    = New-StyledButton "Find Duplicates" (New-Object System.Drawing.Point(170,94)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(130,80,55))
$btnSetSPN     = New-StyledButton "Register SPN"  (New-Object System.Drawing.Point(328,94)) (New-Object System.Drawing.Size(150,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,110,55))

$lblSPNNote = New-StyledLabel "Duplicate SPNs are the #1 cause of Kerberos authentication failures. A KDC cannot determine which account to use." `
    -Location (New-Object System.Drawing.Point(12,138)) -Size (New-Object System.Drawing.Size(460,30)) -Font $fontSmall -ForeColor $clrWarn

$lblSPNTip = New-StyledLabel "Use 'setspn -X -F' for forest-wide duplicate scan (requires Domain Admin)." `
    -Location (New-Object System.Drawing.Point(12,170)) -Size (New-Object System.Drawing.Size(460,18)) -Font $fontSmall -ForeColor $clrTextDim

$gbSPNCtrl.Controls.AddRange(@($btnSPNSearch,$btnSPNDupe,$btnSetSPN,$lblSPNNote,$lblSPNTip))

$rtbSPN = New-RichOutput (New-Object System.Drawing.Point(10,260)) (New-Object System.Drawing.Size(1010,380))
$tabSPN.Controls.Add($rtbSPN)

$btnSPNSearch.Add_Click({
    Invoke-WithSpinner -Btn $btnSPNSearch -Action {
        $target = $tbSPNTarget.Text.Trim()
        $rtbSPN.Clear()
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        Append-Output $rtbSPN "SPN Search — Type: $($cbSPNType.Text)" $clrAccent
        Append-Output $rtbSPN "Domain: $domain  |  $(Get-Date -Format 'HH:mm:ss')" $clrTextDim
        Append-Output $rtbSPN ""

        try {
            $adParams = @{ Server = $domain }
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }

            switch ($cbSPNType.SelectedIndex) {
                0 { # By Account
                    if (-not $target) { Append-Output $rtbSPN "[!] Enter account name." $clrWarn; return }
                    $obj = Get-ADObject -Filter { SamAccountName -eq $target } -Properties servicePrincipalName @adParams -ErrorAction Stop
                    if ($obj) {
                        Append-Output $rtbSPN "Account: $($obj.DistinguishedName)" $clrSuccess
                        if ($obj.servicePrincipalName) {
                            foreach ($spn in $obj.servicePrincipalName) { Append-Output $rtbSPN "  → $spn" $clrText }
                        } else { Append-Output $rtbSPN "  No SPNs registered." $clrWarn }
                    } else { Append-Output $rtbSPN "[X] Account not found: $target" $clrError }
                }
                1 { # By SPN String
                    if (-not $target) { Append-Output $rtbSPN "[!] Enter SPN string." $clrWarn; return }
                    $results = Get-ADObject -Filter { servicePrincipalName -like $target } -Properties servicePrincipalName,SamAccountName @adParams
                    if ($results) {
                        foreach ($r in $results) {
                            Append-Output $rtbSPN "Account: $($r.SamAccountName)" $clrSuccess
                            foreach ($spn in $r.servicePrincipalName | Where-Object { $_ -like $target }) {
                                Append-Output $rtbSPN "  → $spn" $clrText
                            }
                        }
                    } else { Append-Output $rtbSPN "No objects found with SPN matching: $target" $clrWarn }
                }
                2 { # By Service Class
                    if (-not $target) { Append-Output $rtbSPN "[!] Enter service class (e.g. HTTP)." $clrWarn; return }
                    $filter = "$target/*"
                    $results = Get-ADObject -Filter { servicePrincipalName -like $filter } -Properties servicePrincipalName,SamAccountName @adParams
                    foreach ($r in $results) {
                        Append-Output $rtbSPN "  [$($r.SamAccountName)]" $clrAccent
                        foreach ($spn in $r.servicePrincipalName | Where-Object { $_ -like $filter }) {
                            Append-Output $rtbSPN "    → $spn" $clrText
                        }
                    }
                }
                3 { # Duplicate SPNs
                    Append-Output $rtbSPN "Running: setspn -X -F (forest-wide duplicate scan)..." $clrWarn
                    $out = setspn -X -F 2>&1
                    $dupe = $false
                    foreach ($l in $out) {
                        if ($l -match "^Checking") { Append-Output $rtbSPN $l $clrTextDim }
                        elseif ($l -match "duplicate") { Append-Output $rtbSPN $l $clrError; $dupe = $true }
                        elseif ($l -match "found") { Append-Output $rtbSPN $l $(if ($l -match "^0") { $clrSuccess } else { $clrError }) }
                        else { Append-Output $rtbSPN $l $clrText }
                    }
                    if (-not $dupe) { Append-Output $rtbSPN "[+] No duplicate SPNs detected." $clrSuccess }
                }
                4 { # All SPNs for domain
                    Append-Output $rtbSPN "Enumerating all SPNs in domain (may take a moment)..." $clrTextDim
                    $results = Get-ADObject -Filter { servicePrincipalName -like "*" } -Properties servicePrincipalName,SamAccountName,ObjectClass @adParams
                    $count = 0
                    foreach ($r in $results | Sort-Object SamAccountName) {
                        Append-Output $rtbSPN "  [$($r.ObjectClass)] $($r.SamAccountName)" $clrAccent
                        foreach ($spn in $r.servicePrincipalName | Sort-Object) {
                            Append-Output $rtbSPN "    $spn" $clrTextDim
                            $count++
                        }
                    }
                    Append-Output $rtbSPN "" ; Append-Output $rtbSPN "Total SPNs: $count" $clrText
                }
            }
        } catch {
            Append-Output $rtbSPN "[X] SPN search failed: $($_.Exception.Message)" $clrError
            Append-Output $rtbSPN "    Ensure ActiveDirectory module is available and connected." $clrTextDim
        }
        Write-Log "SPN search completed: $($cbSPNType.Text) → $target"
    }
})

$btnSPNDupe.Add_Click({
    Invoke-WithSpinner -Btn $btnSPNDupe -Action {
        $rtbSPN.Clear()
        Append-Output $rtbSPN "setspn -X -F   (Forest Duplicate Scan)" $clrAccent
        $out = setspn -X -F 2>&1
        foreach ($l in $out) {
            $c = if ($l -match "duplicate|Duplicate") { $clrError } elseif ($l -match "^0 duplicate") { $clrSuccess } else { $clrTextDim }
            Append-Output $rtbSPN $l $c
        }
        Write-Log "setspn -X -F executed"
    }
})

$btnSetSPN.Add_Click({
    $spnVal  = [Microsoft.VisualBasic.Interaction]::InputBox("Enter SPN to register:`n(e.g. HTTP/server.domain.com)", "Register SPN", "HTTP/")
    $acctVal = [Microsoft.VisualBasic.Interaction]::InputBox("Account to register SPN on:`n(e.g. DOMAIN\svcAccount)", "Target Account", "$env:USERDOMAIN\")
    if ($spnVal -and $acctVal) {
        $rtbSPN.Clear()
        Append-Output $rtbSPN "Registering SPN: $spnVal → $acctVal" $clrAccent
        $out = setspn -S $spnVal $acctVal 2>&1
        foreach ($l in $out) {
            $c = if ($l -match "Updated|Registering|Success") { $clrSuccess } elseif ($l -match "Error|Duplicate|exist") { $clrError } else { $clrText }
            Append-Output $rtbSPN $l $c
        }
        Write-Log "setspn -S $spnVal $acctVal"
    }
})

Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 6 — ACCOUNT CHECK
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Account

$tabAccount = New-TabPage "👤 Account Check"
$tabs.TabPages.Add($tabAccount)

$gbAccCtrl = New-GroupBox -Text "User / Computer Account Kerberos Status" `
    -Location (New-Object System.Drawing.Point(10,10)) -Size (New-Object System.Drawing.Size(490,180))
$tabAccount.Controls.Add($gbAccCtrl)

$lblAccName = New-StyledLabel "Account Name (sAMAccountName):" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(240,20))
$tbAccName  = New-StyledTextBox -Location (New-Object System.Drawing.Point(250,26)) -Size (New-Object System.Drawing.Size(220,22)) -Text $env:USERNAME

$lblAccType = New-StyledLabel "Account Type:" (New-Object System.Drawing.Point(12,60)) (New-Object System.Drawing.Size(180,20))
$rbUser     = New-Object System.Windows.Forms.RadioButton
$rbUser.Text = "User"; $rbUser.Location = New-Object System.Drawing.Point(200,58); $rbUser.Size = New-Object System.Drawing.Size(80,20)
$rbUser.ForeColor = $clrText; $rbUser.BackColor = [System.Drawing.Color]::Transparent; $rbUser.Font = $fontUI; $rbUser.Checked = $true
$rbComp     = New-Object System.Windows.Forms.RadioButton
$rbComp.Text = "Computer"; $rbComp.Location = New-Object System.Drawing.Point(290,58); $rbComp.Size = New-Object System.Drawing.Size(90,20)
$rbComp.ForeColor = $clrText; $rbComp.BackColor = [System.Drawing.Color]::Transparent; $rbComp.Font = $fontUI

$btnAccCheck = New-StyledButton "Check Account" (New-Object System.Drawing.Point(12,94)) (New-Object System.Drawing.Size(160,30))
$btnAccLock  = New-StyledButton "Lockout Status" (New-Object System.Drawing.Point(180,94)) (New-Object System.Drawing.Size(160,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))
$btnAccKerbConst = New-StyledButton "Kerberos Constraints" (New-Object System.Drawing.Point(348,94)) (New-Object System.Drawing.Size(130,30)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))

$lblAccNote = New-StyledLabel "Checks: enabled status, password expiry, Kerberos flags, preauth, delegation settings, group memberships." `
    -Location (New-Object System.Drawing.Point(12,136)) -Size (New-Object System.Drawing.Size(460,18)) -Font $fontSmall -ForeColor $clrTextDim

$gbAccCtrl.Controls.AddRange(@($lblAccName,$tbAccName,$lblAccType,$rbUser,$rbComp,$btnAccCheck,$btnAccLock,$btnAccKerbConst,$lblAccNote))

$rtbAcc = New-RichOutput (New-Object System.Drawing.Point(10,200)) (New-Object System.Drawing.Size(1010,440))
$tabAccount.Controls.Add($rtbAcc)

$btnAccCheck.Add_Click({
    Invoke-WithSpinner -Btn $btnAccCheck -Action {
        $acct   = $tbAccName.Text.Trim()
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        if (-not $acct) { Append-Output $rtbAcc "[!] Enter an account name." $clrWarn; return }
        $rtbAcc.Clear()
        Append-Output $rtbAcc "Account Kerberos Analysis: $acct @ $domain" $clrAccent
        Append-Output $rtbAcc "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" $clrTextDim
        Append-Output $rtbAcc ""

        try {
            $adParams = @{ Server = $domain }
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }

            if ($rbUser.Checked) {
                $props = @('Enabled','LockedOut','PasswordExpired','PasswordNeverExpires','PasswordLastSet',
                           'LastLogonDate','BadLogonCount','BadPwdCount','DoesNotRequirePreAuth',
                           'TrustedForDelegation','TrustedToAuthForDelegation','AccountNotDelegated',
                           'ServicePrincipalNames','msDS-SupportedEncryptionTypes',
                           'UserAccountControl','MemberOf','LogonWorkstations')
                $obj = Get-ADUser -Identity $acct -Properties $props @adParams -ErrorAction Stop

                Write-Result $rtbAcc "Account Enabled"         $(if ($obj.Enabled)           { 'PASS' } else { 'FAIL' }) ""
                Write-Result $rtbAcc "Account Locked Out"      $(if ($obj.LockedOut)          { 'FAIL' } else { 'PASS' }) $(if ($obj.LockedOut) { 'LOCKED — use Unlock-ADAccount' } else { 'Not locked' })
                Write-Result $rtbAcc "Password Expired"        $(if ($obj.PasswordExpired)    { 'FAIL' } else { 'PASS' }) $(if ($obj.PasswordExpired) { 'Password must be reset' } else { '' })
                Write-Result $rtbAcc "Password Never Expires"  $(if ($obj.PasswordNeverExpires) { 'WARN' } else { 'PASS' }) $(if ($obj.PasswordNeverExpires) { 'May indicate service account - review' } else { '' })
                Write-Result $rtbAcc "Kerberos Pre-Auth Reqd"  $(if ($obj.DoesNotRequirePreAuth) { 'WARN' } else { 'PASS' }) $(if ($obj.DoesNotRequirePreAuth) { 'PREAUTHNOTREQUIRED set — AS-REP Roasting risk' } else { 'Pre-auth required (secure)' })
                Write-Result $rtbAcc "Unconstrained Delegation" $(if ($obj.TrustedForDelegation) { 'WARN' } else { 'PASS' }) $(if ($obj.TrustedForDelegation) { 'Unconstrained delegation enabled — HIGH RISK' } else { 'Not set' })
                Write-Result $rtbAcc "Constrained Delegation"  $(if ($obj.TrustedToAuthForDelegation) { 'WARN' } else { 'PASS' }) $(if ($obj.TrustedToAuthForDelegation) { 'Protocol Transition (S4U2Self) enabled' } else { 'Not set' })
                Write-Result $rtbAcc "Not Delegatable"         $(if ($obj.AccountNotDelegated) { 'PASS' } else { 'WARN' }) $(if ($obj.AccountNotDelegated) { 'Protected — cannot be delegated' } else { 'Account can be delegated' })

                # Encryption types
                $encFlags = $obj.'msDS-SupportedEncryptionTypes'
                $encStr = @()
                if ($encFlags -band 1)  { $encStr += "DES-CBC-CRC" }
                if ($encFlags -band 2)  { $encStr += "DES-CBC-MD5" }
                if ($encFlags -band 4)  { $encStr += "RC4-HMAC" }
                if ($encFlags -band 8)  { $encStr += "AES128" }
                if ($encFlags -band 16) { $encStr += "AES256" }
                $encDisplay = if ($encStr) { $encStr -join ", " } else { "Default (RC4/AES per domain policy)" }
                $encStat = if ($encStr -and ($encStr | Where-Object { $_ -match "DES" })) { 'WARN' } else { 'PASS' }
                Write-Result $rtbAcc "Enc Types (msDS-SupportedEncryptionTypes)" $encStat $encDisplay

                Append-Output $rtbAcc ""
                Append-Output $rtbAcc "── Account Details ─────────────────────" $clrAccent
                Append-Output $rtbAcc "  Password Last Set : $($obj.PasswordLastSet)"
                Append-Output $rtbAcc "  Last Logon        : $($obj.LastLogonDate)"
                Append-Output $rtbAcc "  Bad Password Count: $($obj.BadPwdCount)"
                Append-Output $rtbAcc "  SPNs Registered   : $($obj.ServicePrincipalNames.Count)"
                if ($obj.ServicePrincipalNames) {
                    foreach ($spn in $obj.ServicePrincipalNames) { Append-Output $rtbAcc "    → $spn" $clrTextDim }
                }
                Append-Output $rtbAcc "  Group Memberships : $($obj.MemberOf.Count) groups"

            } else {
                # Computer account
                $obj = Get-ADComputer -Identity $acct -Properties Enabled,TrustedForDelegation,ServicePrincipalNames,'msDS-SupportedEncryptionTypes',OperatingSystem,LastLogonDate @adParams -ErrorAction Stop
                Write-Result $rtbAcc "Account Enabled"     $(if ($obj.Enabled)              { 'PASS' } else { 'FAIL' }) ""
                Write-Result $rtbAcc "Unconstrained Deleg" $(if ($obj.TrustedForDelegation) { 'WARN' } else { 'PASS' }) $(if ($obj.TrustedForDelegation) { 'Set — review if required' } else { 'Not set' })
                Append-Output $rtbAcc ""
                Append-Output $rtbAcc "  OS          : $($obj.OperatingSystem)"
                Append-Output $rtbAcc "  Last Logon  : $($obj.LastLogonDate)"
                Append-Output $rtbAcc "  SPNs        : $($obj.ServicePrincipalNames.Count)"
                foreach ($spn in $obj.ServicePrincipalNames) { Append-Output $rtbAcc "    → $spn" $clrTextDim }
            }
        } catch {
            Append-Output $rtbAcc "[X] Account check failed: $($_.Exception.Message)" $clrError
        }
        Write-Log "Account check completed: $acct"
    }
})

$btnAccLock.Add_Click({
    Invoke-WithSpinner -Btn $btnAccLock -Action {
        $acct = $tbAccName.Text.Trim()
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $rtbAcc.Clear()
        Append-Output $rtbAcc "Lockout Status: $acct" $clrAccent
        try {
            $adParams = @{ Server = $domain }
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
            $obj = Get-ADUser -Identity $acct -Properties LockedOut,BadPwdCount,BadPasswordTime,LastBadPasswordAttempt,PasswordLastSet @adParams
            Write-Result $rtbAcc "Locked Out"           $(if ($obj.LockedOut) { 'FAIL' } else { 'PASS' }) ""
            Append-Output $rtbAcc "  Bad Password Count  : $($obj.BadPwdCount)"
            Append-Output $rtbAcc "  Last Bad Attempt    : $($obj.LastBadPasswordAttempt)"
            Append-Output $rtbAcc "  Password Last Set   : $($obj.PasswordLastSet)"

            if ($obj.LockedOut) {
                $unlockBtn = [System.Windows.Forms.MessageBox]::Show(
                    "Account $acct is LOCKED.`nUnlock now?","Account Locked",'YesNo','Warning')
                if ($unlockBtn -eq 'Yes') {
                    Unlock-ADAccount -Identity $acct @adParams
                    Append-Output $rtbAcc "[+] Account unlocked successfully." $clrSuccess
                    Write-Log "Account unlocked: $acct"
                }
            }
        } catch {
            Append-Output $rtbAcc "[X] Lockout check failed: $($_.Exception.Message)" $clrError
        }
    }
})

$btnAccKerbConst.Add_Click({
    Invoke-WithSpinner -Btn $btnAccKerbConst -Action {
        $acct = $tbAccName.Text.Trim()
        $domain = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $rtbAcc.Clear()
        Append-Output $rtbAcc "Kerberos Delegation & Constraints: $acct" $clrAccent
        try {
            $adParams = @{ Server = $domain }
            if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
            $obj = Get-ADUser -Identity $acct -Properties 'msDS-AllowedToDelegateTo','msDS-AllowedToActOnBehalfOfOtherIdentity',TrustedForDelegation,TrustedToAuthForDelegation @adParams

            Append-Output $rtbAcc "── Delegation Configuration ─────────" $clrAccent
            Append-Output $rtbAcc "  Unconstrained Delegation : $($obj.TrustedForDelegation)"
            Append-Output $rtbAcc "  Protocol Transition (S4U2Self): $($obj.TrustedToAuthForDelegation)"
            $delegTo = $obj.'msDS-AllowedToDelegateTo'
            if ($delegTo) {
                Append-Output $rtbAcc "  Constrained Delegation To:" $clrWarn
                foreach ($svc in $delegTo) { Append-Output $rtbAcc "    → $svc" $clrTextDim }
            } else { Append-Output $rtbAcc "  Constrained Delegation: None configured" }

            $rbaced = $obj.'msDS-AllowedToActOnBehalfOfOtherIdentity'
            if ($rbaced) {
                Append-Output $rtbAcc "  RBAC Delegation (Resource-Based): Configured" $clrWarn
            } else { Append-Output $rtbAcc "  RBAC Delegation: Not configured" }

        } catch {
            Append-Output $rtbAcc "[X] Delegation check failed: $($_.Exception.Message)" $clrError
        }
    }
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 7 — EVENT LOG ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Event Log

$tabEvents = New-TabPage "📋 Event Logs"
$tabs.TabPages.Add($tabEvents)

$gbEvtCtrl = New-GroupBox -Text "Kerberos Event Log Query" `
    -Location (New-Object System.Drawing.Point(10,10)) -Size (New-Object System.Drawing.Size(490,210))
$tabEvents.Controls.Add($gbEvtCtrl)

$lblEvtSource = New-StyledLabel "Source Computer (blank=local):" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(220,20))
$tbEvtSource  = New-StyledTextBox -Location (New-Object System.Drawing.Point(240,26)) -Size (New-Object System.Drawing.Size(230,22))

$lblEvtHours  = New-StyledLabel "Look back (hours):" (New-Object System.Drawing.Point(12,60)) (New-Object System.Drawing.Size(220,20))
$nudHours     = New-Object System.Windows.Forms.NumericUpDown
$nudHours.Location = New-Object System.Drawing.Point(240,58); $nudHours.Size = New-Object System.Drawing.Size(80,22)
$nudHours.Minimum = 1; $nudHours.Maximum = 168; $nudHours.Value = 24
$nudHours.BackColor = $clrInput; $nudHours.ForeColor = $clrText; $nudHours.Font = $fontUI

$btnEvtKerb   = New-StyledButton "Kerberos Errors"   (New-Object System.Drawing.Point(12,96))  (New-Object System.Drawing.Size(150,28))
$btnEvtLogon  = New-StyledButton "Logon Failures"    (New-Object System.Drawing.Point(170,96)) (New-Object System.Drawing.Size(150,28)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))
$btnEvtAll    = New-StyledButton "All Kerberos"      (New-Object System.Drawing.Point(328,96)) (New-Object System.Drawing.Size(150,28)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,110))
$btnEvtKDC    = New-StyledButton "KDC Events (DC)"   (New-Object System.Drawing.Point(12,132)) (New-Object System.Drawing.Size(150,28)) -BgColor ([System.Drawing.Color]::FromArgb(110,80,55))
$btnEvtCopy   = New-StyledButton "Copy Results"      (New-Object System.Drawing.Point(170,132)) (New-Object System.Drawing.Size(150,28)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,55))

$lblEvtRef = New-StyledLabel "Key IDs: 4768 (TGT Req), 4769 (ST Req), 4771 (Pre-Auth Fail), 4776 (NTLM Fallback), 4625 (Logon Fail)" `
    -Location (New-Object System.Drawing.Point(12,170)) -Size (New-Object System.Drawing.Size(460,18)) -Font $fontSmall -ForeColor $clrTextDim

$gbEvtCtrl.Controls.AddRange(@($lblEvtSource,$tbEvtSource,$lblEvtHours,$nudHours,$btnEvtKerb,$btnEvtLogon,$btnEvtAll,$btnEvtKDC,$btnEvtCopy,$lblEvtRef))

$rtbEvents = New-RichOutput (New-Object System.Drawing.Point(10,230)) (New-Object System.Drawing.Size(1010,410))
$tabEvents.Controls.Add($rtbEvents)

function Query-SecurityLog {
    param([string]$Computer, [int]$Hours, [int[]]$EventIDs, [string]$Label)
    $rtbEvents.Clear()
    $since = (Get-Date).AddHours(-$Hours)
    Append-Output $rtbEvents "Event Query: $Label" $clrAccent
    Append-Output $rtbEvents "Source: $(if($Computer){ $Computer } else { 'LocalHost' })  |  Last $Hours hours  |  Since: $($since.ToString('yyyy-MM-dd HH:mm:ss'))" $clrTextDim
    Append-Output $rtbEvents ""

    $filter = @{
        LogName   = 'Security'
        Id        = $EventIDs
        StartTime = $since
    }
    $params = @{ FilterHashtable = $filter; MaxEvents = 200; ErrorAction = 'Stop' }
    if ($Computer) { $params.ComputerName = $Computer; if ($script:SavedCred) { $params.Credential = $script:SavedCred } }

    try {
        $events = Get-WinEvent @params | Sort-Object TimeCreated -Descending
        if ($events) {
            Append-Output $rtbEvents "Found $($events.Count) events (showing newest first)" $clrWarn
            Append-Output $rtbEvents ""
            foreach ($evt in $events) {
                $idColor = switch ($evt.Id) {
                    4768 { $clrText }    # TGT request
                    4769 { $clrText }    # Service ticket
                    4771 { $clrError }   # Pre-auth failure
                    4776 { $clrWarn }    # NTLM fallback
                    4625 { $clrError }   # Logon failure
                    default { $clrText }
                }
                Append-Output $rtbEvents "[$($evt.TimeCreated.ToString('MM/dd HH:mm:ss'))] ID:$($evt.Id) — $($evt.Message.Split("`n")[0].Trim())" $idColor
                Write-Log "EVT $($evt.Id) @ $($evt.TimeCreated)"
            }
        } else {
            Append-Output $rtbEvents "[i] No events found for the selected criteria." $clrTextDim
        }
    } catch [System.UnauthorizedAccessException] {
        Append-Output $rtbEvents "[X] Access denied — run as administrator or use alternate credentials." $clrError
    } catch {
        Append-Output $rtbEvents "[X] Query failed: $($_.Exception.Message)" $clrError
        Append-Output $rtbEvents "    Ensure audit policies are enabled: auditpol /get /category:'Account Logon'" $clrTextDim
    }
}

$btnEvtKerb.Add_Click({
    Invoke-WithSpinner -Btn $btnEvtKerb -Action {
        Query-SecurityLog -Computer $tbEvtSource.Text.Trim() -Hours $nudHours.Value `
            -EventIDs @(4771, 4768, 4769) -Label "Kerberos Auth Errors (4771) + TGT/ST Requests"
    }
})

$btnEvtLogon.Add_Click({
    Invoke-WithSpinner -Btn $btnEvtLogon -Action {
        Query-SecurityLog -Computer $tbEvtSource.Text.Trim() -Hours $nudHours.Value `
            -EventIDs @(4625, 4776, 4648) -Label "Logon Failures (4625) + NTLM (4776) + Explicit Creds (4648)"
    }
})

$btnEvtAll.Add_Click({
    Invoke-WithSpinner -Btn $btnEvtAll -Action {
        Query-SecurityLog -Computer $tbEvtSource.Text.Trim() -Hours $nudHours.Value `
            -EventIDs @(4768,4769,4770,4771,4772,4773,4774,4776,4625,4648) -Label "All Kerberos + Logon Events"
    }
})

$btnEvtKDC.Add_Click({
    Invoke-WithSpinner -Btn $btnEvtKDC -Action {
        $src = $tbEvtSource.Text.Trim()
        $rtbEvents.Clear()
        Append-Output $rtbEvents "KDC / System Kerberos Events" $clrAccent
        $since = (Get-Date).AddHours(-$nudHours.Value)
        $params = @{ FilterHashtable = @{ LogName='System'; ProviderName='Microsoft-Windows-Security-Kerberos'; StartTime=$since }; MaxEvents=100; ErrorAction='Stop' }
        if ($src) { $params.ComputerName = $src }
        try {
            $events = Get-WinEvent @params | Sort-Object TimeCreated -Descending
            foreach ($evt in $events) {
                $c = if ($evt.LevelDisplayName -match 'Error|Critical') { $clrError } elseif ($evt.LevelDisplayName -match 'Warning') { $clrWarn } else { $clrText }
                Append-Output $rtbEvents "[$($evt.TimeCreated.ToString('MM/dd HH:mm:ss'))] [$($evt.LevelDisplayName)] $($evt.Message.Split("`n")[0].Trim())" $c
            }
            if (-not $events) { Append-Output $rtbEvents "[i] No KDC system events found." $clrTextDim }
        } catch {
            Append-Output $rtbEvents "[X] $($_.Exception.Message)" $clrError
        }
    }
})

$btnEvtCopy.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($rtbEvents.Text)
    $btnEvtCopy.Text = "✔ Copied!"
    Start-Sleep -Milliseconds 1500
    $btnEvtCopy.Text = "Copy Results"
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  TAB 8 — FULL DIAGNOSTIC REPORT
# ════════════════════════════════════════════════════════════════════════════
#region TAB: Full Report

$tabReport = New-TabPage "📊 Full Report"
$tabs.TabPages.Add($tabReport)

$gbRptCtrl = New-GroupBox -Text "Automated Kerberos Health Report" `
    -Location (New-Object System.Drawing.Point(10,10)) -Size (New-Object System.Drawing.Size(490,130))
$tabReport.Controls.Add($gbRptCtrl)

$lblRptAcct = New-StyledLabel "Account to diagnose (optional):" (New-Object System.Drawing.Point(12,28)) (New-Object System.Drawing.Size(230,20))
$tbRptAcct  = New-StyledTextBox -Location (New-Object System.Drawing.Point(250,26)) -Size (New-Object System.Drawing.Size(220,22)) -Text $env:USERNAME

$btnRunReport = New-StyledButton "▶  Run Full Report" (New-Object System.Drawing.Point(12,62)) (New-Object System.Drawing.Size(200,34))
$btnSaveRpt   = New-StyledButton "💾 Save to File"    (New-Object System.Drawing.Point(220,62)) (New-Object System.Drawing.Size(160,34)) -BgColor ([System.Drawing.Color]::FromArgb(55,80,55))
$btnCopyRpt   = New-StyledButton "Copy Report"        (New-Object System.Drawing.Point(388,62)) (New-Object System.Drawing.Size(90,34))  -BgColor ([System.Drawing.Color]::FromArgb(55,55,90))

$gbRptCtrl.Controls.AddRange(@($lblRptAcct,$tbRptAcct,$btnRunReport,$btnSaveRpt,$btnCopyRpt))

$rtbReport = New-RichOutput (New-Object System.Drawing.Point(10,150)) (New-Object System.Drawing.Size(1010,490))
$tabReport.Controls.Add($rtbReport)

$btnRunReport.Add_Click({
    Invoke-WithSpinner -Btn $btnRunReport -Action {
        $rtbReport.Clear()
        $domain  = Get-ActiveDomain; if (-not $domain) { $domain = $env:USERDNSDOMAIN }
        $acct    = $tbRptAcct.Text.Trim()
        $adMod   = $null -ne (Get-Module -ListAvailable -Name ActiveDirectory)
        $ts      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        Append-Output $rtbReport "╔══════════════════════════════════════════════════════════════════╗" $clrAccent
        Append-Output $rtbReport "║     KERBEROS AUTHENTICATION HEALTH REPORT                        ║" $clrAccent
        Append-Output $rtbReport "╚══════════════════════════════════════════════════════════════════╝" $clrAccent
        Append-Output $rtbReport "  Generated : $ts"
        Append-Output $rtbReport "  Computer  : $env:COMPUTERNAME ($env:USERDOMAIN\$env:USERNAME)"
        Append-Output $rtbReport "  Target    : $domain"
        Append-Output $rtbReport "  Account   : $acct"
        Append-Output $rtbReport ""

        # Section 1: Environment
        Append-Output $rtbReport "── [1] ENVIRONMENT ─────────────────────────────────────────" $clrAccent
        $adModStat = if ($adMod) { '[+] ActiveDirectory module available' } else { '[!] ActiveDirectory module NOT available' }
        $adModClr  = if ($adMod) { $clrSuccess } else { $clrWarn }
        Append-Output $rtbReport $adModStat $adModClr
        try {
            $dns = [System.Net.Dns]::GetHostAddresses($domain)
            Write-Result $rtbReport "DNS Resolution for $domain" 'PASS' "$($dns.Count) address(es) found"
        } catch {
            Write-Result $rtbReport "DNS Resolution for $domain" 'FAIL' $_.Exception.Message
        }

        # Section 2: Kerberos Ports
        Append-Output $rtbReport "`n── [2] KERBEROS PORT CONNECTIVITY ──────────────────────────" $clrAccent
        $kerbPorts = @{ 88='KDC'; 464='kpasswd'; 389='LDAP'; 636='LDAPS'; 53='DNS' }
        foreach ($kp in $kerbPorts.GetEnumerator()) {
            try {
                $r = Test-NetConnection -ComputerName $domain -Port $kp.Key -WarningAction SilentlyContinue -ErrorAction Stop
                Write-Result $rtbReport "Port $($kp.Key) ($($kp.Value))" $(if($r.TcpTestSucceeded){'PASS'}else{'FAIL'}) ""
            } catch { Write-Result $rtbReport "Port $($kp.Key)" 'FAIL' $_.Exception.Message }
        }

        # Section 3: Time Sync
        Append-Output $rtbReport "`n── [3] TIME SYNCHRONIZATION ─────────────────────────────" $clrAccent
        $w32 = w32tm /query /status 2>&1
        $src = ($w32 | Select-String "Source") -replace ".*Source:\s*",""
        $str = ($w32 | Select-String "Stratum") -replace ".*Stratum:\s*",""
        Append-Output $rtbReport "  NTP Source  : $src"
        Append-Output $rtbReport "  Stratum     : $str"

        # Section 4: Kerberos Tickets
        Append-Output $rtbReport "`n── [4] KERBEROS TICKETS ─────────────────────────────────" $clrAccent
        $kl = klist 2>&1
        $tktCount = ($kl | Select-String "^#\d+").Count
        $hasTGT   = $kl | Select-String "krbtgt"
        Write-Result $rtbReport "Ticket Cache" $(if($tktCount -gt 0){'PASS'}else{'WARN'}) "$tktCount ticket(s) cached"
        Write-Result $rtbReport "TGT Present" $(if($hasTGT){'PASS'}else{'WARN'}) $(if($hasTGT){'TGT found in cache'}else{'No TGT — may need kinit'})
        $desTkts = $kl | Select-String "DES"
        if ($desTkts) { Write-Result $rtbReport "DES Encryption Detected" 'WARN' "Insecure DES tickets found — review domain enc policy" }

        # Section 5: nltest
        Append-Output $rtbReport "`n── [5] DC LOCATOR / SECURE CHANNEL ─────────────────────" $clrAccent
        $nl = nltest /dsgetdc:$domain 2>&1
        $dcLine = $nl | Select-String "\\\\"
        Write-Result $rtbReport "DC Discovery" $(if($dcLine){'PASS'}else{'FAIL'}) ($dcLine.ToString().Trim())
        $sc = nltest /sc_query:$domain 2>&1
        $scOK = $sc | Select-String "SUCCESS|CONNECTED"
        Write-Result $rtbReport "Secure Channel" $(if($scOK){'PASS'}else{'FAIL'}) ($scOK.ToString().Trim())

        # Section 6: Account (if AD module available)
        if ($adMod -and $acct) {
            Append-Output $rtbReport "`n── [6] ACCOUNT ANALYSIS ($acct) ──────────────────────" $clrAccent
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                $adParams = @{ Server = $domain }
                if ($script:SavedCred) { $adParams.Credential = $script:SavedCred }
                $props = @('Enabled','LockedOut','PasswordExpired','DoesNotRequirePreAuth','TrustedForDelegation','BadPwdCount')
                $obj = Get-ADUser -Identity $acct -Properties $props @adParams -ErrorAction Stop
                Write-Result $rtbReport "Enabled"             $(if($obj.Enabled){'PASS'}else{'FAIL'}) ""
                Write-Result $rtbReport "Locked Out"          $(if($obj.LockedOut){'FAIL'}else{'PASS'}) ""
                Write-Result $rtbReport "Password Expired"    $(if($obj.PasswordExpired){'FAIL'}else{'PASS'}) ""
                Write-Result $rtbReport "Pre-Auth Required"   $(if($obj.DoesNotRequirePreAuth){'WARN'}else{'PASS'}) $(if($obj.DoesNotRequirePreAuth){'PREAUTHNOTREQUIRED — AS-REP Roasting risk'}else{'Secure'})
                Write-Result $rtbReport "Uncons. Delegation"  $(if($obj.TrustedForDelegation){'WARN'}else{'PASS'}) $(if($obj.TrustedForDelegation){'HIGH RISK — review immediately'}else{'Not configured'})
            } catch {
                Append-Output $rtbReport "  [X] Account lookup failed: $($_.Exception.Message)" $clrError
            }
        } elseif (-not $adMod) {
            Append-Output $rtbReport "  [!] Skipped — ActiveDirectory module not available." $clrWarn
        }

        # Section 7: SPN Duplicates
        Append-Output $rtbReport "`n── [7] DUPLICATE SPN CHECK ──────────────────────────────" $clrAccent
        $spnOut = setspn -X -F 2>&1 | Where-Object { $_ -match "duplicate|found|Checking" } | Select-Object -First 5
        foreach ($l in $spnOut) {
            $c = if ($l -match "duplicate") { $clrError } elseif ($l -match "^0") { $clrSuccess } else { $clrTextDim }
            Append-Output $rtbReport "  $l" $c
        }

        Append-Output $rtbReport ""
        Append-Output $rtbReport "══════════════════════════════════════════════════════════════════" $clrAccent
        Append-Output $rtbReport "  Report complete — saved to: $LogFile" $clrTextDim
        Append-Output $rtbReport "══════════════════════════════════════════════════════════════════" $clrAccent
        Write-Log "Full diagnostic report completed for $domain / $acct"
    }
})

$btnSaveRpt.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = "Text Files (*.txt)|*.txt|Log Files (*.log)|*.log|All Files (*.*)|*.*"
    $sfd.FileName = "KerberosReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $sfd.InitialDirectory = $LogDir
    if ($sfd.ShowDialog() -eq 'OK') {
        $rtbReport.Text | Out-File -FilePath $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Report saved to:`n$($sfd.FileName)", "Saved", 'OK', 'Information')
        Write-Log "Report saved to $($sfd.FileName)"
    }
})

$btnCopyRpt.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($rtbReport.Text)
    $btnCopyRpt.Text = "✔ Copied"
    Start-Sleep -Milliseconds 1500
    $btnCopyRpt.Text = "Copy Report"
})

#endregion

# ════════════════════════════════════════════════════════════════════════════
#  FINAL LAYOUT & LAUNCH
# ════════════════════════════════════════════════════════════════════════════

$form.Add_Shown({
    $statusLabel.Text = "Ready — select Connection tab to configure domain settings."
    # Auto-load connection info
    $rtbConnInfo.Clear()
    Append-Output $rtbConnInfo "Configure settings above, then click '▶ Apply & Connect'." $clrTextDim
    Append-Output $rtbConnInfo "" 
    Append-Output $rtbConnInfo "Detected environment:" $clrAccent
    Append-Output $rtbConnInfo "  Computer : $env:COMPUTERNAME"
    Append-Output $rtbConnInfo "  User     : $env:USERDOMAIN\$env:USERNAME"
    Append-Output $rtbConnInfo "  Domain   : $env:USERDNSDOMAIN"
    Append-Output $rtbConnInfo ""
    Append-Output $rtbConnInfo "Log file: $LogFile" $clrTextDim
})

$form.Add_FormClosing({
    Write-Log "KerberosTroubleshooter closed."
    $fontMono.Dispose(); $fontUI.Dispose(); $fontUIB.Dispose(); $fontTitle.Dispose(); $fontSmall.Dispose()
})

[System.Windows.Forms.Application]::Run($form)
