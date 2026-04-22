#Requires -Version 5.1
<#
.SYNOPSIS
    Add a T1 account to Administrators and Remote Desktop Users groups on remote servers.

.DESCRIPTION
    GUI-based PowerShell ISE script for AD Tier 1 admins to add an account to
    multiple remote servers' local Administrators and Remote Desktop Users groups.
    Uses provided T1 credentials for all remote operations.

.NOTES
    - Run from PowerShell ISE
    - Requires network access to target servers
    - Credentials must have permission to modify local groups on target servers
    - GitHub URL - https://github.com/Henchman33/AD/blob/main/Users/Add%20User%20Account%20to%20Server%20Admin%20and%20RDP%20Groups/Add-User-Admin-RDP-GUI.ps1
    - Users/Add User Account to Server Admin and RDP Groups/Add-User-Admin-RDP-GUI.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region ── Build the Form ──────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text            = "T1 Account — Remote Server Group Manager"
$form.Size            = New-Object System.Drawing.Size(1024, 768)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 40)
$form.ForeColor       = [System.Drawing.Color]::White
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# ── Helper: styled label ──────────────────────────────────────────────────────
function New-Label($text, $x, $y, $w = 180, $h = 20) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.Size     = New-Object System.Drawing.Size($w, $h)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 255)
    return $lbl
}

# ── Helper: styled textbox ────────────────────────────────────────────────────
function New-TextBox($x, $y, $w = 380, $h = 24) {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location  = New-Object System.Drawing.Point($x, $y)
    $tb.Size      = New-Object System.Drawing.Size($w, $h)
    $tb.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 60)
    $tb.ForeColor = [System.Drawing.Color]::White
    $tb.BorderStyle = "FixedSingle"
    return $tb
}

#region ── Section 1: Credentials ─────────────────────────────────────────────

$panelCreds = New-Object System.Windows.Forms.Panel
$panelCreds.Location  = New-Object System.Drawing.Point(10, 10)
$panelCreds.Size      = New-Object System.Drawing.Size(580, 145)
$panelCreds.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 55)
$panelCreds.BorderStyle = "FixedSingle"

$lblCredTitle = New-Object System.Windows.Forms.Label
$lblCredTitle.Text      = "  🔐  YOUR T1 ADMIN CREDENTIALS"
$lblCredTitle.Location  = New-Object System.Drawing.Point(0, 0)
$lblCredTitle.Size      = New-Object System.Drawing.Size(580, 28)
$lblCredTitle.BackColor = [System.Drawing.Color]::FromArgb(60, 80, 140)
$lblCredTitle.ForeColor = [System.Drawing.Color]::White
$lblCredTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelCreds.Controls.Add($lblCredTitle)

$panelCreds.Controls.Add((New-Label "Domain\T1-Username:" 10 40))
$txtAdminUser = New-TextBox 200 38 360
$txtAdminUser.PlaceholderText = "DOMAIN\t1-yourname"
$panelCreds.Controls.Add($txtAdminUser)

$panelCreds.Controls.Add((New-Label "Password:" 10 78))
$txtAdminPass = New-TextBox 200 76 360
$txtAdminPass.UseSystemPasswordChar = $true
$panelCreds.Controls.Add($txtAdminPass)

$panelCreds.Controls.Add((New-Label "Account to Add:" 10 112))
$txtTargetAccount = New-TextBox 200 110 360
$txtTargetAccount.PlaceholderText = "DOMAIN\t1-account (account to add to servers)"
$panelCreds.Controls.Add($txtTargetAccount)

$form.Controls.Add($panelCreds)

#endregion

#region ── Section 2: Server List ─────────────────────────────────────────────

$panelServers = New-Object System.Windows.Forms.Panel
$panelServers.Location  = New-Object System.Drawing.Point(10, 165)
$panelServers.Size      = New-Object System.Drawing.Size(580, 230)
$panelServers.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 55)
$panelServers.BorderStyle = "FixedSingle"

$lblSrvTitle = New-Object System.Windows.Forms.Label
$lblSrvTitle.Text      = "  🖥️  TARGET SERVERS  (one per line or comma-separated)"
$lblSrvTitle.Location  = New-Object System.Drawing.Point(0, 0)
$lblSrvTitle.Size      = New-Object System.Drawing.Size(580, 28)
$lblSrvTitle.BackColor = [System.Drawing.Color]::FromArgb(60, 80, 140)
$lblSrvTitle.ForeColor = [System.Drawing.Color]::White
$lblSrvTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelServers.Controls.Add($lblSrvTitle)

$txtServers = New-Object System.Windows.Forms.TextBox
$txtServers.Location   = New-Object System.Drawing.Point(10, 38)
$txtServers.Size       = New-Object System.Drawing.Size(555, 155)
$txtServers.Multiline  = $true
$txtServers.ScrollBars = "Vertical"
$txtServers.BackColor  = [System.Drawing.Color]::FromArgb(45, 45, 60)
$txtServers.ForeColor  = [System.Drawing.Color]::White
$txtServers.BorderStyle = "FixedSingle"
$txtServers.PlaceholderText = "SERVER01`r`nSERVER02`r`nSERVER03"
$panelServers.Controls.Add($txtServers)

$form.Controls.Add($panelServers)

#endregion

#region ── Section 3: Group Selection ────────────────────────────────────────

$panelGroups = New-Object System.Windows.Forms.Panel
$panelGroups.Location  = New-Object System.Drawing.Point(10, 405)
$panelGroups.Size      = New-Object System.Drawing.Size(580, 80)
$panelGroups.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 55)
$panelGroups.BorderStyle = "FixedSingle"

$lblGrpTitle = New-Object System.Windows.Forms.Label
$lblGrpTitle.Text      = "  ✅  GROUPS TO ADD ACCOUNT TO"
$lblGrpTitle.Location  = New-Object System.Drawing.Point(0, 0)
$lblGrpTitle.Size      = New-Object System.Drawing.Size(580, 28)
$lblGrpTitle.BackColor = [System.Drawing.Color]::FromArgb(60, 80, 140)
$lblGrpTitle.ForeColor = [System.Drawing.Color]::White
$lblGrpTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelGroups.Controls.Add($lblGrpTitle)

$chkAdmins = New-Object System.Windows.Forms.CheckBox
$chkAdmins.Text      = "Administrators"
$chkAdmins.Location  = New-Object System.Drawing.Point(20, 40)
$chkAdmins.Size      = New-Object System.Drawing.Size(180, 24)
$chkAdmins.Checked   = $true
$chkAdmins.ForeColor = [System.Drawing.Color]::White
$panelGroups.Controls.Add($chkAdmins)

$chkRDP = New-Object System.Windows.Forms.CheckBox
$chkRDP.Text      = "Remote Desktop Users"
$chkRDP.Location  = New-Object System.Drawing.Point(220, 40)
$chkRDP.Size      = New-Object System.Drawing.Size(200, 24)
$chkRDP.Checked   = $true
$chkRDP.ForeColor = [System.Drawing.Color]::White
$panelGroups.Controls.Add($chkRDP)

$form.Controls.Add($panelGroups)

#endregion

#region ── Section 4: Run Button ─────────────────────────────────────────────

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text      = "▶  ADD ACCOUNT TO SERVERS"
$btnRun.Location  = New-Object System.Drawing.Point(10, 495)
$btnRun.Size      = New-Object System.Drawing.Size(580, 40)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(60, 120, 60)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = "Flat"
$btnRun.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnRun.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnRun)

#endregion

#region ── Section 5: Output Log ──────────────────────────────────────────────

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text      = "  📋  RESULTS LOG"
$lblLog.Location  = New-Object System.Drawing.Point(10, 545)
$lblLog.Size      = New-Object System.Drawing.Size(580, 22)
$lblLog.BackColor = [System.Drawing.Color]::FromArgb(60, 80, 140)
$lblLog.ForeColor = [System.Drawing.Color]::White
$lblLog.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location   = New-Object System.Drawing.Point(10, 568)
$txtLog.Size       = New-Object System.Drawing.Size(580, 90)
$txtLog.BackColor  = [System.Drawing.Color]::FromArgb(15, 15, 25)
$txtLog.ForeColor  = [System.Drawing.Color]::FromArgb(150, 255, 150)
$txtLog.ReadOnly   = $true
$txtLog.BorderStyle = "None"
$txtLog.Font       = New-Object System.Drawing.Font("Consolas", 8.5)
$form.Controls.Add($txtLog)

# Helper to write to the log with color
function Write-Log {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(150,255,150)
    )
    $txtLog.SelectionStart  = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor  = $Color
    $txtLog.AppendText("$Message`n")
    $txtLog.ScrollToCaret()
    $form.Refresh()
}

#endregion

#region ── Button Click Logic ────────────────────────────────────────────────

$btnRun.Add_Click({

    $txtLog.Clear()

    # ── Validate inputs ───────────────────────────────────────────────────────
    $adminUser   = $txtAdminUser.Text.Trim()
    $adminPass   = $txtAdminPass.Text
    $targetAcct  = $txtTargetAccount.Text.Trim()

    if (-not $adminUser -or -not $adminPass) {
        Write-Log "⛔  Please enter your T1 credentials (username & password)." ([System.Drawing.Color]::FromArgb(255,80,80))
        return
    }
    if (-not $targetAcct) {
        Write-Log "⛔  Please enter the account to add (e.g. DOMAIN\t1-account)." ([System.Drawing.Color]::FromArgb(255,80,80))
        return
    }
    if (-not $chkAdmins.Checked -and -not $chkRDP.Checked) {
        Write-Log "⛔  Please select at least one group." ([System.Drawing.Color]::FromArgb(255,80,80))
        return
    }

    # ── Parse server list ─────────────────────────────────────────────────────
    $rawServers = $txtServers.Text -split "[\r\n,]" |
                  ForEach-Object { $_.Trim() } |
                  Where-Object   { $_ -ne "" } |
                  Select-Object  -Unique

    if ($rawServers.Count -eq 0) {
        Write-Log "⛔  Please enter at least one server name." ([System.Drawing.Color]::FromArgb(255,80,80))
        return
    }

    # ── Build credential object ───────────────────────────────────────────────
    $securePass = ConvertTo-SecureString $adminPass -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($adminUser, $securePass)

    # ── Resolve domain\account into ADSI-compatible parts ────────────────────
    # targetAcct expected as DOMAIN\username
    if ($targetAcct -match "^(.+)\\(.+)$") {
        $targetDomain = $Matches[1]
        $targetUser   = $Matches[2]
    } else {
        Write-Log "⛔  Account format must be DOMAIN\username." ([System.Drawing.Color]::FromArgb(255,80,80))
        return
    }

    # ── Groups to process ────────────────────────────────────────────────────
    $groupsToProcess = @()
    if ($chkAdmins.Checked) { $groupsToProcess += "Administrators" }
    if ($chkRDP.Checked)    { $groupsToProcess += "Remote Desktop Users" }

    Write-Log "═══════════════════════════════════════════" ([System.Drawing.Color]::FromArgb(100,120,200))
    Write-Log "  Starting — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" ([System.Drawing.Color]::FromArgb(200,200,255))
    Write-Log "  Account  : $targetAcct" ([System.Drawing.Color]::FromArgb(200,200,255))
    Write-Log "  Groups   : $($groupsToProcess -join ', ')" ([System.Drawing.Color]::FromArgb(200,200,255))
    Write-Log "  Servers  : $($rawServers.Count)" ([System.Drawing.Color]::FromArgb(200,200,255))
    Write-Log "═══════════════════════════════════════════" ([System.Drawing.Color]::FromArgb(100,120,200))

    # ── Scriptblock to run on each remote server ──────────────────────────────
    $remoteScript = {
        param($Groups, $TargetDomain, $TargetUser)

        $results = @()

        foreach ($groupName in $Groups) {
            try {
                # Get the local group
                $group = [ADSI]"WinNT://$env:COMPUTERNAME/$groupName,group"

                # Check if already a member
                $members   = @($group.Invoke("Members"))
                $alreadyIn = $false

                foreach ($member in $members) {
                    $memberName = $member.GetType().InvokeMember("Name", "GetProperty", $null, $member, $null)
                    if ($memberName -eq $TargetUser) {
                        $alreadyIn = $true
                        break
                    }
                }

                if ($alreadyIn) {
                    $results += [PSCustomObject]@{
                        Group  = $groupName
                        Status = "ALREADY_MEMBER"
                        Error  = $null
                    }
                } else {
                    # Add the account
                    $group.Add("WinNT://$TargetDomain/$TargetUser,user")
                    $results += [PSCustomObject]@{
                        Group  = $groupName
                        Status = "ADDED"
                        Error  = $null
                    }
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    Group  = $groupName
                    Status = "ERROR"
                    Error  = $_.Exception.Message
                }
            }
        }
        return $results
    }

    # ── Iterate each server ───────────────────────────────────────────────────
    foreach ($server in $rawServers) {

        Write-Log "`n▶  $server" ([System.Drawing.Color]::FromArgb(255,220,80))

        try {
            # Test connectivity first (quick WinRM/ping check)
            $pingResult = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $pingResult) {
                Write-Log "   ⚠  Unreachable (ping failed) — skipping." ([System.Drawing.Color]::FromArgb(255,140,40))
                continue
            }

            $invokeParams = @{
                ComputerName = $server
                Credential   = $credential
                ScriptBlock  = $remoteScript
                ArgumentList = $groupsToProcess, $targetDomain, $targetUser
                ErrorAction  = "Stop"
            }

            $results = Invoke-Command @invokeParams

            foreach ($r in $results) {
                switch ($r.Status) {
                    "ADDED"          { Write-Log ("   ✔  [{0}] Added successfully." -f $r.Group) ([System.Drawing.Color]::FromArgb(80,220,80))  }
                    "ALREADY_MEMBER" { Write-Log ("   ℹ  [{0}] Already a member."   -f $r.Group) ([System.Drawing.Color]::FromArgb(180,180,80)) }
                    "ERROR"          { Write-Log ("   ✘  [{0}] Error: {1}"          -f $r.Group, $r.Error) ([System.Drawing.Color]::FromArgb(255,80,80)) }
                }
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Log "   ✘  Remote connection failed: $errMsg" ([System.Drawing.Color]::FromArgb(255,80,80))
        }
    }

    Write-Log "`n═══════════════════════════════════════════" ([System.Drawing.Color]::FromArgb(100,120,200))
    Write-Log "  Completed — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" ([System.Drawing.Color]::FromArgb(200,200,255))
    Write-Log "═══════════════════════════════════════════" ([System.Drawing.Color]::FromArgb(100,120,200))

    # Clear password from memory
    $securePass.Dispose()
    $adminPass = $null
    [System.GC]::Collect()
})

#endregion

# ── Launch the form ────────────────────────────────────────────────────────────
[void]$form.ShowDialog()
