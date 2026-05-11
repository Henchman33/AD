# Active Directory Kerberos Troubleshooting Tool (PowerShell GUI)

```powershell
<#
.SYNOPSIS
    Active Directory / Kerberos Troubleshooting GUI Tool

.DESCRIPTION
    GUI-based troubleshooting utility for Active Directory Kerberos diagnostics.

    Features:
    - Connect to different domains
    - Use alternate credentials
    - Validate DC connectivity
    - DNS resolution testing
    - Kerberos ticket inspection
    - SPN lookup
    - Time synchronization validation
    - LDAP binding test
    - Secure channel validation
    - Port testing
    - Export results

.NOTES
    Author: OpenAI ChatGPT
    Requires: PowerShell 5.1+
    Modules: ActiveDirectory
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

#====================================================
# Global Variables
#====================================================

$Script:Credential = $null
$Script:Results = @()

#====================================================
# Functions
#====================================================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    $txtResults.AppendText($line + [Environment]::NewLine)
    $txtResults.SelectionStart = $txtResults.Text.Length
    $txtResults.ScrollToCaret()

    $Script:Results += $line
}

function Get-SelectedCredential {
    if ($Script:Credential -ne $null) {
        return $Script:Credential
    }

    return $null
}

function Test-DomainControllerConnectivity {
    param(
        [string]$DomainController
    )

    Write-Log "Testing ICMP connectivity to $DomainController"

    try {
        $ping = Test-Connection -ComputerName $DomainController -Count 2 -Quiet -ErrorAction Stop

        if ($ping) {
            Write-Log "Ping successful to $DomainController" "SUCCESS"
        }
        else {
            Write-Log "Ping failed to $DomainController" "ERROR"
        }
    }
    catch {
        Write-Log "Ping test failed: $($_.Exception.Message)" "ERROR"
    }
}

function Test-DnsResolution {
    param(
        [string]$DomainName
    )

    Write-Log "Testing DNS SRV records for $DomainName"

    try {
        $records = Resolve-DnsName -Type SRV "_kerberos._tcp.$DomainName" -ErrorAction Stop

        foreach ($record in $records) {
            Write-Log "Found Kerberos SRV: $($record.NameTarget) Port $($record.Port)" "SUCCESS"
        }
    }
    catch {
        Write-Log "DNS SRV lookup failed: $($_.Exception.Message)" "ERROR"
    }

    try {
        $dcRecords = Resolve-DnsName -Type SRV "_ldap._tcp.dc._msdcs.$DomainName" -ErrorAction Stop

        foreach ($record in $dcRecords) {
            Write-Log "Found LDAP SRV: $($record.NameTarget) Port $($record.Port)" "SUCCESS"
        }
    }
    catch {
        Write-Log "LDAP SRV lookup failed: $($_.Exception.Message)" "ERROR"
    }
}

function Test-KerberosTickets {
    Write-Log "Enumerating Kerberos tickets"

    try {
        $tickets = klist

        foreach ($line in $tickets) {
            Write-Log $line
        }
    }
    catch {
        Write-Log "Failed to enumerate Kerberos tickets: $($_.Exception.Message)" "ERROR"
    }
}

function Purge-KerberosTickets {
    Write-Log "Purging Kerberos tickets"

    try {
        klist purge | Out-Null
        Write-Log "Kerberos tickets purged successfully" "SUCCESS"
    }
    catch {
        Write-Log "Failed to purge Kerberos tickets: $($_.Exception.Message)" "ERROR"
    }
}

function Test-TimeSynchronization {
    param(
        [string]$DomainController
    )

    Write-Log "Checking time synchronization against $DomainController"

    try {
        $output = w32tm /stripchart /computer:$DomainController /samples:3 /dataonly

        foreach ($line in $output) {
            Write-Log $line
        }

        Write-Log "Time synchronization check completed" "SUCCESS"
    }
    catch {
        Write-Log "Time synchronization test failed: $($_.Exception.Message)" "ERROR"
    }
}

function Test-PortConnectivity {
    param(
        [string]$Server
    )

    $ports = @(
        53,
        88,
        123,
        135,
        389,
        445,
        464,
        636,
        3268,
        3269
    )

    foreach ($port in $ports) {
        try {
            $result = Test-NetConnection -ComputerName $Server -Port $port -WarningAction SilentlyContinue

            if ($result.TcpTestSucceeded) {
                Write-Log "Port $port open on $Server" "SUCCESS"
            }
            else {
                Write-Log "Port $port closed/unreachable on $Server" "ERROR"
            }
        }
        catch {
            Write-Log "Port test failed for $port : $($_.Exception.Message)" "ERROR"
        }
    }
}

function Test-LdapBind {
    param(
        [string]$DomainController,
        [PSCredential]$Credential
    )

    Write-Log "Testing LDAP bind to $DomainController"

    try {
        if ($Credential) {
            $username = $Credential.UserName
            $password = $Credential.GetNetworkCredential().Password

            $entry = New-Object System.DirectoryServices.DirectoryEntry(
                "LDAP://$DomainController",
                $username,
                $password
            )
        }
        else {
            $entry = New-Object System.DirectoryServices.DirectoryEntry(
                "LDAP://$DomainController"
            )
        }

        $null = $entry.NativeObject

        Write-Log "LDAP bind successful" "SUCCESS"
    }
    catch {
        Write-Log "LDAP bind failed: $($_.Exception.Message)" "ERROR"
    }
}

function Get-SPNInformation {
    param(
        [string]$AccountName,
        [string]$DomainName,
        [PSCredential]$Credential
    )

    Write-Log "Searching SPNs for account: $AccountName"

    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $params = @{
            Filter = "SamAccountName -eq '$AccountName'"
            Properties = 'ServicePrincipalName'
            Server = $DomainName
        }

        if ($Credential) {
            $params.Credential = $Credential
        }

        $account = Get-ADUser @params -ErrorAction SilentlyContinue

        if (-not $account) {
            $account = Get-ADComputer @params -ErrorAction SilentlyContinue
        }

        if ($account) {
            if ($account.ServicePrincipalName.Count -gt 0) {
                foreach ($spn in $account.ServicePrincipalName) {
                    Write-Log "SPN: $spn" "SUCCESS"
                }
            }
            else {
                Write-Log "No SPNs found for account" "INFO"
            }
        }
        else {
            Write-Log "Account not found" "ERROR"
        }
    }
    catch {
        Write-Log "SPN query failed: $($_.Exception.Message)" "ERROR"
    }
}

function Test-SecureChannel {
    param(
        [string]$DomainName,
        [PSCredential]$Credential
    )

    Write-Log "Testing machine secure channel"

    try {
        if ($Credential) {
            $result = Test-ComputerSecureChannel -Server $DomainName -Credential $Credential -Verbose
        }
        else {
            $result = Test-ComputerSecureChannel -Server $DomainName -Verbose
        }

        if ($result) {
            Write-Log "Secure channel healthy" "SUCCESS"
        }
        else {
            Write-Log "Secure channel broken" "ERROR"
        }
    }
    catch {
        Write-Log "Secure channel test failed: $($_.Exception.Message)" "ERROR"
    }
}

function Export-Results {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveDialog.Title = "Export Troubleshooting Results"
    $saveDialog.FileName = "ADKerberosResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        try {
            $Script:Results | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            Write-Log "Results exported to $($saveDialog.FileName)" "SUCCESS"
        }
        catch {
            Write-Log "Export failed: $($_.Exception.Message)" "ERROR"
        }
    }
}

#====================================================
# GUI Creation
#====================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory Kerberos Troubleshooting Tool"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

#====================================================
# Connection Group
#====================================================

$grpConnection = New-Object System.Windows.Forms.GroupBox
$grpConnection.Text = "Connection Settings"
$grpConnection.Location = New-Object System.Drawing.Point(10,10)
$grpConnection.Size = New-Object System.Drawing.Size(1160,120)
$form.Controls.Add($grpConnection)

$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "Domain:"
$lblDomain.Location = New-Object System.Drawing.Point(20,35)
$lblDomain.AutoSize = $true
$grpConnection.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(100,30)
$txtDomain.Size = New-Object System.Drawing.Size(250,25)
$grpConnection.Controls.Add($txtDomain)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller:"
$lblDC.Location = New-Object System.Drawing.Point(380,35)
$lblDC.AutoSize = $true
$grpConnection.Controls.Add($lblDC)

$txtDC = New-Object System.Windows.Forms.TextBox
$txtDC.Location = New-Object System.Drawing.Point(520,30)
$txtDC.Size = New-Object System.Drawing.Size(250,25)
$grpConnection.Controls.Add($txtDC)

$btnCredential = New-Object System.Windows.Forms.Button
$btnCredential.Text = "Set Credentials"
$btnCredential.Location = New-Object System.Drawing.Point(800,28)
$btnCredential.Size = New-Object System.Drawing.Size(150,30)
$grpConnection.Controls.Add($btnCredential)

$lblCredStatus = New-Object System.Windows.Forms.Label
$lblCredStatus.Text = "Using Current Credentials"
$lblCredStatus.Location = New-Object System.Drawing.Point(20,75)
$lblCredStatus.Size = New-Object System.Drawing.Size(500,25)
$grpConnection.Controls.Add($lblCredStatus)

#====================================================
# Operations Group
#====================================================

$grpOperations = New-Object System.Windows.Forms.GroupBox
$grpOperations.Text = "Diagnostics"
$grpOperations.Location = New-Object System.Drawing.Point(10,140)
$grpOperations.Size = New-Object System.Drawing.Size(1160,220)
$form.Controls.Add($grpOperations)

$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Text = "Test Connectivity"
$btnPing.Location = New-Object System.Drawing.Point(20,35)
$btnPing.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnPing)

$btnDNS = New-Object System.Windows.Forms.Button
$btnDNS.Text = "Test DNS"
$btnDNS.Location = New-Object System.Drawing.Point(220,35)
$btnDNS.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnDNS)

$btnPorts = New-Object System.Windows.Forms.Button
$btnPorts.Text = "Test Ports"
$btnPorts.Location = New-Object System.Drawing.Point(420,35)
$btnPorts.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnPorts)

$btnLDAP = New-Object System.Windows.Forms.Button
$btnLDAP.Text = "Test LDAP Bind"
$btnLDAP.Location = New-Object System.Drawing.Point(620,35)
$btnLDAP.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnLDAP)

$btnSecure = New-Object System.Windows.Forms.Button
$btnSecure.Text = "Test Secure Channel"
$btnSecure.Location = New-Object System.Drawing.Point(820,35)
$btnSecure.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnSecure)

$btnTickets = New-Object System.Windows.Forms.Button
$btnTickets.Text = "View Kerberos Tickets"
$btnTickets.Location = New-Object System.Drawing.Point(20,95)
$btnTickets.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnTickets)

$btnPurge = New-Object System.Windows.Forms.Button
$btnPurge.Text = "Purge Tickets"
$btnPurge.Location = New-Object System.Drawing.Point(220,95)
$btnPurge.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnPurge)

$btnTime = New-Object System.Windows.Forms.Button
$btnTime.Text = "Check Time Sync"
$btnTime.Location = New-Object System.Drawing.Point(420,95)
$btnTime.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnTime)

$btnSPN = New-Object System.Windows.Forms.Button
$btnSPN.Text = "Lookup SPNs"
$btnSPN.Location = New-Object System.Drawing.Point(620,95)
$btnSPN.Size = New-Object System.Drawing.Size(180,40)
$grpOperations.Controls.Add($btnSPN)

$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "Run Full Diagnostics"
$btnAll.Location = New-Object System.Drawing.Point(820,95)
$btnAll.Size = New-Object System.Drawing.Size(180,40)
$btnAll.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$grpOperations.Controls.Add($btnAll)

#====================================================
# SPN Lookup
#====================================================

$lblSPN = New-Object System.Windows.Forms.Label
$lblSPN.Text = "SPN Account:"
$lblSPN.Location = New-Object System.Drawing.Point(20,165)
$lblSPN.AutoSize = $true
$grpOperations.Controls.Add($lblSPN)

$txtSPN = New-Object System.Windows.Forms.TextBox
$txtSPN.Location = New-Object System.Drawing.Point(130,160)
$txtSPN.Size = New-Object System.Drawing.Size(300,25)
$grpOperations.Controls.Add($txtSPN)

#====================================================
# Results Window
#====================================================

$grpResults = New-Object System.Windows.Forms.GroupBox
$grpResults.Text = "Results"
$grpResults.Location = New-Object System.Drawing.Point(10,370)
$grpResults.Size = New-Object System.Drawing.Size(1160,370)
$form.Controls.Add($grpResults)

$txtResults = New-Object System.Windows.Forms.RichTextBox
$txtResults.Location = New-Object System.Drawing.Point(10,25)
$txtResults.Size = New-Object System.Drawing.Size(1140,300)
$txtResults.BackColor = [System.Drawing.Color]::Black
$txtResults.ForeColor = [System.Drawing.Color]::LightGreen
$txtResults.Font = New-Object System.Drawing.Font("Consolas",10)
$txtResults.ReadOnly = $true
$grpResults.Controls.Add($txtResults)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Results"
$btnExport.Location = New-Object System.Drawing.Point(10,330)
$btnExport.Size = New-Object System.Drawing.Size(160,30)
$grpResults.Controls.Add($btnExport)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear Results"
$btnClear.Location = New-Object System.Drawing.Point(190,330)
$btnClear.Size = New-Object System.Drawing.Size(160,30)
$grpResults.Controls.Add($btnClear)

#====================================================
# Event Handlers
#====================================================

$btnCredential.Add_Click({
    try {
        $cred = Get-Credential -Message "Enter alternate domain credentials"

        if ($cred) {
            $Script:Credential = $cred
            $lblCredStatus.Text = "Using Credentials: $($cred.UserName)"
            Write-Log "Alternate credentials configured" "SUCCESS"
        }
    }
    catch {
        Write-Log "Credential prompt failed: $($_.Exception.Message)" "ERROR"
    }
})

$btnPing.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDC.Text)) {
        Write-Log "Please specify a domain controller" "ERROR"
        return
    }

    Test-DomainControllerConnectivity -DomainController $txtDC.Text
})

$btnDNS.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) {
        Write-Log "Please specify a domain" "ERROR"
        return
    }

    Test-DnsResolution -DomainName $txtDomain.Text
})

$btnPorts.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDC.Text)) {
        Write-Log "Please specify a domain controller" "ERROR"
        return
    }

    Test-PortConnectivity -Server $txtDC.Text
})

$btnLDAP.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDC.Text)) {
        Write-Log "Please specify a domain controller" "ERROR"
        return
    }

    Test-LdapBind -DomainController $txtDC.Text -Credential (Get-SelectedCredential)
})

$btnSecure.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) {
        Write-Log "Please specify a domain" "ERROR"
        return
    }

    Test-SecureChannel -DomainName $txtDomain.Text -Credential (Get-SelectedCredential)
})

$btnTickets.Add_Click({
    Test-KerberosTickets
})

$btnPurge.Add_Click({
    Purge-KerberosTickets
})

$btnTime.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtDC.Text)) {
        Write-Log "Please specify a domain controller" "ERROR"
        return
    }

    Test-TimeSynchronization -DomainController $txtDC.Text
})

$btnSPN.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSPN.Text)) {
        Write-Log "Please specify an account name" "ERROR"
        return
    }

    if ([string]::IsNullOrWhiteSpace($txtDomain.Text)) {
        Write-Log "Please specify a domain" "ERROR"
        return
    }

    Get-SPNInformation -AccountName $txtSPN.Text -DomainName $txtDomain.Text -Credential (Get-SelectedCredential)
})

$btnAll.Add_Click({
    Write-Log "================================================="
    Write-Log "Starting Full Active Directory Kerberos Diagnostics"
    Write-Log "================================================="

    if (-not [string]::IsNullOrWhiteSpace($txtDC.Text)) {
        Test-DomainControllerConnectivity -DomainController $txtDC.Text
        Test-PortConnectivity -Server $txtDC.Text
        Test-TimeSynchronization -DomainController $txtDC.Text
        Test-LdapBind -DomainController $txtDC.Text -Credential (Get-SelectedCredential)
    }

    if (-not [string]::IsNullOrWhiteSpace($txtDomain.Text)) {
        Test-DnsResolution -DomainName $txtDomain.Text
        Test-SecureChannel -DomainName $txtDomain.Text -Credential (Get-SelectedCredential)
    }

    Test-KerberosTickets

    Write-Log "================================================="
    Write-Log "Diagnostics Completed"
    Write-Log "================================================="
})

$btnExport.Add_Click({
    Export-Results
})

$btnClear.Add_Click({
    $txtResults.Clear()
    $Script:Results = @()
})

#====================================================
# Startup
#====================================================

Write-Log "Active Directory Kerberos Troubleshooting Tool Started" "SUCCESS"
Write-Log "Running as: $([Environment]::UserName)"
Write-Log "Computer: $([Environment]::MachineName)"

#====================================================
# Launch GUI
#====================================================

[void]$form.ShowDialog()
```

## Recommended Enhancements

### Optional Features You Can Add

* Multi-threaded background jobs for faster scans
* Kerberos event log parser
* NTLM fallback detection
* Trust relationship validation
* Replication status checks
* GPO validation
* Forest-wide SPN duplicate scanning
* Certificate/KDC validation
* Live DC health dashboard
* Integrated packet capture launcher
* HTML reporting
* Dark/light mode toggle
* Saved connection profiles
* Azure AD / Entra hybrid checks
* RDP connectivity testing
* WinRM validation

## Recommended Execution

Run PowerShell as Administrator:

```powershell
Set-ExecutionPolicy Bypass -Scope Process
.\ADKerberosTool.ps1
```

## Required RSAT Module

Install RSAT Active Directory tools if missing:

```powershell
Get-WindowsCapability -Name RSAT.ActiveDirectory* -Online | Add-WindowsCapability -Online
```
