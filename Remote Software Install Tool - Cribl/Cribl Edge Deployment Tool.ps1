Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Cribl Edge Deployment Tool (PsExec)"
$form.Size = New-Object System.Drawing.Size(1050, 850)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Consolas", 9)

# Credentials Group
$credGroup = New-Object System.Windows.Forms.GroupBox
$credGroup.Location = New-Object System.Drawing.Point(10, 10)
$credGroup.Size = New-Object System.Drawing.Size(1015, 110)
$credGroup.Text = "Credentials & PsExec Configuration"
$form.Controls.Add($credGroup)

$currentCredRadio = New-Object System.Windows.Forms.RadioButton
$currentCredRadio.Location = New-Object System.Drawing.Point(15, 25)
$currentCredRadio.Size = New-Object System.Drawing.Size(200, 20)
$currentCredRadio.Text = "Use Current Credentials"
$currentCredRadio.Checked = $true
$credGroup.Controls.Add($currentCredRadio)

$altCredRadio = New-Object System.Windows.Forms.RadioButton
$altCredRadio.Location = New-Object System.Drawing.Point(15, 50)
$altCredRadio.Size = New-Object System.Drawing.Size(150, 20)
$altCredRadio.Text = "Use Alternate Credentials"
$credGroup.Controls.Add($altCredRadio)

# Domain\Username input
$domainUserLabel = New-Object System.Windows.Forms.Label
$domainUserLabel.Location = New-Object System.Drawing.Point(170, 48)
$domainUserLabel.Size = New-Object System.Drawing.Size(120, 20)
$domainUserLabel.Text = "Domain\Username:"
$credGroup.Controls.Add($domainUserLabel)

$domainUserInput = New-Object System.Windows.Forms.TextBox
$domainUserInput.Location = New-Object System.Drawing.Point(295, 48)
$domainUserInput.Size = New-Object System.Drawing.Size(200, 20)
$domainUserInput.Enabled = $false
$credGroup.Controls.Add($domainUserInput)

# Password input
$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Location = New-Object System.Drawing.Point(505, 48)
$passwordLabel.Size = New-Object System.Drawing.Size(65, 20)
$passwordLabel.Text = "Password:"
$credGroup.Controls.Add($passwordLabel)

$passwordInput = New-Object System.Windows.Forms.MaskedTextBox
$passwordInput.Location = New-Object System.Drawing.Point(575, 48)
$passwordInput.Size = New-Object System.Drawing.Size(200, 20)
$passwordInput.PasswordChar = '*'
$passwordInput.Enabled = $false
$credGroup.Controls.Add($passwordInput)

# PsExec path
$psexecLabel = New-Object System.Windows.Forms.Label
$psexecLabel.Location = New-Object System.Drawing.Point(15, 78)
$psexecLabel.Size = New-Object System.Drawing.Size(100, 20)
$psexecLabel.Text = "PsExec Location:"
$credGroup.Controls.Add($psexecLabel)

$psexecPath = New-Object System.Windows.Forms.TextBox
$psexecPath.Location = New-Object System.Drawing.Point(120, 75)
$psexecPath.Size = New-Object System.Drawing.Size(350, 20)
$psexecPath.Text = "C:\PsTools\PsExec.exe"
$credGroup.Controls.Add($psexecPath)

$psexecBrowse = New-Object System.Windows.Forms.Button
$psexecBrowse.Location = New-Object System.Drawing.Point(475, 73)
$psexecBrowse.Size = New-Object System.Drawing.Size(30, 23)
$psexecBrowse.Text = "..."
$credGroup.Controls.Add($psexecBrowse)

$downloadPsexec = New-Object System.Windows.Forms.Button
$downloadPsexec.Location = New-Object System.Drawing.Point(510, 73)
$downloadPsexec.Size = New-Object System.Drawing.Size(120, 23)
$downloadPsexec.Text = "Download PsExec"
$credGroup.Controls.Add($downloadPsexec)

$credStatus = New-Object System.Windows.Forms.Label
$credStatus.Location = New-Object System.Drawing.Point(640, 78)
$credStatus.Size = New-Object System.Drawing.Size(350, 20)
$credStatus.Text = "Using current credentials"
$credStatus.ForeColor = [System.Drawing.Color]::Gray
$credGroup.Controls.Add($credStatus)

# Server Input Group (Left Panel)
$serverGroup = New-Object System.Windows.Forms.GroupBox
$serverGroup.Location = New-Object System.Drawing.Point(10, 130)
$serverGroup.Size = New-Object System.Drawing.Size(510, 220)
$serverGroup.Text = "Server Names (one per line)"
$form.Controls.Add($serverGroup)

$serverInput = New-Object System.Windows.Forms.TextBox
$serverInput.Location = New-Object System.Drawing.Point(15, 25)
$serverInput.Size = New-Object System.Drawing.Size(480, 155)
$serverInput.Multiline = $true
$serverInput.ScrollBars = "Vertical"
$serverInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$serverGroup.Controls.Add($serverInput)

# Import/Export buttons for servers
$importServers = New-Object System.Windows.Forms.Button
$importServers.Location = New-Object System.Drawing.Point(15, 188)
$importServers.Size = New-Object System.Drawing.Size(100, 23)
$importServers.Text = "Import from File"
$serverGroup.Controls.Add($importServers)

$exportServers = New-Object System.Windows.Forms.Button
$exportServers.Location = New-Object System.Drawing.Point(120, 188)
$exportServers.Size = New-Object System.Drawing.Size(100, 23)
$exportServers.Text = "Export Servers"
$serverGroup.Controls.Add($exportServers)

$clearServers = New-Object System.Windows.Forms.Button
$clearServers.Location = New-Object System.Drawing.Point(395, 188)
$clearServers.Size = New-Object System.Drawing.Size(100, 23)
$clearServers.Text = "Clear"
$serverGroup.Controls.Add($clearServers)

# Options Group
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Location = New-Object System.Drawing.Point(10, 360)
$optionsGroup.Size = New-Object System.Drawing.Size(510, 100)
$optionsGroup.Text = "Installation Options"
$form.Controls.Add($optionsGroup)

$checkInstalledRadio = New-Object System.Windows.Forms.RadioButton
$checkInstalledRadio.Location = New-Object System.Drawing.Point(15, 25)
$checkInstalledRadio.Size = New-Object System.Drawing.Size(480, 20)
$checkInstalledRadio.Text = "Install Only (skip if already installed)"
$checkInstalledRadio.Checked = $true
$optionsGroup.Controls.Add($checkInstalledRadio)

$uninstallReinstallRadio = New-Object System.Windows.Forms.RadioButton
$uninstallReinstallRadio.Location = New-Object System.Drawing.Point(15, 50)
$uninstallReinstallRadio.Size = New-Object System.Drawing.Size(480, 20)
$uninstallReinstallRadio.Text = "Uninstall if present, then Install"
$optionsGroup.Controls.Add($uninstallReinstallRadio)

$forceReinstallCheck = New-Object System.Windows.Forms.CheckBox
$forceReinstallCheck.Location = New-Object System.Drawing.Point(15, 75)
$forceReinstallCheck.Size = New-Object System.Drawing.Size(480, 20)
$forceReinstallCheck.Text = "Force reinstall even if already installed (ignore version check)"
$optionsGroup.Controls.Add($forceReinstallCheck)

# Output/Progress Group (Right Panel)
$outputGroup = New-Object System.Windows.Forms.GroupBox
$outputGroup.Location = New-Object System.Drawing.Point(530, 130)
$outputGroup.Size = New-Object System.Drawing.Size(495, 330)
$outputGroup.Text = "Output/Progress"
$form.Controls.Add($outputGroup)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(15, 25)
$outputBox.Size = New-Object System.Drawing.Size(465, 295)
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::LimeGreen
$outputGroup.Controls.Add($outputBox)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 470)
$progressBar.Size = New-Object System.Drawing.Size(1015, 20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 495)
$statusLabel.Size = New-Object System.Drawing.Size(600, 20)
$statusLabel.Text = "Ready"
$form.Controls.Add($statusLabel)

# Script Command Group
$scriptGroup = New-Object System.Windows.Forms.GroupBox
$scriptGroup.Location = New-Object System.Drawing.Point(10, 520)
$scriptGroup.Size = New-Object System.Drawing.Size(1015, 120)
$scriptGroup.Text = "Installation Command (this will be run locally on each remote server)"
$form.Controls.Add($scriptGroup)

$scriptCommand = New-Object System.Windows.Forms.TextBox
$scriptCommand.Location = New-Object System.Drawing.Point(15, 20)
$scriptCommand.Size = New-Object System.Drawing.Size(985, 65)
$scriptCommand.Multiline = $true
$scriptCommand.ScrollBars = "Vertical"
$scriptCommand.Font = New-Object System.Drawing.Font("Consolas", 9)
$scriptCommand.Text = @'
msiexec /i "https://cdn.cribl.io/dl/4.17.1/cribl-4.17.1-b862732f-win32-x64.msi" /qn MODE="mode-managed-edge" HOSTNAME="XXX.cribl.cloud" PORT="4200" FLEET="XXX" AUTH="XXX" TLS="true" USERNAME="LocalSystem" APPLICATIONROOTDIRECTORY="C:\Program Files\Cribl\" /log "%SYSTEMROOT%\Temp\cribl-msiexec-install.log"
'@
$scriptGroup.Controls.Add($scriptCommand)

$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(15, 90)
$infoLabel.Size = New-Object System.Drawing.Size(985, 20)
$infoLabel.Text = "Note: This command will be executed using PsExec. Use %SYSTEMROOT% instead of `$env:SYSTEMROOT for remote execution."
$infoLabel.ForeColor = [System.Drawing.Color]::Gray
$scriptGroup.Controls.Add($infoLabel)

# Control Buttons
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Location = New-Object System.Drawing.Point(10, 650)
$checkButton.Size = New-Object System.Drawing.Size(150, 40)
$checkButton.Text = "Check for Cribl"
$checkButton.BackColor = [System.Drawing.Color]::LightBlue
$checkButton.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($checkButton)

$uninstallButton = New-Object System.Windows.Forms.Button
$uninstallButton.Location = New-Object System.Drawing.Point(170, 650)
$uninstallButton.Size = New-Object System.Drawing.Size(150, 40)
$uninstallButton.Text = "Uninstall Cribl"
$uninstallButton.BackColor = [System.Drawing.Color]::LightSalmon
$uninstallButton.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($uninstallButton)

$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Location = New-Object System.Drawing.Point(330, 650)
$executeButton.Size = New-Object System.Drawing.Size(150, 40)
$executeButton.Text = "Deploy Cribl"
$executeButton.BackColor = [System.Drawing.Color]::LightGreen
$executeButton.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($executeButton)

$testButton = New-Object System.Windows.Forms.Button
$testButton.Location = New-Object System.Drawing.Point(490, 650)
$testButton.Size = New-Object System.Drawing.Size(150, 40)
$testButton.Text = "Test Connection"
$form.Controls.Add($testButton)

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(650, 650)
$stopButton.Size = New-Object System.Drawing.Size(150, 40)
$stopButton.Text = "Stop Execution"
$stopButton.Enabled = $false
$form.Controls.Add($stopButton)

$logButton = New-Object System.Windows.Forms.Button
$logButton.Location = New-Object System.Drawing.Point(810, 650)
$logButton.Size = New-Object System.Drawing.Size(100, 40)
$logButton.Text = "View Log"
$form.Controls.Add($logButton)

$exportLogButton = New-Object System.Windows.Forms.Button
$exportLogButton.Location = New-Object System.Drawing.Point(920, 650)
$exportLogButton.Size = New-Object System.Drawing.Size(100, 40)
$exportLogButton.Text = "Export Report"
$form.Controls.Add($exportLogButton)

# Global variables
$script:alternateCredential = $null
$script:isRunning = $false
$script:logFilePath = ""
$script:results = @()

# Helper function to get server list
function Get-ServerList {
    $rawText = $serverInput.Text
    
    if ([string]::IsNullOrWhiteSpace($rawText)) {
        return @()
    }
    
    # Split by newlines properly using regex
    $servers = $rawText -split '\r?\n' | 
               Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | 
               ForEach-Object { $_.Trim() } |
               Where-Object { $_ -ne "" }
    
    return @($servers)
}

# Functions
function Write-OutputBox {
    param([string]$Message, [string]$Color = "LimeGreen")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $outputBox.SelectionColor = [System.Drawing.Color]::Gray
    $outputBox.AppendText("[$timestamp] ")
    $outputBox.SelectionColor = [System.Drawing.Color]::$Color
    $outputBox.AppendText("$Message`r`n")
    $outputBox.ScrollToCaret()
    
    if ($script:logFilePath) {
        "[$timestamp] $Message" | Out-File -FilePath $script:logFilePath -Append
    }
}

function Test-PsExecExists {
    $path = $psexecPath.Text.Trim()
    if (-not (Test-Path $path)) {
        throw "PsExec not found at: $path`r`nPlease download PsExec from Microsoft Sysinternals or click 'Download PsExec' button."
    }
    return $path
}

# Create a batch file for the command to execute
function New-CommandBatchFile {
    param([string]$Command)
    
    $tempDir = Join-Path $env:TEMP "CriblDeploy"
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    }
    
    $batchFile = Join-Path $tempDir "cribl_command_$(Get-Date -Format 'yyyyMMdd_HHmmss').cmd"
    
    # Create batch file with the command and logging
    @"
@echo off
echo [%date% %time%] Starting Cribl deployment > "%SYSTEMROOT%\Temp\cribl_deploy.log"
echo Command: $Command >> "%SYSTEMROOT%\Temp\cribl_deploy.log"
$Command >> "%SYSTEMROOT%\Temp\cribl_deploy.log" 2>&1
echo Exit Code: %ERRORLEVEL% >> "%SYSTEMROOT%\Temp\cribl_deploy.log"
echo [%date% %time%] Deployment completed with exit code: %ERRORLEVEL% >> "%SYSTEMROOT%\Temp\cribl_deploy.log"
exit %ERRORLEVEL%
"@ | Out-File -FilePath $batchFile -Encoding ASCII
    
    return $batchFile
}

# Check if Cribl is installed using PsExec
function Test-CriblInstalledPsExec {
    param(
        [string]$ComputerName,
        [string]$PsExecPath
    )
    
    $checkScript = @'
@echo off
if exist "C:\Program Files\Cribl\bin\cribl.exe" (
    echo INSTALLED:C:\Program Files\Cribl
    exit /b 0
)
if exist "C:\Program Files (x86)\Cribl\bin\cribl.exe" (
    echo INSTALLED:C:\Program Files (x86)\Cribl
    exit /b 0
)
reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Cribl" 2>nul | find "DisplayName" >nul
if %ERRORLEVEL% EQU 0 (
    echo INSTALLED:Registry
    exit /b 0
)
reg query "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Cribl" 2>nul | find "DisplayName" >nul
if %ERRORLEVEL% EQU 0 (
    echo INSTALLED:Registry
    exit /b 0
)
echo NOT_INSTALLED
exit /b 1
'@
    
    $tempDir = Join-Path $env:TEMP "CriblDeploy"
    $checkFile = Join-Path $tempDir "check_cribl_$(Get-Date -Format 'yyyyMMdd_HHmmss').cmd"
    $checkScript | Out-File -FilePath $checkFile -Encoding ASCII
    
    $arguments = @("\\$ComputerName")
    
    if ($altCredRadio.Checked -and $script:alternateCredential) {
        $arguments += "-u", $script:alternateCredential.UserName
        $arguments += "-p", $script:alternateCredential.GetNetworkCredential().Password
    }
    
    $arguments += "-s", "-h", "-accepteula", "cmd", "/c", $checkFile
    
    try {
        $result = & $PsExecPath $arguments 2>&1
        return $result
    }
    catch {
        throw "PsExec check failed: $_"
    }
    finally {
        Remove-Item $checkFile -Force -ErrorAction SilentlyContinue
    }
}

# Uninstall Cribl using PsExec
function Uninstall-CriblPsExec {
    param(
        [string]$ComputerName,
        [string]$PsExecPath
    )
    
    $uninstallScript = @'
@echo off
echo [%date% %time%] Starting Cribl uninstall > "%SYSTEMROOT%\Temp\cribl_uninstall.log"

REM Try to find product code from registry
for /f "tokens=*" %%a in ('reg query "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Cribl" 2^>nul ^| find "{"') do (
    set PRODUCT_CODE=%%a
    goto :found
)
for /f "tokens=*" %%a in ('reg query "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" /s /f "Cribl" 2^>nul ^| find "{"') do (
    set PRODUCT_CODE=%%a
    goto :found
)

:notfound
echo No MSI product code found, attempting manual removal >> "%SYSTEMROOT%\Temp\cribl_uninstall.log"

REM Stop services
sc stop CriblEdge 2>nul
sc delete CriblEdge 2>nul
sc stop Cribl 2>nul
sc delete Cribl 2>nul

REM Remove directories
if exist "C:\Program Files\Cribl" rmdir /s /q "C:\Program Files\Cribl"
if exist "C:\Program Files (x86)\Cribl" rmdir /s /q "C:\Program Files (x86)\Cribl"
if exist "%PROGRAMDATA%\Cribl" rmdir /s /q "%PROGRAMDATA%\Cribl"

echo Manual removal completed >> "%SYSTEMROOT%\Temp\cribl_uninstall.log"
exit /b 0

:found
echo Found product code: %PRODUCT_CODE% >> "%SYSTEMROOT%\Temp\cribl_uninstall.log"
msiexec /x %PRODUCT_CODE% /qn /norestart /log "%SYSTEMROOT%\Temp\cribl_uninstall_msi.log"
echo Uninstall exit code: %ERRORLEVEL% >> "%SYSTEMROOT%\Temp\cribl_uninstall.log"
exit /b %ERRORLEVEL%
'@
    
    $tempDir = Join-Path $env:TEMP "CriblDeploy"
    $uninstallFile = Join-Path $tempDir "uninstall_cribl_$(Get-Date -Format 'yyyyMMdd_HHmmss').cmd"
    $uninstallScript | Out-File -FilePath $uninstallFile -Encoding ASCII
    
    $arguments = @("\\$ComputerName")
    
    if ($altCredRadio.Checked -and $script:alternateCredential) {
        $arguments += "-u", $script:alternateCredential.UserName
        $arguments += "-p", $script:alternateCredential.GetNetworkCredential().Password
    }
    
    $arguments += "-s", "-h", "-accepteula", "cmd", "/c", $uninstallFile
    
    try {
        $result = & $PsExecPath $arguments 2>&1
        return $result
    }
    catch {
        throw "PsExec uninstall failed: $_"
    }
    finally {
        Remove-Item $uninstallFile -Force -ErrorAction SilentlyContinue
    }
}

# Deploy Cribl using PsExec
function Deploy-CriblPsExec {
    param(
        [string]$ComputerName,
        [string]$PsExecPath,
        [string]$Command
    )
    
    $batchFile = New-CommandBatchFile -Command $Command
    
    $arguments = @("\\$ComputerName")
    
    if ($altCredRadio.Checked -and $script:alternateCredential) {
        $arguments += "-u", $script:alternateCredential.UserName
        $arguments += "-p", $script:alternateCredential.GetNetworkCredential().Password
    }
    
    $arguments += "-s", "-h", "-accepteula", "cmd", "/c", $batchFile
    
    try {
        $result = & $PsExecPath $arguments 2>&1
        return $result
    }
    catch {
        throw "PsExec deployment failed: $_"
    }
    finally {
        Remove-Item $batchFile -Force -ErrorAction SilentlyContinue
    }
}

# Event Handlers
$psexecBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PsExec.exe|PsExec.exe|Executable files (*.exe)|*.exe|All files (*.*)|*.*"
    $openFileDialog.Title = "Select PsExec.exe"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $psexecPath.Text = $openFileDialog.FileName
        Write-OutputBox "PsExec path set to: $($openFileDialog.FileName)" "Cyan"
    }
})

$downloadPsexec.Add_Click({
    Write-OutputBox "Opening PsTools download page..." "Yellow"
    Start-Process "https://learn.microsoft.com/en-us/sysinternals/downloads/psexec"
    Write-OutputBox "Download PsTools, extract PsExec.exe, and set the path above" "Cyan"
})

$altCredRadio.Add_CheckedChanged({
    $isAlt = $altCredRadio.Checked
    $domainUserInput.Enabled = $isAlt
    $passwordInput.Enabled = $isAlt
    
    if (-not $isAlt) {
        $domainUserInput.Text = ""
        $passwordInput.Text = ""
        $script:alternateCredential = $null
        $credStatus.Text = "Using current credentials"
        $credStatus.ForeColor = [System.Drawing.Color]::Gray
    }
    else {
        $credStatus.Text = "Enter alternate domain credentials"
        $credStatus.ForeColor = [System.Drawing.Color]::Blue
    }
})

$currentCredRadio.Add_CheckedChanged({
    if ($currentCredRadio.Checked) {
        $script:alternateCredential = $null
        $credStatus.Text = "Using current credentials"
        $credStatus.ForeColor = [System.Drawing.Color]::Gray
    }
})

$importServers.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $servers = Get-Content $openFileDialog.FileName
        $serverInput.Text = ($servers -join "`r`n").Trim()
        Write-OutputBox "Imported servers from: $($openFileDialog.FileName)" "Cyan"
    }
})

$exportServers.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
    $saveFileDialog.DefaultExt = "txt"
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $serverInput.Text | Out-File -FilePath $saveFileDialog.FileName
        Write-OutputBox "Exported servers to: $($saveFileDialog.FileName)" "Cyan"
    }
})

$clearServers.Add_Click({
    $serverInput.Clear()
    Write-OutputBox "Server list cleared" "Gray"
})

$testButton.Add_Click({
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        Write-OutputBox "No servers specified" "Red"
        return
    }
    
    try {
        $psexec = Test-PsExecExists
    }
    catch {
        Write-OutputBox $_.Exception.Message "Red"
        return
    }
    
    if ($altCredRadio.Checked) {
        $domainUser = $domainUserInput.Text.Trim()
        $password = $passwordInput.Text
        
        if ([string]::IsNullOrEmpty($domainUser) -or [string]::IsNullOrEmpty($password)) {
            Write-OutputBox "Please enter both Domain\Username and Password" "Red"
            return
        }
        
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $script:alternateCredential = New-Object System.Management.Automation.PSCredential($domainUser, $securePassword)
    }
    
    Write-OutputBox "`r`n=== Testing Connections with PsExec ===" "White"
    
    foreach ($server in $servers) {
        Write-OutputBox "Testing $server..." "Yellow"
        
        try {
            $arguments = @("\\$server")
            
            if ($script:alternateCredential) {
                $arguments += "-u", $script:alternateCredential.UserName
                $arguments += "-p", $script:alternateCredential.GetNetworkCredential().Password
            }
            
            $arguments += "-accepteula", "cmd", "/c", "echo OK"
            
            $result = & $psexec $arguments 2>&1 | Out-String
            
            if ($result -match "OK") {
                Write-OutputBox "✓ $server - Connection successful" "Green"
            }
            else {
                Write-OutputBox "✗ $server - Connection failed: $result" "Red"
            }
        }
        catch {
            Write-OutputBox "✗ $server - Error: $_" "Red"
        }
    }
})

$checkButton.Add_Click({
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        Write-OutputBox "No servers specified" "Red"
        return
    }
    
    try {
        $psexec = Test-PsExecExists
    }
    catch {
        Write-OutputBox $_.Exception.Message "Red"
        return
    }
    
    if ($altCredRadio.Checked -and -not $script:alternateCredential) {
        $domainUser = $domainUserInput.Text.Trim()
        $password = $passwordInput.Text
        if ([string]::IsNullOrEmpty($domainUser) -or [string]::IsNullOrEmpty($password)) {
            Write-OutputBox "Please enter both Domain\Username and Password" "Red"
            return
        }
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $script:alternateCredential = New-Object System.Management.Automation.PSCredential($domainUser, $securePassword)
    }
    
    Write-OutputBox "`r`n=== Checking for Cribl Installation ===" "White"
    
    foreach ($server in $servers) {
        Write-OutputBox "Checking $server..." "Yellow"
        
        try {
            $result = Test-CriblInstalledPsExec -ComputerName $server -PsExecPath $psexec
            $resultString = $result | Out-String
            
            if ($resultString -match "INSTALLED:(.+)") {
                Write-OutputBox "✓ $server - Cribl IS installed ($($matches[1].Trim()))" "Yellow"
            }
            else {
                Write-OutputBox "○ $server - Cribl NOT installed" "Cyan"
            }
        }
        catch {
            Write-OutputBox "✗ $server - Check failed: $_" "Red"
        }
    }
})

$uninstallButton.Add_Click({
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        Write-OutputBox "No servers specified" "Red"
        return
    }
    
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to uninstall Cribl from $($servers.Count) server(s)?`r`n`r`nServers: $($servers -join ', ')",
        "Confirm Uninstall",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($confirmResult -ne 'Yes') { return }
    
    try {
        $psexec = Test-PsExecExists
    }
    catch {
        Write-OutputBox $_.Exception.Message "Red"
        return
    }
    
    if ($altCredRadio.Checked -and -not $script:alternateCredential) {
        $domainUser = $domainUserInput.Text.Trim()
        $password = $passwordInput.Text
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $script:alternateCredential = New-Object System.Management.Automation.PSCredential($domainUser, $securePassword)
    }
    
    $script:logFilePath = Join-Path $env:TEMP "CriblUninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-OutputBox "`r`n=== Starting Cribl Uninstall ===" "White"
    
    $script:isRunning = $true
    $uninstallButton.Enabled = $false
    $checkButton.Enabled = $false
    $executeButton.Enabled = $false
    $stopButton.Enabled = $true
    $progressBar.Visible = $true
    $progressBar.Maximum = $servers.Count
    $progressBar.Value = 0
    
    $script:results = @()
    
    for ($i = 0; $i -lt $servers.Count; $i++) {
        if (-not $script:isRunning) { break }
        
        $server = $servers[$i]
        Write-OutputBox "[$($i+1)/$($servers.Count)] Uninstalling from $server..." "Yellow"
        
        try {
            $result = Uninstall-CriblPsExec -ComputerName $server -PsExecPath $psexec
            $resultString = $result | Out-String
            
            if ($resultString -match "exit code: 0") {
                Write-OutputBox "✓ $server - Uninstall successful" "Green"
                $script:results += [PSCustomObject]@{Server=$server; Status="Success"; Details=$resultString}
            }
            else {
                Write-OutputBox "⚠ $server - Uninstall completed with warnings" "Yellow"
                $script:results += [PSCustomObject]@{Server=$server; Status="Warning"; Details=$resultString}
            }
        }
        catch {
            Write-OutputBox "✗ $server - Error: $_" "Red"
            $script:results += [PSCustomObject]@{Server=$server; Status="Error"; Details=$_.Exception.Message}
        }
        
        $progressBar.Value = $i + 1
        $statusLabel.Text = "Uninstalling: $server ($($i+1)/$($servers.Count))"
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    $successCount = ($script:results | Where-Object {$_.Status -eq "Success"}).Count
    Write-OutputBox "`r`n=== Complete: $successCount/$($servers.Count) successful ===" "White"
    
    $script:isRunning = $false
    $uninstallButton.Enabled = $true
    $checkButton.Enabled = $true
    $executeButton.Enabled = $true
    $stopButton.Enabled = $false
    $progressBar.Visible = $false
    $statusLabel.Text = "Ready"
})

$executeButton.Add_Click({
    if ($script:isRunning) {
        Write-OutputBox "Execution already in progress" "Red"
        return
    }
    
    $servers = Get-ServerList
    if ($servers.Count -eq 0) {
        Write-OutputBox "No servers specified" "Red"
        return
    }
    
    try {
        $psexec = Test-PsExecExists
    }
    catch {
        Write-OutputBox $_.Exception.Message "Red"
        return
    }
    
    if ($altCredRadio.Checked -and -not $script:alternateCredential) {
        $domainUser = $domainUserInput.Text.Trim()
        $password = $passwordInput.Text
        if ([string]::IsNullOrEmpty($domainUser) -or [string]::IsNullOrEmpty($password)) {
            Write-OutputBox "Please enter both Domain\Username and Password" "Red"
            return
        }
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $script:alternateCredential = New-Object System.Management.Automation.PSCredential($domainUser, $securePassword)
    }
    
    $command = $scriptCommand.Text.Trim()
    if ([string]::IsNullOrEmpty($command)) {
        Write-OutputBox "No installation command specified" "Red"
        return
    }
    
    $script:logFilePath = Join-Path $env:TEMP "CriblDeployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-OutputBox "`r`n=== Starting Cribl Deployment with PsExec ===" "White"
    Write-OutputBox "Target servers: $($servers.Count)" "Cyan"
    Write-OutputBox "Command: $command" "Gray"
    Write-OutputBox "Log file: $script:logFilePath" "Cyan"
    
    $script:isRunning = $true
    $executeButton.Enabled = $false
    $uninstallButton.Enabled = $false
    $checkButton.Enabled = $false
    $testButton.Enabled = $false
    $stopButton.Enabled = $true
    $progressBar.Visible = $true
    $progressBar.Maximum = $servers.Count
    $progressBar.Value = 0
    
    $statusLabel.Text = "Deploying..."
    
    $script:results = @()
    $shouldUninstall = $uninstallReinstallRadio.Checked
    $forceReinstall = $forceReinstallCheck.Checked
    
    for ($i = 0; $i -lt $servers.Count; $i++) {
        if (-not $script:isRunning) { break }
        
        $server = $servers[$i]
        Write-OutputBox "`r`n[$($i+1)/$($servers.Count)] Processing $server..." "Yellow"
        
        try {
            # Check if already installed
            if (-not $forceReinstall) {
                $criblCheck = Test-CriblInstalledPsExec -ComputerName $server -PsExecPath $psexec
                $criblCheckString = $criblCheck | Out-String
                
                if ($criblCheckString -match "INSTALLED") {
                    Write-OutputBox "$server - Cribl already installed" "Yellow"
                    
                    if ($shouldUninstall) {
                        Write-OutputBox "$server - Uninstalling existing Cribl..." "Yellow"
                        $uninstallResult = Uninstall-CriblPsExec -ComputerName $server -PsExecPath $psexec
                        Write-OutputBox "✓ $server - Uninstall completed" "Green"
                    }
                    elseif ($checkInstalledRadio.Checked) {
                        Write-OutputBox "○ $server - Skipping (Install Only mode)" "Cyan"
                        $script:results += [PSCustomObject]@{Server=$server; Status="Skipped"; Details="Already installed"}
                        continue
                    }
                }
            }
            
            # Deploy Cribl
            Write-OutputBox "$server - Deploying Cribl..." "Yellow"
            
            $deployResult = Deploy-CriblPsExec -ComputerName $server -PsExecPath $psexec -Command $command
            $deployResultString = $deployResult | Out-String
            
            if ($deployResultString -match "exit code: 0") {
                Write-OutputBox "✓ $server - Deployment successful" "Green"
                $script:results += [PSCustomObject]@{Server=$server; Status="Success"; Details="Deployed successfully"}
            }
            elseif ($deployResultString -match "exit code:") {
                Write-OutputBox "⚠ $server - Deployment completed with non-zero exit code" "Yellow"
                $script:results += [PSCustomObject]@{Server=$server; Status="Warning"; Details=$deployResultString}
            }
            else {
                Write-OutputBox "? $server - Deployment status unknown" "Yellow"
                $script:results += [PSCustomObject]@{Server=$server; Status="Unknown"; Details=$deployResultString}
            }
        }
        catch {
            Write-OutputBox "✗ $server - Error: $_" "Red"
            $script:results += [PSCustomObject]@{Server=$server; Status="Error"; Details=$_.Exception.Message}
        }
        
        $progressBar.Value = $i + 1
        $statusLabel.Text = "Processing: $server ($($i+1)/$($servers.Count))"
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Summary
    Write-OutputBox "`r`n========================================" "White"
    Write-OutputBox "=== Deployment Summary ===" "White"
    Write-OutputBox "========================================" "White"
    
    $successCount = ($script:results | Where-Object {$_.Status -eq "Success"}).Count
    $skipCount = ($script:results | Where-Object {$_.Status -eq "Skipped"}).Count
    $failCount = ($script:results | Where-Object {$_.Status -in @("Error","Failed")}).Count
    $warnCount = ($script:results | Where-Object {$_.Status -eq "Warning"}).Count
    
    Write-OutputBox "✓ Success: $successCount" "Green"
    Write-OutputBox "○ Skipped: $skipCount" "Cyan"
    Write-OutputBox "⚠ Warnings: $warnCount" "Yellow"
    Write-OutputBox "✗ Failed: $failCount" "Red"
    Write-OutputBox "Total: $($script:results.Count)" "White"
    
    if ($failCount -gt 0) {
        Write-OutputBox "`r`nFailed servers:" "Red"
        $script:results | Where-Object {$_.Status -in @("Error","Failed")} | ForEach-Object {
            Write-OutputBox "  - $($_.Server)" "Red"
        }
    }
    
    Write-OutputBox "`r`nLog file: $script:logFilePath" "Cyan"
    Write-OutputBox "Remote logs: %SYSTEMROOT%\Temp\cribl_deploy.log on each server" "Gray"
    
    $script:isRunning = $false
    $executeButton.Enabled = $true
    $uninstallButton.Enabled = $true
    $checkButton.Enabled = $true
    $testButton.Enabled = $true
    $stopButton.Enabled = $false
    $progressBar.Visible = $false
    $statusLabel.Text = "Completed"
})

$stopButton.Add_Click({
    if ($script:isRunning) {
        $script:isRunning = $false
        Write-OutputBox "`r`n⚠ Execution stopped by user" "Yellow"
        $statusLabel.Text = "Stopped by user"
    }
})

$logButton.Add_Click({
    if ($script:logFilePath -and (Test-Path $script:logFilePath)) {
        Start-Process notepad.exe $script:logFilePath
    }
    else {
        Write-OutputBox "No log file available. Execute an operation first." "Yellow"
    }
})

$exportLogButton.Add_Click({
    if ($script:results.Count -eq 0) {
        Write-OutputBox "No results to export. Execute an operation first." "Yellow"
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.DefaultExt = "csv"
    $saveFileDialog.FileName = "CriblDeployment_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $script:results | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
        Write-OutputBox "Report exported to: $($saveFileDialog.FileName)" "Cyan"
    }
})

$form.Add_FormClosing({
    if ($script:isRunning) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Execution is in progress. Are you sure you want to exit?",
            "Confirm Exit",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq 'No') {
            $_.Cancel = $true
        }
    }
})

# Create temp directory
$tempDir = Join-Path $env:TEMP "CriblDeploy"
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

# Show the form
Write-OutputBox "Cribl Edge Deployment Tool v3.0 (PsExec Edition)" "Cyan"
Write-OutputBox "========================================" "Cyan"
Write-OutputBox "This tool uses PsExec for remote deployment" "Cyan"
Write-OutputBox "1. Download PsExec from Microsoft Sysinternals" "Gray"
Write-OutputBox "2. Set PsExec path above" "Gray"
Write-OutputBox "3. Enter server names (one per line)" "Gray"
Write-OutputBox "4. Configure credentials if needed" "Gray"
Write-OutputBox "5. Test connections first!" "Gray"
Write-OutputBox "6. Deploy Cribl" "Gray"
Write-OutputBox "========================================`r`n" "Cyan"
$form.ShowDialog()
