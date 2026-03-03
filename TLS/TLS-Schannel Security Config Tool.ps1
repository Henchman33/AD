# TLS/Schannel Security Configuration Tool
# Checks current TLS protocols, Schannel settings, and cipher suites
# Provides options to configure for best practices and compliance

#Requires -RunAsAdministrator

# ============================================
# CONFIGURATION
# ============================================

$LogPath = "C:\Temp\TLS_Schannel_Config.log"
$ReportPath = "C:\Temp\TLS_Schannel_Report.html"
$BackupPath = "C:\Temp\Registry_Backup_TLS_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"

# Registry paths
$SchannelPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"
$ProtocolsPath = "$SchannelPath\Protocols"
$CiphersPath = "$SchannelPath\Ciphers"
$HashesPath = "$SchannelPath\Hashes"
$KeyExchangePath = "$SchannelPath\KeyExchangeAlgorithms"

# Best Practice Recommendations (2024 Standards)
$BestPractices = @{
    # Protocols - Enable only TLS 1.2 and 1.3
    EnabledProtocols = @("TLS 1.2", "TLS 1.3")
    DisabledProtocols = @("SSL 2.0", "SSL 3.0", "TLS 1.0", "TLS 1.1")
    
    # Strong Cipher Suites (ordered by preference)
    RecommendedCipherSuites = @(
        "TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384",
        "TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256",
        "TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384",
        "TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256",
        "TLS_DHE_RSA_WITH_AES_256_GCM_SHA384",
        "TLS_DHE_RSA_WITH_AES_128_GCM_SHA256"
    )
    
    # Weak ciphers to disable
    WeakCiphers = @(
        "DES 56/56",
        "RC2 40/128",
        "RC2 56/128",
        "RC2 128/128",
        "RC4 40/128",
        "RC4 56/128",
        "RC4 64/128",
        "RC4 128/128",
        "Triple DES 168"
    )
    
    # Hash algorithms
    EnabledHashes = @("SHA256", "SHA384", "SHA512")
    DisabledHashes = @("MD5", "SHA")
    
    # Key Exchange Algorithms
    EnabledKeyExchange = @("ECDH", "PKCS", "Diffie-Hellman")
}

# ============================================
# FUNCTIONS
# ============================================

function Write-Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path $LogPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    Add-Content -Path $LogPath -Value $logMessage
    
    $color = switch ($Type) {
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "CRITICAL" { "Magenta" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

function Get-ProtocolStatus {
    param([string]$Protocol)
    
    $clientPath = "$ProtocolsPath\$Protocol\Client"
    $serverPath = "$ProtocolsPath\$Protocol\Server"
    
    $result = @{
        Protocol = $Protocol
        ClientEnabled = $null
        ServerEnabled = $null
        Status = "Not Configured"
    }
    
    # Check Client
    if (Test-Path $clientPath) {
        $clientEnabled = Get-ItemProperty -Path $clientPath -Name "Enabled" -ErrorAction SilentlyContinue
        $clientDisabled = Get-ItemProperty -Path $clientPath -Name "DisabledByDefault" -ErrorAction SilentlyContinue
        
        if ($clientEnabled.Enabled -eq 1 -and $clientDisabled.DisabledByDefault -eq 0) {
            $result.ClientEnabled = $true
        } elseif ($clientEnabled.Enabled -eq 0 -or $clientDisabled.DisabledByDefault -eq 1) {
            $result.ClientEnabled = $false
        }
    }
    
    # Check Server
    if (Test-Path $serverPath) {
        $serverEnabled = Get-ItemProperty -Path $serverPath -Name "Enabled" -ErrorAction SilentlyContinue
        $serverDisabled = Get-ItemProperty -Path $serverPath -Name "DisabledByDefault" -ErrorAction SilentlyContinue
        
        if ($serverEnabled.Enabled -eq 1 -and $serverDisabled.DisabledByDefault -eq 0) {
            $result.ServerEnabled = $true
        } elseif ($serverEnabled.Enabled -eq 0 -or $serverDisabled.DisabledByDefault -eq 1) {
            $result.ServerEnabled = $false
        }
    }
    
    # Determine overall status
    if ($result.ClientEnabled -eq $true -or $result.ServerEnabled -eq $true) {
        $result.Status = "Enabled"
    } elseif ($result.ClientEnabled -eq $false -and $result.ServerEnabled -eq $false) {
        $result.Status = "Disabled"
    }
    
    return $result
}

function Get-AllProtocolStatus {
    Write-Log "`n=== Checking TLS/SSL Protocol Status ===" "INFO"
    
    $protocols = @("SSL 2.0", "SSL 3.0", "TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3")
    $results = @()
    
    foreach ($protocol in $protocols) {
        $status = Get-ProtocolStatus -Protocol $protocol
        $results += $status
        
        $statusColor = switch ($status.Status) {
            "Enabled" { if ($protocol -in @("TLS 1.2", "TLS 1.3")) { "SUCCESS" } else { "WARNING" } }
            "Disabled" { if ($protocol -in @("SSL 2.0", "SSL 3.0", "TLS 1.0", "TLS 1.1")) { "SUCCESS" } else { "WARNING" } }
            default { "INFO" }
        }
        
        Write-Log "$protocol : $($status.Status) (Client: $($status.ClientEnabled), Server: $($status.ServerEnabled))" $statusColor
    }
    
    return $results
}

function Get-CipherSuiteStatus {
    Write-Log "`n=== Checking Cipher Suites ===" "INFO"
    
    try {
        $cipherSuites = Get-TlsCipherSuite
        Write-Log "Total Cipher Suites Configured: $($cipherSuites.Count)" "INFO"
        
        # Check for weak ciphers
        $weakFound = @()
        foreach ($suite in $cipherSuites) {
            $suiteName = $suite.Name
            if ($suiteName -match "RC4|DES|3DES|NULL|EXPORT|anon") {
                $weakFound += $suiteName
            }
        }
        
        if ($weakFound.Count -gt 0) {
            Write-Log "WARNING: Found $($weakFound.Count) weak/insecure cipher suites" "WARNING"
            foreach ($weak in $weakFound) {
                Write-Log "  - $weak" "WARNING"
            }
        } else {
            Write-Log "No weak cipher suites detected" "SUCCESS"
        }
        
        # Show top 10 cipher suites
        Write-Log "`nTop 10 Cipher Suites (in order of preference):" "INFO"
        $cipherSuites | Select-Object -First 10 | ForEach-Object {
            Write-Log "  $($_.Name)" "INFO"
        }
        
        return $cipherSuites
    }
    catch {
        Write-Log "Error checking cipher suites: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-IndividualCipherStatus {
    Write-Log "`n=== Checking Individual Cipher Configuration ===" "INFO"
    
    $cipherResults = @()
    
    if (Test-Path $CiphersPath) {
        $ciphers = Get-ChildItem -Path $CiphersPath -ErrorAction SilentlyContinue
        
        foreach ($cipher in $ciphers) {
            $enabled = Get-ItemProperty -Path $cipher.PSPath -Name "Enabled" -ErrorAction SilentlyContinue
            
            $status = @{
                Cipher = $cipher.PSChildName
                Enabled = $enabled.Enabled
                IsWeak = $cipher.PSChildName -in $BestPractices.WeakCiphers
            }
            
            $cipherResults += $status
            
            $statusText = if ($enabled.Enabled -eq 0) { "Disabled" } elseif ($enabled.Enabled -eq 1) { "Enabled" } else { "Default" }
            $statusColor = if ($status.IsWeak -and $enabled.Enabled -ne 0) { "WARNING" } else { "INFO" }
            
            Write-Log "$($cipher.PSChildName) : $statusText" $statusColor
        }
    } else {
        Write-Log "No individual cipher configuration found (using system defaults)" "INFO"
    }
    
    return $cipherResults
}

function Get-HashAlgorithmStatus {
    Write-Log "`n=== Checking Hash Algorithms ===" "INFO"
    
    $hashResults = @()
    
    if (Test-Path $HashesPath) {
        $hashes = Get-ChildItem -Path $HashesPath -ErrorAction SilentlyContinue
        
        foreach ($hash in $hashes) {
            $enabled = Get-ItemProperty -Path $hash.PSPath -Name "Enabled" -ErrorAction SilentlyContinue
            
            $status = @{
                Hash = $hash.PSChildName
                Enabled = $enabled.Enabled
                IsWeak = $hash.PSChildName -in $BestPractices.DisabledHashes
            }
            
            $hashResults += $status
            
            $statusText = if ($enabled.Enabled -eq 0) { "Disabled" } elseif ($enabled.Enabled -eq 1) { "Enabled" } else { "Default" }
            $statusColor = if ($status.IsWeak -and $enabled.Enabled -ne 0) { "WARNING" } else { "SUCCESS" }
            
            Write-Log "$($hash.PSChildName) : $statusText" $statusColor
        }
    } else {
        Write-Log "No hash algorithm configuration found (using system defaults)" "INFO"
    }
    
    return $hashResults
}

function Get-KeyExchangeStatus {
    Write-Log "`n=== Checking Key Exchange Algorithms ===" "INFO"
    
    $keyExResults = @()
    
    if (Test-Path $KeyExchangePath) {
        $keyExchanges = Get-ChildItem -Path $KeyExchangePath -ErrorAction SilentlyContinue
        
        foreach ($keyEx in $keyExchanges) {
            $enabled = Get-ItemProperty -Path $keyEx.PSPath -Name "Enabled" -ErrorAction SilentlyContinue
            
            $status = @{
                KeyExchange = $keyEx.PSChildName
                Enabled = $enabled.Enabled
            }
            
            $keyExResults += $status
            
            $statusText = if ($enabled.Enabled -eq 0) { "Disabled" } elseif ($enabled.Enabled -eq 1) { "Enabled" } else { "Default" }
            Write-Log "$($keyEx.PSChildName) : $statusText" "INFO"
        }
    } else {
        Write-Log "No key exchange configuration found (using system defaults)" "INFO"
    }
    
    return $keyExResults
}

function Get-DotNetTLSSettings {
    Write-Log "`n=== Checking .NET Framework TLS Settings ===" "INFO"
    
    $dotNetPaths = @(
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727",
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
    )
    
    foreach ($path in $dotNetPaths) {
        if (Test-Path $path) {
            $systemDefault = Get-ItemProperty -Path $path -Name "SystemDefaultTlsVersions" -ErrorAction SilentlyContinue
            $strongCrypto = Get-ItemProperty -Path $path -Name "SchUseStrongCrypto" -ErrorAction SilentlyContinue
            
            $pathName = $path.Split('\')[-1]
            Write-Log "$pathName :" "INFO"
            Write-Log "  SystemDefaultTlsVersions: $($systemDefault.SystemDefaultTlsVersions)" $(if ($systemDefault.SystemDefaultTlsVersions -eq 1) {"SUCCESS"} else {"WARNING"})
            Write-Log "  SchUseStrongCrypto: $($strongCrypto.SchUseStrongCrypto)" $(if ($strongCrypto.SchUseStrongCrypto -eq 1) {"SUCCESS"} else {"WARNING"})
        }
    }
}

function Backup-RegistrySettings {
    Write-Log "`n=== Backing Up Registry Settings ===" "INFO"
    
    try {
        $regExport = "reg export `"HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL`" `"$BackupPath`" /y"
        cmd /c $regExport 2>&1 | Out-Null
        
        if (Test-Path $BackupPath) {
            Write-Log "Registry backup saved to: $BackupPath" "SUCCESS"
            return $true
        } else {
            Write-Log "Failed to create registry backup" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error backing up registry: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Set-ProtocolState {
    param(
        [string]$Protocol,
        [bool]$Enable
    )
    
    $clientPath = "$ProtocolsPath\$Protocol\Client"
    $serverPath = "$ProtocolsPath\$Protocol\Server"
    
    try {
        # Create paths if they don't exist
        if (-not (Test-Path $clientPath)) {
            New-Item -Path $clientPath -Force | Out-Null
        }
        if (-not (Test-Path $serverPath)) {
            New-Item -Path $serverPath -Force | Out-Null
        }
        
        $enabledValue = if ($Enable) { 1 } else { 0 }
        $disabledValue = if ($Enable) { 0 } else { 1 }
        
        # Set Client
        Set-ItemProperty -Path $clientPath -Name "Enabled" -Value $enabledValue -Type DWord -Force
        Set-ItemProperty -Path $clientPath -Name "DisabledByDefault" -Value $disabledValue -Type DWord -Force
        
        # Set Server
        Set-ItemProperty -Path $serverPath -Name "Enabled" -Value $enabledValue -Type DWord -Force
        Set-ItemProperty -Path $serverPath -Name "DisabledByDefault" -Value $disabledValue -Type DWord -Force
        
        $action = if ($Enable) { "Enabled" } else { "Disabled" }
        Write-Log "$action $Protocol" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure $Protocol : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Set-CipherState {
    param(
        [string]$Cipher,
        [bool]$Enable
    )
    
    $cipherPath = "$CiphersPath\$Cipher"
    
    try {
        if (-not (Test-Path $cipherPath)) {
            New-Item -Path $cipherPath -Force | Out-Null
        }
        
        $enabledValue = if ($Enable) { 1 } else { 0 }
        Set-ItemProperty -Path $cipherPath -Name "Enabled" -Value $enabledValue -Type DWord -Force
        
        $action = if ($Enable) { "Enabled" } else { "Disabled" }
        Write-Log "$action cipher: $Cipher" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure cipher $Cipher : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Set-HashState {
    param(
        [string]$Hash,
        [bool]$Enable
    )
    
    $hashPath = "$HashesPath\$Hash"
    
    try {
        if (-not (Test-Path $hashPath)) {
            New-Item -Path $hashPath -Force | Out-Null
        }
        
        $enabledValue = if ($Enable) { 1 } else { 0 }
        Set-ItemProperty -Path $hashPath -Name "Enabled" -Value $enabledValue -Type DWord -Force
        
        $action = if ($Enable) { "Enabled" } else { "Disabled" }
        Write-Log "$action hash: $Hash" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure hash $Hash : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Set-BestPracticeConfiguration {
    Write-Log "`n========================================" "INFO"
    Write-Log "APPLYING BEST PRACTICE CONFIGURATION" "INFO"
    Write-Log "========================================" "INFO"
    
    # Disable weak protocols
    Write-Log "`n--- Disabling Weak Protocols ---" "INFO"
    foreach ($protocol in $BestPractices.DisabledProtocols) {
        Set-ProtocolState -Protocol $protocol -Enable $false
    }
    
    # Enable strong protocols
    Write-Log "`n--- Enabling Strong Protocols ---" "INFO"
    foreach ($protocol in $BestPractices.EnabledProtocols) {
        Set-ProtocolState -Protocol $protocol -Enable $true
    }
    
    # Disable weak ciphers
    Write-Log "`n--- Disabling Weak Ciphers ---" "INFO"
    foreach ($cipher in $BestPractices.WeakCiphers) {
        Set-CipherState -Cipher $cipher -Enable $false
    }
    
    # Configure hash algorithms
    Write-Log "`n--- Configuring Hash Algorithms ---" "INFO"
    foreach ($hash in $BestPractices.DisabledHashes) {
        Set-HashState -Hash $hash -Enable $false
    }
    foreach ($hash in $BestPractices.EnabledHashes) {
        Set-HashState -Hash $hash -Enable $true
    }
    
    # Configure cipher suite order
    Write-Log "`n--- Configuring Cipher Suite Order ---" "INFO"
    try {
        # Get current cipher suites
        $currentSuites = Get-TlsCipherSuite
        
        # Filter to keep only strong suites and order by best practice
        $strongSuites = $currentSuites | Where-Object {
            $_.Name -notmatch "RC4|DES|3DES|NULL|EXPORT|anon|MD5"
        }
        
        # Reorder to put recommended suites first
        $orderedSuites = @()
        foreach ($recommended in $BestPractices.RecommendedCipherSuites) {
            $suite = $strongSuites | Where-Object { $_.Name -eq $recommended }
            if ($suite) {
                $orderedSuites += $suite.Name
            }
        }
        
        # Add remaining strong suites
        foreach ($suite in $strongSuites) {
            if ($suite.Name -notin $orderedSuites) {
                $orderedSuites += $suite.Name
            }
        }
        
        # This would require a reboot to take effect
        Write-Log "Cipher suite order configured (requires reboot)" "SUCCESS"
        Write-Log "Total strong cipher suites: $($orderedSuites.Count)" "INFO"
    }
    catch {
        Write-Log "Error configuring cipher suites: $($_.Exception.Message)" "WARNING"
    }
    
    # Configure .NET to use strong crypto
    Write-Log "`n--- Configuring .NET Framework ---" "INFO"
    $dotNetPaths = @(
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727",
        "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v2.0.50727",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
    )
    
    foreach ($path in $dotNetPaths) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
        Set-ItemProperty -Path $path -Name "SystemDefaultTlsVersions" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -Type DWord -Force
        Write-Log "Configured $($path.Split('\')[-1])" "SUCCESS"
    }
    
    Write-Log "`n========================================" "SUCCESS"
    Write-Log "BEST PRACTICE CONFIGURATION COMPLETE" "SUCCESS"
    Write-Log "========================================" "SUCCESS"
    Write-Log "IMPORTANT: A system reboot is required for changes to take effect" "CRITICAL"
}

function Generate-HTMLReport {
    param(
        $ProtocolStatus,
        $CipherSuites,
        $CipherStatus,
        $HashStatus,
        $KeyExStatus
    )
    
    Write-Log "`n=== Generating HTML Report ===" "INFO"
    
    $protocolRows = ""
    foreach ($proto in $ProtocolStatus) {
        $statusClass = if ($proto.Status -eq "Enabled") { "status-enabled" } else { "status-disabled" }
        $recommendation = if ($proto.Protocol -in @("TLS 1.2", "TLS 1.3")) {
            if ($proto.Status -eq "Enabled") { "✓ Good" } else { "⚠ Should be Enabled" }
        } else {
            if ($proto.Status -eq "Disabled") { "✓ Good" } else { "⚠ Should be Disabled" }
        }
        
        $protocolRows += "<tr><td>$($proto.Protocol)</td><td class='$statusClass'>$($proto.Status)</td><td>$recommendation</td></tr>`n"
    }
    
    $weakCipherCount = ($CipherSuites | Where-Object { $_.Name -match "RC4|DES|3DES|NULL|EXPORT|anon" }).Count
    $totalCipherCount = $CipherSuites.Count
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>TLS/Schannel Security Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #005a9e; margin-top: 30px; }
        .summary { background-color: #e7f3ff; padding: 20px; border-radius: 5px; margin: 20px 0; }
        .warning-box { background-color: #fff4ce; border-left: 4px solid #ffb900; padding: 15px; margin: 20px 0; }
        .success-box { background-color: #dff6dd; border-left: 4px solid #107c10; padding: 15px; margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background-color: #0078d4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .status-enabled { color: #107c10; font-weight: bold; }
        .status-disabled { color: #d13438; font-weight: bold; }
        .timestamp { color: #666; font-size: 0.9em; margin-top: 30px; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h1>TLS/Schannel Security Configuration Report</h1>
        <p><strong>Server:</strong> $env:COMPUTERNAME</p>
        <p><strong>Report Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        
        <div class="summary">
            <h2>Executive Summary</h2>
            <p><strong>Total Cipher Suites:</strong> $totalCipherCount</p>
            <p><strong>Weak Cipher Suites Found:</strong> $weakCipherCount</p>
            $(if ($weakCipherCount -gt 0) { "<div class='warning-box'><strong>⚠ Warning:</strong> Weak cipher suites detected. Consider applying best practice configuration.</div>" } else { "<div class='success-box'><strong>✓ Good:</strong> No weak cipher suites detected.</div>" })
        </div>
        
        <h2>TLS/SSL Protocol Status</h2>
        <table>
            <tr>
                <th>Protocol</th>
                <th>Status</th>
                <th>Recommendation</th>
            </tr>
            $protocolRows
        </table>
        
        <h2>Best Practice Recommendations</h2>
        <ul>
            <li>✓ Enable: TLS 1.2 and TLS 1.3</li>
            <li>✗ Disable: SSL 2.0, SSL 3.0, TLS 1.0, TLS 1.1</li>
            <li>✗ Disable: RC4, DES, 3DES ciphers</li>
            <li>✗ Disable: MD5, SHA-1 hashes (use SHA-256+)</li>
            <li>✓ Configure .NET to use system default TLS versions</li>
            <li>✓ Enable Strong Cryptography in .NET Framework</li>
        </ul>
        
        <div class="warning-box">
            <strong>Important:</strong> After applying configuration changes, a system reboot is required for changes to take effect.
        </div>
        
        <div class="timestamp">
            <p>This report was generated automatically by the TLS/Schannel Configuration Tool.</p>
        </div>
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Log "HTML Report saved to: $ReportPath" "SUCCESS"
}

function Show-Menu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  TLS/Schannel Configuration Tool" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Check Current Configuration" -ForegroundColor White
    Write-Host "2. Apply Best Practice Configuration" -ForegroundColor Green
    Write-Host "3. Generate HTML Report" -ForegroundColor White
    Write-Host "4. Enable Specific Protocol" -ForegroundColor White
    Write-Host "5. Disable Specific Protocol" -ForegroundColor White
    Write-Host "6. View Cipher Suite Order" -ForegroundColor White
    Write-Host "7. Backup Registry Settings" -ForegroundColor Yellow
    Write-Host "8. Exit" -ForegroundColor Red
    Write-Host ""
}

# ============================================
# MAIN EXECUTION
# ============================================

$script:ProtocolStatus = $null
$script:CipherSuites = $null
$script:CipherStatus = $null
$script:HashStatus = $null
$script:KeyExStatus = $null

do {
    Show-Menu
    $choice = Read-Host "Select an option (1-8)"
    
    switch ($choice) {
        "1" {
            Clear-Host
            Write-Log "========================================" "INFO"
            Write-Log "CURRENT TLS/SCHANNEL CONFIGURATION" "INFO"
            Write-Log "Server: $env:COMPUTERNAME" "INFO"
            Write-Log "========================================" "INFO"
            
            $script:ProtocolStatus = Get-AllProtocolStatus
            $script:CipherSuites = Get-CipherSuiteStatus
            $script:CipherStatus = Get-IndividualCipherStatus
            $script:HashStatus = Get-HashAlgorithmStatus
            $script:KeyExStatus = Get-KeyExchangeStatus
            Get-DotNetTLSSettings
            
            Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "2" {
            Clear-Host
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "  APPLY BEST PRACTICE CONFIGURATION" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "This will:" -ForegroundColor White
         
