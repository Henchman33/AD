<# Below is a single, all-in-one PowerShell ISE–ready hardening script that aligns with IIS Crypto “Best Practices” (modern / secure baseline) and is safe for:
•	Windows Server 2016
•	Windows Server 2019
•	Windows Server 2022
•	Windows Server 2025
It does all of the following in one run:
✅ Enables TLS 1.2 (all supported OS)
✅ Enables TLS 1.3 (effective on 2022 / 2025)
❌ Disables SSL 2.0, SSL 3.0, TLS 1.0, TLS 1.1
❌ Disables weak ciphers (RC4, DES, 3DES, NULL, EXPORT)
❌ Disables weak hashes (MD5, SHA-1)
❌ Disables weak key exchange (RSA < 2048, Diffie-Hellman < 2048)
✅ Enables strong cipher suites only
✅ Configures cipher suite order
✅ Forces .NET to use system default TLS / strong crypto
🔁 Reboot required
________________________________________
🔐 ONE COMPLETE POWERSHELL ISE SCRIPT
Run as Administrator
# =====================================================================================
# Windows Server TLS / SCHANNEL Hardening Script
# IIS Crypto "Best Practices" equivalent (Modern Secure Baseline)
# Supports: Server 2016, 2019, 2022, 2025
# Run in PowerShell ISE as Administrator
# REBOOT REQUIRED
# =====================================================================================
#>
Write-Host "Starting TLS / SCHANNEL hardening..." -ForegroundColor Cyan

# ------------------------------------------------
# 1. Protocol Configuration
# ------------------------------------------------
$protocols = @{
    "SSL 2.0" = 0
    "SSL 3.0" = 0
    "TLS 1.0" = 0
    "TLS 1.1" = 0
    "TLS 1.2" = 1
    "TLS 1.3" = 1
}

foreach ($protocol in $protocols.Keys) {
    foreach ($role in @("Client","Server")) {
        $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$protocol\$role"
        New-Item -Path $path -Force | Out-Null

        if ($protocols[$protocol] -eq 1) {
            New-ItemProperty -Path $path -Name "Enabled" -Value 1 -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -Path $path -Name "DisabledByDefault" -Value 0 -PropertyType DWORD -Force | Out-Null
        } else {
            New-ItemProperty -Path $path -Name "Enabled" -Value 0 -PropertyType DWORD -Force | Out-Null
            New-ItemProperty -Path $path -Name "DisabledByDefault" -Value 1 -PropertyType DWORD -Force | Out-Null
        }
    }
}

# ------------------------------------------------
# 2. Cipher Configuration
# ------------------------------------------------
$disableCiphers = @(
    "NULL",
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

foreach ($cipher in $disableCiphers) {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\$cipher"
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Enabled" -Value 0 -PropertyType DWORD -Force | Out-Null
}

# Enable strong AES ciphers
$enableCiphers = @(
    "AES 128/128",
    "AES 256/256"
)

foreach ($cipher in $enableCiphers) {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\$cipher"
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Enabled" -Value 1 -PropertyType DWORD -Force | Out-Null
}

# ------------------------------------------------
# 3. Hash Configuration
# ------------------------------------------------
$disableHashes = @("MD5","SHA")
foreach ($hash in $disableHashes) {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes\$hash"
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Enabled" -Value 0 -PropertyType DWORD -Force | Out-Null
}

# Enable SHA-256 / SHA-384 / SHA-512
$enableHashes = @("SHA256","SHA384","SHA512")
foreach ($hash in $enableHashes) {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Hashes\$hash"
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Enabled" -Value 1 -PropertyType DWORD -Force | Out-Null
}

# ------------------------------------------------
# 4. Key Exchange Algorithms
# ------------------------------------------------
$keyExchanges = @{
    "Diffie-Hellman" = @{ Enabled = 1; MinKeyBitLength = 2048 }
    "ECDH"           = @{ Enabled = 1 }
    "PKCS"           = @{ Enabled = 1; MinKeyBitLength = 2048 }
}

foreach ($kx in $keyExchanges.Keys) {
    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\KeyExchangeAlgorithms\$kx"
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Enabled" -Value $keyExchanges[$kx].Enabled -PropertyType DWORD -Force | Out-Null

    if ($keyExchanges[$kx].ContainsKey("MinKeyBitLength")) {
        New-ItemProperty -Path $path -Name "MinKeyBitLength" -Value $keyExchanges[$kx].MinKeyBitLength -PropertyType DWORD -Force | Out-Null
    }
}

# ------------------------------------------------
# 5. Cipher Suite Order (Modern / Secure)
# ------------------------------------------------
$cipherSuites = @(
    "TLS_AES_256_GCM_SHA384",
    "TLS_AES_128_GCM_SHA256",
    "TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384",
    "TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256",
    "TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384",
    "TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256"
)

$cipherSuitePath = "HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002"
New-Item -Path $cipherSuitePath -Force | Out-Null
New-ItemProperty -Path $cipherSuitePath -Name "Functions" -Value $cipherSuites -PropertyType MultiString -Force | Out-Null

# ------------------------------------------------
# 6. .NET Framework Strong Crypto
# ------------------------------------------------
$dotNetPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
)

foreach ($path in $dotNetPaths) {
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $path -Name "SystemDefaultTlsVersions" -Value 1 -PropertyType DWORD -Force | Out-Null
}

Write-Host "TLS / SCHANNEL hardening complete." -ForegroundColor Green
Write-Host "REBOOT THE SERVER to apply changes." -ForegroundColor Yellow
<#
________________________________________
⚠️ IMPORTANT OPERATIONAL NOTES
🔁 Reboot Required
SCHANNEL settings do not apply until reboot
🧪 Compatibility Check
This configuration:
•	BREAKS legacy clients (TLS 1.0 / 1.1)
•	Requires modern browsers, .NET 4.6+, Java 8+, OpenSSL 1.1+
•	Is safe for IIS, RDP, WinRM, LDAP-S, SQL (modern drivers)
🟢 TLS 1.3 Behavior
•	Server 2016/2019 → ignored (not implemented)
•	Server 2022/2025 → fully active
________________________________________
✅ MATCHES IIS CRYPTO “BEST PRACTICES” (MODERN)
Area	Status
SSL 2.0 / 3.0	❌ Disabled
TLS 1.0 / 1.1	❌ Disabled
TLS 1.2	✅ Enabled
TLS 1.3	✅ Enabled
RC4 / 3DES	❌ Disabled
SHA-1	❌ Disabled
PFS (ECDHE)	✅ Enforced
AES-GCM	✅ Enforced
.NET Strong Crypto	✅ Enabled
#>
