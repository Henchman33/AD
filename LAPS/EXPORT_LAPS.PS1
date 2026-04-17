<#
.SYNOPSIS
    Secure LAPS export with logging, cleanup, parallel processing, and email alerting

.NOTES
    - Legacy LAPS (ms-Mcs-AdmPwd)
    - Output path: C:\ADLAPS
    - Email is summary ONLY (no passwords)
#>

param(
    [Parameter(Mandatory=$true)]
    [string[]]$SearchBases,

    [int]$RetentionDays = 7,
    [int]$ThrottleLimit = 3,

    # Email Settings
    [string]$SmtpServer = "smtp.yourdomain.com",
    [string]$From = "laps-report@yourdomain.com",
    [string]$To = "admins@yourdomain.com",
    [switch]$UseSsl,
    [pscredential]$SmtpCredential
)

# Paths
$BasePath = "C:\ADLAPS"
$LogPath  = "$BasePath\Logs"

Import-Module ActiveDirectory -ErrorAction Stop

# Create directories
foreach ($Path in @($BasePath, $LogPath)) {
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

# Secure folder
try {
    icacls $BasePath /inheritance:r | Out-Null
    icacls $BasePath /grant "Domain Admins:(OI)(CI)F" | Out-Null
} catch {
    Write-Warning "Failed to apply NTFS permissions"
}

# Timestamp
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile = "$BasePath\LAPS_ServerPasswords_$Timestamp.csv"
$LogFile = "$LogPath\LAPS_Export_$Timestamp.log"

# Logging
function Write-Log {
    param ($Message, $Level = "INFO")
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Time [$Level] $Message"
    Add-Content -Path $LogFile -Value $Entry
    Write-Host $Entry
}

Write-Log "===== LAPS Export Started ====="

# Cleanup
Write-Log "Cleanup started (>$RetentionDays days)"
Get-ChildItem $BasePath -Filter *.csv -ErrorAction SilentlyContinue |
    Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays)} |
    ForEach-Object {
        Remove-Item $_.FullName -Force
        Write-Log "Deleted old CSV: $($_.Name)"
    }

Get-ChildItem $LogPath -Filter *.log -ErrorAction SilentlyContinue |
    Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays)} |
    ForEach-Object {
        Remove-Item $_.FullName -Force
        Write-Log "Deleted old LOG: $($_.Name)"
    }

# Parallel job block
$ScriptBlock = {
    param($Base)
    Import-Module ActiveDirectory

    $Output = @()

    try {
        $Computers = Get-ADComputer `
            -SearchBase $Base `
            -Filter {OperatingSystem -like "*Server*"} `
            -Properties Name,OperatingSystem,ms-Mcs-AdmPwd,ms-Mcs-AdmPwdExpirationTime

        foreach ($Computer in $Computers) {

            $Expiration = if ($Computer."ms-Mcs-AdmPwdExpirationTime") {
                [DateTime]::FromFileTime($Computer."ms-Mcs-AdmPwdExpirationTime")
            } else {
                $null
            }

            $Output += [PSCustomObject]@{
                ServerName     = $Computer.Name
                OS             = $Computer.OperatingSystem
                Password       = $Computer."ms-Mcs-AdmPwd"
                ExpirationTime = $Expiration
                SourceOU       = $Base
            }
        }

        return $Output

    } catch {
        return "ERROR::$Base::$($_.Exception.Message)"
    }
}

# Start jobs
$Jobs = @()

foreach ($Base in $SearchBases) {
    Write-Log "Queueing OU: $Base"
    $Jobs += Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Base

    while (@($Jobs | Where-Object {$_.State -eq "Running"}).Count -ge $ThrottleLimit) {
        Start-Sleep -Seconds 2
    }
}

# Collect results
$Results = @()
$Errors = 0

foreach ($Job in $Jobs) {
    $Data = Receive-Job -Job $Job -Wait

    if ($Data -is [string] -and $Data -like "ERROR::*") {
        $Errors++
        Write-Log "Job failure: $Data" "ERROR"
    } else {
        $Results += $Data
        Write-Log "Job completed with $($Data.Count) records"
    }

    Remove-Job $Job
}

# Export CSV
try {
    $Results | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
    Write-Log "CSV exported: $CsvFile"
} catch {
    Write-Log "CSV export failed: $_" "ERROR"
    exit
}

# Summary
$Total = $Results.Count
Write-Log "Total servers: $Total"
Write-Log "Errors: $Errors"
Write-Log "===== LAPS Export Completed ====="

# =========================
# EMAIL ALERTING (SECURE)
# =========================

$Status = if ($Errors -gt 0) { "Completed with Errors" } else { "Success" }

$Subject = "LAPS Export - $Status - $Timestamp"

$Body = @"
LAPS Export Summary

Status: $Status
Timestamp: $Timestamp

Total Servers: $Total
Errors: $Errors

CSV Location:
$CsvFile

Log Location:
$LogFile
"@

try {
    $MailParams = @{
        SmtpServer = $SmtpServer
        From       = $From
        To         = $To
        Subject    = $Subject
        Body       = $Body
    }

    if ($UseSsl) {
        $MailParams.UseSsl = $true
    }

    if ($SmtpCredential) {
        $MailParams.Credential = $SmtpCredential
    }

    Send-MailMessage @MailParams

    Write-Log "Email notification sent"
} catch {
    Write-Log "Email failed: $_" "ERROR"
}
