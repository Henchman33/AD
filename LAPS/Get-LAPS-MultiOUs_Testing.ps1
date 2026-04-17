<#
.SYNOPSIS
    LAPS compliance and health report for server objects across multiple OUs

    [string[]]$SearchBases = @(
    "OU=Servers,DC=domain,DC=com",
    "OU=DMZ,DC=domain,DC=com",
    "OU=SQL,DC=domain,DC=com",
    "OU=Legacy,DC=domain,DC=com"
)
#>

param(
    [string[]]$SearchBases = @(
        "OU=Servers,DC=domain,DC=com",
        "OU=DMZ,DC=domain,DC=com"
        "OU=DMZ,DC=domain,DC=com",
        "OU=SQL,DC=domain,DC=com",
        "OU=Legacy,DC=domain,DC=com"
    ),

    [string]$ExportPath = "C:\Secure\LAPSReports\LAPS_Audit.csv",
    [int]$StaleDays = 30
)

# Ensure export directory exists
$ExportDir = Split-Path $ExportPath
if (!(Test-Path $ExportDir)) {
    New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null
}

Import-Module ActiveDirectory

$Now = Get-Date
$StaleThreshold = $Now.AddDays(-$StaleDays)

Write-Host "Running LAPS audit across multiple OUs..." -ForegroundColor Cyan

$AllServers = @()

foreach ($OU in $SearchBases) {

    Write-Host "Querying OU: $OU"

    try {
        $Servers = Get-ADComputer `
            -SearchBase $OU `
            -Filter {
                Enabled -eq $true -and OperatingSystem -like "*Server*"
            } `
            -Properties Name,
                         DistinguishedName,
                         OperatingSystem,
                         ms-Mcs-AdmPwd,
                         ms-Mcs-AdmPwdExpirationTime,
                         PasswordLastSet,
                         LastLogonDate

        foreach ($Server in $Servers) {
            # Add OU context
            $Server | Add-Member -NotePropertyName SourceOU -NotePropertyValue $OU
        }

        $AllServers += $Servers
    }
    catch {
        Write-Warning "Failed to query OU: $OU"
    }
}

$Results = foreach ($Server in $AllServers) {

    $LapsPassword = $Server.'ms-Mcs-AdmPwd'
    $ExpirationRaw = $Server.'ms-Mcs-AdmPwdExpirationTime'

    # Convert expiration
    $Expiration = $null
    if ($ExpirationRaw) {
        $Expiration = [DateTime]::FromFileTime($ExpirationRaw)
    }

    # Status checks
    $MissingLAPS = -not $LapsPassword

    $Expired = $false
    if ($Expiration -and $Expiration -lt $Now) {
        $Expired = $true
    }

    $Stale = $false
    if ($Server.PasswordLastSet -and $Server.PasswordLastSet -lt $StaleThreshold) {
        $Stale = $true
    }

    # Build status string
    $Status = @()
    if ($MissingLAPS) { $Status += "MissingLAPS" }
    if ($Expired) { $Status += "ExpiredPassword" }
    if ($Stale) { $Status += "StalePassword" }

    if ($Status.Count -eq 0) {
        $Status = "Healthy"
    } else {
        $Status = $Status -join "; "
    }

    [PSCustomObject]@{
        ServerName        = $Server.Name
        SourceOU          = $Server.SourceOU
        OperatingSystem   = $Server.OperatingSystem
        LAPSConfigured    = -not $MissingLAPS
        PasswordExpiration= $Expiration
        PasswordLastSet   = $Server.PasswordLastSet
        LastLogonDate     = $Server.LastLogonDate
        Status            = $Status
    }
}

# Export report
$Results | Sort-Object Status, ServerName | Export-Csv `
    -Path $ExportPath `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host "Report exported to: $ExportPath" -ForegroundColor Green
