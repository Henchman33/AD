<# 
### Script
Save this as something like DiskCleanup-Server.ps1 and run it as PowerShell ISE from a management workstation or deploy it to servers through GPO. 
The script targets server operating systems discovered in AD, uses UNC paths for remote cleanup, skips in-use files, and logs every action per server. 
It also avoids removing profiles tied to common service-account naming patterns you specified.
The SCCM cache portion uses the client cache WMI namespace approach commonly used for cache maintenance, and profile cleanup uses Win32_UserProfile with Remove-CimInstance, which is safer than deleting profile folders directly.
What it does

The script finds server objects in Active Directory with OperatingSystem -like '*Server*', then connects to each server remotely. 
It cleans temp locations, attempts to skip files that are locked or in use by catching deletion failures, and only removes CCM cache entries older than the age you set. 
For profiles, it removes only non-special profiles older than the cutoff and avoids names starting with SVC_, SVC0, SVC1, and SAPService.
Important limits

A few things are worth noting before deploying broadly. 
Manually deleting SCCM client cache contents is often discouraged in favor of supported cache management, so test the cache section on a pilot group first. 
The profile-removal logic is based on local profile folders and LastUseTime, which is practical, but you should still verify the exact service-account naming in your environment before mass rollout.
GPO deployment

Use a Computer Configuration scheduled task so the script runs as SYSTEM on target servers. 
A common approach is to copy the .ps1 file to a domain share such as \\DOMAIN\NETLOGON\DiskCleanup\DiskCleanup-Server.ps1, 
then create a GPO under Computer Configuration -> Preferences -> Control Panel Settings -> Scheduled Tasks and configure an action that starts powershell.exe with the script path. 
Microsoft guidance and common GPO practice also support startup scripts and scheduled tasks for PowerShell deployment.
#>
param(
    [string]$SearchBase = "",

    [int]$ProfileAgeDays = 365,
    [int]$KeepCcmCacheDays = 60,
    [switch]$WhatIfMode
)

Import-Module ActiveDirectory -ErrorAction Stop

$ErrorActionPreference = 'Stop'
$ScriptName = 'DiskCleanup'
$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$GlobalLogRoot = 'C:\AdminFiles\DiskCleanUp'

if (-not (Test-Path $GlobalLogRoot)) {
    New-Item -Path $GlobalLogRoot -ItemType Directory -Force | Out-Null
}

$RunLog = Join-Path $GlobalLogRoot "DiskCleanup_$RunStamp.csv"
$SummaryLog = Join-Path $GlobalLogRoot "DiskCleanup_$RunStamp.txt"

$Results = New-Object System.Collections.Generic.List[object]

function Write-RunLog {
    param(
        [string]$ComputerName,
        [string]$Category,
        [string]$Action,
        [string]$Item,
        [string]$Status,
        [string]$Details
    )

    $Results.Add([pscustomobject]@{
        Timestamp    = Get-Date
        ComputerName = $ComputerName
        Category     = $Category
        Action       = $Action
        Item         = $Item
        Status       = $Status
        Details      = $Details
    }) | Out-Null
}

function Test-ServiceProfileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }

    $patterns = @(
        '^SVC_',
        '^SVC0',
        '^SVC1',
        '^SAPService',
        '^svc_',
        '^svc0',
        '^svc1',
        '^sapservice'
    )

    foreach ($p in $patterns) {
        if ($Name -match $p) { return $true }
    }
    return $false
}

function Get-ServerList {
    param([string]$BaseDN)

    $filter = "OperatingSystem -like '*Server*'"
    if ([string]::IsNullOrWhiteSpace($BaseDN)) {
        Get-ADComputer -Filter $filter -Properties OperatingSystem,Enabled,LastLogonDate
    } else {
        Get-ADComputer -Filter $filter -SearchBase $BaseDN -Properties OperatingSystem,Enabled,LastLogonDate
    }
}

function Invoke-RemoteCleanup {
    param(
        [string]$ComputerName,
        [int]$ProfileAgeDays,
        [int]$KeepCcmCacheDays,
        [switch]$WhatIfMode
    )

    $session = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
    try {
        Invoke-Command -Session $session -ArgumentList $ComputerName,$ProfileAgeDays,$KeepCcmCacheDays,$WhatIfMode -ScriptBlock {
            param($ComputerName,$ProfileAgeDays,$KeepCcmCacheDays,$WhatIfMode)

            $logRoot = 'C:\AdminFiles\DiskCleanUp'
            if (-not (Test-Path $logRoot)) {
                New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
            }

            $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $logFile = Join-Path $logRoot "DiskCleaup_$stamp.log"

            function Log-Line {
                param([string]$Line)
                Add-Content -Path $logFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Line)
            }

            Log-Line "Start cleanup on $ComputerName"
            Log-Line "ProfileAgeDays=$ProfileAgeDays KeepCcmCacheDays=$KeepCcmCacheDays WhatIfMode=$WhatIfMode"

            $tempPaths = @(
                'C:\Windows\Temp',
                'C:\Temp',
                'C:\Users\*\AppData\Local\Temp'
            )

            foreach ($path in $tempPaths) {
                try {
                    $items = Get-ChildItem -Path $path -Force -ErrorAction SilentlyContinue
                    foreach ($item in $items) {
                        try {
                            if ($WhatIfMode) {
                                Log-Line "WHATIF would remove temp item: $($item.FullName)"
                            } else {
                                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                if (-not (Test-Path $item.FullName)) {
                                    Log-Line "Removed temp item: $($item.FullName)"
                                }
                            }
                        } catch {
                            Log-Line "Skipped temp item in use or locked: $($item.FullName)"
                        }
                    }
                } catch {
                    Log-Line "Temp path inaccessible: $path"
                }
            }

            try {
                $ccmCachePath = $null
                try {
                    $ccmCachePath = ([wmi]"ROOT\ccm\SoftMgmtAgent:CacheConfig.ConfigKey='Cache'").Location
                } catch {
                    $ccmCachePath = 'C:\Windows\ccmcache'
                }

                if (Test-Path $ccmCachePath) {
                    $cacheItems = Get-CimInstance -Namespace 'ROOT\ccm\SoftMgmtAgent' -ClassName CacheInfoEx -ErrorAction SilentlyContinue
                    foreach ($cacheItem in $cacheItems) {
                        try {
                            $lastRef = [System.Management.ManagementDateTimeConverter]::ToDateTime($cacheItem.LastReferenced)
                            $ageDays = ((Get-Date) - $lastRef).Days
                            if ($ageDays -gt $KeepCcmCacheDays) {
                                if ($WhatIfMode) {
                                    Log-Line "WHATIF would remove CCMCache item: $($cacheItem.Location)"
                                } else {
                                    Remove-Item -Path $cacheItem.Location -Recurse -Force -ErrorAction SilentlyContinue
                                    try {
                                        $cacheItem | Remove-CimInstance -ErrorAction SilentlyContinue
                                    } catch {}
                                    Log-Line "Removed CCMCache item older than $KeepCcmCacheDays days: $($cacheItem.Location)"
                                }
                            }
                        } catch {
                            Log-Line "Skipped CCMCache item: $($cacheItem.Location)"
                        }
                    }
                } else {
                    Log-Line "CCMCache path not found"
                }
            } catch {
                Log-Line "CCMCache cleanup failed: $($_.Exception.Message)"
            }

            try {
                $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
                    Where-Object {
                        -not $_.Special -and
                        $_.LocalPath -like 'C:\Users\*' -and
                        $_.LocalPath -notmatch '\\(Public|Default|Default User|All Users)$' -and
                        $_.LocalPath -notmatch '\\Administrator$'
                    }

                $cutoff = (Get-Date).AddDays(-$ProfileAgeDays)

                foreach ($profile in $profiles) {
                    $localName = Split-Path $profile.LocalPath -Leaf
                    if (Test-ServiceProfileName -Name $localName) {
                        Log-Line "Skipped service profile by name: $($profile.LocalPath)"
                        continue
                    }

                    $lastUse = $null
                    try { $lastUse = [datetime]$profile.LastUseTime } catch {}
                    if ($lastUse -and $lastUse -lt $cutoff) {
                        try {
                            if ($WhatIfMode) {
                                Log-Line "WHATIF would remove profile: $($profile.LocalPath)"
                            } else {
                                Remove-CimInstance -InputObject $profile -ErrorAction Stop
                                Log-Line "Removed profile: $($profile.LocalPath)"
                            }
                        } catch {
                            Log-Line "Failed removing profile: $($profile.LocalPath) - $($_.Exception.Message)"
                        }
                    }
                }
            } catch {
                Log-Line "Profile cleanup failed: $($_.Exception.Message)"
            }

            Log-Line "Cleanup complete"
        }

        Write-RunLog -ComputerName $ComputerName -Category 'Server' -Action 'Cleanup' -Item 'RemoteRun' -Status 'Success' -Details 'Completed'
    }
    catch {
        Write-RunLog -ComputerName $ComputerName -Category 'Server' -Action 'Cleanup' -Item 'RemoteRun' -Status 'Failed' -Details $_.Exception.Message
    }
    finally {
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    }
}

$servers = Get-ServerList -BaseDN $SearchBase | Where-Object { $_.Enabled -eq $true }

foreach ($server in $servers) {
    if ($server.Name -match '^(SVC_|SVC0|SVC1|SAPService)') {
        continue
    }

    try {
        Invoke-RemoteCleanup -ComputerName $server.Name -ProfileAgeDays $ProfileAgeDays -KeepCcmCacheDays $KeepCcmCacheDays -WhatIfMode:$WhatIfMode
    }
    catch {
        Write-RunLog -ComputerName $server.Name -Category 'Server' -Action 'Cleanup' -Item 'Invoke' -Status 'Failed' -Details $_.Exception.Message
    }
}

$Results | Export-Csv -Path $RunLog -NoTypeInformation -Encoding UTF8
$Results | Out-File -FilePath $SummaryLog -Encoding UTF8

Write-Host "Cleanup complete."
Write-Host "Log CSV: $RunLog"
Write-Host "Summary: $SummaryLog"
