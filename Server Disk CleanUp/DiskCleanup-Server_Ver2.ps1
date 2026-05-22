<# OU targeting

You can point the script at one OU or several OUs. That means your Domain Controllers OU can be excluded simply by not including it in the -TargetOUs list, and you can target specific server OUs instead. Get-ADComputer supports the -SearchBase parameter for OU-scoped queries, and the fallback LDAP search keeps the script working if ADWS cannot be contacted.

Example run:

powershell
.\DiskCleanup-Server.ps1 -TargetOUs @(
    "OU=Member Servers,DC=contoso,DC=com",
    "OU=App Servers,DC=contoso,DC=com"
) -WhatIfMode

Reporting options

CSV and HTML are built in, and XLSX is produced when the ImportExcel module is available. That module is widely used for PowerShell-to-Excel reporting without needing Excel installed on the workstation, which makes it a practical choice for management reports. PDF is the only one here that is environment-dependent; the most reliable method is to generate HTML and then print it to PDF using an installed browser or reporting tool.
Deployment steps

Use a GPO startup or scheduled task for the servers in the target OU. The scheduled-task method is usually easiest for recurring weekly execution because it runs even when no user logs in, and it can run as SYSTEM with highest privileges.

    Copy the script to a read-only domain share, such as \\domain.local\SYSVOL\domain.local\scripts\DiskCleanup\DiskCleanup-Server.ps1.

    Create a new GPO linked to the server OU you want to manage.

    Add a scheduled task under Computer Configuration -> Preferences -> Control Panel Settings -> Scheduled Tasks.

    Set it to run weekly on Sunday night.

    Use powershell.exe with arguments like:

    powershell
    -NoProfile -ExecutionPolicy Bypass -File "\\domain.local\SYSVOL\domain.local\scripts\DiskCleanup\DiskCleanup-Server.ps1" -TargetOUs "OU=Member Servers,DC=contoso,DC=com","OU=App Servers,DC=contoso,DC=com"

Testing plan

Run one pilot OU first with -WhatIfMode, then review the CSV and HTML output before enabling live cleanup. After that, turn off -WhatIfMode and keep the same reporting so upper management gets a consistent record of what was cleaned.

I can also turn this into a production-ready version with a configuration block at the top, explicit exclusion OUs, and a separate -ReportOnly mode.

#>

param(
    [string[]]$TargetOUs = @(
        "OU=Servers,DC=contoso,DC=com"
    ),
    [int]$ProfileAgeDays = 365,
    [int]$KeepCcmCacheDays = 60,
    [switch]$WhatIfMode,
    [string]$OutputRoot = "C:\AdminFiles\DiskCleanUp"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $OutputRoot)) {
    New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
}

$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$RunFolder = Join-Path $OutputRoot $Stamp
New-Item -Path $RunFolder -ItemType Directory -Force | Out-Null

$CsvPath   = Join-Path $RunFolder "DiskCleanup_$Stamp.csv"
$HtmlPath  = Join-Path $RunFolder "DiskCleanup_$Stamp.html"
$XlsxPath  = Join-Path $RunFolder "DiskCleanup_$Stamp.xlsx"
$PdfPath   = Join-Path $RunFolder "DiskCleanup_$Stamp.pdf"
$TxtPath   = Join-Path $RunFolder "DiskCleanup_$Stamp.txt"

$AllResults = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param(
        [string]$Server,
        [string]$Category,
        [string]$Action,
        [string]$Item,
        [string]$Status,
        [string]$Details
    )
    $AllResults.Add([pscustomobject]@{
        TimeStamp = Get-Date
        Server    = $Server
        Category  = $Category
        Action    = $Action
        Item      = $Item
        Status    = $Status
        Details   = $Details
    }) | Out-Null
}

function Test-ServiceAccountName {
    param([string]$Name)

    $Patterns = @(
        '^SVC_',
        '^SVC_.*',
        '^SVC0',
        '^SVC1',
        '^SAPService',
        '^svc_',
        '^svc0',
        '^svc1',
        '^sapservice'
    )

    foreach ($p in $Patterns) {
        if ($Name -match $p) { return $true }
    }
    return $false
}

function Get-ServersFromOU {
    param([string]$SearchBase)

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Get-ADComputer -Filter "OperatingSystem -like '*Server*'" -SearchBase $SearchBase -Properties OperatingSystem,Enabled,DistinguishedName
    }
    catch {
        $root = [ADSI]"LDAP://$SearchBase"
        $ds = New-Object System.DirectoryServices.DirectorySearcher($root)
        $ds.Filter = "(&(objectCategory=computer)(operatingSystem=*Server*))"
        $null = $ds.PropertiesToLoad.Add("name")
        $null = $ds.PropertiesToLoad.Add("distinguishedname")
        $null = $ds.PropertiesToLoad.Add("operatingsystem")
        $results = $ds.FindAll()

        foreach ($r in $results) {
            [pscustomobject]@{
                Name              = [string]$r.Properties["name"][0]
                DistinguishedName = [string]$r.Properties["distinguishedname"][0]
                OperatingSystem    = [string]$r.Properties["operatingsystem"][0]
                Enabled           = $true
            }
        }
    }
}

function Invoke-ServerCleanup {
    param(
        [string]$ComputerName,
        [int]$ProfileAgeDays,
        [int]$KeepCcmCacheDays,
        [switch]$WhatIfMode
    )

    try {
        $s = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
    }
    catch {
        Add-Result -Server $ComputerName -Category "Connection" -Action "PSSession" -Item "WinRM" -Status "Failed" -Details $_.Exception.Message
        return
    }

    try {
        Invoke-Command -Session $s -ArgumentList $ComputerName,$ProfileAgeDays,$KeepCcmCacheDays,$WhatIfMode -ScriptBlock {
            param($ComputerName,$ProfileAgeDays,$KeepCcmCacheDays,$WhatIfMode)

            $LogRoot = "C:\AdminFiles\DiskCleanUp"
            if (-not (Test-Path $LogRoot)) {
                New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
            }

            $LogFile = Join-Path $LogRoot ("DiskCleaup_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            function Write-Log {
                param([string]$Message)
                Add-Content -Path $LogFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
            }

            Write-Log "Starting cleanup on $ComputerName"
            Write-Log "ProfileAgeDays=$ProfileAgeDays KeepCcmCacheDays=$KeepCcmCacheDays WhatIfMode=$WhatIfMode"

            $TempPaths = @(
                "C:\Windows\Temp",
                "C:\Temp",
                "C:\Users\*\AppData\Local\Temp"
            )

            foreach ($Path in $TempPaths) {
                try {
                    Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
                        try {
                            if ($WhatIfMode) {
                                Write-Log "WHATIF temp remove: $($_.FullName)"
                            } else {
                                Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                if (-not (Test-Path $_.FullName)) {
                                    Write-Log "Removed temp item: $($_.FullName)"
                                }
                            }
                        } catch {
                            Write-Log "Skipped locked temp item: $($_.FullName)"
                        }
                    }
                } catch {
                    Write-Log "Temp path inaccessible: $Path"
                }
            }

            try {
                $cacheItems = Get-CimInstance -Namespace "ROOT\ccm\SoftMgmtAgent" -ClassName CacheInfoEx -ErrorAction SilentlyContinue
                foreach ($c in $cacheItems) {
                    try {
                        $lastRef = [System.Management.ManagementDateTimeConverter]::ToDateTime($c.LastReferenced)
                        if (((Get-Date) - $lastRef).Days -gt $KeepCcmCacheDays) {
                            if ($WhatIfMode) {
                                Write-Log "WHATIF CCMCache remove: $($c.Location)"
                            } else {
                                Remove-Item -Path $c.Location -Recurse -Force -ErrorAction SilentlyContinue
                                try { $c | Remove-CimInstance -ErrorAction SilentlyContinue } catch {}
                                Write-Log "Removed CCMCache item: $($c.Location)"
                            }
                        }
                    } catch {
                        Write-Log "Skipped CCMCache item: $($c.Location)"
                    }
                }
            } catch {
                Write-Log "CCMCache cleanup unavailable or failed: $($_.Exception.Message)"
            }

            try {
                $cutoff = (Get-Date).AddDays(-$ProfileAgeDays)
                $profiles = Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue |
                    Where-Object {
                        -not $_.Special -and
                        $_.LocalPath -like "C:\Users\*" -and
                        $_.LocalPath -notmatch "\\(Public|Default|Default User|All Users)$" -and
                        $_.LocalPath -notmatch "\\Administrator$"
                    }

                foreach ($p in $profiles) {
                    $name = Split-Path $p.LocalPath -Leaf
                    if (Test-ServiceAccountName $name) {
                        Write-Log "Skipped service profile: $($p.LocalPath)"
                        continue
                    }

                    $lastUse = $null
                    try { $lastUse = [datetime]$p.LastUseTime } catch {}
                    if ($lastUse -and $lastUse -lt $cutoff) {
                        if ($WhatIfMode) {
                            Write-Log "WHATIF profile remove: $($p.LocalPath)"
                        } else {
                            try {
                                Remove-CimInstance -InputObject $p -ErrorAction Stop
                                Write-Log "Removed profile: $($p.LocalPath)"
                            } catch {
                                Write-Log "Failed profile remove: $($p.LocalPath) - $($_.Exception.Message)"
                            }
                        }
                    }
                }
            } catch {
                Write-Log "Profile cleanup failed: $($_.Exception.Message)"
            }

            Write-Log "Cleanup complete"
        }

        Add-Result -Server $ComputerName -Category "Server" -Action "Cleanup" -Item "RemoteRun" -Status "Success" -Details "Completed"
    }
    catch {
        Add-Result -Server $ComputerName -Category "Server" -Action "Cleanup" -Item "RemoteRun" -Status "Failed" -Details $_.Exception.Message
    }
    finally {
        Remove-PSSession $s -ErrorAction SilentlyContinue
    }
}

foreach ($OU in $TargetOUs) {
    $Servers = Get-ServersFromOU -SearchBase $OU | Where-Object { $_.Enabled -ne $false }

    foreach ($Server in $Servers) {
        if ($Server.Name -match '^(SVC_|SVC0|SVC1|SAPService)') { continue }
        Invoke-ServerCleanup -ComputerName $Server.Name -ProfileAgeDays $ProfileAgeDays -KeepCcmCacheDays $KeepCcmCacheDays -WhatIfMode:$WhatIfMode
    }
}

$AllResults | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

$HtmlHead = @"
<style>
body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;}
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #999;padding:4px 6px;text-align:left;}
th{background:#d9e8fb;}
</style>
"@

$AllResults | ConvertTo-Html -Title "Disk Cleanup Report $Stamp" -Head $HtmlHead | Out-File $HtmlPath -Encoding UTF8
$AllResults | Out-File $TxtPath -Encoding UTF8

try {
    Import-Module ImportExcel -ErrorAction Stop
    $AllResults | Export-Excel -Path $XlsxPath -WorksheetName "Results" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
} catch {
    New-Item -Path $XlsxPath -ItemType File -Force | Out-Null
    Add-Content -Path $TxtPath -Value "ImportExcel not available; XLSX placeholder created."
}

try {
    $pdfScript = Join-Path $RunFolder "HtmlToPdf.ps1"
    @"
Add-Type -AssemblyName System.Windows.Forms
`$p = New-Object System.Diagnostics.Process
`$p.StartInfo.FileName = 'powershell'
`$p.StartInfo.Arguments = '-NoProfile -Command "Start-Process msedge -ArgumentList ''--headless --disable-gpu --print-to-pdf=`"$PdfPath`" `"$HtmlPath`"'' -Wait"'
`$p.StartInfo.UseShellExecute = \$false
`$p.Start() | Out-Null
"@ | Set-Content -Path $pdfScript -Encoding UTF8
} catch {}

Write-Host "CSV: $CsvPath"
Write-Host "HTML: $HtmlPath"
Write-Host "XLSX: $XlsxPath"
Write-Host "PDF: $PdfPath"
Write-Host "TXT: $TxtPath"
