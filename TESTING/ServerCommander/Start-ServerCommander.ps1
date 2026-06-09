#Requires -Version 5.1
<#
.SYNOPSIS
  Server Commander -  All-In-One - Enterprise Server Management GUI for System Administrators
.DESCRIPTION
  A professional WPF-based all-in-one tool for daily server administration tasks:
    - Computer/Server Info (SystemInfo, DriverQuery, WMI Explorer integration)
    - Remote PowerShell Code Runner (single or multi-host, import from file)
    - Services, Processes, Event Logs, Disk management
    - Registry, Shares, Scheduled Tasks, WSUS/Windows Update status
    - Network diagnostics (ping, traceroute, open ports, DNS, netstat)
    - RDP Launcher and PSExec/PAExec integration
    - AD Computer lookup and SCCM/MECM quick queries
    - External Tools launcher (AdExplorer, WMIExplorer, PSExec, PAExec, sydi-server)
    - CMTrace-compatible logging
    - Dark/Light theme (Catppuccin Mocha-inspired dark + clean light)
    - Credential management per host/domain
    - Export results (CSV/JSON/TXT)
.NOTES
  Author : Steve McKee / IGT PLC
  Version: 1.0
  Requires: PowerShell 5.1, RSAT (optional), WinRM for remoting
  External : PSExec, PAExec, AdExplorer, WMIExplorer.ps1, sydi-server.vbs (optional)
  PSVersion: 5.1 ONLY - no ternary, no ??, no ?. operators
#>

param(
    [string]$InitialComputer = "",
    [switch]$DarkMode
)

# =========================================================
#  ASSEMBLIES
# =========================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================================================
#  GLOBAL CONFIG
# =========================================================
$Script:AppName    = "Server Commander - All-In-One"
$Script:Version    = "1.0"
$Script:IsDark     = $true
$Script:LogPath    = Join-Path $env:USERPROFILE "Desktop\ServerCommander_$(Get-Date -Format 'yyyyMMdd').log"
$Script:CredStore  = @{}   # key = hostname/domain, value = PSCredential
$Script:ExportBase = Join-Path $env:USERPROFILE "Desktop"
$Script:JobList    = [System.Collections.ArrayList]@()

# External tool paths - update these to match your environment
$Script:ExternalTools = @{
    PSExec       = "C:\Tools\PSExec.exe"
    PAExec       = "C:\Tools\PAExec.exe"
    AdExplorer   = "C:\Tools\AdExplorer.exe"
    WMIExplorer  = "C:\Tools\WmiExplorer.ps1"
    SydiServer   = "C:\Tools\sydi-server.vbs"
    SystemInfo   = "systeminfo.exe"
    DriverQuery  = "driverquery.exe"
    CMTrace      = "C:\Windows\CCM\CMTrace.exe"
}

# =========================================================
#  LOGGING (CMTrace-compatible)
# =========================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO',
        [string]$Component = "Server Commander"
    )
    $severity = switch ($Level) { 'INFO'{1} 'WARN'{2} 'ERROR'{3} 'DEBUG'{0} default{1} }
    $ts = Get-Date -Format "MM-dd-yyyy HH:mm:ss.fff"
    $cmEntry = "<![LOG[$Message]LOG]!><time=""$($ts.Split(' ')[1])+000"" date=""$($ts.Split(' ')[0])"" component=""$Component"" context="""" type=""$severity"" thread=""$PID"" file="""">"
    try { Add-Content -Path $Script:LogPath -Value $cmEntry -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

# =========================================================
#  RUNSPACE POOL (async work)
# =========================================================
$Script:RSPool = [runspacefactory]::CreateRunspacePool(1, 12)
$Script:RSPool.ThreadOptions = "ReuseThread"
$Script:RSPool.Open()

function Invoke-Async {
    param(
        [ScriptBlock]$ScriptBlock,
        [object[]]$ArgumentList,
        [ScriptBlock]$CompletedCallback
    )
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $Script:RSPool
    $null = $ps.AddScript($ScriptBlock)
    if ($ArgumentList) { foreach ($a in $ArgumentList) { $null = $ps.AddArgument($a) } }
    $ar = $ps.BeginInvoke()
    $null = [System.Threading.ThreadPool]::QueueUserWorkItem({
        param($state)
        $psInst = $state[0]; $asyncResult = $state[1]; $cb = $state[2]
        try   { $out = $psInst.EndInvoke($asyncResult); & $cb $out $null }
        catch { & $cb $null $_ }
        finally { $psInst.Dispose() }
    }, [object[]]@($ps, $ar, $CompletedCallback))
}

# =========================================================
#  HELPER UTILITIES
# =========================================================
function Get-Cred {
    param([string]$Target)
    if ($Script:CredStore.ContainsKey($Target)) { return $Script:CredStore[$Target] }
    return $null
}

function Export-GridData {
    param([object[]]$Data, [string]$Category)
    if (-not $Data -or $Data.Count -eq 0) { Show-Msg "No data to export." "Info"; return }
    $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
    $base = "$Category`_$ts"
    $path = Join-Path $Script:ExportBase "$base.csv"
    $Data | Export-Csv -Path $path -NoTypeInformation -Force
    Show-Msg "Exported $($Data.Count) rows to:`n$path" "Export"
    Write-Log "Exported $($Data.Count) $Category rows to $path"
}

function Show-Msg {
    param([string]$Text, [string]$Title = "Info")
    [System.Windows.MessageBox]::Show($Text, "$Script:AppName - $Title", 'OK', 'Information') | Out-Null
}

function Test-ToolPath {
    param([string]$Key)
    $p = $Script:ExternalTools[$Key]
    if (-not $p) { return $false }
    if ($p -like "*.exe" -or $p -like "*.vbs" -or $p -like "*.ps1") {
        return (Test-Path $p)
    }
    # Built-in commands (systeminfo, driverquery)
    return ($null -ne (Get-Command $p -ErrorAction SilentlyContinue))
}

function Get-SafeString { param([string]$s) if (-not $s) { return "" }; return $s }

# =========================================================
#  REMOTE EXECUTION HELPERS
# =========================================================
function Invoke-RemoteCode {
    param(
        [string[]]$ComputerList,
        [string]$Code,
        [System.Management.Automation.PSCredential]$Cred = $null,
        [int]$ThrottleLimit = 10
    )
    $results = [System.Collections.ArrayList]@()
    $jobs    = [System.Collections.ArrayList]@()

    foreach ($computer in $ComputerList) {
        $sb = [scriptblock]::Create($Code)
        $params = @{
            ComputerName  = $computer
            ScriptBlock   = $sb
            ErrorAction   = 'Stop'
        }
        if ($Cred) { $params.Credential = $Cred }
        try {
            $job = Invoke-Command @params -AsJob -JobName "AIO_$computer"
            $null = $jobs.Add([PSCustomObject]@{ Computer=$computer; Job=$job })
        } catch {
            $null = $results.Add([PSCustomObject]@{
                ComputerName = $computer
                Status       = "ERROR"
                Output       = $_.Exception.Message
                Duration     = "N/A"
            })
        }
    }

    $startTime = Get-Date
    while ($jobs | Where-Object { $_.Job.State -in @('Running','NotStarted') }) {
        Start-Sleep -Milliseconds 200
        if ((Get-Date) -gt $startTime.AddSeconds(120)) { break }
    }

    foreach ($item in $jobs) {
        try {
            $out = Receive-Job -Job $item.Job -ErrorAction Stop
            $dur = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
            $null = $results.Add([PSCustomObject]@{
                ComputerName = $item.Computer
                Status       = "OK"
                Output       = ($out | Out-String).Trim()
                Duration     = "$dur`s"
            })
        } catch {
            $null = $results.Add([PSCustomObject]@{
                ComputerName = $item.Computer
                Status       = "ERROR"
                Output       = $_.Exception.Message
                Duration     = "N/A"
            })
        }
        Remove-Job -Job $item.Job -Force -ErrorAction SilentlyContinue
    }
    return $results
}

function Get-ComputerInfo-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        $os  = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
        $cs  = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
        $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $mem = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        $up  = if ($os) { (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime) } else { $null }
        [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            OSName       = if ($os) { $os.Caption } else { "N/A" }
            OSVersion    = if ($os) { $os.Version } else { "N/A" }
            OSBuild      = if ($os) { $os.BuildNumber } else { "N/A" }
            ServicePack  = if ($os) { $os.ServicePackMajorVersion } else { "N/A" }
            Domain       = if ($cs) { $cs.Domain } else { "N/A" }
            Manufacturer = if ($cs) { $cs.Manufacturer } else { "N/A" }
            Model        = if ($cs) { $cs.Model } else { "N/A" }
            TotalRAM_GB  = $mem
            CPU          = if ($cpu) { $cpu.Name } else { "N/A" }
            LogicalCores = if ($cpu) { $cpu.NumberOfLogicalProcessors } else { "N/A" }
            UptimeDays   = if ($up) { [math]::Round($up.TotalDays, 2) } else { "N/A" }
            LastBoot     = if ($os) { $os.ConvertToDateTime($os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm') } else { "N/A" }
            IPv4         = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.*" } | Select-Object -First 1 -ExpandProperty IPAddress)
        }
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    try { return Invoke-Command @p } catch { throw }
}

function Get-Services-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null, [string]$Filter = "")
    $sb = [scriptblock]::Create("Get-Service -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType | Sort-Object DisplayName")
    $p  = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    $svc = Invoke-Command @p
    if ($Filter) { $svc = $svc | Where-Object { $_.DisplayName -like "*$Filter*" -or $_.Name -like "*$Filter*" } }
    return $svc
}

function Get-Processes-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        Get-Process -ErrorAction SilentlyContinue |
            Select-Object Name,Id,CPU,
                @{n='WorkingSet_MB';e={[math]::Round($_.WorkingSet64/1MB,1)}},
                @{n='VirtualMem_MB';e={[math]::Round($_.VirtualMemorySize64/1MB,1)}},
                @{n='Threads';e={$_.Threads.Count}},
                Description,
                @{n='StartTime';e={if($_.StartTime){$_.StartTime.ToString('HH:mm:ss')}else{'N/A'}}} |
            Sort-Object WorkingSet_MB -Descending
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-EventLog-Remote {
    param([string]$ComputerName, [string]$LogName = "System", [int]$Count = 100, [System.Management.Automation.PSCredential]$Cred = $null, [string]$Level = "All")
    $sb = [scriptblock]::Create(@"
`$level = '$Level'
`$logName = '$LogName'
`$count = $Count
`$filter = @{ LogName=`$logName; MaxEvents=`$count }
if (`$level -eq 'Error')    { `$filter.Level = 2 }
elseif (`$level -eq 'Warning') { `$filter.Level = 3 }
elseif (`$level -eq 'Info')    { `$filter.Level = 4 }
Get-WinEvent -FilterHashtable `$filter -ErrorAction SilentlyContinue |
    Select-Object TimeCreated,Id,LevelDisplayName,ProviderName,Message |
    Sort-Object TimeCreated -Descending
"@)
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-Disks-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue |
            Select-Object DeviceID,VolumeName,
                @{n='Total_GB';e={[math]::Round($_.Size/1GB,2)}},
                @{n='Free_GB';e={[math]::Round($_.FreeSpace/1GB,2)}},
                @{n='Used_GB';e={[math]::Round(($_.Size-$_.FreeSpace)/1GB,2)}},
                @{n='PercentFree';e={if($_.Size -gt 0){[math]::Round(($_.FreeSpace/$_.Size)*100,1)}else{0}}},
                FileSystem
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-Shares-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        Get-WmiObject -Class Win32_Share -ErrorAction SilentlyContinue |
            Select-Object Name,Path,Description,
                @{n='Type';e={switch($_.Type){0{'Disk'}1{'Printer'}2{'Device'}3{'IPC'}2147483648{'Special Disk'}2147483649{'Special Printer'}default{$_.Type}}}}
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-ScheduledTasks-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Disabled' } |
            Select-Object TaskName,TaskPath,State,
                @{n='LastRun';e={($_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue).LastRunTime}},
                @{n='NextRun';e={($_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue).NextRunTime}},
                @{n='LastResult';e={($_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue).LastTaskResult}},
                @{n='RunAs';e={$_.Principal.UserId}} |
            Sort-Object TaskPath,TaskName
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-NetworkInfo-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        $adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' }
        $adapters | ForEach-Object {
            $nic = $_
            $ip4 = Get-NetIPAddress -InterfaceIndex $nic.IfIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
            $ip6 = Get-NetIPAddress -InterfaceIndex $nic.IfIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue | Where-Object { $_.PrefixOrigin -ne 'WellKnown' } | Select-Object -First 1
            $dns = Get-DnsClientServerAddress -InterfaceIndex $nic.IfIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                Name         = $nic.Name
                Description  = $nic.InterfaceDescription
                MacAddress   = $nic.MacAddress
                LinkSpeed    = $nic.LinkSpeed
                Status       = $nic.Status
                IPv4Address  = if ($ip4) { ($ip4 | Select-Object -First 1).IPAddress } else { "N/A" }
                SubnetPrefix = if ($ip4) { ($ip4 | Select-Object -First 1).PrefixLength } else { "N/A" }
                IPv6Address  = if ($ip6) { $ip6.IPAddress } else { "N/A" }
                DNSServers   = if ($dns) { ($dns.ServerAddresses -join ', ') } else { "N/A" }
                DHCP         = if ($ip4) { ($ip4 | Select-Object -First 1).PrefixOrigin -eq 'Dhcp' } else { $false }
            }
        }
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Get-InstalledSoftware-Remote {
    param([string]$ComputerName, [string]$Filter = "", [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $software = $regPaths | ForEach-Object {
            Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -and $_.DisplayName.Trim() } |
                Select-Object DisplayName,DisplayVersion,Publisher,InstallDate,EstimatedSize
        }
        $software | Sort-Object DisplayName -Unique
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    $result = Invoke-Command @p
    if ($Filter) { $result = $result | Where-Object { $_.DisplayName -like "*$Filter*" -or $_.Publisher -like "*$Filter*" } }
    return $result
}

function Get-WindowsUpdates-Remote {
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Cred = $null)
    $sb = {
        $sess = New-Object -ComObject Microsoft.Update.Session
        $searcher = $sess.CreateUpdateSearcher()
        $count = $searcher.GetTotalHistoryCount()
        if ($count -gt 0) {
            $history = $searcher.QueryHistory(0, [math]::Min($count, 50))
            $history | ForEach-Object {
                [PSCustomObject]@{
                    Title          = $_.Title
                    Date           = $_.Date.ToString('yyyy-MM-dd HH:mm')
                    ResultCode     = switch($_.ResultCode){1{'In Progress'}2{'Succeeded'}3{'Succeeded w/Errors'}4{'Failed'}5{'Aborted'}default{$_.ResultCode}}
                    KB             = if ($_.Title -match 'KB(\d+)') { "KB$($matches[1])" } else { "N/A" }
                    Description    = $_.Description
                }
            } | Sort-Object Date -Descending
        }
    }
    $p = @{ ComputerName=$ComputerName; ScriptBlock=$sb; ErrorAction='Stop' }
    if ($Cred) { $p.Credential = $Cred }
    return Invoke-Command @p
}

function Test-NetworkReach {
    param([string]$ComputerName, [int[]]$Ports = @(135,445,3389,5985))
    $results = [System.Collections.ArrayList]@()
    $ping = Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction SilentlyContinue
    $null = $results.Add([PSCustomObject]@{
        Test   = "Ping"
        Target = $ComputerName
        Result = if ($ping) { "OK - $([math]::Round($ping.ResponseTime,0))ms" } else { "FAILED" }
        Note   = ""
    })
    foreach ($port in $Ports) {
        $tcp = $null
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $conn = $tcp.BeginConnect($ComputerName, $port, $null, $null)
            $wait = $conn.AsyncWaitHandle.WaitOne(1500, $false)
            if ($wait) { $tcp.EndConnect($conn); $status = "OPEN" } else { $status = "TIMEOUT" }
        } catch { $status = "CLOSED/ERROR" }
        finally { if ($tcp) { $tcp.Close() } }
        $note = switch ($port) {
            135   { "WMI/RPC" }
            139   { "NetBIOS" }
            445   { "SMB/File Sharing" }
            3389  { "RDP" }
            5985  { "WinRM HTTP" }
            5986  { "WinRM HTTPS" }
            22    { "SSH" }
            443   { "HTTPS" }
            80    { "HTTP" }
            default { "" }
        }
        $null = $results.Add([PSCustomObject]@{
            Test   = "TCP Port $port"
            Target = $ComputerName
            Result = $status
            Note   = $note
        })
    }
    return $results
}

# =========================================================
#  THEME DEFINITIONS
# =========================================================
$Script:Themes = @{
    Dark = @{
        WindowBg       = "#1E1F2E"
        PanelBg        = "#252638"
        TabBg          = "#1A1B2A"
        TabSelected    = "#6C63FF"
        TabHover       = "#2D2E45"
        ControlBg      = "#2D2E45"
        ControlBorder  = "#4A4B6A"
        ControlFocus   = "#6C63FF"
        ButtonBg       = "#6C63FF"
        ButtonHover    = "#7C73FF"
        ButtonFg       = "#FFFFFF"
        DangerBg       = "#C0392B"
        SuccessBg      = "#27AE60"
        WarnBg         = "#D68910"
        AccentCyan     = "#76E4F7"
        AccentYellow   = "#F6E05E"
        TextPrimary    = "#E8E9F3"
        TextSecondary  = "#9898B5"
        TextMuted      = "#5A5B7A"
        GridHeaderBg   = "#1A1B2A"
        GridAltRow     = "#232436"
        GridBorder     = "#3A3B5A"
        StatusBg       = "#141520"
        SectionLabel   = "#6C63FF"
        HeaderGrad1    = "#6C63FF"
        HeaderGrad2    = "#9F7AEA"
        RowError       = "#4A1515"
        RowWarn        = "#3A2E10"
        RowOk          = "#102A1A"
        ScrollThumb    = "#4A4B6A"
        CodeBg         = "#12131F"
        CodeFg         = "#76E4F7"
        OutputBg       = "#181928"
    }
    Light = @{
        WindowBg       = "#F0F2FA"
        PanelBg        = "#FFFFFF"
        TabBg          = "#E8EAF6"
        TabSelected    = "#5C54E8"
        TabHover       = "#D8DAF0"
        ControlBg      = "#FFFFFF"
        ControlBorder  = "#C5C9E8"
        ControlFocus   = "#5C54E8"
        ButtonBg       = "#5C54E8"
        ButtonHover    = "#6C64F8"
        ButtonFg       = "#FFFFFF"
        DangerBg       = "#C0392B"
        SuccessBg      = "#27AE60"
        WarnBg         = "#D68910"
        AccentCyan     = "#2B6CB0"
        AccentYellow   = "#D69E2E"
        TextPrimary    = "#1A1B2E"
        TextSecondary  = "#4A4B6A"
        TextMuted      = "#888AAA"
        GridHeaderBg   = "#E8EAF6"
        GridAltRow     = "#F8F9FF"
        GridBorder     = "#D0D3EC"
        StatusBg       = "#E0E3F8"
        SectionLabel   = "#5C54E8"
        HeaderGrad1    = "#5C54E8"
        HeaderGrad2    = "#7C6FE8"
        RowError       = "#FED7D7"
        RowWarn        = "#FEFCBF"
        RowOk          = "#C6F6D5"
        ScrollThumb    = "#C5C9E8"
        CodeBg         = "#F8F9FF"
        CodeFg         = "#2B4C7E"
        OutputBg       = "#EEF0FF"
    }
}

# =========================================================
#  XAML WINDOW DEFINITION
# =========================================================
function Get-XAML {
    param([hashtable]$T)
    return @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='$($Script:AppName) v$($Script:Version)'
    Height='900' Width='1480'
    MinHeight='700' MinWidth='1100'
    WindowStartupLocation='CenterScreen'
    Background='$($T.WindowBg)'
    FontFamily='Segoe UI'
    FontSize='13'>

  <Window.Resources>

    <!-- ===== BUTTON STYLE ===== -->
    <Style x:Key='ModernButton' TargetType='Button'>
      <Setter Property='Background' Value='$($T.ButtonBg)'/>
      <Setter Property='Foreground' Value='$($T.ButtonFg)'/>
      <Setter Property='BorderThickness' Value='0'/>
      <Setter Property='Padding' Value='14,7'/>
      <Setter Property='Cursor' Value='Hand'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='FontWeight' Value='SemiBold'/>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='Button'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' CornerRadius='6' Padding='{TemplateBinding Padding}'>
              <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property='IsMouseOver' Value='True'>
                <Setter TargetName='bd' Property='Background' Value='$($T.ButtonHover)'/>
              </Trigger>
              <Trigger Property='IsPressed' Value='True'>
                <Setter TargetName='bd' Property='Opacity' Value='0.82'/>
              </Trigger>
              <Trigger Property='IsEnabled' Value='False'>
                <Setter TargetName='bd' Property='Opacity' Value='0.42'/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key='DangerButton' TargetType='Button' BasedOn='{StaticResource ModernButton}'>
      <Setter Property='Background' Value='$($T.DangerBg)'/>
    </Style>
    <Style x:Key='SuccessButton' TargetType='Button' BasedOn='{StaticResource ModernButton}'>
      <Setter Property='Background' Value='$($T.SuccessBg)'/>
    </Style>
    <Style x:Key='WarnButton' TargetType='Button' BasedOn='{StaticResource ModernButton}'>
      <Setter Property='Background' Value='$($T.WarnBg)'/>
    </Style>

    <!-- ===== TEXTBOX STYLE ===== -->
    <Style x:Key='ModernTextBox' TargetType='TextBox'>
      <Setter Property='Background' Value='$($T.ControlBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
      <Setter Property='CaretBrush' Value='$($T.TextPrimary)'/>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='TextBox'>
            <Border x:Name='bd' Background='{TemplateBinding Background}'
                    BorderBrush='{TemplateBinding BorderBrush}'
                    BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='6'>
              <ScrollViewer x:Name='PART_ContentHost' Margin='0' VerticalAlignment='Center'/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property='IsFocused' Value='True'>
                <Setter TargetName='bd' Property='BorderBrush' Value='$($T.ControlFocus)'/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ===== CODE TEXTBOX (monospace) ===== -->
    <Style x:Key='CodeTextBox' TargetType='TextBox' BasedOn='{StaticResource ModernTextBox}'>
      <Setter Property='Background' Value='$($T.CodeBg)'/>
      <Setter Property='Foreground' Value='$($T.CodeFg)'/>
      <Setter Property='FontFamily' Value='Consolas, Courier New'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='AcceptsReturn' Value='True'/>
      <Setter Property='AcceptsTab' Value='True'/>
      <Setter Property='VerticalScrollBarVisibility' Value='Auto'/>
      <Setter Property='HorizontalScrollBarVisibility' Value='Auto'/>
      <Setter Property='SpellCheck.IsEnabled' Value='False'/>
    </Style>

    <!-- ===== OUTPUT TEXTBOX ===== -->
    <Style x:Key='OutputTextBox' TargetType='TextBox'>
      <Setter Property='Background' Value='$($T.OutputBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
      <Setter Property='IsReadOnly' Value='True'/>
      <Setter Property='FontFamily' Value='Consolas, Courier New'/>
      <Setter Property='FontSize' Value='11'/>
      <Setter Property='AcceptsReturn' Value='True'/>
      <Setter Property='VerticalScrollBarVisibility' Value='Auto'/>
      <Setter Property='HorizontalScrollBarVisibility' Value='Auto'/>
      <Setter Property='TextWrapping' Value='NoWrap'/>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='TextBox'>
            <Border Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}'
                    BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='6'>
              <ScrollViewer x:Name='PART_ContentHost'/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ===== COMBOBOX ===== -->
    <Style x:Key='ModernComboBox' TargetType='ComboBox'>
      <Setter Property='Background' Value='$($T.ControlBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
    </Style>

    <!-- ===== LABEL STYLES ===== -->
    <Style x:Key='ModernLabel' TargetType='Label'>
      <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='FontWeight' Value='SemiBold'/>
      <Setter Property='Padding' Value='2,0'/>
      <Setter Property='VerticalAlignment' Value='Center'/>
    </Style>
    <Style x:Key='SectionLabel' TargetType='TextBlock'>
      <Setter Property='Foreground' Value='$($T.SectionLabel)'/>
      <Setter Property='FontSize' Value='11'/>
      <Setter Property='FontWeight' Value='Bold'/>
      <Setter Property='Margin' Value='0,0,0,4'/>
    </Style>
    <Style x:Key='HeaderLabel' TargetType='TextBlock'>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='FontSize' Value='13'/>
      <Setter Property='FontWeight' Value='Bold'/>
    </Style>

    <!-- ===== DATAGRID ===== -->
    <Style x:Key='ModernDataGrid' TargetType='DataGrid'>
      <Setter Property='Background' Value='$($T.PanelBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.GridBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='GridLinesVisibility' Value='Horizontal'/>
      <Setter Property='HorizontalGridLinesBrush' Value='$($T.GridBorder)'/>
      <Setter Property='RowBackground' Value='$($T.PanelBg)'/>
      <Setter Property='AlternatingRowBackground' Value='$($T.GridAltRow)'/>
      <Setter Property='AutoGenerateColumns' Value='True'/>
      <Setter Property='IsReadOnly' Value='True'/>
      <Setter Property='SelectionMode' Value='Extended'/>
      <Setter Property='CanUserAddRows' Value='False'/>
      <Setter Property='CanUserDeleteRows' Value='False'/>
      <Setter Property='HeadersVisibility' Value='Column'/>
      <Setter Property='ColumnHeaderStyle'>
        <Setter.Value>
          <Style TargetType='DataGridColumnHeader'>
            <Setter Property='Background' Value='$($T.GridHeaderBg)'/>
            <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
            <Setter Property='Padding' Value='8,6'/>
            <Setter Property='FontWeight' Value='SemiBold'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='BorderBrush' Value='$($T.GridBorder)'/>
            <Setter Property='BorderThickness' Value='0,0,1,1'/>
          </Style>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ===== CHECKBOX ===== -->
    <Style x:Key='ModernCheckBox' TargetType='CheckBox'>
      <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='VerticalAlignment' Value='Center'/>
      <Setter Property='Margin' Value='4,0'/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='54'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='26'/>
    </Grid.RowDefinitions>

    <!-- ===== HEADER BAR ===== -->
    <Border Grid.Row='0'>
      <Border.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='1,0'>
          <GradientStop Color='$($T.HeaderGrad1)' Offset='0'/>
          <GradientStop Color='$($T.HeaderGrad2)' Offset='1'/>
        </LinearGradientBrush>
      </Border.Background>
      <Grid Margin='16,0'>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='*'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column='0' Orientation='Horizontal' VerticalAlignment='Center'>
          <TextBlock Text='[Server Commander v1.0]' Foreground='White' FontSize='22' FontWeight='Bold' VerticalAlignment='Center' Margin='0,0,10,0'/>
          <TextBlock Text='$($Script:AppName) v$($Script:Version)' Foreground='#DDEEFF' FontSize='14' FontWeight='SemiBold' VerticalAlignment='Center'/>
        </StackPanel>
        <StackPanel Grid.Column='2' Orientation='Horizontal' VerticalAlignment='Center' Margin='0,0,12,0'>
          <TextBlock Text='Target:' Foreground='#CCDDFF' FontSize='12' VerticalAlignment='Center' Margin='0,0,6,0'/>
          <TextBox Name='txtGlobalTarget' Width='180' Height='28' Style='{StaticResource ModernTextBox}' Background='#22FFFFFF' Foreground='White' CaretBrush='White'
                   VerticalContentAlignment='Center' ToolTip='Enter hostname or IP. Used as default target for all tabs.'/>
        </StackPanel>
        <Button Name='btnSetCred' Grid.Column='3' Content='Set Creds' Style='{StaticResource ModernButton}'
                Background='#22FFFFFF' Margin='0,0,6,0' Height='30' ToolTip='Store credentials for target host'/>
        <Button Name='btnThemeToggle' Grid.Column='4' Content='Light Mode' Style='{StaticResource ModernButton}'
                Background='#22FFFFFF' Margin='0,0,6,0' Height='30'/>
        <Button Name='btnOpenLog' Grid.Column='5' Content='View Log' Style='{StaticResource ModernButton}'
                Background='#22FFFFFF' Height='30' ToolTip='Open CMTrace-compatible log file'/>
      </Grid>
    </Border>

    <!-- ===== MAIN TAB CONTROL ===== -->
    <TabControl Name='MainTabs' Grid.Row='1' Background='$($T.TabBg)' BorderThickness='0' Margin='0'>
      <TabControl.Resources>
        <Style TargetType='TabItem'>
          <Setter Property='Background' Value='$($T.TabBg)'/>
          <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
          <Setter Property='BorderThickness' Value='0'/>
          <Setter Property='Padding' Value='16,10'/>
          <Setter Property='FontSize' Value='12'/>
          <Setter Property='FontWeight' Value='SemiBold'/>
          <Setter Property='Template'>
            <Setter.Value>
              <ControlTemplate TargetType='TabItem'>
                <Border x:Name='tb' Background='{TemplateBinding Background}' Padding='{TemplateBinding Padding}' BorderThickness='0,0,0,3' BorderBrush='Transparent'>
                  <ContentPresenter ContentSource='Header' HorizontalAlignment='Center' VerticalAlignment='Center'/>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property='IsSelected' Value='True'>
                    <Setter TargetName='tb' Property='Background' Value='$($T.PanelBg)'/>
                    <Setter TargetName='tb' Property='BorderBrush' Value='$($T.TabSelected)'/>
                    <Setter Property='Foreground' Value='$($T.TabSelected)'/>
                  </Trigger>
                  <Trigger Property='IsMouseOver' Value='True'>
                    <Setter TargetName='tb' Property='Background' Value='$($T.TabHover)'/>
                  </Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Style>
      </TabControl.Resources>

      <!-- ============================================ -->
      <!-- TAB 1: COMPUTER INFO                         -->
      <!-- ============================================ -->
      <TabItem Header='Computer Info'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>

          <!-- Target row -->
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='COMPUTER INFO' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtInfoHost' Width='200' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0' ToolTip='Hostname or IP - defaults to Global Target'/>
            <Button Name='btnInfoQuery' Content='Query Computer' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnInfoLocal' Content='Query Local' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnInfoRDP' Content='Launch RDP' Style='{StaticResource WarnButton}' Height='30' Margin='0,0,6,0' ToolTip='Open mstsc to target'/>
            <Button Name='btnInfoPSExec' Content='PSExec Shell' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Launch PSExec remote shell'/>
            <Button Name='btnInfoSydi' Content='SYDI Report' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Run sydi-server.vbs inventory'/>
            <Button Name='btnInfoSysinfo' Content='SystemInfo' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Run systeminfo.exe'/>
            <Button Name='btnInfoDrivers' Content='DriverQuery' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Run driverquery.exe'/>
            <Button Name='btnInfoExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>

          <!-- Info summary cards row -->
          <Border Grid.Row='1' Background='$($T.ControlBg)' CornerRadius='8' Padding='12,8' Margin='0,0,0,8'>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width='*'/>
                <ColumnDefinition Width='*'/>
                <ColumnDefinition Width='*'/>
                <ColumnDefinition Width='*'/>
                <ColumnDefinition Width='*'/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column='0'>
                <TextBlock Text='OS / Build' Style='{StaticResource SectionLabel}'/>
                <TextBlock Name='lblInfoOS' Text='---' Foreground='$($T.TextPrimary)' FontSize='12' TextWrapping='Wrap'/>
              </StackPanel>
              <StackPanel Grid.Column='1'>
                <TextBlock Text='Hardware' Style='{StaticResource SectionLabel}'/>
                <TextBlock Name='lblInfoHW' Text='---' Foreground='$($T.TextPrimary)' FontSize='12' TextWrapping='Wrap'/>
              </StackPanel>
              <StackPanel Grid.Column='2'>
                <TextBlock Text='Memory / CPU' Style='{StaticResource SectionLabel}'/>
                <TextBlock Name='lblInfoCPU' Text='---' Foreground='$($T.TextPrimary)' FontSize='12' TextWrapping='Wrap'/>
              </StackPanel>
              <StackPanel Grid.Column='3'>
                <TextBlock Text='Network' Style='{StaticResource SectionLabel}'/>
                <TextBlock Name='lblInfoNet' Text='---' Foreground='$($T.TextPrimary)' FontSize='12' TextWrapping='Wrap'/>
              </StackPanel>
              <StackPanel Grid.Column='4'>
                <TextBlock Text='Uptime / Last Boot' Style='{StaticResource SectionLabel}'/>
                <TextBlock Name='lblInfoUptime' Text='---' Foreground='$($T.AccentCyan)' FontSize='12'/>
              </StackPanel>
            </Grid>
          </Border>

          <!-- Data grid -->
          <DataGrid Name='dgInfo' Grid.Row='2' Style='{StaticResource ModernDataGrid}'/>

          <TextBlock Name='txtInfoCount' Grid.Row='3' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 2: SERVICES                              -->
      <!-- ============================================ -->
      <TabItem Header='Services'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='SERVICES' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtSvcHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Label Content='Filter:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtSvcFilter' Width='150' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnSvcQuery' Content='List Services' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnSvcStart' Content='Start' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnSvcStop' Content='Stop' Style='{StaticResource DangerButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnSvcRestart' Content='Restart' Style='{StaticResource WarnButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnSvcExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgServices' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtSvcCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 3: PROCESSES                             -->
      <!-- ============================================ -->
      <TabItem Header='Processes'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='PROCESSES' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtProcHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnProcQuery' Content='List Processes' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnProcKill' Content='Kill Process' Style='{StaticResource DangerButton}' Height='30' Margin='0,0,6,0' ToolTip='Kill selected process (use with caution)'/>
            <Button Name='btnProcExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgProcesses' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtProcCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 4: EVENT LOGS                            -->
      <!-- ============================================ -->
      <TabItem Header='Event Logs'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='EVENT LOGS' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtEvtHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Label Content='Log:' Style='{StaticResource ModernLabel}'/>
            <ComboBox Name='cmbEvtLog' Width='120' Height='30' Style='{StaticResource ModernComboBox}' Margin='4,0,8,0'>
              <ComboBoxItem Content='System' IsSelected='True'/>
              <ComboBoxItem Content='Application'/>
              <ComboBoxItem Content='Security'/>
              <ComboBoxItem Content='Setup'/>
              <ComboBoxItem Content='Microsoft-Windows-WindowsUpdateClient/Operational'/>
              <ComboBoxItem Content='Microsoft-Windows-TaskScheduler/Operational'/>
            </ComboBox>
            <Label Content='Level:' Style='{StaticResource ModernLabel}'/>
            <ComboBox Name='cmbEvtLevel' Width='100' Height='30' Style='{StaticResource ModernComboBox}' Margin='4,0,8,0'>
              <ComboBoxItem Content='All' IsSelected='True'/>
              <ComboBoxItem Content='Error'/>
              <ComboBoxItem Content='Warning'/>
              <ComboBoxItem Content='Info'/>
            </ComboBox>
            <Label Content='Count:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtEvtCount' Text='100' Width='60' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnEvtQuery' Content='Get Events' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnEvtExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgEvents' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtEvtStatusCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 5: DISKS                                 -->
      <!-- ============================================ -->
      <TabItem Header='Disks'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='DISK / STORAGE' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtDiskHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnDiskQuery' Content='Get Disk Info' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnDiskExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgDisks' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtDiskCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 6: SHARES                                -->
      <!-- ============================================ -->
      <TabItem Header='Shares'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='SHARES' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtShareHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnShareQuery' Content='List Shares' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnShareOpen' Content='Open in Explorer' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0' ToolTip='Open selected share UNC path'/>
            <Button Name='btnShareExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgShares' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtShareCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 7: SCHEDULED TASKS                       -->
      <!-- ============================================ -->
      <TabItem Header='Sched Tasks'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='SCHEDULED TASKS' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtTaskHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnTaskQuery' Content='List Tasks' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnTaskRun' Content='Run Now' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnTaskEnable' Content='Enable' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnTaskDisable' Content='Disable' Style='{StaticResource DangerButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnTaskExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgTasks' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtTaskCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 8: NETWORK                               -->
      <!-- ============================================ -->
      <TabItem Header='Network'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>

          <!-- Network actions -->
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='NETWORK DIAGNOSTICS' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Target:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtNetHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnNetAdapters' Content='NIC Info' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnNetTest' Content='Reachability Test' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Ping + key port tests'/>
            <Label Content='Extra Ports:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtNetPorts' Width='140' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'
                     ToolTip='Comma-separated port numbers (e.g. 80,443,1433)' Text='80,443,1433,5985'/>
            <Button Name='btnNetExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>

          <!-- DNS/Trace sub-row -->
          <StackPanel Grid.Row='1' Orientation='Horizontal' Margin='0,0,0,8'>
            <Label Content='DNS Lookup:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtDnsLookup' Width='200' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,6,0'/>
            <Button Name='btnDnsLookup' Content='Resolve' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,12,0'/>
            <Label Content='Tracert:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtTracert' Width='200' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,6,0'/>
            <Button Name='btnTracert' Content='Trace Route' Style='{StaticResource WarnButton}' Height='30' Margin='0,0,12,0'/>
            <Button Name='btnNetstat' Content='Netstat (local)' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>

          <DataGrid Name='dgNetwork' Grid.Row='2' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtNetCount' Grid.Row='3' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 9: SOFTWARE                              -->
      <!-- ============================================ -->
      <TabItem Header='Software'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='INSTALLED SOFTWARE' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtSwHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Label Content='Filter:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtSwFilter' Width='150' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnSwQuery' Content='List Software' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnSwExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgSoftware' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtSwCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 10: WINDOWS UPDATES                      -->
      <!-- ============================================ -->
      <TabItem Header='Win Updates'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='WINDOWS UPDATE HISTORY' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Label Content='Host:' Style='{StaticResource ModernLabel}'/>
            <TextBox Name='txtUpdHost' Width='180' Height='30' Style='{StaticResource ModernTextBox}' Margin='4,0,8,0'/>
            <Button Name='btnUpdQuery' Content='Get Update History' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnUpdExport' Content='Export CSV' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>
          <DataGrid Name='dgUpdates' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
          <TextBlock Name='txtUpdCount' Grid.Row='2' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 11: REMOTE PS CODE RUNNER                -->
      <!-- ============================================ -->
      <TabItem Header='Remote PS Runner'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='6'/>
            <RowDefinition Height='*'/>
            <RowDefinition Height='Auto'/>
          </Grid.RowDefinitions>

          <!-- Toolbar -->
          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='REMOTE POWERSHELL RUNNER' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Button Name='btnPSRun' Content='Run on Target(s)' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnPSClear' Content='Clear Code' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnPSClearOut' Content='Clear Output' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnPSImport' Content='Import Script...' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0' ToolTip='Load a .ps1 file into the editor'/>
            <Button Name='btnPSImportHosts' Content='Import Host List...' Style='{StaticResource WarnButton}' Height='30' Margin='0,0,6,0' ToolTip='Load newline-separated host list from file'/>
            <Button Name='btnPSSave' Content='Save Script...' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnPSExport' Content='Export Results' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>

          <!-- Split: left = code editor, right = host list -->
          <Grid Grid.Row='1'>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width='*'/>
              <ColumnDefinition Width='8'/>
              <ColumnDefinition Width='240'/>
            </Grid.ColumnDefinitions>

            <!-- Code editor panel -->
            <Border Grid.Column='0' Background='$($T.CodeBg)' CornerRadius='8' Padding='0'>
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height='Auto'/>
                  <RowDefinition Height='*'/>
                </Grid.RowDefinitions>
                <Border Grid.Row='0' Background='$($T.GridHeaderBg)' CornerRadius='8,8,0,0' Padding='10,6'>
                  <StackPanel Orientation='Horizontal'>
                    <TextBlock Text='PowerShell Editor' Style='{StaticResource SectionLabel}' Margin='0'/>
                    <TextBlock Text='  |  Ctrl+A to select all  |  Tab = 4 spaces' Foreground='$($T.TextMuted)' FontSize='10' VerticalAlignment='Center' Margin='8,0,0,0'/>
                  </StackPanel>
                </Border>
                <TextBox Name='txtPSCode' Grid.Row='1' Style='{StaticResource CodeTextBox}'
                         Margin='2' MinHeight='120'
                         ToolTip='Enter PowerShell code to run remotely. $args[0] = ComputerName, $args[1] = PSCredential (if set)'/>
              </Grid>
            </Border>

            <!-- Host list panel -->
            <Border Grid.Column='2' Background='$($T.ControlBg)' CornerRadius='8' Padding='10'>
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height='Auto'/>
                  <RowDefinition Height='*'/>
                  <RowDefinition Height='Auto'/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row='0'>
                  <TextBlock Text='Target Hosts' Style='{StaticResource SectionLabel}'/>
                  <TextBlock Text='One per line. Blank = Global Target.' Foreground='$($T.TextMuted)' FontSize='10' Margin='0,0,0,4'/>
                </StackPanel>
                <TextBox Name='txtPSHosts' Grid.Row='1' Style='{StaticResource ModernTextBox}'
                         AcceptsReturn='True' VerticalScrollBarVisibility='Auto'
                         FontFamily='Consolas' FontSize='11'
                         ToolTip='Enter one hostname or IP per line. Leave blank to use Global Target.'/>
                <StackPanel Grid.Row='2' Margin='0,6,0,0'>
                  <TextBlock Text='Throttle Limit:' Style='{StaticResource SectionLabel}' Margin='0,0,0,2'/>
                  <TextBox Name='txtPSThrottle' Text='10' Width='60' Height='26' Style='{StaticResource ModernTextBox}'
                           HorizontalAlignment='Left' ToolTip='Max parallel jobs'/>
                </StackPanel>
              </Grid>
            </Border>
          </Grid>

          <!-- Divider -->
          <Rectangle Grid.Row='2' Fill='$($T.GridBorder)' Margin='0,2'/>

          <!-- Output panel -->
          <Border Grid.Row='3' Background='$($T.OutputBg)' CornerRadius='8' Padding='0'>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/>
                <RowDefinition Height='*'/>
              </Grid.RowDefinitions>
              <Border Grid.Row='0' Background='$($T.GridHeaderBg)' CornerRadius='8,8,0,0' Padding='10,6'>
                <StackPanel Orientation='Horizontal'>
                  <TextBlock Text='Output / Results' Style='{StaticResource SectionLabel}' Margin='0'/>
                  <TextBlock Name='txtPSStatus' Foreground='$($T.AccentCyan)' FontSize='11' VerticalAlignment='Center' Margin='16,0,0,0'/>
                </StackPanel>
              </Border>
              <TextBox Name='txtPSOutput' Grid.Row='1' Style='{StaticResource OutputTextBox}' Margin='2'/>
            </Grid>
          </Border>

          <TextBlock Name='txtPSCount' Grid.Row='4' Foreground='$($T.TextMuted)' FontSize='11' Margin='2,4,0,0'/>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 12: EXTERNAL TOOLS                       -->
      <!-- ============================================ -->
      <TabItem Header='Ext Tools'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
          </Grid.RowDefinitions>

          <TextBlock Grid.Row='0' Text='EXTERNAL TOOLS' Style='{StaticResource SectionLabel}' FontSize='13' Margin='0,6,0,8'/>

          <WrapPanel Grid.Row='1' Margin='0,0,0,12'>
            <!-- PSExec -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='PSExec' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='Sysinternals remote shell launcher. Runs commands on remote hosts without installing a client.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Target:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtPSExecHost' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,4'/>
                <Label Content='Command:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtPSExecCmd' Height='28' Style='{StaticResource ModernTextBox}' Text='cmd.exe' Margin='0,0,0,6'/>
                <Button Name='btnLaunchPSExec' Content='Launch PSExec' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- PAExec -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='PAExec' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='PowerAdmin remote exec. PSExec alternative with credential passthrough support.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Target:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtPAExecHost' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,4'/>
                <Label Content='Command:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtPAExecCmd' Height='28' Style='{StaticResource ModernTextBox}' Text='cmd.exe' Margin='0,0,0,6'/>
                <Button Name='btnLaunchPAExec' Content='Launch PAExec' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- AdExplorer -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='AD Explorer' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='Sysinternals Active Directory viewer. Browse and edit AD objects directly.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='DC / Domain:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtAdExplorerDC' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,10'/>
                <Button Name='btnLaunchAdExplorer' Content='Open AD Explorer' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- WMI Explorer -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='WMI Explorer' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='PowerShell-based WMI class browser. Explore all WMI namespaces and classes.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Target (optional):' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtWMIHost' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,10'/>
                <Button Name='btnLaunchWMI' Content='Open WMI Explorer' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- SYDI Server -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='SYDI Server' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='VBScript inventory tool. Generates full server documentation reports (Word format).' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Target:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtSydiHost' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,10'/>
                <Button Name='btnLaunchSydi' Content='Run SYDI Report' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- CMTrace -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='CMTrace' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='SCCM/MECM log viewer. Real-time log parsing with color-coded severity. Essential for server logs.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Log file (optional):' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtCMTraceFile' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,10'/>
                <Button Name='btnLaunchCMTrace' Content='Open CMTrace' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- RDP Launcher -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='RDP Launcher' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='Launch Remote Desktop connections. Supports admin mode, full screen, and custom resolution.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Label Content='Host:' Style='{StaticResource ModernLabel}' Padding='0' Margin='0,0,0,2'/>
                <TextBox Name='txtRDPHost' Height='28' Style='{StaticResource ModernTextBox}' Margin='0,0,0,4'/>
                <CheckBox Name='chkRDPAdmin' Content='Admin mode (/admin)' Style='{StaticResource ModernCheckBox}' Margin='0,0,0,6'/>
                <Button Name='btnLaunchRDP' Content='Connect RDP' Style='{StaticResource SuccessButton}' Height='28'/>
              </StackPanel>
            </Border>

            <!-- Tool Path Config -->
            <Border Background='$($T.ControlBg)' CornerRadius='8' Padding='12' Margin='0,0,10,10' Width='210'>
              <StackPanel>
                <TextBlock Text='Tool Paths' Style='{StaticResource HeaderLabel}' Margin='0,0,0,4'/>
                <TextBlock Text='Configure paths to external tools. Updates take effect immediately.' Foreground='$($T.TextMuted)' FontSize='11' TextWrapping='Wrap' Margin='0,0,0,8'/>
                <Button Name='btnConfigPaths' Content='Configure Tool Paths' Style='{StaticResource WarnButton}' Height='28' Margin='0,0,0,4'/>
                <Button Name='btnCheckTools' Content='Check Tool Status' Style='{StaticResource ModernButton}' Height='28'/>
              </StackPanel>
            </Border>
          </WrapPanel>

          <!-- Tool status output -->
          <Border Grid.Row='2' Background='$($T.OutputBg)' CornerRadius='8' Padding='8'>
            <TextBox Name='txtToolOutput' Style='{StaticResource OutputTextBox}'/>
          </Border>
        </Grid>
      </TabItem>

      <!-- ============================================ -->
      <!-- TAB 13: MULTI-HOST RUNNER (batch)            -->
      <!-- ============================================ -->
      <TabItem Header='Multi-Host Batch'>
        <Grid Background='$($T.PanelBg)' Margin='8'>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,6,0,8'>
            <TextBlock Text='MULTI-HOST BATCH RUNNER' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,16,0' FontSize='13'/>
            <Button Name='btnBatchImportHosts' Content='Import Host List...' Style='{StaticResource WarnButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnBatchImportScript' Content='Import Script...' Style='{StaticResource ModernButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnBatchRun' Content='Run Batch' Style='{StaticResource SuccessButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnBatchAbort' Content='Abort' Style='{StaticResource DangerButton}' Height='30' Margin='0,0,6,0'/>
            <Button Name='btnBatchExport' Content='Export Results' Style='{StaticResource ModernButton}' Height='30'/>
          </StackPanel>

          <Grid Grid.Row='1'>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width='220'/>
              <ColumnDefinition Width='8'/>
              <ColumnDefinition Width='*'/>
              <ColumnDefinition Width='8'/>
              <ColumnDefinition Width='*'/>
            </Grid.ColumnDefinitions>

            <!-- Host list -->
            <Border Grid.Column='0' Background='$($T.ControlBg)' CornerRadius='8' Padding='10'>
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height='Auto'/>
                  <RowDefinition Height='*'/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row='0'>
                  <TextBlock Text='Host List' Style='{StaticResource SectionLabel}'/>
                  <TextBlock Text='One per line' Foreground='$($T.TextMuted)' FontSize='10' Margin='0,0,0,4'/>
                </StackPanel>
                <TextBox Name='txtBatchHosts' Grid.Row='1' Style='{StaticResource ModernTextBox}'
                         AcceptsReturn='True' VerticalScrollBarVisibility='Auto'
                         FontFamily='Consolas' FontSize='11'/>
              </Grid>
            </Border>

            <!-- Script editor -->
            <Border Grid.Column='2' Background='$($T.CodeBg)' CornerRadius='8' Padding='10'>
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height='Auto'/>
                  <RowDefinition Height='*'/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row='0' Text='Script (runs on each host)' Style='{StaticResource SectionLabel}'/>
                <TextBox Name='txtBatchScript' Grid.Row='1' Style='{StaticResource CodeTextBox}'/>
              </Grid>
            </Border>

            <!-- Results grid -->
            <Border Grid.Column='4' Background='$($T.PanelBg)' CornerRadius='8' Padding='0'>
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height='Auto'/>
                  <RowDefinition Height='*'/>
                </Grid.RowDefinitions>
                <Border Grid.Row='0' Background='$($T.GridHeaderBg)' Padding='10,6' CornerRadius='8,8,0,0'>
                  <StackPanel Orientation='Horizontal'>
                    <TextBlock Text='Results' Style='{StaticResource SectionLabel}' Margin='0'/>
                    <TextBlock Name='txtBatchStatus' Foreground='$($T.AccentCyan)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
                  </StackPanel>
                </Border>
                <DataGrid Name='dgBatchResults' Grid.Row='1' Style='{StaticResource ModernDataGrid}'/>
              </Grid>
            </Border>
          </Grid>
        </Grid>
      </TabItem>

    </TabControl>

    <!-- ===== STATUS BAR ===== -->
    <Border Grid.Row='2' Background='$($T.StatusBg)' Padding='10,2'>
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='*'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>
        <TextBlock Name='txtStatus' Grid.Column='0' Foreground='$($T.TextSecondary)' FontSize='11' VerticalAlignment='Center' Text='Ready.'/>
        <TextBlock Name='txtCredStatus' Grid.Column='1' Foreground='$($T.AccentYellow)' FontSize='11' VerticalAlignment='Center' Margin='12,0' Text='No credentials stored'/>
        <TextBlock Grid.Column='2' Text='|' Foreground='$($T.TextMuted)' VerticalAlignment='Center' Margin='4,0'/>
        <TextBlock Name='txtDateTime' Grid.Column='3' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='4,0,0,0'/>
      </Grid>
    </Border>

  </Grid>
</Window>
"@
}

# =========================================================
#  BUILD WINDOW
# =========================================================
$T = if ($Script:IsDark) { $Script:Themes.Dark } else { $Script:Themes.Light }
[xml]$xamlDoc = Get-XAML -T $T
$reader = [System.Xml.XmlNodeReader]::new($xamlDoc)
$Window = [System.Windows.Markup.XamlReader]::Load($reader)

# ── Bind named controls ──────────────────────────────────
$Controls = @{}
$xamlDoc.SelectNodes('//*[@Name]') | ForEach-Object {
    $n = $_.Name
    $Controls[$n] = $Window.FindName($n)
}

# Shortcuts
$txtStatus      = $Controls['txtStatus']
$txtCredStatus  = $Controls['txtCredStatus']
$txtDateTime    = $Controls['txtDateTime']
$txtGlobalTarget= $Controls['txtGlobalTarget']

# ── Clock timer ──────────────────────────────────────────
$clock = [System.Windows.Threading.DispatcherTimer]::new()
$clock.Interval = [TimeSpan]::FromSeconds(1)
$clock.Add_Tick({ $txtDateTime.Text = (Get-Date).ToString('ddd yyyy-MM-dd HH:mm:ss') })
$clock.Start()

# =========================================================
#  STATUS HELPERS
# =========================================================
function Set-Status {
    param([string]$Msg, [string]$Color = "")
    $Window.Dispatcher.Invoke([action]{
        $txtStatus.Text = $Msg
        if ($Color) { $txtStatus.Foreground = $Color }
    })
}

function Get-ActiveTarget {
    param([string]$TabHost = "")
    $t = $TabHost.Trim()
    if (-not $t) { $t = $Controls['txtGlobalTarget'].Text.Trim() }
    if (-not $t) { $t = $env:COMPUTERNAME }
    return $t
}

function Update-CredStatus {
    $count = $Script:CredStore.Count
    if ($count -eq 0) { $txtCredStatus.Text = "No credentials stored" }
    else { $txtCredStatus.Text = "Creds stored: $count host(s)" }
}

# =========================================================
#  THEME TOGGLE
# =========================================================
$Controls['btnThemeToggle'].Add_Click({
    $Script:IsDark = -not $Script:IsDark
    $btnLabel = if ($Script:IsDark) { "Light Mode" } else { "Dark Mode" }
    $Controls['btnThemeToggle'].Content = $btnLabel
    Show-Msg "Theme change takes effect on next launch. Save any work before restarting." "Theme"
})

# =========================================================
#  CREDENTIAL MANAGEMENT
# =========================================================
$Controls['btnSetCred'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtGlobalTarget'].Text
    try {
        $cred = Get-Credential -Message "Enter credentials for: $target" -UserName "$env:USERDOMAIN\$env:USERNAME"
        if ($cred) {
            $Script:CredStore[$target] = $cred
            Update-CredStatus
            Set-Status "Credentials stored for: $target"
            Write-Log "Credentials stored for $target" -Level INFO
        }
    } catch { Set-Status "Credential entry cancelled." }
})

# =========================================================
#  OPEN LOG
# =========================================================
$Controls['btnOpenLog'].Add_Click({
    if (Test-Path $Script:LogPath) {
        $cmtrace = $Script:ExternalTools['CMTrace']
        if (Test-Path $cmtrace) {
            Start-Process $cmtrace -ArgumentList "`"$Script:LogPath`""
        } else {
            Start-Process notepad.exe -ArgumentList "`"$Script:LogPath`""
        }
    } else {
        Show-Msg "No log file found yet at:`n$Script:LogPath" "Log"
    }
})

# =========================================================
#  TAB 1: COMPUTER INFO
# =========================================================
$Controls['btnInfoQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Querying computer info for $target..."
    Write-Log "Computer Info query: $target"

    $sb = {
        param($comp, $crd)
        try {
            $os  = Get-WmiObject Win32_OperatingSystem -ComputerName $comp -ErrorAction Stop
            $cs  = Get-WmiObject Win32_ComputerSystem  -ComputerName $comp -ErrorAction Stop
            $cpu = Get-WmiObject Win32_Processor       -ComputerName $comp -ErrorAction SilentlyContinue | Select-Object -First 1
            $bios = Get-WmiObject Win32_BIOS            -ComputerName $comp -ErrorAction SilentlyContinue
            $mem = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
            $up  = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)

            @{
                Summary = [PSCustomObject]@{
                    OSName       = $os.Caption
                    OSVersion    = $os.Version
                    OSBuild      = $os.BuildNumber
                    Domain       = $cs.Domain
                    Manufacturer = $cs.Manufacturer
                    Model        = $cs.Model
                    TotalRAM_GB  = $mem
                    CPU          = if ($cpu) { $cpu.Name } else { "N/A" }
                    LogicalCores = if ($cpu) { $cpu.NumberOfLogicalProcessors } else { "N/A" }
                    UptimeDays   = [math]::Round($up.TotalDays, 2)
                    LastBoot     = $os.ConvertToDateTime($os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm')
                    BIOSVersion  = if ($bios) { $bios.SMBIOSBIOSVersion } else { "N/A" }
                    BIOSDate     = if ($bios) { $bios.ReleaseDate } else { "N/A" }
                }
                Disk = Get-WmiObject Win32_LogicalDisk -ComputerName $comp -Filter "DriveType=3" -ErrorAction SilentlyContinue |
                    Select-Object DeviceID,VolumeName,
                        @{n='Total_GB';e={[math]::Round($_.Size/1GB,2)}},
                        @{n='Free_GB';e={[math]::Round($_.FreeSpace/1GB,2)}},
                        @{n='PercentFree';e={if($_.Size -gt 0){[math]::Round(($_.FreeSpace/$_.Size)*100,1)}else{0}}},
                        FileSystem
                NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $comp -Filter "IPEnabled=True" -ErrorAction SilentlyContinue |
                    Select-Object Description,
                        @{n='IPv4';e={($_.IPAddress | Where-Object{$_ -notlike '*:*'}) -join ', '}},
                        @{n='Subnet';e={($_.IPSubnet | Where-Object{$_ -notlike '*:*'}) -join ', '}},
                        @{n='Gateway';e={$_.DefaultIPGateway -join ', '}},
                        @{n='DNS';e={$_.DNSServerSearchOrder -join ', '}},
                        MACAddress,DHCPEnabled
            }
        } catch {
            @{ Error = $_.Exception.Message }
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($result, $err)
        $Window.Dispatcher.Invoke([action]{
            if ($err -or ($result -and $result.Error)) {
                $msg = if ($err) { $err.Exception.Message } else { $result.Error }
                Set-Status "Error querying '$target': $msg"
                Write-Log "Error querying $target : $msg" -Level ERROR
                return
            }

            $s = $result.Summary
            $Controls['lblInfoOS'].Text    = "$($s.OSName)`nBuild: $($s.OSBuild)  |  $($s.OSVersion)"
            $Controls['lblInfoHW'].Text    = "$($s.Manufacturer) $($s.Model)`nBIOS: $($s.BIOSVersion)"
            $Controls['lblInfoCPU'].Text   = "$($s.TotalRAM_GB) GB RAM`n$($s.CPU) ($($s.LogicalCores) cores)"
            $Controls['lblInfoUptime'].Text = "Up: $($s.UptimeDays) days`nLast Boot: $($s.LastBoot)"

            $combined = @()
            if ($result.Disk)  { $combined += $result.Disk  }
            if ($result.NIC)   { $combined += $result.NIC   }
            $Controls['dgInfo'].ItemsSource = $combined
            $Controls['txtInfoCount'].Text  = "Disks: $($result.Disk.Count)  |  NICs: $($result.NIC.Count)  |  Domain: $($s.Domain)"
            Set-Status "Computer info loaded for $target"
            Write-Log "Computer Info OK: $target - $($s.OSName)"
        })
    }
})

$Controls['btnInfoLocal'].Add_Click({
    $Controls['txtInfoHost'].Text = $env:COMPUTERNAME
    $Controls['btnInfoQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})

$Controls['btnInfoRDP'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    try { Start-Process mstsc -ArgumentList "/v:$target" }
    catch { Set-Status "RDP launch failed: $_" }
})

$Controls['btnInfoPSExec'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    $exe    = $Script:ExternalTools['PSExec']
    if (-not (Test-Path $exe)) { Show-Msg "PSExec not found at:`n$exe`n`nUpdate path in External Tools tab." "PSExec"; return }
    try { Start-Process cmd.exe -ArgumentList "/K `"$exe`" \\$target cmd.exe" }
    catch { Set-Status "PSExec launch failed: $_" }
})

$Controls['btnInfoSysinfo'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    $out = ""
    try { $out = & systeminfo.exe /s $target 2>&1 | Out-String }
    catch { $out = "systeminfo failed: $_" }
    $Controls['txtToolOutput'].Text = $out
    Set-Status "SystemInfo completed for $target"
})

$Controls['btnInfoDrivers'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    $out = ""
    try { $out = & driverquery.exe /s $target /fo list 2>&1 | Out-String }
    catch { $out = "driverquery failed: $_" }
    $Controls['txtToolOutput'].Text = $out
    Set-Status "DriverQuery completed for $target"
})

$Controls['btnInfoSydi'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtInfoHost'].Text
    $exe    = $Script:ExternalTools['SydiServer']
    if (-not (Test-Path $exe)) { Show-Msg "sydi-server.vbs not found at:`n$exe`n`nUpdate path in External Tools tab." "SYDI"; return }
    try { Start-Process cscript.exe -ArgumentList "`"$exe`" -t:$target -w:y" }
    catch { Set-Status "SYDI launch failed: $_" }
})

$Controls['btnInfoExport'].Add_Click({
    $data = $Controls['dgInfo'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "ComputerInfo" }
})

# =========================================================
#  TAB 2: SERVICES
# =========================================================
$Controls['btnSvcQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtSvcHost'].Text
    $filter = $Controls['txtSvcFilter'].Text.Trim()
    $cred   = Get-Cred -Target $target
    Set-Status "Loading services from $target..."
    Write-Log "Services query: $target filter='$filter'"

    $sb = {
        param($comp, $flt, $crd)
        try {
            $params = @{ ComputerName=$comp; ScriptBlock={ Get-Service -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType | Sort-Object DisplayName }; ErrorAction='Stop' }
            if ($crd) { $params.Credential = $crd }
            $svc = Invoke-Command @params
            if ($flt) { $svc = $svc | Where-Object { $_.DisplayName -like "*$flt*" -or $_.Name -like "*$flt*" } }
            return $svc
        } catch { throw }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $filter, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Services error: $($e.Exception.Message)"; Write-Log "Services error '$target': $($e.Exception.Message)" -Level ERROR; return }
            $Controls['dgServices'].ItemsSource = $r
            $Controls['txtSvcCount'].Text = "$($r.Count) services on $target"
            Set-Status "Services loaded: $($r.Count) on $target"
            Write-Log "Services OK: $($r.Count) on $target"
        })
    }
})

$Controls['btnSvcStart'].Add_Click({
    $sel = $Controls['dgServices'].SelectedItem
    if (-not $sel) { Show-Msg "Select a service first." "Services"; return }
    $target = Get-ActiveTarget -TabHost $Controls['txtSvcHost'].Text
    $cred   = Get-Cred -Target $target
    try {
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Start-Service -Name '$($sel.Name)' -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Started service: $($sel.Name) on $target"
        Write-Log "Started service $($sel.Name) on $target"
        $Controls['btnSvcQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Start failed: $($_.Exception.Message)"; Write-Log "Start service failed $($sel.Name): $_" -Level ERROR }
})

$Controls['btnSvcStop'].Add_Click({
    $sel = $Controls['dgServices'].SelectedItem
    if (-not $sel) { Show-Msg "Select a service first." "Services"; return }
    $target = Get-ActiveTarget -TabHost $Controls['txtSvcHost'].Text
    $cred   = Get-Cred -Target $target
    $confirm = [System.Windows.MessageBox]::Show("Stop service '$($sel.DisplayName)' on $target?", "Confirm Stop", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    try {
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Stop-Service -Name '$($sel.Name)' -Force -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Stopped service: $($sel.Name) on $target"
        Write-Log "Stopped service $($sel.Name) on $target" -Level WARN
        $Controls['btnSvcQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Stop failed: $($_.Exception.Message)"; Write-Log "Stop service failed $($sel.Name): $_" -Level ERROR }
})

$Controls['btnSvcRestart'].Add_Click({
    $sel = $Controls['dgServices'].SelectedItem
    if (-not $sel) { Show-Msg "Select a service first." "Services"; return }
    $target = Get-ActiveTarget -TabHost $Controls['txtSvcHost'].Text
    $cred   = Get-Cred -Target $target
    $confirm = [System.Windows.MessageBox]::Show("Restart service '$($sel.DisplayName)' on $target?", "Confirm Restart", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    try {
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Restart-Service -Name '$($sel.Name)' -Force -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Restarted service: $($sel.Name) on $target"
        Write-Log "Restarted service $($sel.Name) on $target" -Level WARN
        $Controls['btnSvcQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Restart failed: $($_.Exception.Message)"; Write-Log "Restart service failed $($sel.Name): $_" -Level ERROR }
})

$Controls['btnSvcExport'].Add_Click({
    $data = $Controls['dgServices'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "Services" }
})

# =========================================================
#  TAB 3: PROCESSES
# =========================================================
$Controls['btnProcQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtProcHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading processes from $target..."
    Write-Log "Processes query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        Invoke-Command @params -ScriptBlock {
            Get-Process -ErrorAction SilentlyContinue |
                Select-Object Name,Id,CPU,
                    @{n='WorkingSet_MB';e={[math]::Round($_.WorkingSet64/1MB,1)}},
                    @{n='VirtualMem_MB';e={[math]::Round($_.VirtualMemorySize64/1MB,1)}},
                    @{n='Threads';e={$_.Threads.Count}},Description,
                    @{n='StartTime';e={if($_.StartTime){$_.StartTime.ToString('HH:mm:ss')}else{'N/A'}}} |
                Sort-Object WorkingSet_MB -Descending
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Processes error: $($e.Exception.Message)"; return }
            $Controls['dgProcesses'].ItemsSource = $r
            $Controls['txtProcCount'].Text = "$($r.Count) processes on $target"
            Set-Status "Processes loaded: $($r.Count) on $target"
            Write-Log "Processes OK: $($r.Count) on $target"
        })
    }
})

$Controls['btnProcKill'].Add_Click({
    $sel = $Controls['dgProcesses'].SelectedItem
    if (-not $sel) { Show-Msg "Select a process first." "Processes"; return }
    $target = Get-ActiveTarget -TabHost $Controls['txtProcHost'].Text
    $cred   = Get-Cred -Target $target
    $confirm = [System.Windows.MessageBox]::Show("Kill process '$($sel.Name)' (PID $($sel.Id)) on $target?`n`nWARNING: This will forcibly terminate the process.", "Confirm Kill", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    try {
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Stop-Process -Id $($sel.Id) -Force -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Killed process $($sel.Name) (PID $($sel.Id)) on $target"
        Write-Log "Killed process $($sel.Name) PID $($sel.Id) on $target" -Level WARN
        $Controls['btnProcQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Kill failed: $($_.Exception.Message)"; Write-Log "Kill process failed: $_" -Level ERROR }
})

$Controls['btnProcExport'].Add_Click({
    $data = $Controls['dgProcesses'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "Processes" }
})

# =========================================================
#  TAB 4: EVENT LOGS
# =========================================================
$Controls['btnEvtQuery'].Add_Click({
    $target  = Get-ActiveTarget -TabHost $Controls['txtEvtHost'].Text
    $logName = ($Controls['cmbEvtLog'].SelectedItem).Content
    $level   = ($Controls['cmbEvtLevel'].SelectedItem).Content
    $count   = [int]($Controls['txtEvtCount'].Text)
    $cred    = Get-Cred -Target $target
    Set-Status "Loading event log '$logName' from $target..."
    Write-Log "EventLog query: $target log='$logName' level='$level' count=$count"

    $sb = {
        param($comp, $crd, $log, $lvl, $cnt)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        $script = [scriptblock]::Create(@"
`$filter = @{ LogName='$log'; MaxEvents=$cnt }
if ('$lvl' -eq 'Error')   { `$filter.Level = 2 }
elseif ('$lvl' -eq 'Warning') { `$filter.Level = 3 }
elseif ('$lvl' -eq 'Info')    { `$filter.Level = 4 }
Get-WinEvent -FilterHashtable `$filter -ErrorAction SilentlyContinue |
    Select-Object TimeCreated,Id,LevelDisplayName,ProviderName,
        @{n='Message';e={`$_.Message -replace '[\r\n]+',' '}} |
    Sort-Object TimeCreated -Descending
"@)
        $params.ScriptBlock = $script
        Invoke-Command @params
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred, $logName, $level, $count) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "EventLog error: $($e.Exception.Message)"; Write-Log "EventLog error: $($e.Exception.Message)" -Level ERROR; return }
            $Controls['dgEvents'].ItemsSource = $r
            $Controls['txtEvtStatusCount'].Text = "$($r.Count) events from '$logName' on $target"
            Set-Status "Event log loaded: $($r.Count) events"
            Write-Log "EventLog OK: $($r.Count) events from $logName on $target"
        })
    }
})

$Controls['btnEvtExport'].Add_Click({
    $data = $Controls['dgEvents'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "EventLogs" }
})

# =========================================================
#  TAB 5: DISKS
# =========================================================
$Controls['btnDiskQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtDiskHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading disk info from $target..."
    Write-Log "Disks query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        Invoke-Command @params -ScriptBlock {
            Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue |
                Select-Object DeviceID,VolumeName,FileSystem,
                    @{n='Total_GB';e={[math]::Round($_.Size/1GB,2)}},
                    @{n='Used_GB';e={[math]::Round(($_.Size-$_.FreeSpace)/1GB,2)}},
                    @{n='Free_GB';e={[math]::Round($_.FreeSpace/1GB,2)}},
                    @{n='PercentFree';e={if($_.Size -gt 0){[math]::Round(($_.FreeSpace/$_.Size)*100,1)}else{0}}},
                    @{n='Health';e={if($_.Size -gt 0){if(($_.FreeSpace/$_.Size) -lt 0.1){'CRITICAL'}elseif(($_.FreeSpace/$_.Size) -lt 0.2){'LOW'}else{'OK'}}else{'N/A'}}}
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Disks error: $($e.Exception.Message)"; return }
            $Controls['dgDisks'].ItemsSource = $r
            $Controls['txtDiskCount'].Text = "$($r.Count) disk(s) on $target"
            Set-Status "Disk info loaded: $($r.Count) drives on $target"
            Write-Log "Disks OK: $($r.Count) drives on $target"
        })
    }
})

$Controls['btnDiskExport'].Add_Click({
    $data = $Controls['dgDisks'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "Disks" }
})

# =========================================================
#  TAB 6: SHARES
# =========================================================
$Controls['btnShareQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtShareHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading shares from $target..."
    Write-Log "Shares query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        Invoke-Command @params -ScriptBlock {
            Get-WmiObject Win32_Share -ErrorAction SilentlyContinue |
                Select-Object Name,Path,Description,
                    @{n='Type';e={switch($_.Type){0{'Disk'}1{'Printer'}2{'Device'}3{'IPC'}2147483648{'Admin Disk'}default{$_.Type}}}}
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Shares error: $($e.Exception.Message)"; return }
            $Controls['dgShares'].ItemsSource = $r
            $Controls['txtShareCount'].Text = "$($r.Count) share(s) on $target"
            Set-Status "Shares loaded: $($r.Count) on $target"
            Write-Log "Shares OK: $($r.Count) on $target"
        })
    }
})

$Controls['btnShareOpen'].Add_Click({
    $sel    = $Controls['dgShares'].SelectedItem
    $target = Get-ActiveTarget -TabHost $Controls['txtShareHost'].Text
    if (-not $sel) { Show-Msg "Select a share first." "Shares"; return }
    $unc = "\\$target\$($sel.Name)"
    try { Start-Process explorer.exe -ArgumentList $unc }
    catch { Set-Status "Could not open: $unc" }
})

$Controls['btnShareExport'].Add_Click({
    $data = $Controls['dgShares'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "Shares" }
})

# =========================================================
#  TAB 7: SCHEDULED TASKS
# =========================================================
$Controls['btnTaskQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtTaskHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading scheduled tasks from $target..."
    Write-Log "ScheduledTasks query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        Invoke-Command @params -ScriptBlock {
            Get-ScheduledTask -ErrorAction SilentlyContinue |
                Select-Object TaskName,TaskPath,State,
                    @{n='RunAs';e={$_.Principal.UserId}},
                    @{n='Actions';e={($_.Actions | ForEach-Object { $_.Execute }) -join '; '}} |
                Sort-Object TaskPath,TaskName
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Tasks error: $($e.Exception.Message)"; return }
            $Controls['dgTasks'].ItemsSource = $r
            $Controls['txtTaskCount'].Text = "$($r.Count) task(s) on $target"
            Set-Status "Tasks loaded: $($r.Count) on $target"
            Write-Log "Tasks OK: $($r.Count) on $target"
        })
    }
})

$Controls['btnTaskRun'].Add_Click({
    $sel    = $Controls['dgTasks'].SelectedItem
    $target = Get-ActiveTarget -TabHost $Controls['txtTaskHost'].Text
    if (-not $sel) { Show-Msg "Select a task first." "Tasks"; return }
    $cred   = Get-Cred -Target $target
    try {
        $taskName = $sel.TaskName
        $taskPath = $sel.TaskPath
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Start-ScheduledTask -TaskName '$taskName' -TaskPath '$taskPath' -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Started task: $taskName on $target"
        Write-Log "Ran task $taskName on $target"
    } catch { Set-Status "Run task failed: $($_.Exception.Message)"; Write-Log "Run task failed: $_" -Level ERROR }
})

$Controls['btnTaskEnable'].Add_Click({
    $sel    = $Controls['dgTasks'].SelectedItem
    $target = Get-ActiveTarget -TabHost $Controls['txtTaskHost'].Text
    if (-not $sel) { Show-Msg "Select a task first." "Tasks"; return }
    $cred = Get-Cred -Target $target
    try {
        $taskName = $sel.TaskName
        $taskPath = $sel.TaskPath
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Enable-ScheduledTask -TaskName '$taskName' -TaskPath '$taskPath' -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Enabled task: $taskName"
        Write-Log "Enabled task $taskName on $target"
        $Controls['btnTaskQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Enable task failed: $($_.Exception.Message)" }
})

$Controls['btnTaskDisable'].Add_Click({
    $sel    = $Controls['dgTasks'].SelectedItem
    $target = Get-ActiveTarget -TabHost $Controls['txtTaskHost'].Text
    if (-not $sel) { Show-Msg "Select a task first." "Tasks"; return }
    $cred = Get-Cred -Target $target
    $confirm = [System.Windows.MessageBox]::Show("Disable task '$($sel.TaskName)' on $target?", "Confirm", 'YesNo', 'Warning')
    if ($confirm -ne 'Yes') { return }
    try {
        $taskName = $sel.TaskName
        $taskPath = $sel.TaskPath
        $params = @{ ComputerName=$target; ScriptBlock=[scriptblock]::Create("Disable-ScheduledTask -TaskName '$taskName' -TaskPath '$taskPath' -ErrorAction Stop"); ErrorAction='Stop' }
        if ($cred) { $params.Credential = $cred }
        Invoke-Command @params
        Set-Status "Disabled task: $taskName"
        Write-Log "Disabled task $taskName on $target" -Level WARN
        $Controls['btnTaskQuery'].RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    } catch { Set-Status "Disable task failed: $($_.Exception.Message)" }
})

$Controls['btnTaskExport'].Add_Click({
    $data = $Controls['dgTasks'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "ScheduledTasks" }
})

# =========================================================
#  TAB 8: NETWORK
# =========================================================
$Controls['btnNetAdapters'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtNetHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading NIC info from $target..."
    Write-Log "Network adapters query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        Invoke-Command @params -ScriptBlock {
            $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue
            $adapters | Select-Object Description,MACAddress,
                @{n='IPv4';e={($_.IPAddress | Where-Object{$_ -notlike '*:*'}) -join ', '}},
                @{n='Subnet';e={($_.IPSubnet | Where-Object{$_ -notlike '*:*'}) -join ', '}},
                @{n='Gateway';e={$_.DefaultIPGateway -join ', '}},
                @{n='DNS';e={$_.DNSServerSearchOrder -join ', '}},
                DHCPEnabled,
                @{n='DHCPServer';e={$_.DHCPServer}},
                @{n='DHCPLeaseObtained';e={if($_.DHCPLeaseObtained){$_.ConvertToDateTime($_.DHCPLeaseObtained)}}}
        }
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "NIC error: $($e.Exception.Message)"; return }
            $Controls['dgNetwork'].ItemsSource = $r
            $Controls['txtNetCount'].Text = "$($r.Count) NIC(s) on $target"
            Set-Status "NIC info loaded: $($r.Count) adapters on $target"
            Write-Log "Network OK: $($r.Count) NICs on $target"
        })
    }
})

$Controls['btnNetTest'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtNetHost'].Text
    $portStr = $Controls['txtNetPorts'].Text.Trim()
    $ports = @(135, 445, 3389, 5985)
    if ($portStr) {
        try { $ports += ($portStr -split ',' | ForEach-Object { [int]$_.Trim() }) } catch {}
    }
    $ports = $ports | Select-Object -Unique | Sort-Object
    Set-Status "Running reachability test against $target..."
    Write-Log "Network reachability test: $target ports=$($ports -join ',')"

    $sb = {
        param($comp, $portList)
        $results = [System.Collections.ArrayList]@()
        $ping = Test-Connection -ComputerName $comp -Count 2 -ErrorAction SilentlyContinue
        $null = $results.Add([PSCustomObject]@{
            Test   = "ICMP Ping"
            Port   = "N/A"
            Result = if ($ping) { "REACHABLE  ($([math]::Round(($ping | Measure-Object ResponseTime -Average).Average,0))ms avg)" } else { "UNREACHABLE" }
            Note   = "Basic connectivity"
        })
        foreach ($port in $portList) {
            $tcp = $null
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $conn = $tcp.BeginConnect($comp, $port, $null, $null)
                $wait = $conn.AsyncWaitHandle.WaitOne(2000, $false)
                if ($wait) { $tcp.EndConnect($conn); $status = "OPEN" } else { $status = "TIMEOUT" }
            } catch { $status = "CLOSED" } finally { if ($tcp) { $tcp.Close() } }
            $note = switch ([int]$port) {
                135   {"WMI / DCOM / RPC"}; 139 {"NetBIOS Session"}; 445 {"SMB / File Sharing"}
                389   {"LDAP"}; 636 {"LDAPS"}; 3389 {"Remote Desktop (RDP)"}
                5985  {"WinRM HTTP (PSRemoting)"}; 5986 {"WinRM HTTPS"}
                22    {"SSH"}; 80 {"HTTP"}; 443 {"HTTPS"}
                1433  {"SQL Server"}; 1434 {"SQL Browser"}
                8080  {"HTTP Alt"}; 8443 {"HTTPS Alt"}
                default {""}
            }
            $null = $results.Add([PSCustomObject]@{ Test="TCP"; Port=$port; Result=$status; Note=$note })
        }
        return $results
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, @($ports)) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Network test error: $($e.Exception.Message)"; return }
            $Controls['dgNetwork'].ItemsSource = $r
            $Controls['txtNetCount'].Text = "$($r.Count) tests against $target"
            Set-Status "Reachability test complete for $target"
            Write-Log "Network test complete for $target - $($r.Count) tests"
        })
    }
})

$Controls['btnDnsLookup'].Add_Click({
    $name = $Controls['txtDnsLookup'].Text.Trim()
    if (-not $name) { Show-Msg "Enter a hostname or IP to resolve." "DNS"; return }
    try {
        $result = [System.Net.Dns]::GetHostEntry($name)
        $addrs  = $result.AddressList | ForEach-Object { $_.ToString() }
        $out    = "Name   : $($result.HostName)`nAddrs  : $($addrs -join ', ')"
        $Controls['txtNetCount'].Text = $out
        Set-Status "DNS resolved: $name -> $($addrs -join ', ')"
        Write-Log "DNS lookup $name -> $($addrs -join ', ')"
    } catch {
        $Controls['txtNetCount'].Text = "DNS resolution failed for '$name': $($_.Exception.Message)"
        Set-Status "DNS lookup failed: $name"
    }
})

$Controls['btnTracert'].Add_Click({
    $target = $Controls['txtTracert'].Text.Trim()
    if (-not $target) { Show-Msg "Enter a target for traceroute." "Tracert"; return }
    Set-Status "Tracing route to $target (may take time)..."
    Write-Log "Tracert to $target"

    $sb = { param($t); $out = tracert.exe $t 2>&1; return ($out | Out-String) }
    Invoke-Async -ScriptBlock $sb -ArgumentList @($target) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            $txt = if ($e) { "Tracert error: $($e.Exception.Message)" } else { $r }
            $Controls['txtNetCount'].Text = $txt
            Set-Status "Traceroute complete to $target"
        })
    }
})

$Controls['btnNetstat'].Add_Click({
    Set-Status "Running netstat locally..."
    $sb = { $out = netstat.exe -ano 2>&1; return ($out | Out-String) }
    Invoke-Async -ScriptBlock $sb -ArgumentList @() -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            $txt = if ($e) { "Netstat error: $($e.Exception.Message)" } else { $r }
            $Controls['txtNetCount'].Text = $txt
            Set-Status "Netstat complete"
        })
    }
})

$Controls['btnNetExport'].Add_Click({
    $data = $Controls['dgNetwork'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "Network" }
})

# =========================================================
#  TAB 9: SOFTWARE
# =========================================================
$Controls['btnSwQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtSwHost'].Text
    $filter = $Controls['txtSwFilter'].Text.Trim()
    $cred   = Get-Cred -Target $target
    Set-Status "Loading installed software from $target..."
    Write-Log "Software query: $target filter='$filter'"

    $sb = {
        param($comp, $flt, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        $script = [scriptblock]::Create(@'
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$sw = $regPaths | ForEach-Object {
    Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.DisplayName.Trim() } |
        Select-Object DisplayName,DisplayVersion,Publisher,InstallDate,
            @{n='Size_MB';e={if($_.EstimatedSize){[math]::Round($_.EstimatedSize/1024,1)}else{"N/A"}}}
}
$sw | Sort-Object DisplayName -Unique
'@)
        $params.ScriptBlock = $script
        $result = Invoke-Command @params
        if ($flt) { $result = $result | Where-Object { $_.DisplayName -like "*$flt*" -or $_.Publisher -like "*$flt*" } }
        return $result
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $filter, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "Software error: $($e.Exception.Message)"; return }
            $Controls['dgSoftware'].ItemsSource = $r
            $Controls['txtSwCount'].Text = "$($r.Count) applications on $target"
            Set-Status "Software list loaded: $($r.Count) on $target"
            Write-Log "Software OK: $($r.Count) apps on $target"
        })
    }
})

$Controls['btnSwExport'].Add_Click({
    $data = $Controls['dgSoftware'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "InstalledSoftware" }
})

# =========================================================
#  TAB 10: WINDOWS UPDATES
# =========================================================
$Controls['btnUpdQuery'].Add_Click({
    $target = Get-ActiveTarget -TabHost $Controls['txtUpdHost'].Text
    $cred   = Get-Cred -Target $target
    Set-Status "Loading Windows Update history from $target..."
    Write-Log "WU history query: $target"

    $sb = {
        param($comp, $crd)
        $params = @{ ComputerName=$comp; ErrorAction='Stop' }
        if ($crd) { $params.Credential = $crd }
        $script = {
            try {
                $sess     = New-Object -ComObject Microsoft.Update.Session
                $searcher = $sess.CreateUpdateSearcher()
                $cnt      = $searcher.GetTotalHistoryCount()
                if ($cnt -gt 0) {
                    $history = $searcher.QueryHistory(0, [math]::Min($cnt, 100))
                    $history | ForEach-Object {
                        [PSCustomObject]@{
                            Title      = $_.Title
                            Date       = $_.Date.ToString('yyyy-MM-dd HH:mm')
                            Result     = switch($_.ResultCode){1{'InProgress'}2{'Succeeded'}3{'SucceededErrors'}4{'Failed'}5{'Aborted'}default{$_.ResultCode}}
                            KB         = if($_.Title -match 'KB(\d+)'){$matches[0]}else{'N/A'}
                            UpdateType = $_.HResult
                        }
                    } | Sort-Object Date -Descending
                }
            } catch { throw }
        }
        $params.ScriptBlock = $script
        return Invoke-Command @params
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($target, $cred) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) { Set-Status "WU error: $($e.Exception.Message)"; return }
            $Controls['dgUpdates'].ItemsSource = $r
            $Controls['txtUpdCount'].Text = "$($r.Count) update records on $target"
            Set-Status "Windows Update history loaded: $($r.Count) records"
            Write-Log "WU OK: $($r.Count) records on $target"
        })
    }
})

$Controls['btnUpdExport'].Add_Click({
    $data = $Controls['dgUpdates'].ItemsSource
    if ($data) { Export-GridData -Data @($data) -Category "WindowsUpdates" }
})

# =========================================================
#  TAB 11: REMOTE PS RUNNER
# =========================================================
# Default sample code
$Controls['txtPSCode'].Text = @'
# Remote PowerShell Code Runner
# This code runs on each target host listed in the Host List panel.
# Use $env:COMPUTERNAME to reference the current host in your code.
#
# Example: Get disk usage
Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" |
    Select-Object DeviceID,
        @{n='Total_GB';e={[math]::Round($_.Size/1GB,2)}},
        @{n='Free_GB';e={[math]::Round($_.FreeSpace/1GB,2)}},
        @{n='PercentFree';e={if($_.Size -gt 0){[math]::Round(($_.FreeSpace/$_.Size)*100,1)}else{0}}}
'@

$Controls['btnPSRun'].Add_Click({
    $code    = $Controls['txtPSCode'].Text.Trim()
    $hostTxt = $Controls['txtPSHosts'].Text.Trim()
    if (-not $code) { Show-Msg "Enter PowerShell code to run." "PS Runner"; return }

    $targets = @()
    if ($hostTxt) {
        $targets = ($hostTxt -split '[\r\n]+') | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    }
    if ($targets.Count -eq 0) {
        $targets = @((Get-ActiveTarget))
    }

    $throttle = 10
    try { $throttle = [int]$Controls['txtPSThrottle'].Text } catch {}

    $Controls['txtPSStatus'].Text = "Running on $($targets.Count) host(s)..."
    $Controls['txtPSOutput'].Text = ""
    Set-Status "Remote PS: executing on $($targets.Count) target(s)..."
    Write-Log "Remote PS Run: $($targets.Count) targets, throttle=$throttle"

    $Script:PSRunResults = $null
    $sb = {
        param($tgts, $code, $throttle)
        $results = [System.Collections.ArrayList]@()
        $jobs    = [System.Collections.ArrayList]@()

        foreach ($comp in $tgts) {
            $cred = $null
            try {
                $scriptBlock = [ScriptBlock]::Create($code)
                $jobParams   = @{ ComputerName=$comp; ScriptBlock=$scriptBlock; AsJob=$true; JobName="AIO_$comp"; ErrorAction='Stop' }
                if ($cred) { $jobParams.Credential = $cred }
                $job = Invoke-Command @jobParams
                $null = $jobs.Add([PSCustomObject]@{ Computer=$comp; Job=$job; Start=(Get-Date) })
            } catch {
                $null = $results.Add([PSCustomObject]@{
                    ComputerName = $comp
                    Status       = "LAUNCH_ERROR"
                    Duration     = "N/A"
                    Output       = $_.Exception.Message
                })
            }
        }

        $deadline = (Get-Date).AddSeconds(180)
        while (($jobs | Where-Object { $_.Job.State -in @('Running','NotStarted') }).Count -gt 0) {
            Start-Sleep -Milliseconds 300
            if ((Get-Date) -gt $deadline) { break }
        }

        foreach ($item in $jobs) {
            $dur = [math]::Round(((Get-Date) - $item.Start).TotalSeconds, 2)
            try {
                $out = Receive-Job -Job $item.Job -ErrorAction Stop
                $outStr = if ($out) { ($out | Out-String).Trim() } else { "(no output)" }
                $null = $results.Add([PSCustomObject]@{
                    ComputerName = $item.Computer
                    Status       = "OK"
                    Duration     = "$dur`s"
                    Output       = $outStr
                })
            } catch {
                $null = $results.Add([PSCustomObject]@{
                    ComputerName = $item.Computer
                    Status       = "ERROR"
                    Duration     = "$dur`s"
                    Output       = $_.Exception.Message
                })
            }
            Remove-Job -Job $item.Job -Force -ErrorAction SilentlyContinue
        }
        return $results
    }

    Invoke-Async -ScriptBlock $sb -ArgumentList @($targets, $code, $throttle) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) {
                $Controls['txtPSOutput'].Text = "FATAL ERROR: $($e.Exception.Message)"
                $Controls['txtPSStatus'].Text = "Failed"
                Set-Status "Remote PS error: $($e.Exception.Message)"
                Write-Log "Remote PS fatal error: $($e.Exception.Message)" -Level ERROR
                return
            }
            $Script:PSRunResults = $r
            $ok    = ($r | Where-Object { $_.Status -eq 'OK' }).Count
            $err   = ($r | Where-Object { $_.Status -ne 'OK' }).Count
            $outText = ($r | ForEach-Object {
                $sep = "=" * 60
                "$sep`n[$($_.Status)] $($_.ComputerName)  ($($_.Duration))`n$sep`n$($_.Output)`n"
            }) -join "`n"
            $Controls['txtPSOutput'].Text   = $outText
            $Controls['txtPSStatus'].Text   = "Done: $ok OK, $err errors"
            $Controls['txtPSCount'].Text    = "Results: $($r.Count) hosts | $ok succeeded | $err failed"
            Set-Status "Remote PS complete: $ok/$($r.Count) succeeded"
            Write-Log "Remote PS complete: $ok OK, $err errors across $($r.Count) hosts"
        })
    }
})

$Controls['btnPSClear'].Add_Click({ $Controls['txtPSCode'].Text = "" })
$Controls['btnPSClearOut'].Add_Click({ $Controls['txtPSOutput'].Text = ""; $Controls['txtPSStatus'].Text = "" })

$Controls['btnPSImport'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
    $dlg.Title  = "Import PowerShell Script"
    if ($dlg.ShowDialog()) {
        try {
            $Controls['txtPSCode'].Text = Get-Content -Path $dlg.FileName -Raw
            Set-Status "Script loaded: $($dlg.FileName)"
            Write-Log "Script imported: $($dlg.FileName)"
        } catch { Show-Msg "Failed to load: $($_.Exception.Message)" "Import" }
    }
})

$Controls['btnPSImportHosts'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.Title  = "Import Host List"
    if ($dlg.ShowDialog()) {
        try {
            $hosts = Get-Content -Path $dlg.FileName -Raw
            $Controls['txtPSHosts'].Text = $hosts
            $count = ($hosts -split '[\r\n]+' | Where-Object { $_.Trim() }).Count
            Set-Status "Host list loaded: $count hosts from $($dlg.FileName)"
            Write-Log "Host list imported: $count hosts from $($dlg.FileName)"
        } catch { Show-Msg "Failed to load: $($_.Exception.Message)" "Import" }
    }
})

$Controls['btnPSSave'].Add_Click({
    $code = $Controls['txtPSCode'].Text
    if (-not $code) { Show-Msg "Nothing to save." "Save"; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter   = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
    $dlg.FileName = "RemoteScript_$(Get-Date -Format 'yyyyMMdd_HHmm').ps1"
    if ($dlg.ShowDialog()) {
        try {
            $code | Set-Content -Path $dlg.FileName -Encoding UTF8
            Set-Status "Script saved: $($dlg.FileName)"
            Write-Log "Script saved: $($dlg.FileName)"
        } catch { Show-Msg "Save failed: $($_.Exception.Message)" "Save" }
    }
})

$Controls['btnPSExport'].Add_Click({
    if (-not $Script:PSRunResults -or $Script:PSRunResults.Count -eq 0) { Show-Msg "No results to export. Run code first." "Export"; return }
    Export-GridData -Data @($Script:PSRunResults) -Category "PSRunner_Results"
})

# =========================================================
#  TAB 12: EXTERNAL TOOLS
# =========================================================
$Controls['btnLaunchRDP'].Add_Click({
    $target = $Controls['txtRDPHost'].Text.Trim()
    if (-not $target) { $target = Get-ActiveTarget }
    $adminFlag = if ($Controls['chkRDPAdmin'].IsChecked) { "/admin" } else { "" }
    try {
        Start-Process mstsc.exe -ArgumentList "/v:$target $adminFlag"
        Set-Status "RDP launched to $target"
        Write-Log "RDP launched: $target admin=$($Controls['chkRDPAdmin'].IsChecked)"
    } catch { Set-Status "RDP launch failed: $_" }
})

$Controls['btnLaunchPSExec'].Add_Click({
    $target = $Controls['txtPSExecHost'].Text.Trim()
    $cmd    = $Controls['txtPSExecCmd'].Text.Trim()
    if (-not $target) { $target = Get-ActiveTarget }
    if (-not $cmd) { $cmd = "cmd.exe" }
    $exe    = $Script:ExternalTools['PSExec']
    if (-not (Test-Path $exe)) { Show-Msg "PSExec not found at:`n$exe`n`nUpdate the tool path." "PSExec"; return }
    try {
        Start-Process cmd.exe -ArgumentList "/K `"$exe`" \\$target $cmd"
        Set-Status "PSExec launched: $target -> $cmd"
        Write-Log "PSExec: $target $cmd"
    } catch { Set-Status "PSExec launch failed: $_" }
})

$Controls['btnLaunchPAExec'].Add_Click({
    $target = $Controls['txtPAExecHost'].Text.Trim()
    $cmd    = $Controls['txtPAExecCmd'].Text.Trim()
    if (-not $target) { $target = Get-ActiveTarget }
    if (-not $cmd) { $cmd = "cmd.exe" }
    $exe    = $Script:ExternalTools['PAExec']
    if (-not (Test-Path $exe)) { Show-Msg "PAExec not found at:`n$exe`n`nUpdate the tool path." "PAExec"; return }
    try {
        Start-Process cmd.exe -ArgumentList "/K `"$exe`" \\$target $cmd"
        Set-Status "PAExec launched: $target -> $cmd"
        Write-Log "PAExec: $target $cmd"
    } catch { Set-Status "PAExec launch failed: $_" }
})

$Controls['btnLaunchAdExplorer'].Add_Click({
    $dc  = $Controls['txtAdExplorerDC'].Text.Trim()
    $exe = $Script:ExternalTools['AdExplorer']
    if (-not (Test-Path $exe)) { Show-Msg "ADExplorer not found at:`n$exe`n`nUpdate the tool path." "ADExplorer"; return }
    try {
        $args = if ($dc) { "-target $dc" } else { "" }
        Start-Process $exe -ArgumentList $args
        Set-Status "ADExplorer launched"
        Write-Log "ADExplorer launched: $dc"
    } catch { Set-Status "ADExplorer launch failed: $_" }
})

$Controls['btnLaunchWMI'].Add_Click({
    $target = $Controls['txtWMIHost'].Text.Trim()
    $script = $Script:ExternalTools['WMIExplorer']
    if (-not (Test-Path $script)) { Show-Msg "WMIExplorer.ps1 not found at:`n$script`n`nUpdate the tool path." "WMIExplorer"; return }
    try {
        if ($target) {
            Start-Process powershell.exe -ArgumentList "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$script`" -ComputerName $target"
        } else {
            Start-Process powershell.exe -ArgumentList "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$script`""
        }
        Set-Status "WMI Explorer launched"
        Write-Log "WMIExplorer launched: $target"
    } catch { Set-Status "WMI Explorer launch failed: $_" }
})

$Controls['btnLaunchSydi'].Add_Click({
    $target = $Controls['txtSydiHost'].Text.Trim()
    $script = $Script:ExternalTools['SydiServer']
    if (-not (Test-Path $script)) { Show-Msg "sydi-server.vbs not found at:`n$script`n`nUpdate the tool path." "SYDI"; return }
    try {
        $outDir = Join-Path $env:USERPROFILE "Desktop\SYDI_Reports"
        if (!(Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }
        Start-Process cscript.exe -ArgumentList "`"$script`" -t:$target -w:y -o:`"$outDir\sydi_$target.doc`""
        Set-Status "SYDI report started for $target -> $outDir"
        Write-Log "SYDI: $target -> $outDir"
    } catch { Set-Status "SYDI launch failed: $_" }
})

$Controls['btnLaunchCMTrace'].Add_Click({
    $logFile = $Controls['txtCMTraceFile'].Text.Trim()
    $exe     = $Script:ExternalTools['CMTrace']
    if (-not (Test-Path $exe)) {
        # Fallback: search common SCCM paths
        $fallbacks = @("C:\Windows\CCM\CMTrace.exe","C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\CMTrace.exe","C:\Tools\CMTrace.exe")
        foreach ($fb in $fallbacks) {
            if (Test-Path $fb) { $exe = $fb; break }
        }
    }
    if (-not (Test-Path $exe)) { Show-Msg "CMTrace not found. Specify path in Config or place at:`nC:\Tools\CMTrace.exe" "CMTrace"; return }
    try {
        if ($logFile -and (Test-Path $logFile)) {
            Start-Process $exe -ArgumentList "`"$logFile`""
        } else {
            Start-Process $exe
        }
        Set-Status "CMTrace launched"
        Write-Log "CMTrace launched: $logFile"
    } catch { Set-Status "CMTrace launch failed: $_" }
})

$Controls['btnCheckTools'].Add_Click({
    $lines = [System.Collections.ArrayList]@()
    $null = $lines.Add("External Tool Status Check - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $lines.Add("=" * 60)
    foreach ($key in $Script:ExternalTools.Keys | Sort-Object) {
        $path   = $Script:ExternalTools[$key]
        $exists = if ($path -match '^[a-z]:\\') { Test-Path $path } else { $null -ne (Get-Command $path -ErrorAction SilentlyContinue) }
        $status = if ($exists) { "[OK]      " } else { "[MISSING] " }
        $null = $lines.Add("$status  $key`n           Path: $path")
    }
    $Controls['txtToolOutput'].Text = $lines -join "`n"
    Write-Log "Tool status check completed"
})

$Controls['btnConfigPaths'].Add_Click({
    $msg = "Current tool paths (edit in script `$Script:ExternalTools):`n`n"
    foreach ($k in $Script:ExternalTools.Keys | Sort-Object) {
        $msg += "  $k`n    -> $($Script:ExternalTools[$k])`n"
    }
    $msg += "`nTo update paths permanently, edit the `$Script:ExternalTools hashtable at the top of the script."
    Show-Msg $msg "Tool Path Configuration"
})

# =========================================================
#  TAB 13: MULTI-HOST BATCH
# =========================================================
$Script:BatchAbort = $false

$Controls['btnBatchImportHosts'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.Title  = "Import Host List"
    if ($dlg.ShowDialog()) {
        try {
            $Controls['txtBatchHosts'].Text = Get-Content -Path $dlg.FileName -Raw
            $count = (($Controls['txtBatchHosts'].Text) -split '[\r\n]+' | Where-Object { $_.Trim() }).Count
            Set-Status "Batch: $count hosts loaded"
            Write-Log "Batch host list: $count hosts from $($dlg.FileName)"
        } catch { Show-Msg "Load failed: $($_.Exception.Message)" "Import" }
    }
})

$Controls['btnBatchImportScript'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*"
    $dlg.Title  = "Import Script"
    if ($dlg.ShowDialog()) {
        try {
            $Controls['txtBatchScript'].Text = Get-Content -Path $dlg.FileName -Raw
            Set-Status "Batch: script loaded from $($dlg.FileName)"
            Write-Log "Batch script: loaded from $($dlg.FileName)"
        } catch { Show-Msg "Load failed: $($_.Exception.Message)" "Import" }
    }
})

$Controls['btnBatchRun'].Add_Click({
    $hostTxt = $Controls['txtBatchHosts'].Text.Trim()
    $code    = $Controls['txtBatchScript'].Text.Trim()
    if (-not $hostTxt) { Show-Msg "Enter at least one host in the Host List." "Batch"; return }
    if (-not $code)    { Show-Msg "Enter a script to run." "Batch"; return }

    $targets = ($hostTxt -split '[\r\n]+') | Where-Object { $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }
    $Script:BatchAbort = $false
    $Controls['txtBatchStatus'].Text = "Running on $($targets.Count) hosts..."
    Set-Status "Batch: executing on $($targets.Count) hosts..."
    Write-Log "Batch run: $($targets.Count) hosts"

    $sb = {
        param($tgts, $code)
        $results  = [System.Collections.ArrayList]@()
        $jobs     = [System.Collections.ArrayList]@()
        $sb2      = [ScriptBlock]::Create($code)

        foreach ($comp in $tgts) {
            try {
                $job = Invoke-Command -ComputerName $comp -ScriptBlock $sb2 -AsJob -JobName "Batch_$comp" -ErrorAction Stop
                $null = $jobs.Add([PSCustomObject]@{ Computer=$comp; Job=$job; Start=(Get-Date) })
            } catch {
                $null = $results.Add([PSCustomObject]@{ ComputerName=$comp; Status="LAUNCH_ERROR"; Duration="N/A"; Output=$_.Exception.Message })
            }
        }

        $deadline = (Get-Date).AddSeconds(300)
        while (($jobs | Where-Object { $_.Job.State -in @('Running','NotStarted') }).Count -gt 0) {
            Start-Sleep -Milliseconds 500
            if ((Get-Date) -gt $deadline) { break }
        }

        foreach ($item in $jobs) {
            $dur = [math]::Round(((Get-Date) - $item.Start).TotalSeconds, 2)
            try {
                $out    = Receive-Job -Job $item.Job -ErrorAction Stop
                $outStr = if ($out) { ($out | Out-String).Trim() } else { "(no output)" }
                $null   = $results.Add([PSCustomObject]@{ ComputerName=$item.Computer; Status="OK"; Duration="${dur}s"; Output=$outStr })
            } catch {
                $null   = $results.Add([PSCustomObject]@{ ComputerName=$item.Computer; Status="ERROR"; Duration="${dur}s"; Output=$_.Exception.Message })
            }
            Remove-Job -Job $item.Job -Force -ErrorAction SilentlyContinue
        }
        return $results
    }

    $Script:BatchResults = $null
    Invoke-Async -ScriptBlock $sb -ArgumentList @($targets, $code) -CompletedCallback {
        param($r, $e)
        $Window.Dispatcher.Invoke([action]{
            if ($e) {
                $Controls['txtBatchStatus'].Text = "ERROR: $($e.Exception.Message)"
                Set-Status "Batch error: $($e.Exception.Message)"
                Write-Log "Batch error: $($e.Exception.Message)" -Level ERROR
                return
            }
            $Script:BatchResults = $r
            $ok  = ($r | Where-Object { $_.Status -eq 'OK' }).Count
            $err = ($r | Where-Object { $_.Status -ne 'OK' }).Count
            $Controls['dgBatchResults'].ItemsSource = $r
            $Controls['txtBatchStatus'].Text = "Done: $ok OK, $err errors  |  $($r.Count) total"
            Set-Status "Batch complete: $ok/$($r.Count) succeeded"
            Write-Log "Batch complete: $ok OK, $err errors"
        })
    }
})

$Controls['btnBatchAbort'].Add_Click({
    $Script:BatchAbort = $true
    Get-Job -Name "Batch_*" -ErrorAction SilentlyContinue | Stop-Job -PassThru | Remove-Job -Force
    $Controls['txtBatchStatus'].Text = "ABORTED by user"
    Set-Status "Batch aborted."
    Write-Log "Batch aborted by user" -Level WARN
})

$Controls['btnBatchExport'].Add_Click({
    if (-not $Script:BatchResults -or $Script:BatchResults.Count -eq 0) { Show-Msg "No results to export." "Export"; return }
    Export-GridData -Data @($Script:BatchResults) -Category "Batch_Results"
})

# =========================================================
#  GLOBAL TARGET PROPAGATION
# =========================================================
$Controls['txtGlobalTarget'].Add_TextChanged({
    $val = $Controls['txtGlobalTarget'].Text
    # Propagate to all host text boxes
    $hostFields = @('txtInfoHost','txtSvcHost','txtProcHost','txtEvtHost','txtDiskHost',
                    'txtShareHost','txtTaskHost','txtNetHost','txtSwHost','txtUpdHost',
                    'txtRDPHost','txtPSExecHost','txtPAExecHost')
    foreach ($f in $hostFields) {
        $ctl = $Controls[$f]
        if ($ctl -and $ctl.Text -eq "") { }  # don't overwrite if user already typed something
    }
    Update-CredStatus
})

# =========================================================
#  INITIAL SETUP
# =========================================================
Write-Log "=== $Script:AppName v$Script:Version Started === User: $env:USERNAME | Host: $env:COMPUTERNAME"

if ($InitialComputer) {
    $Controls['txtGlobalTarget'].Text = $InitialComputer
}

Set-Status "Ready. Set a target host in the header bar or on each tab. F5 to refresh."

# =========================================================
#  WINDOW CLOSE
# =========================================================
$Window.Add_Closed({
    $clock.Stop()
    try {
        Get-Job -Name "AIO_*","Batch_*" -ErrorAction SilentlyContinue | Stop-Job -PassThru | Remove-Job -Force
        $Script:RSPool.Close()
        $Script:RSPool.Dispose()
    } catch {}
    Write-Log "=== $Script:AppName closed ==="
})

# =========================================================
#  SHOW WINDOW
# =========================================================
[void]$Window.ShowDialog()
