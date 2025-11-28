<# READ ME
Save as AdminGUI.ps1 and open it with PowerShell 7 as Administrator.
#>

<#
.SYNOPSIS
 Modernized for PowerShell 7 (WPF, ThreadJobs)

.DESCRIPTION
  - Modernized WPF single-file GUI for PowerShell 7 on Windows.
  - Uses Start-ThreadJob for async Active Directory queries (AD-native module).
  - STA thread for WPF with Dispatcher updates.
  - Export CSV/JSON/TXT, Cancel active queries, config saved to user profile.
  - Improvements: error handling, cancellation, progress, row coloring, safe JSON parsing.

.PARAMETER ScheduledMode
  If present, will run in headless scheduled mode (limited support).

.EXAMPLE
  pwsh -File .\ADExpert.ps1
#>

param(
    [switch]$ScheduledMode,
    [string]$Preset = "",
    [string]$ExportFolderArg = "",
    [string]$Formats = "csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# -----------------------------
# Config & Globals
# -----------------------------
$Script:AppName = "AD Expert (Async, PS7)"
$Script:ConfigFile = Join-Path $env:USERPROFILE "ADExpert.config.json"
$Script:DefaultExportFolder = if ($ExportFolderArg -and $ExportFolderArg.Trim()) { $ExportFolderArg } else { Join-Path $env:USERPROFILE "Desktop\ADExports" }
If (!(Test-Path $Script:DefaultExportFolder)) { New-Item -Path $Script:DefaultExportFolder -ItemType Directory -Force | Out-Null }

function Save-Config { param([hashtable]$cfg)
    try { $cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $Script:ConfigFile -Encoding UTF8 -Force } catch { Write-Warning "Unable to save config: $_" }
}
function Load-Config {
    if (Test-Path $Script:ConfigFile) {
        try { Get-Content -Path $Script:ConfigFile -Raw | ConvertFrom-Json -ErrorAction Stop } catch { return $null }
    } else { return $null }
}

function Ensure-ModuleLoaded {
    param([string]$Name)
    try {
        if (Get-Module -ListAvailable -Name $Name) {
            return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}

$HasAD  = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO = Ensure-ModuleLoaded -Name GroupPolicy
$HasDFS = Ensure-ModuleLoaded -Name Dfsn
$HasDHCP = Ensure-ModuleLoaded -Name DhcpServer
$HasDNS = Ensure-ModuleLoaded -Name DnsServer

# -----------------------------
# Utility functions
# -----------------------------
function SafeFileName { param([string]$n) if (-not $n) { $n = "results" } return ($n -replace '[^\w\-\._ ]','_').Trim() }

function Export-Results {
    param([object[]]$Results, [string]$Category, [string]$Filter, [string]$ExportPath, [string[]]$Formats)
    if (-not $Results) { return }
    if (!(Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $base = SafeFileName("$Category`_$Filter`_$timestamp")
    foreach ($fmt in $Formats) {
        switch ($fmt.ToLower()) {
            "csv" {
                $file = Join-Path $ExportPath ($base + ".csv")
                try { $Results | Export-Csv -Path $file -NoTypeInformation -Force -Encoding UTF8 } catch { Write-Warning "CSV export failed: $_" }
            }
            "json" {
                $file = Join-Path $ExportPath ($base + ".json")
                try { $Results | ConvertTo-Json -Depth 6 | Set-Content -Path $file -Encoding UTF8 -Force } catch { Write-Warning "JSON export failed: $_" }
            }
            "txt" {
                $file = Join-Path $ExportPath ($base + ".txt")
                try { $Results | Out-String | Set-Content -Path $file -Encoding UTF8 -Force } catch { Write-Warning "TXT export failed: $_" }
            }
            default {
                $file = Join-Path $ExportPath ($base + ".txt")
                try { $Results | Out-String | Set-Content -Path $file -Encoding UTF8 -Force } catch { Write-Warning "Default export failed: $_" }
            }
        }
    }
}

# Safe JSON parse that returns array or empty array
function Convert-JsonTextToObjects {
    param([string]$text)
    if (-not $text) { return @() }
    $t = $text.Trim()
    try {
        if ($t.StartsWith("[") -or $t.StartsWith("{")) {
            $obj = $t | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $obj) { return @() }
            if ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) { return $obj } else { return ,$obj }
        } else {
            return @()
        }
    } catch {
        return @()
    }
}

# Enrich objects with RowClass
function Enrich-WithRowClass {
    param([object[]]$items)
    if (-not $items) { return @() }
    $out = @()
    foreach ($o in $items) {
        $po = [pscustomobject]@{}
        foreach ($p in $o.psobject.properties) { $po | Add-Member -MemberType NoteProperty -Name $p.Name -Value $p.Value -Force }
        $rowClass = ""
        if ($po.PSObject.Properties.Match("LockedOut") -and $po.LockedOut -eq $true) { $rowClass = "LockedOut" }
        elseif ($po.PSObject.Properties.Match("Enabled") -and $po.Enabled -eq $false) { $rowClass = "Disabled" }
        elseif ($po.PSObject.Properties.Match("OperatingSystem") -and $po.OperatingSystem -and $po.OperatingSystem -like "*Server*") { $rowClass = "Server" }
        elseif ($po.PSObject.Properties.Match("IsRODC") -and $po.IsRODC) { $rowClass = "RODC" }
        $po | Add-Member -MemberType NoteProperty -Name "RowClass" -Value $rowClass -Force
        $out += $po
    }
    return $out
}

# -----------------------------
# XAML for WPF (string)
# -----------------------------
$Xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='$($Script:AppName)' Height='760' Width='1150' WindowStartupLocation='CenterScreen'>
  <Window.Resources>
    <Style x:Key='WatermarkTextBox' TargetType='TextBox'>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='TextBox'>
            <Grid>
              <ScrollViewer x:Name='PART_ContentHost' />
              <TextBlock x:Name='Watermark' Text='{TemplateBinding Tag}' Foreground='Gray' Margin='4,2,0,0' IsHitTestVisible='False' Visibility='Collapsed' />
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property='Text' Value=''>
                <Setter TargetName='Watermark' Property='Visibility' Value='Visible' />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin='8'>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto' />
      <RowDefinition Height='Auto' />
      <RowDefinition Height='*' />
      <RowDefinition Height='Auto' />
    </Grid.RowDefinitions>

    <StackPanel Orientation='Horizontal' Grid.Row='0' Margin='0,0,0,8'>
      <Label Content='Domain:' VerticalAlignment='Center' />
      <ComboBox x:Name='cmbDomain' Width='260' Margin='6,0,12,0' />
      <Label Content='DC (optional):' VerticalAlignment='Center' />
      <ComboBox x:Name='cmbDC' Width='260' Margin='6,0,12,0' />
      <Button x:Name='btnRefreshDCs' Content='Refresh DCs' Width='110' Margin='6,0,0,0' />
      <Label Content='Export Folder:' VerticalAlignment='Center' Margin='12,0,0,0' />
      <TextBox x:Name='txtExportFolder' Width='300' Margin='6,0,12,0' Style='{StaticResource WatermarkTextBox}' Tag='C:\ADExports' />
      <Button x:Name='btnBrowse' Content='Browse' Width='70' Margin='0,0,0,0' />
    </StackPanel>

    <StackPanel Orientation='Horizontal' Grid.Row='1' Margin='0,0,0,8'>
      <Label Content='Category:' VerticalAlignment='Center' />
      <ComboBox x:Name='cmbCategory' Width='180' Margin='6,0,12,0' />
      <Label Content='Filter:' VerticalAlignment='Center' />
      <TextBox x:Name='txtFilter' Width='320' Margin='6,0,12,0' Style='{StaticResource WatermarkTextBox}' Tag='LDAP filter or text' />
      <Button x:Name='btnSearch' Content='Search' Width='100' Margin='6,0,0,0' />
      <Button x:Name='btnCancel' Content='Cancel' Width='80' Margin='6,0,0,0' IsEnabled='False' />
      <Button x:Name='btnClear' Content='Clear' Width='80' Margin='6,0,0,0' />
    </StackPanel>

    <TabControl x:Name='tabMain' Grid.Row='2'>
      <TabItem Header='Users'>
        <Grid Margin='6'><DataGrid x:Name='dgUsers' AutoGenerateColumns='True' IsReadOnly='True' /></Grid>
      </TabItem>
      <TabItem Header='Computers'>
        <Grid Margin='6'><DataGrid x:Name='dgComputers' AutoGenerateColumns='True' IsReadOnly='True' /></Grid>
      </TabItem>
      <TabItem Header='Groups'>
        <Grid Margin='6'><DataGrid x:Name='dgGroups' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
      <TabItem Header='GPOs'>
        <Grid Margin='6'><DataGrid x:Name='dgGPOs' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
      <TabItem Header='Subnets'>
        <Grid Margin='6'><DataGrid x:Name='dgSubnets' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
      <TabItem Header='RODCs'>
        <Grid Margin='6'><DataGrid x:Name='dgRODCs' AutoGenerateColumns='True' IsReadOnly='True' /></Grid>
      </TabItem>
      <TabItem Header='DFS'>
        <Grid Margin='6'><DataGrid x:Name='dgDFS' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
      <TabItem Header='DHCP'>
        <Grid Margin='6'><DataGrid x:Name='dgDHCP' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
      <TabItem Header='DNS'>
        <Grid Margin='6'><DataGrid x:Name='dgDNS' AutoGenerateColumns='True' IsReadOnly='True'/></Grid>
      </TabItem>
    </TabControl>

    <StatusBar Grid.Row='3' Margin='0,8,0,0'>
      <StatusBarItem><TextBlock x:Name='txtStatus' Text='Ready'/></StatusBarItem>
      <StatusBarItem><TextBlock x:Name='txtProgress' Text=''/></StatusBarItem>
    </StatusBar>

  </Grid>
</Window>
"@

# -----------------------------
# Run WPF in STA thread
# -----------------------------
# We'll create a scriptblock executed on an STA thread, building UI and wiring events.
$sta = {
    param($XamlString, $ScriptGlobals)

    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
    Add-Type -AssemblyName System.Windows.Forms

    [xml]$xamlXml = $XamlString
    $reader = (New-Object System.Xml.XmlNodeReader $xamlXml)
    try {
        $Window = [Windows.Markup.XamlReader]::Load($reader)
    } catch {
        Write-Error "Failed to load XAML: $_"
        return
    }

    function Get-Ctrl([string]$name) { $Window.FindName($name) }

    # Controls
    $cmbDomain = Get-Ctrl "cmbDomain"; $cmbDC = Get-Ctrl "cmbDC"; $btnRefreshDCs = Get-Ctrl "btnRefreshDCs"; $txtExportFolder = Get-Ctrl "txtExportFolder"; $txtStatus = Get-Ctrl "txtStatus"; $txtProgress = Get-Ctrl "txtProgress"
    $cmbCategory = Get-Ctrl "cmbCategory"; $txtFilter = Get-Ctrl "txtFilter"; $btnSearch = Get-Ctrl "btnSearch"; $btnClear = Get-Ctrl "btnClear"; $btnCancel = Get-Ctrl "btnCancel"; $btnBrowse = Get-Ctrl "btnBrowse"
    $dgUsers = Get-Ctrl "dgUsers"; $dgComputers = Get-Ctrl "dgComputers"; $dgGroups = Get-Ctrl "dgGroups"; $dgGPOs = Get-Ctrl "dgGPOs"; $dgSubnets = Get-Ctrl "dgSubnets"; $dgRODCs = Get-Ctrl "dgRODCs"; $dgDFS = Get-Ctrl "dgDFS"; $dgDHCP = Get-Ctrl "dgDHCP"; $dgDNS = Get-Ctrl "dgDNS"

    # config
    $cfg = $ScriptGlobals.LoadConfig.Invoke()
    if ($cfg -and $cfg.ExportFolder) { $txtExportFolder.Text = $cfg.ExportFolder } else { $txtExportFolder.Text = $ScriptGlobals.DefaultExportFolder }

    # populate categories
    $categories = @("Users","Computers","Groups","GPOs","Subnets","RODCs","DFS","DHCP","DNS")
    $cmbCategory.ItemsSource = $categories; $cmbCategory.SelectedIndex = 0

    # Row style via XAML load (simple)
    $rowStyleXaml = @"
<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='DataGridRow'>
  <Style.Triggers>
    <DataTrigger Binding='{Binding RowClass}' Value='LockedOut'>
      <Setter Property='Background' Value='#FFF8D7D7'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='Disabled'>
      <Setter Property='Background' Value='#FFECECEC'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='Server'>
      <Setter Property='Background' Value='#FFDDEBF7'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='RODC'>
      <Setter Property='Background' Value='#FFFFF8D9'/>
    </DataTrigger>
  </Style.Triggers>
</Style>
"@
    [xml]$sx = $rowStyleXaml
    $rdr = (New-Object System.Xml.XmlNodeReader $sx)
    try { $rowStyle = [Windows.Markup.XamlReader]::Load($rdr) } catch { $rowStyle = $null }
    if ($rowStyle) {
        $dgUsers.RowStyle = $rowStyle; $dgComputers.RowStyle = $rowStyle; $dgRODCs.RowStyle = $rowStyle
    }

    # Jobs tracking
    $script:ActiveJobs = @()
    $script:JobEvents = @()

    # Helper: add job and register callback
    function Start-AndRegisterJob {
        param([ScriptBlock]$ScriptBlock, [hashtable]$Args, [ScriptBlock]$OnComplete)
        $job = Start-ThreadJob -ScriptBlock $ScriptBlock -ArgumentList ($Args) -RunAs32:$false
        $script:ActiveJobs += $job
        $btnCancel.IsEnabled = $true
        # register event to detect completion
        $evt = Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
            param($sender,$e)
            if ($sender.State -eq 'Completed' -or $sender.State -eq 'Failed' -or $sender.State -eq 'Stopped') {
                try {
                    $out = if ($sender.State -eq 'Completed') { Receive-Job -Job $sender -Keep | Out-String } else { $null }
                    & $OnComplete $sender $out $sender.ChildJobs[0].JobStateInfo.Reason
                } catch {
                    & $OnComplete $sender $null $_
                } finally {
                    # cleanup
                    try { Unregister-Event -SourceIdentifier $event.SubscriptionId -ErrorAction SilentlyContinue } catch {}
                    try { Remove-Job -Job $sender -Force -ErrorAction SilentlyContinue } catch {}
                    # remove from active jobs
                    [void]($script:ActiveJobs = $script:ActiveJobs | Where-Object { $_.Id -ne $sender.Id })
                    if ($script:ActiveJobs.Count -eq 0) { $btnCancel.IsEnabled = $false; $txtProgress.Text = "" }
                }
            }
        }
        $script:JobEvents += $evt
        return $job
    }

    # Helper to stop active jobs
    function Stop-ActiveJobs {
        foreach ($j in $script:ActiveJobs) {
            try { Stop-Job -Job $j -Force -ErrorAction SilentlyContinue } catch {}
            try { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue } catch {}
        }
        $script:ActiveJobs = @()
        $btnCancel.IsEnabled = $false
        $txtStatus.Text = "Cancelled."
        $txtProgress.Text = ""
    }

    # Domain discovery
    function Get-ForestDomainsLocal {
        try {
            if (-not ($ScriptGlobals.HasAD)) { return @() }
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue
            $f = Get-ADForest -ErrorAction Stop
            return $f.Domains
        } catch { return @() }
    }

    function Get-DomainControllersLocal {
        param([string]$Domain)
        try {
            if (-not ($ScriptGlobals.HasAD)) { return @() }
            Import-Module ActiveDirectory -ErrorAction SilentlyContinue
            return Get-ADDomainController -Filter * -Server $Domain -ErrorAction SilentlyContinue
        } catch { return @() }
    }

    # Refresh domains & DCs
    function Refresh-Domains {
        $txtStatus.Text = "Loading domains..."
        $domains = Get-ForestDomainsLocal
        if ($domains) {
            $cmbDomain.ItemsSource = $domains
            $cmbDomain.SelectedIndex = 0
            $txtStatus.Text = "Domains loaded."
            Refresh-DCs
        } else {
            $cmbDomain.ItemsSource = @()
            $txtStatus.Text = "No domains / AD module not available."
        }
    }

    function Refresh-DCs {
        $domain = $cmbDomain.SelectedItem
        $cmbDC.ItemsSource = @()
        if (-not $domain) { return }
        $txtStatus.Text = "Refreshing DCs..."
        try {
            $dcs = Get-DomainControllersLocal -Domain $domain
            if ($dcs) { $cmbDC.ItemsSource = $dcs | ForEach-Object { $_.HostName }; $cmbDC.SelectedIndex = 0 }
            $txtStatus.Text = "DCs refreshed."
        } catch { $txtStatus.Text = "Unable to enumerate DCs: $($_.Exception.Message)" }
    }

    $btnRefreshDCs.Add_Click({ Refresh-DCs })
    try { Refresh-Domains } catch {}

    # Browse folder
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.SelectedPath = $txtExportFolder.Text
        if ($dlg.ShowDialog() -eq 'OK') { $txtExportFolder.Text = $dlg.SelectedPath }
    })

    # Search worker scriptblocks (run inside thread jobs)
    $sb_SearchUsers = {
        param($args)
        $f = $args.Filter; $s = $args.Server
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        $props = @("Name","sAMAccountName","DistinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","userPrincipalName","mail","OperatingSystem")
        try {
            if ($f -and $f.Trim()) {
                if ($f -match '^\(|\=|\&|\|') {
                    $res = Get-ADUser -LDAPFilter $f -Properties $props -Server $s -ErrorAction SilentlyContinue
                } else {
                    $like = "*$f*"
                    $res = Get-ADUser -Filter { Name -like $like -or sAMAccountName -like $like -or mail -like $like -or userPrincipalName -like $like } -Properties $props -Server $s -ErrorAction SilentlyContinue
                }
            } else {
                $res = Get-ADUser -Filter * -Properties $props -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,sAMAccountName,Enabled,LockedOut,LastLogonDate,DistinguishedName | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchComputers = {
        param($args)
        $f = $args.Filter; $s = $args.Server
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        $props = @("Name","OperatingSystem","OperatingSystemVersion","DistinguishedName","whenCreated","LastLogonDate")
        try {
            if ($f -and $f.Trim()) {
                if ($f -match '^\(|\=|\&|\|') {
                    $res = Get-ADComputer -LDAPFilter $f -Properties $props -Server $s -ErrorAction SilentlyContinue
                } else {
                    $like = "*$f*"
                    $res = Get-ADComputer -Filter { Name -like $like -or OperatingSystem -like $like } -Properties $props -Server $s -ErrorAction SilentlyContinue
                }
            } else {
                $res = Get-ADComputer -Filter * -Properties $props -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,DistinguishedName | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchGroups = {
        param($args)
        $f = $args.Filter; $s = $args.Server
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            if ($f -and $f.Trim()) {
                $like = "*$f*"
                $res = Get-ADGroup -Filter { Name -like $like } -Properties member,GroupScope -Server $s -ErrorAction SilentlyContinue
            } else {
                $res = Get-ADGroup -Filter * -Properties member,GroupScope -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,GroupScope,@{n='Members';e={$_.member -join '; '}} | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchGPOs = {
        param($args)
        $f = $args.Filter
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        try {
            $g = Get-GPO -All -ErrorAction SilentlyContinue
            if ($f -and $f.Trim()) { $g = $g | Where-Object { $_.DisplayName -like "*$f*" } }
            $g | Select-Object DisplayName,Id,Owner,CreationTime,ModificationTime | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchSubnets = {
        param($args)
        $s = $args.Server
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            $cn = (Get-ADRootDSE -Server $s).configurationNamingContext
            $base = "CN=Subnets,CN=Sites,$cn"
            $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -Server $s -ErrorAction SilentlyContinue
            $subnets | Select-Object Name,@{n='Location';e={$_.location}},@{n='DN';e={$_.DistinguishedName}} | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchRODCs = {
        param($args)
        $s = $args.Server
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            $dcs = Get-ADDomainController -Filter * -Server $s -ErrorAction SilentlyContinue
            $ro = $dcs | Where-Object { $_.IsReadOnly -eq $true } | Select-Object HostName,Site,OperatingSystem,@{n='IsRODC';e={$true}}
            $ro | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchDFS = {
        param($args)
        Import-Module Dfsn -ErrorAction SilentlyContinue
        try {
            $r = Get-DfsnRoot -ErrorAction SilentlyContinue
            if (-not $r) { @{ Note = "No DFS roots or module missing" } | ConvertTo-Json; return }
            $r | Select-Object Path,State,Type | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchDHCP = {
        param($args)
        $s = $args.Server
        Import-Module DhcpServer -ErrorAction SilentlyContinue
        try {
            $scopes = Get-DhcpServerv4Scope -ComputerName $s -ErrorAction SilentlyContinue
            if (-not $scopes) { @{ Note = "No DHCP scopes or module missing" } | ConvertTo-Json; return }
            $scopes | Select-Object ScopeId,Name,StartRange,EndRange,State | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    $sb_SearchDNS = {
        param($args)
        $s = $args.Server
        Import-Module DnsServer -ErrorAction SilentlyContinue
        try {
            $z = Get-DnsServerZone -ComputerName $s -ErrorAction SilentlyContinue
            if (-not $z) { @{ Note = "No zones or module missing" } | ConvertTo-Json; return }
            $z | Select-Object ZoneName,ZoneType,IsDsIntegrated | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    }

    # Helpers to kick off searches and process results
    function DoUserSearch {
        param([string]$filter, [string]$server, [bool]$allDomains)
        $txtStatus.Text = "Searching users..."
        $txtProgress.Text = "Running user search..."
        if ($allDomains) {
            $domains = Get-ForestDomainsLocal
            $agg = New-Object System.Collections.ArrayList
            $count = 0
            foreach ($d in $domains) {
                $dcs = Get-DomainControllersLocal -Domain $d
                $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
                $args = @{ Filter = $filter; Server = $target }
                $OnComplete = {
                    param($job,$out,$reason)
                    [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                        if ($reason) {
                            $txtStatus.Text = "Error querying $target: $($reason.Exception.Message)"
                            return
                        }
                        $objs = Convert-JsonTextToObjects $out
                        if ($objs) {
                            foreach ($o in $objs) { [void]$agg.Add($o) }
                            $en = Enrich-WithRowClass -items $agg
                            $dgUsers.ItemsSource = $en
                            $count = $en.Count
                            $txtStatus.Text = "Users: $count results (partial)"
                        }
                    }))
                }
                Start-AndRegisterJob -ScriptBlock $sb_SearchUsers -Args $args -OnComplete $OnComplete | Out-Null
            }
        } else {
            $args = @{ Filter = $filter; Server = $server }
            $OnComplete = {
                param($job,$out,$reason)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($reason) { $txtStatus.Text = "User search failed: $($reason.Exception.Message)"; $dgUsers.ItemsSource = @(); return }
                    $objs = Convert-JsonTextToObjects $out
                    $en = Enrich-WithRowClass -items $objs
                    $dgUsers.ItemsSource = $en
                    $txtStatus.Text = "Users: $($en.Count) results."
                }))
            }
            Start-AndRegisterJob -ScriptBlock $sb_SearchUsers -Args $args -OnComplete $OnComplete | Out-Null
        }
    }

    function DoComputerSearch {
        param([string]$filter, [string]$server, [bool]$allDomains)
        $txtStatus.Text = "Searching computers..."
        $txtProgress.Text = "Running computer search..."
        if ($allDomains) {
            $domains = Get-ForestDomainsLocal
            $agg = New-Object System.Collections.ArrayList
            foreach ($d in $domains) {
                $dcs = Get-DomainControllersLocal -Domain $d
                $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
                $args = @{ Filter = $filter; Server = $target }
                $OnComplete = {
                    param($job,$out,$reason)
                    [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                        if ($reason) { $txtStatus.Text = "Error: $($reason.Exception.Message)"; return }
                        $objs = Convert-JsonTextToObjects $out
                        if ($objs) {
                            foreach ($o in $objs) { [void]$agg.Add($o) }
                            $en = Enrich-WithRowClass -items $agg
                            $dgComputers.ItemsSource = $en
                            $txtStatus.Text = "Computers: $($en.Count) results (partial)"
                        }
                    }))
                }
                Start-AndRegisterJob -ScriptBlock $sb_SearchComputers -Args $args -OnComplete $OnComplete | Out-Null
            }
        } else {
            $args = @{ Filter = $filter; Server = $server }
            $OnComplete = {
                param($job,$out,$reason)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($reason) { $txtStatus.Text = "Computer search failed: $($reason.Exception.Message)"; $dgComputers.ItemsSource = @(); return }
                    $objs = Convert-JsonTextToObjects $out
                    $en = Enrich-WithRowClass -items $objs
                    $dgComputers.ItemsSource = $en
                    $txtStatus.Text = "Computers: $($en.Count) results."
                }))
            }
            Start-AndRegisterJob -ScriptBlock $sb_SearchComputers -Args $args -OnComplete $OnComplete | Out-Null
        }
    }

    function DoGroupSearch {
        param([string]$filter, [string]$server, [bool]$allDomains)
        $txtStatus.Text = "Searching groups..."
        $txtProgress.Text = "Running group search..."
        if ($allDomains) {
            $domains = Get-ForestDomainsLocal
            $agg = New-Object System.Collections.ArrayList
            foreach ($d in $domains) {
                $dcs = Get-DomainControllersLocal -Domain $d
                $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
                $args = @{ Filter = $filter; Server = $target }
                $OnComplete = {
                    param($job,$out,$reason)
                    [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                        if ($reason) { $txtStatus.Text = "Error: $($reason.Exception.Message)"; return }
                        $objs = Convert-JsonTextToObjects $out
                        if ($objs) {
                            foreach ($o in $objs) { [void]$agg.Add($o) }
                            $dgGroups.ItemsSource = $agg
                            $txtStatus.Text = "Groups: $($agg.Count) results (partial)"
                        }
                    }))
                }
                Start-AndRegisterJob -ScriptBlock $sb_SearchGroups -Args $args -OnComplete $OnComplete | Out-Null
            }
        } else {
            $args = @{ Filter = $filter; Server = $server }
            $OnComplete = {
                param($job,$out,$reason)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($reason) { $txtStatus.Text = "Group search failed: $($reason.Exception.Message)"; $dgGroups.ItemsSource = @(); return }
                    $objs = Convert-JsonTextToObjects $out
                    $dgGroups.ItemsSource = $objs
                    $txtStatus.Text = "Groups: $($objs.Count) results."
                }))
            }
            Start-AndRegisterJob -ScriptBlock $sb_SearchGroups -Args $args -OnComplete $OnComplete | Out-Null
        }
    }

    function DoGPOs { param([string]$filter)
        $txtStatus.Text = "Searching GPOs..."; $txtProgress.Text = "Running GPO query..."
        $args = @{ Filter = $filter }
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "GPO search error: $($reason.Exception.Message)"; $dgGPOs.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $dgGPOs.ItemsSource = $objs
                $txtStatus.Text = "GPOs: $($objs.Count) results."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchGPOs -Args $args -OnComplete $OnComplete | Out-Null
    }

    function DoSubnets { param([string]$server)
        $txtStatus.Text = "Retrieving subnets..."; $txtProgress.Text = "Running subnets query..."
        $args = @{ Server = $server }
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "Subnets retrieval failed: $($reason.Exception.Message)"; $dgSubnets.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $dgSubnets.ItemsSource = $objs
                $txtStatus.Text = "Subnets: $($objs.Count) results."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchSubnets -Args $args -OnComplete $OnComplete | Out-Null
    }

    function DoRODCs { param([string]$server)
        $txtStatus.Text = "Finding RODCs..."; $txtProgress.Text = "Running RODC query..."
        $args = @{ Server = $server }
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "RODC search failed: $($reason.Exception.Message)"; $dgRODCs.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $en = Enrich-WithRowClass -items $objs
                $dgRODCs.ItemsSource = $en
                $txtStatus.Text = "RODCs: $($en.Count) results."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchRODCs -Args $args -OnComplete $OnComplete | Out-Null
    }

    function DoDFS {
        $txtStatus.Text = "Querying DFS..."; $txtProgress.Text = "Running DFS query..."
        $args = @{}
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "DFS query failed: $($reason.Exception.Message)"; $dgDFS.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $dgDFS.ItemsSource = $objs
                $txtStatus.Text = "DFS: $($objs.Count) items."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchDFS -Args $args -OnComplete $OnComplete | Out-Null
    }

    function DoDHCP { param([string]$server)
        if (-not $server) { $txtStatus.Text = "Specify DHCP server."; return }
        $txtStatus.Text = "Querying DHCP..."; $txtProgress.Text = "Running DHCP query..."
        $args = @{ Server = $server }
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "DHCP query failed: $($reason.Exception.Message)"; $dgDHCP.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $dgDHCP.ItemsSource = $objs
                $txtStatus.Text = "DHCP: $($objs.Count) items."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchDHCP -Args $args -OnComplete $OnComplete | Out-Null
    }

    function DoDNS { param([string]$server)
        if (-not $server) { $txtStatus.Text = "Specify DNS server."; return }
        $txtStatus.Text = "Querying DNS..."; $txtProgress.Text = "Running DNS query..."
        $args = @{ Server = $server }
        $OnComplete = {
            param($job,$out,$reason)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($reason) { $txtStatus.Text = "DNS query failed: $($reason.Exception.Message)"; $dgDNS.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $out
                $dgDNS.ItemsSource = $objs
                $txtStatus.Text = "DNS zones: $($objs.Count)."
            }))
        }
        Start-AndRegisterJob -ScriptBlock $sb_SearchDNS -Args $args -OnComplete $OnComplete | Out-Null
    }

    # Wire UI events
    $btnSearch.Add_Click({
        $cat = $cmbCategory.SelectedItem
        $filter = $txtFilter.Text
        $domain = $cmbDomain.SelectedItem
        $dc = $cmbDC.SelectedItem
        $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
        $btnCancel.IsEnabled = $true
        switch ($cat) {
            "Users" { DoUserSearch -filter $filter -server $server -allDomains:$false }
            "Computers" { DoComputerSearch -filter $filter -server $server -allDomains:$false }
            "Groups" { DoGroupSearch -filter $filter -server $server -allDomains:$false }
            "GPOs" { DoGPOs -filter $filter }
            "Subnets" { DoSubnets -server $server }
            "RODCs" { DoRODCs -server $server }
            "DFS" { DoDFS }
            "DHCP" { DoDHCP -server $txtExportFolder.Text.Trim() }   # placeholder: or extend UI to specify server
            "DNS" { DoDNS -server $txtExportFolder.Text.Trim() }     # same
            default { $txtStatus.Text = "Unknown category" }
        }
    })

    $btnCancel.Add_Click({ Stop-ActiveJobs })

    $btnClear.Add_Click({
        $txtFilter.Text = ""
        $dgUsers.ItemsSource = @()
        $dgComputers.ItemsSource = @()
        $dgGroups.ItemsSource = @()
        $dgGPOs.ItemsSource = @()
        $dgSubnets.ItemsSource = @()
        $dgRODCs.ItemsSource = @()
        $dgDFS.ItemsSource = @()
        $dgDHCP.ItemsSource = @()
        $dgDNS.ItemsSource = @()
        $txtStatus.Text = "Cleared."
    })

    # Export active tab
    function Export-ActiveTabLocal {
        $sel = $Window.FindName("tabMain").SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("No tab selected.","Export") | Out-Null; return }
        $header = $sel.Header
        switch ($header) {
            "Users" { $items = $dgUsers.ItemsSource }
            "Computers" { $items = $dgComputers.ItemsSource }
            "Groups" { $items = $dgGroups.ItemsSource }
            "GPOs" { $items = $dgGPOs.ItemsSource }
            "Subnets" { $items = $dgSubnets.ItemsSource }
            "RODCs" { $items = $dgRODCs.ItemsSource }
            "DFS" { $items = $dgDFS.ItemsSource }
            "DHCP" { $items = $dgDHCP.ItemsSource }
            "DNS" { $items = $dgDNS.ItemsSource }
            default { $items = @() }
        }
        if (-not $items -or $items.Count -eq 0) { [System.Windows.MessageBox]::Show("No results to export.","Export") | Out-Null; return }
        $formats = $ScriptGlobals.Formats.Split(',') | ForEach-Object { $_.Trim() }
        $path = $txtExportFolder.Text.Trim(); if (-not $path) { $path = $ScriptGlobals.DefaultExportFolder }
        Export-Results -Results ($items | ForEach-Object { $_ }) -Category $header -Filter $txtFilter.Text -ExportPath $path -Formats $formats
        $ScriptGlobals.SaveConfig.Invoke(@{ ExportFolder = $path; Formats = $formats })
        $txtStatus.Text = "Exported $($items.Count) items to $path"
    }

    # Ctrl+E => export
    $Window.Add_KeyDown({
        param($sender,$e)
        if ($e.KeyboardDevice.Modifiers -eq 'Control' -and $e.Key -eq 'E') { Export-ActiveTabLocal }
    })

    # Add right-click context menu to data grids for export (lightweight)
    $addContext = {
        param($dg)
        $ctx = New-Object System.Windows.Controls.ContextMenu
        $m = New-Object System.Windows.Controls.MenuItem
        $m.Header = "Export..."
        $m.Add_Click({ Export-ActiveTabLocal })
        $ctx.Items.Add($m)
        $dg.ContextMenu = $ctx
    }
    $addContext.Invoke($dgUsers); $addContext.Invoke($dgComputers); $addContext.Invoke($dgGroups)

    # When window closes, clean jobs and events
    $Window.Add_Closed({
        Stop-ActiveJobs
        foreach ($ev in $script:JobEvents) { try { Unregister-Event -SubscriptionId $ev.SubscriptionId -ErrorAction SilentlyContinue } catch {} }
    })

    # Show dialog
    $Window.ShowDialog() | Out-Null
}

# -----------------------------
# Start STA thread with GUI
# -----------------------------
# Prepare helper delegates to pass into STA block (functions cannot be passed directly across runspace boundaries, so pass scriptblock wrappers)
$scriptGlobals = @{
    LoadConfig = { Load-Config }
    SaveConfig = { param($h) Save-Config $h }
    DefaultExportFolder = $Script:DefaultExportFolder
    Formats = $Formats
    HasAD = $HasAD
    LoadConfig.Invoke = { Load-Config } # for usage inside
}

# Start STA thread
$thread = [System.Threading.Thread]::new([System.Threading.ThreadStart]{ param() })
# We will create a thread start with a scriptblock wrapper using Powershell's Create and Invoke
$sb = {
    param($XamlString, $globals)
    & $using:sta $XamlString $globals
}

# Use .NET ThreadStart with parameter via closure
$thread = New-Object System.Threading.Thread( ([System.Threading.ThreadStart]{ & $using:sta $Xaml $scriptGlobals }) )
$thread.SetApartmentState('STA')
$thread.IsBackground = $false
$thread.Start()
$thread.Join()  # Wait for UI to close before exiting the script

# End of script
