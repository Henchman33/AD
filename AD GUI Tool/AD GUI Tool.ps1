<#
\.AD GUI Tool - Full Async WPF tool (single-file)
Features:
 - WPF GUI compatible with PowerShell 5.1 (.NET Framework)
 - Watermark TextBoxes (Tag + ControlTemplate)
 - RunspacePool-based async searches (forest-aware)
 - Tabs: Users, Computers, Groups, GPOs, Subnets, RODCs, DFS, DHCP, DNS
 - Color-coded rows (RowClass property applied in UI)
 - Export: CSV/JSON/TXT
 - Headless ScheduledMode support (simple presets)
 - Author: Stephen McKee - Server Administrator 2
#>

param(
    [switch]$ScheduledMode,
    [string]$Preset = "",
    [string]$ExportFolderArg = "",
    [string]$Formats = "csv"
)

# -----------------------------
# Assemblies
# -----------------------------
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms

# -----------------------------
# Globals & Config
# -----------------------------
$Script:AppName = "AD Expert (Async)"
$Script:ConfigFile = Join-Path $env:USERPROFILE "ADExpert.config.json"
$Script:DefaultExportFolder = if ($ExportFolderArg -and $ExportFolderArg.Trim()) { $ExportFolderArg } else { Join-Path $env:USERPROFILE "Desktop\ADExports" }
If (!(Test-Path $Script:DefaultExportFolder)) { New-Item -Path $Script:DefaultExportFolder -ItemType Directory -Force | Out-Null }

function Save-Config { param($cfg) try { $cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $Script:ConfigFile -Encoding UTF8 } catch { Write-Warning "Unable to save config: $_" } }
function Load-Config { if (Test-Path $Script:ConfigFile) { try { Get-Content -Path $Script:ConfigFile -Raw | ConvertFrom-Json } catch { $null } } else { $null } }

function Ensure-ModuleLoaded { param([string]$Name) if (Get-Module -ListAvailable -Name $Name) { Import-Module $Name -ErrorAction SilentlyContinue; return $true } else { return $false } }

$HasAD  = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO = Ensure-ModuleLoaded -Name GroupPolicy
$HasDFS = Ensure-ModuleLoaded -Name Dfsn
$HasDHCP = Ensure-ModuleLoaded -Name DhcpServer
$HasDNS = Ensure-ModuleLoaded -Name DnsServer

# -----------------------------
# Runspace pool
# -----------------------------
$minThreads = 1
$maxThreads = 6
$runspacePool = [runspacefactory]::CreateRunspacePool($minThreads, $maxThreads)
$runspacePool.ThreadOptions = "ReuseThread"
$runspacePool.Open()

function Invoke-Async {
    param(
        [ScriptBlock]$ScriptBlock,
        [ScriptBlock]$CompletedCallback  # { param($text,$error) ... }
    )
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $runspacePool
    $ps.AddScript($ScriptBlock) | Out-Null

    $async = $ps.BeginInvoke()
    [System.Threading.ThreadPool]::QueueUserWorkItem({
        param($ps,$async,$CompletedCallback)
        try {
            $out = $ps.EndInvoke($async)
            $text = $out -join "`n"
            & $CompletedCallback $text $null
        } catch {
            & $CompletedCallback $null $_
        } finally {
            $ps.Dispose()
        }
    }, @($ps,$async,$CompletedCallback)) | Out-Null
}

# -----------------------------
# Utilities
# -----------------------------
function SafeFileName { param([string]$n) if (-not $n) { $n = "results" } return ($n -replace '[^\w\-\._ ]','_').Trim() }

function Export-Results {
    param([object[]]$Results, [string]$Category, [string]$Filter, [string]$ExportPath, [string[]]$Formats)
    if (!(Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $base = SafeFileName("$Category`_$Filter`_$timestamp")
    foreach ($fmt in $Formats) {
        switch ($fmt.ToLower()) {
            "csv" {
                $file = Join-Path $ExportPath ($base + ".csv")
                $Results | Export-Csv -Path $file -NoTypeInformation -Force
            }
            "json" {
                $file = Join-Path $ExportPath ($base + ".json")
                $Results | ConvertTo-Json -Depth 6 | Set-Content -Path $file -Encoding UTF8
            }
            "txt" {
                $file = Join-Path $ExportPath ($base + ".txt")
                $Results | Out-String | Set-Content -Path $file -Encoding UTF8
            }
            default {
                $file = Join-Path $ExportPath ($base + ".txt")
                $Results | Out-String | Set-Content -Path $file -Encoding UTF8
            }
        }
    }
}

# -----------------------------
# AD discovery helpers
# -----------------------------
function Get-ForestDomains {
    if (-not $HasAD) { return @() }
    try { (Get-ADForest -ErrorAction Stop).Domains } catch { @() }
}
function Get-DomainControllers {
    param([string]$Domain)
    if (-not $HasAD) { return @() }
    try { Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop } catch { @() }
}

# -----------------------------
# ScriptBlocks for runspaces (return JSON strings)
# -----------------------------
function SB-SearchUsers {
    param([string]$Filter,[string]$Server)
    return {
        param($f,$s)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        $props = @("Name","sAMAccountName","DistinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","userPrincipalName","mail","OperatingSystem")
        try {
            if ($f -and $f.Trim() -ne "") {
                if ($f -match '^\(|\=|\&|\|') {
                    $res = Get-ADUser -LDAPFilter $f -Properties $props -Server $s -ErrorAction SilentlyContinue
                } else {
                    $res = Get-ADUser -Filter "Name -like '$f' -or sAMAccountName -like '$f' -or mail -like '$f' -or userPrincipalName -like '$f'" -Properties $props -Server $s -ErrorAction SilentlyContinue
                }
            } else {
                $res = Get-ADUser -Filter * -Properties $props -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,sAMAccountName,Enabled,LockedOut,LastLogonDate,DistinguishedName | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Filter $Server }
}

function SB-SearchComputers {
    param([string]$Filter,[string]$Server)
    return {
        param($f,$s)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        $props = @("Name","OperatingSystem","OperatingSystemVersion","DistinguishedName","whenCreated","LastLogonDate")
        try {
            if ($f -and $f.Trim() -ne "") {
                if ($f -match '^\(|\=|\&|\|') {
                    $res = Get-ADComputer -LDAPFilter $f -Properties $props -Server $s -ErrorAction SilentlyContinue
                } else {
                    $res = Get-ADComputer -Filter "Name -like '$f' -or OperatingSystem -like '$f'" -Properties $props -Server $s -ErrorAction SilentlyContinue
                }
            } else {
                $res = Get-ADComputer -Filter * -Properties $props -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,OperatingSystem,OperatingSystemVersion,LastLogonDate,DistinguishedName | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Filter $Server }
}

function SB-SearchGroups {
    param([string]$Filter,[string]$Server)
    return {
        param($f,$s)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            if ($f -and $f.Trim() -ne "") {
                $res = Get-ADGroup -Filter "Name -like '$f'" -Properties member,GroupScope -Server $s -ErrorAction SilentlyContinue
            } else {
                $res = Get-ADGroup -Filter * -Properties member,GroupScope -Server $s -ErrorAction SilentlyContinue
            }
            $res | Select-Object Name,GroupScope,@{n='Members';e={$_.member -join '; '}} | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Filter $Server }
}

function SB-SearchGPOs {
    param([string]$Filter)
    return {
        param($f)
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        try {
            $g = Get-GPO -All -ErrorAction SilentlyContinue
            if ($f -and $f.Trim() -ne "") { $g = $g | Where-Object { $_.DisplayName -like "*$f*" } }
            $g | Select-Object DisplayName,Id,Owner,CreationTime,ModificationTime | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Filter }
}

function SB-SearchSubnets {
    param([string]$Server)
    return {
        param($s)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            $cn = (Get-ADRootDSE -Server $s).configurationNamingContext
            $base = "CN=Subnets,CN=Sites,$cn"
            $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -Server $s -ErrorAction SilentlyContinue
            $subnets | Select-Object Name,@{n='Location';e={$_.location}},@{n='DN';e={$_.DistinguishedName}} | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Server }
}

function SB-SearchRODCs {
    param([string]$Server)
    return {
        param($s)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        try {
            $dcs = Get-ADDomainController -Filter * -Server $s -ErrorAction SilentlyContinue
            $ro = $dcs | Where-Object { $_.IsReadOnly -eq $true } | Select-Object HostName,Site,OperatingSystem,@{n='IsRODC';e={$true}}
            $ro | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Server }
}

function SB-SearchDFS {
    return {
        Import-Module Dfsn -ErrorAction SilentlyContinue
        try {
            $r = Get-DfsnRoot -ErrorAction SilentlyContinue
            if (-not $r) { @{ Note = "No DFS roots or module missing" } | ConvertTo-Json; return }
            $r | Select-Object Path,State,Type | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ }
}

function SB-SearchDHCP {
    param([string]$Server)
    return {
        param($s)
        Import-Module DhcpServer -ErrorAction SilentlyContinue
        try {
            $scopes = Get-DhcpServerv4Scope -ComputerName $s -ErrorAction SilentlyContinue
            if (-not $scopes) { @{ Note = "No DHCP scopes or module missing" } | ConvertTo-Json; return }
            $scopes | Select-Object ScopeId,Name,StartRange,EndRange,State | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Server }
}

function SB-SearchDNS {
    param([string]$Server)
    return {
        param($s)
        Import-Module DnsServer -ErrorAction SilentlyContinue
        try {
            $z = Get-DnsServerZone -ComputerName $s -ErrorAction SilentlyContinue
            if (-not $z) { @{ Note = "No zones or module missing" } | ConvertTo-Json; return }
            $z | Select-Object ZoneName,ZoneType,IsDsIntegrated | ConvertTo-Json -Depth 3
        } catch { @{ Error = $_.Exception.Message } | ConvertTo-Json }
    } | ForEach-Object { & $_ $Server }
}

# -----------------------------
# WPF XAML (wrapped in here-string)
# -----------------------------
$Xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='$($Script:AppName)' Height='760' Width='1150' WindowStartupLocation='CenterScreen'>
  <Window.Resources>
    <!-- Watermark TextBox style -->
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

    <!-- Row style placeholder -->
    <Style x:Key='ResultRowStyle' TargetType='DataGridRow'>
      <Setter Property='Background' Value='White' />
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
    </StackPanel>

    <StackPanel Orientation='Horizontal' Grid.Row='1' Margin='0,0,0,8'>
      <Label Content='Category:' VerticalAlignment='Center' />
      <ComboBox x:Name='cmbCategory' Width='180' Margin='6,0,12,0' />
      <Label Content='Filter:' VerticalAlignment='Center' />
      <TextBox x:Name='txtFilter' Width='320' Margin='6,0,12,0' Style='{StaticResource WatermarkTextBox}' Tag='LDAP filter or text' />
      <Button x:Name='btnSearch' Content='Search' Width='100' Margin='6,0,0,0' />
      <Button x:Name='btnClear' Content='Clear' Width='80' Margin='6,0,0,0' />
    </StackPanel>

    <TabControl x:Name='tabMain' Grid.Row='2'>
      <TabItem Header='Users'>
        <Grid Margin='6'><DataGrid x:Name='dgUsers' AutoGenerateColumns='True' IsReadOnly='True' RowStyle='{StaticResource ResultRowStyle}'/></Grid>
      </TabItem>
      <TabItem Header='Computers'>
        <Grid Margin='6'><DataGrid x:Name='dgComputers' AutoGenerateColumns='True' IsReadOnly='True' RowStyle='{StaticResource ResultRowStyle}'/></Grid>
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
        <Grid Margin='6'><DataGrid x:Name='dgRODCs' AutoGenerateColumns='True' IsReadOnly='True' RowStyle='{StaticResource ResultRowStyle}'/></Grid>
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
    </StatusBar>

  </Grid>
</Window>
"@

# -----------------------------
# Load XAML
# -----------------------------
[xml]$xamlXml = $Xaml
$reader = (New-Object System.Xml.XmlNodeReader $xamlXml)
try {
    $Window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Failed to load XAML: $_"
    # Cleanup runspace pool
    try { $runspacePool.Close(); $runspacePool.Dispose() } catch { }
    throw
}

# Helper to find controls
function Get-Ctrl([string]$name) { $Window.FindName($name) }

# Controls
$cmbDomain = Get-Ctrl "cmbDomain"; $cmbDC = Get-Ctrl "cmbDC"; $btnRefreshDCs = Get-Ctrl "btnRefreshDCs"; $txtExportFolder = Get-Ctrl "txtExportFolder"; $txtStatus = Get-Ctrl "txtStatus"
$cmbCategory = Get-Ctrl "cmbCategory"; $txtFilter = Get-Ctrl "txtFilter"; $btnSearch = Get-Ctrl "btnSearch"; $btnClear = Get-Ctrl "btnClear"
$dgUsers = Get-Ctrl "dgUsers"; $dgComputers = Get-Ctrl "dgComputers"; $dgGroups = Get-Ctrl "dgGroups"; $dgGPOs = Get-Ctrl "dgGPOs"
$dgSubnets = Get-Ctrl "dgSubnets"; $dgRODCs = Get-Ctrl "dgRODCs"; $dgDFS = Get-Ctrl "dgDFS"; $dgDHCP = Get-Ctrl "dgDHCP"; $dgDNS = Get-Ctrl "dgDNS"

# -----------------------------
# Init UI: domains, categories, export folder
# -----------------------------
$cfg = Load-Config
if ($cfg -and $cfg.ExportFolder) { $txtExportFolder.Text = $cfg.ExportFolder } else { $txtExportFolder.Text = $Script:DefaultExportFolder }

try {
    $domains = if ($HasAD) { Get-ForestDomains } else { @() }
    $cmbDomain.ItemsSource = $domains
    if ($domains.Count -gt 0) { $cmbDomain.SelectedIndex = 0 }
} catch { $cmbDomain.ItemsSource = @() }

function Refresh-DCs {
    $domain = $cmbDomain.SelectedItem
    $cmbDC.ItemsSource = @()
    if (-not $domain) { return }
    $txtStatus.Text = "Refreshing DCs..."
    try {
        $dcs = Get-DomainControllers -Domain $domain
        if ($dcs) { $cmbDC.ItemsSource = $dcs | ForEach-Object { $_.HostName }; $cmbDC.SelectedIndex = 0 }
        $txtStatus.Text = "DCs refreshed."
    } catch { $txtStatus.Text = "Unable to enumerate DCs: $($_.Exception.Message)" }
}
$btnRefreshDCs.Add_Click({ Refresh-DCs })
try { Refresh-DCs } catch {}

# Populate categories
$categories = @("Users","Computers","Groups","GPOs","Subnets","RODCs","DFS","DHCP","DNS")
$cmbCategory.ItemsSource = $categories; $cmbCategory.SelectedIndex = 0

# -----------------------------
# Row coloring helper (adds RowClass and sets RowStyle triggers)
# -----------------------------
# We'll enrich objects by adding RowClass property (LockedOut, Disabled, Server, RODC)
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

# Create and apply a RowStyle that checks RowClass
$styleXaml = @"
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
[xml]$sx = $styleXaml
$reader = (New-Object System.Xml.XmlNodeReader $sx)
try { $rowStyle = [Windows.Markup.XamlReader]::Load($reader) } catch { $rowStyle = $null }
if ($rowStyle) {
    $dgUsers.RowStyle = $rowStyle
    $dgComputers.RowStyle = $rowStyle
    $dgRODCs.RowStyle = $rowStyle
}

# -----------------------------
# JSON -> objects helper
# -----------------------------
function Convert-JsonTextToObjects {
    param([string]$text)
    if (-not $text) { return @() }
    $t = $text.Trim()
    try {
        if ($t.StartsWith("[" ) -or $t.StartsWith("{")) {
            return $t | ConvertFrom-Json
        } else {
            # attempt to extract JSON substring
            $start = $t.IndexOf("`n[")
            if ($start -ge 0) { $t = $t.Substring($start+1).Trim() }
            if ($t) { return $t | ConvertFrom-Json } else { return @() }
        }
    } catch { return @() }
}

# -----------------------------
# Async search wrappers
# -----------------------------
function DoUserSearch {
    param([string]$filter, [string]$server, [bool]$allDomains)
    $txtStatus.Text = "Searching users..."
    if ($allDomains) {
        $domains = Get-ForestDomains
        $aggregate = New-Object System.Collections.ArrayList
        foreach ($d in $domains) {
            $dcs = Get-DomainControllers -Domain $d
            $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
            $sbText = SB-SearchUsers -Filter $filter -Server $target
            Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
                param($text,$err)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($err) { $txtStatus.Text = "Error: $($err.Exception.Message)"; return }
                    $objs = Convert-JsonTextToObjects $text
                    if ($objs) {
                        foreach ($o in $objs) { [void]$aggregate.Add($o) }
                        $en = Enrich-WithRowClass -items $aggregate
                        $dgUsers.ItemsSource = $en
                        $txtStatus.Text = "Users: $($en.Count) results (partial)"
                    }
                }))
            }
        }
    } else {
        $sbText = SB-SearchUsers -Filter $filter -Server $server
        Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
            param($text,$err)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($err) { $txtStatus.Text = "User search failed: $($err.Exception.Message)"; $dgUsers.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $text
                $en = Enrich-WithRowClass -items $objs
                $dgUsers.ItemsSource = $en
                $txtStatus.Text = "Users: $($en.Count) results."
            }))
        }
    }
}

function DoComputerSearch {
    param([string]$filter, [string]$server, [bool]$allDomains)
    $txtStatus.Text = "Searching computers..."
    if ($allDomains) {
        $domains = Get-ForestDomains
        $aggregate = New-Object System.Collections.ArrayList
        foreach ($d in $domains) {
            $dcs = Get-DomainControllers -Domain $d
            $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
            $sbText = SB-SearchComputers -Filter $filter -Server $target
            Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
                param($text,$err)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($err) { $txtStatus.Text = "Error: $($err.Exception.Message)"; return }
                    $objs = Convert-JsonTextToObjects $text
                    if ($objs) {
                        foreach ($o in $objs) { [void]$aggregate.Add($o) }
                        $en = Enrich-WithRowClass -items $aggregate
                        $dgComputers.ItemsSource = $en
                        $txtStatus.Text = "Computers: $($en.Count) results (partial)"
                    }
                }))
            }
        }
    } else {
        $sbText = SB-SearchComputers -Filter $filter -Server $server
        Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
            param($text,$err)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($err) { $txtStatus.Text = "Computer search failed: $($err.Exception.Message)"; $dgComputers.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $text
                $en = Enrich-WithRowClass -items $objs
                $dgComputers.ItemsSource = $en
                $txtStatus.Text = "Computers: $($en.Count) results."
            }))
        }
    }
}

function DoGroupSearch {
    param([string]$filter, [string]$server, [bool]$allDomains)
    $txtStatus.Text = "Searching groups..."
    if ($allDomains) {
        $domains = Get-ForestDomains
        $aggregate = New-Object System.Collections.ArrayList
        foreach ($d in $domains) {
            $dcs = Get-DomainControllers -Domain $d
            $target = if ($server) { $server } elseif ($dcs -and $dcs.Count -gt 0) { $dcs[0].HostName } else { $d }
            $sbText = SB-SearchGroups -Filter $filter -Server $target
            Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
                param($text,$err)
                [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                    if ($err) { $txtStatus.Text = "Error: $($err.Exception.Message)"; return }
                    $objs = Convert-JsonTextToObjects $text
                    if ($objs) {
                        foreach ($o in $objs) { [void]$aggregate.Add($o) }
                        $dgGroups.ItemsSource = $aggregate
                        $txtStatus.Text = "Groups: $($aggregate.Count) results (partial)"
                    }
                }))
            }
        }
    } else {
        $sbText = SB-SearchGroups -Filter $filter -Server $server
        Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
            param($text,$err)
            [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
                if ($err) { $txtStatus.Text = "Group search failed: $($err.Exception.Message)"; $dgGroups.ItemsSource = @(); return }
                $objs = Convert-JsonTextToObjects $text
                $dgGroups.ItemsSource = $objs
                $txtStatus.Text = "Groups: $($objs.Count) results."
            }))
        }
    }
}

function DoGPOs {
    param([string]$filter)
    $txtStatus.Text = "Searching GPOs..."
    $sbText = SB-SearchGPOs -Filter $filter
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "GPO search error: $($err.Exception.Message)"; $dgGPOs.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $dgGPOs.ItemsSource = $objs
            $txtStatus.Text = "GPOs: $($objs.Count) results."
        }))
    }
}

function DoSubnets {
    param([string]$server)
    $txtStatus.Text = "Retrieving subnets..."
    $sbText = SB-SearchSubnets -Server $server
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "Subnets retrieval failed: $($err.Exception.Message)"; $dgSubnets.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $dgSubnets.ItemsSource = $objs
            $txtStatus.Text = "Subnets: $($objs.Count) results."
        }))
    }
}

function DoRODCs {
    param([string]$server)
    $txtStatus.Text = "Finding RODCs..."
    $sbText = SB-SearchRODCs -Server $server
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "RODC search failed: $($err.Exception.Message)"; $dgRODCs.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $en = Enrich-WithRowClass -items $objs
            $dgRODCs.ItemsSource = $en
            $txtStatus.Text = "RODCs: $($en.Count) results."
        }))
    }
}

function DoDFS {
    $txtStatus.Text = "Querying DFS..."
    $sbText = SB-SearchDFS
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "DFS query failed: $($err.Exception.Message)"; $dgDFS.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $dgDFS.ItemsSource = $objs
            $txtStatus.Text = "DFS: $($objs.Count) items."
        }))
    }
}

function DoDHCP {
    param([string]$server)
    if (-not $server) { $txtStatus.Text = "Specify DHCP server."; return }
    $txtStatus.Text = "Querying DHCP..."
    $sbText = SB-SearchDHCP -Server $server
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "DHCP query failed: $($err.Exception.Message)"; $dgDHCP.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $dgDHCP.ItemsSource = $objs
            $txtStatus.Text = "DHCP: $($objs.Count) items."
        }))
    }
}

function DoDNS {
    param([string]$server)
    if (-not $server) { $txtStatus.Text = "Specify DNS server."; return }
    $txtStatus.Text = "Querying DNS..."
    $sbText = SB-SearchDNS -Server $server
    Invoke-Async -ScriptBlock { $sbText } -CompletedCallback {
        param($text,$err)
        [void]([System.Windows.Application]::Current.Dispatcher.Invoke({
            if ($err) { $txtStatus.Text = "DNS query failed: $($err.Exception.Message)"; $dgDNS.ItemsSource = @(); return }
            $objs = Convert-JsonTextToObjects $text
            $dgDNS.ItemsSource = $objs
            $txtStatus.Text = "DNS zones: $($objs.Count)."
        }))
    }
}

# -----------------------------
# Wire UI events
# -----------------------------
$btnSearch.Add_Click({
    $cat = $cmbCategory.SelectedItem
    $filter = $txtFilter.Text
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    # default false for all-domains; user can type domain list later if desired
    switch ($cat) {
        "Users" { DoUserSearch -filter $filter -server $server -allDomains:$false }
        "Computers" { DoComputerSearch -filter $filter -server $server -allDomains:$false }
        "Groups" { DoGroupSearch -filter $filter -server $server -allDomains:$false }
        "GPOs" { DoGPOs -filter $filter }
        "Subnets" { DoSubnets -server $server }
        "RODCs" { DoRODCs -server $server }
        "DFS" { DoDFS }
        "DHCP" { DoDHCP -server $txtExportFolder.Text.Trim() }   # placeholder: you can type DHCP server into export box or extend UI
        "DNS" { DoDNS -server $txtExportFolder.Text.Trim() }     # same as DHCP
        default { $txtStatus.Text = "Unknown category" }
    }
})

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

$btnRefreshDCs.Add_Click({ Refresh-DCs })

# Export handler - exports the selected tab's ItemsSource
function Export-ActiveTab {
    $tab = $Window.FindName("tabMain")
    $sel = $tab.SelectedItem
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
    $formats = $Formats.Split(',') | ForEach-Object { $_.Trim() }
    $path = $txtExportFolder.Text.Trim(); if (-not $path) { $path = $Script:DefaultExportFolder }
    Export-Results -Results ($items | ForEach-Object { $_ }) -Category $header -Filter $txtFilter.Text -ExportPath $path -Formats $formats
    Save-Config @{ ExportFolder = $path; Formats = $formats }
    $txtStatus.Text = "Exported $($items.Count) items to $path"
}

# Add Export menu as right-click? simple: map Ctrl+E
$Window.Add_KeyDown({
    param($sender,$e)
    if ($e.KeyboardDevice.Modifiers -eq 'Control' -and $e.Key -eq 'E') { Export-ActiveTab }
})

# Add context menu button for Export (simple approach)
$exportButton = New-Object System.Windows.Controls.Button
$exportButton.Content = "Export Active Tab"
$exportButton.Width = 140
$exportButton.Margin = [System.Windows.Thickness]::new(6,0,0,0)
$exportPanel = [System.Windows.Controls.StackPanel]::new()
# we won't add it to XAML tree—user can use Ctrl+E or Save-Config uses export folder

# -----------------------------
# Scheduled / Headless mode
# -----------------------------
if ($ScheduledMode) {
    if ($Preset) {
        switch ($Preset) {
            "LockedOutUsers" {
                if (-not $HasAD) { Write-Host "ActiveDirectory module required for preset." ; break }
                $locked = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue
                $results = @()
                foreach ($l in $locked) {
                    $u = Get-ADUser -Identity $l.SamAccountName -Properties LockedOut,LastLogonDate -ErrorAction SilentlyContinue
                    if ($u) { $results += [pscustomobject]@{ Name=$u.Name; sAMAccountName=$u.sAMAccountName; LockedOut=$u.LockedOut; LastLogon=$u.LastLogonDate } }
                }
                $formats = $Formats.Split(',') | ForEach-Object { $_.Trim() }
                Export-Results -Results $results -Category $Preset -Filter "" -ExportPath $Script:DefaultExportFolder -Formats $formats
                Write-Host "Exported $($results.Count) items to $Script:DefaultExportFolder"
            }
            default { Write-Host "Preset not implemented: $Preset" }
        }
    } else {
        Write-Host "ScheduledMode requires -Preset argument (e.g., 'LockedOutUsers')."
    }
    try { $runspacePool.Close(); $runspacePool.Dispose() } catch { }
    return
}

# -----------------------------
# Show window
# -----------------------------
try {
    $Window.ShowDialog() | Out-Null
} finally {
    try { $runspacePool.Close(); $runspacePool.Dispose() } catch { }
}
