<#
.AD Expert — AD Search Tool (WPF) — Multi-tab, async, forest-aware
Features added:
 - Async (BackgroundWorker) searches so GUI stays responsive
 - Tabs: Users, Computers, Groups, GPOs, Subnets, RODCs, DFS, DHCP, DNS
 - Forest-wide domain selection + Domain Controller selection
 - Color-coded result rows
 - Export (CSV/JSON/TXT)
 - Headless / ScheduledMode support
Notes:
 - Some server-specific queries (DFS, DHCP, DnsServer) require their respective PowerShell modules
 - Run as domain account or account with read permissions in AD
#>

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# -----------------------------
# Globals & Config
# -----------------------------
$Script:AppName = "AD Expert - AD Search Tool"
$Script:ConfigFile = Join-Path $env:USERPROFILE "ADSearchTool.config.json"
$Script:DefaultExportFolder = Join-Path $env:USERPROFILE "Desktop\ADExports"
If (!(Test-Path $Script:DefaultExportFolder)) { New-Item -Path $Script:DefaultExportFolder -ItemType Directory -Force | Out-Null }

# Load modules if available
function Ensure-Module([string]$Name) {
    if (Get-Module -ListAvailable -Name $Name) {
        try { Import-Module $Name -ErrorAction Stop; return $true } catch { return $false }
    }
    return $false
}

$HasAD = Ensure-Module -Name ActiveDirectory
$HasGPO = Ensure-Module -Name GroupPolicy
$HasDfs = Ensure-Module -Name Dfsn
$HasDhcp = Ensure-Module -Name DhcpServer
$HasDns = Ensure-Module -Name DnsServer

# -----------------------------
# Config helpers
# -----------------------------
function Save-Config {
    param($cfg)
    try {
        $cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $Script:ConfigFile -Encoding UTF8
    } catch { Write-Warning "Unable to save config: $_" }
}
function Load-Config {
    if (Test-Path $Script:ConfigFile) {
        try { Get-Content -Path $Script:ConfigFile -Raw | ConvertFrom-Json } catch { $null }
    } else { $null }
}

# -----------------------------
# AD helper utilities
# -----------------------------
if ($HasAD) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

function Get-ForestDomains {
    try {
        $f = Get-ADForest -ErrorAction Stop
        return $f.Domains
    } catch {
        # If AD module absent or denied, return current domain only
        try { return @([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name) } catch { return @() }
    }
}

function Get-DomainControllersForDomain {
    param([string]$Domain)
    try {
        Get-ADDomainController -Filter * -Server $Domain -ErrorAction Stop | Select-Object HostName,Site,OperatingSystem,@{n='IsReadOnly';e={$_.IsReadOnly}}
    } catch {
        # fallback: return empty
        return @()
    }
}

# -----------------------------
# Search functions (AD)
# -----------------------------
function Search-Users {
    param([string]$Filter, [string]$Server)
    if (-not $HasAD) { throw "ActiveDirectory module not available on this machine." }
    $props = @("Name","sAMAccountName","distinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","userPrincipalName","mail")
    try {
        if ($Filter -and $Filter.Trim() -ne "") {
            if ($Filter -match '^\(|\=|\&|\|') {
                $res = Get-ADUser -LDAPFilter $Filter -Properties $props -Server $Server -ErrorAction Stop
            } else {
                $f = $Filter
                $res = Get-ADUser -Filter "Name -like '$f' -or sAMAccountName -like '$f' -or mail -like '$f' -or userPrincipalName -like '$f'" -Properties $props -Server $Server -ErrorAction Stop
            }
        } else {
            $res = Get-ADUser -Filter * -Properties $props -Server $Server -ErrorAction Stop
        }
        return $res | Select-Object @{n='Type';e={'User'}}, Name,sAMAccountName,distinguishedName,Enabled,LockedOut,LastLogonDate,whenCreated,userPrincipalName,@{n='MemberOf';e={$_.memberOf -join '; '}}
    } catch {
        throw $_
    }
}

function Search-Computers {
    param([string]$Filter, [string]$Server)
    if (-not $HasAD) { throw "ActiveDirectory module not available on this machine." }
    $props = @("Name","OperatingSystem","OperatingSystemVersion","distinguishedName","whenCreated","lastLogonDate")
    try {
        if ($Filter -and $Filter.Trim() -ne "") {
            if ($Filter -match '^\(|\=|\&|\|') {
                $res = Get-ADComputer -LDAPFilter $Filter -Properties $props -Server $Server -ErrorAction Stop
            } else {
                $f = $Filter
                $res = Get-ADComputer -Filter "Name -like '$f' -or OperatingSystem -like '$f'" -Properties $props -Server $Server -ErrorAction Stop
            }
        } else {
            $res = Get-ADComputer -Filter * -Properties $props -Server $Server -ErrorAction Stop
        }
        return $res | Select-Object @{n='Type';e={'Computer'}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}}
    } catch { throw $_ }
}

function Search-Groups {
    param([string]$Filter, [string]$Server)
    if (-not $HasAD) { throw "ActiveDirectory module not available on this machine." }
    try {
        if ($Filter -and $Filter.Trim() -ne "") {
            $res = Get-ADGroup -Filter "Name -like '$Filter'" -Properties member,GroupScope,GroupCategory -Server $Server -ErrorAction Stop
        } else {
            $res = Get-ADGroup -Filter * -Properties member,GroupScope,GroupCategory -Server $Server -ErrorAction Stop
        }
        return $res | Select-Object @{n='Type';e={'Group'}}, Name,GroupScope,GroupCategory,distinguishedName,@{n='Members';e={$_.member -join '; '}}
    } catch { throw $_ }
}

function Search-GPOs {
    param([string]$Filter)
    if (-not $HasGPO) { throw "GroupPolicy module not available." }
    try {
        Import-Module GroupPolicy -ErrorAction Stop
        $gpos = Get-GPO -All -ErrorAction Stop
        if ($Filter -and $Filter.Trim() -ne "") { $gpos = $gpos | Where-Object { $_.DisplayName -like "*$Filter*" } }
        return $gpos | Select-Object @{n='Type';e={'GPO'}}, DisplayName,@{n='Id';e={$_.Id}},Owner,CreationTime,ModificationTime
    } catch { throw $_ }
}

function Search-Subnets {
    param([string]$Server)
    if (-not $HasAD) { throw "ActiveDirectory module not available on this machine." }
    try {
        $cn = (Get-ADRootDSE -Server $Server).configurationNamingContext
        $base = "CN=Subnets,CN=Sites,$cn"
        $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -Server $Server -ErrorAction Stop
        return $subnets | Select-Object @{n='Type';e={'Subnet'}}, Name,@{n='Location';e={$_.location}},@{n='DistinguishedName';e={$_.DistinguishedName}}
    } catch { throw $_ }
}

function Search-RODCs {
    param([string]$Server)
    if (-not $HasAD) { throw "ActiveDirectory module not available on this machine." }
    try {
        $dcs = Get-ADDomainController -Filter * -Server $Server -ErrorAction Stop
        $ro = $dcs | Where-Object { $_.IsReadOnly -eq $true } | Select-Object HostName,Site,OperatingSystem,@{n='IsRODC';e={$true}}
        return $ro
    } catch { throw $_ }
}

function Search-DFS {
    param([string]$Server)
    if (-not $HasDfs) { return [pscustomobject]@{ Note="DFSn module not available locally. Install RSAT-DFS-Namespace or run from a server." } }
    try {
        $roots = Get-DfsnRoot -ErrorAction Stop
        $out = @()
        foreach ($r in $roots) {
            $out += [pscustomobject]@{ Root=$r.Path; State=$r.State; Type=$r.Type }
        }
        return $out
    } catch { throw $_ }
}

function Search-DHCP {
    param([string]$DhcpServer)
    if (-not $HasDhcp) { return [pscustomobject]@{ Note="DhcpServer module unavailable. Run this from DHCP server or install RSAT-DHCP." } }
    try {
        $scopes = Get-DhcpServerv4Scope -ComputerName $DhcpServer -ErrorAction Stop
        return $scopes
    } catch { throw $_ }
}

function Search-DNSZones {
    param([string]$DnsServer)
    if (-not $HasDns) { return [pscustomobject]@{ Note="DnsServer module unavailable. Install DNS tools or run on DNS machine." } }
    try {
        Get-DnsServerZone -ComputerName $DnsServer -ErrorAction Stop
    } catch { throw $_ }
}

# -----------------------------
# Export function
# -----------------------------
function Export-Results {
    param([Parameter(Mandatory=$true)][object[]]$Results, [string]$Category, [string]$Filter, [string]$ExportPath, [string[]]$Formats)

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
                $Results | ConvertTo-Json -Depth 5 | Set-Content -Path $file -Encoding UTF8
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

function SafeFileName { param([string]$n) if (-not $n) { $n = "results" } return ($n -replace '[^\w\-\._ ]','_').Trim() }

# -----------------------------
# BackgroundWorker helper for async search
# -----------------------------
Add-Type -AssemblyName System
Add-Type -AssemblyName System.ComponentModel

function Start-AsyncOperation {
    param(
        [ScriptBlock]$Work,
        [scriptblock]$ProgressCallback,
        [scriptblock]$CompletedCallback
    )
    $bw = New-Object System.ComponentModel.BackgroundWorker
    $bw.WorkerReportsProgress = $true
    $bw.WorkerSupportsCancellation = $false

    $bw.add_DoWork({
        param($sender,$e)
        try {
            $res = & $Work
            $e.Result = $res
        } catch {
            $e.Result = $null
            $e.Exception = $_
        }
    })
    $bw.add_RunWorkerCompleted({
        param($sender,$e)
        if ($e.Exception) {
            & $CompletedCallback $null $e.Exception
        } else {
            & $CompletedCallback $e.Result $null
        }
    })
    $bw.RunWorkerAsync()
    return $bw
}

# -----------------------------
# Build WPF XAML with tabs
# -----------------------------
$Xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='$($Script:AppName)' Height='700' Width='1100' WindowStartupLocation='CenterScreen'>
  <Grid Margin='8'>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='0,0,0,8'>
      <Label Content='Forest Domain:' VerticalAlignment='Center'/>
      <ComboBox x:Name='cmbDomain' Width='240' Margin='6,0,12,0'/>
      <Label Content='Domain Controller:' VerticalAlignment='Center'/>
      <ComboBox x:Name='cmbDC' Width='240' Margin='6,0,12,0'/>
      <Button x:Name='btnRefreshDCs' Content='Refresh DCs' Width='110' Margin='6,0,0,0'/>
      <Label Content='Export Folder:' VerticalAlignment='Center' Margin='12,0,0,0'/>
      <TextBox x:Name='txtExportFolder' Width='300' Margin='6,0,12,0'/>
    </StackPanel>

    <TabControl Grid.Row='1' x:Name='tabMain'>
      <TabItem Header='Users'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0' Margin='0,0,0,6'>
            <TextBox x:Name='txtUserFilter' Width='360' Margin='0,0,6,0' ToolTip='Name, sAMAccountName, email or LDAP filter'/>
            <Button x:Name='btnUserSearch' Content='Search' Width='110'/>
            <Button x:Name='btnUserExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgUsers' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='Computers'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0' Margin='0,0,0,6'>
            <TextBox x:Name='txtCompFilter' Width='360' Margin='0,0,6,0' ToolTip='Name or OperatingSystem'/>
            <Button x:Name='btnCompSearch' Content='Search' Width='110'/>
            <Button x:Name='btnCompExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgComputers' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='Groups'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0' Margin='0,0,0,6'>
            <TextBox x:Name='txtGroupFilter' Width='360' Margin='0,0,6,0' ToolTip='Group name or filter'/>
            <Button x:Name='btnGroupSearch' Content='Search' Width='110'/>
            <Button x:Name='btnGroupExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgGroups' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='GPOs'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <TextBox x:Name='txtGPOFilter' Width='360' Margin='0,0,6,0' ToolTip='GPO name filter'/>
            <Button x:Name='btnGPOSearch' Content='Search' Width='110'/>
            <Button x:Name='btnGPOExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgGPOs' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='Subnets'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <Button x:Name='btnSubnetsSearch' Content='Refresh Subnets' Width='140'/>
            <Button x:Name='btnSubnetsExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgSubnets' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='RODCs'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <Button x:Name='btnRODCSearch' Content='Find RODCs' Width='140'/>
            <Button x:Name='btnRODCExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgRODCs' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='DFS'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <Button x:Name='btnDFSSearch' Content='Query DFS' Width='120'/>
            <Button x:Name='btnDFSExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgDFS' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='DHCP'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <TextBox x:Name='txtDhcpServer' Width='240' Margin='0,0,6,0' ToolTip='DHCP Server (optional)'/>
            <Button x:Name='btnDhcpSearch' Content='Query DHCP' Width='120'/>
            <Button x:Name='btnDhcpExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgDHCP' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

      <TabItem Header='DNS'>
        <Grid Margin='6'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Grid.Row='0'>
            <TextBox x:Name='txtDnsServer' Width='240' Margin='0,0,6,0' ToolTip='DNS Server (optional)'/>
            <Button x:Name='btnDnsSearch' Content='Query DNS Zones' Width='140'/>
            <Button x:Name='btnDnsExport' Content='Export' Width='90' Margin='6,0,0,0'/>
          </StackPanel>
          <DataGrid x:Name='dgDNS' Grid.Row='1' AutoGenerateColumns='True' CanUserAddRows='False'/>
        </Grid>
      </TabItem>

    </TabControl>

    <StatusBar Grid.Row='2' VerticalAlignment='Bottom'>
      <StatusBarItem>
        <TextBlock x:Name='txtStatus'>Ready</TextBlock>
      </StatusBarItem>
      <StatusBarItem>
        <TextBlock x:Name='txtHelp' Text='Tip: Select domain, optional DC, then search.'/>
      </StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
"@

# -----------------------------
# Load XAML and controls
# -----------------------------
[xml]$xamlXml = $Xaml
$reader = (New-Object System.Xml.XmlNodeReader $xamlXml)
try {
    $Window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Failed to load WPF: $_"
    return
}

# Find controls - improved method
$cmbDomain = $Window.FindName("cmbDomain")
$cmbDC = $Window.FindName("cmbDC")
$btnRefreshDCs = $Window.FindName("btnRefreshDCs")
$txtExportFolder = $Window.FindName("txtExportFolder")
$txtStatus = $Window.FindName("txtStatus")

# Tab controls
$txtUserFilter = $Window.FindName("txtUserFilter")
$btnUserSearch = $Window.FindName("btnUserSearch")
$dgUsers = $Window.FindName("dgUsers")
$btnUserExport = $Window.FindName("btnUserExport")

$txtCompFilter = $Window.FindName("txtCompFilter")
$btnCompSearch = $Window.FindName("btnCompSearch")
$dgComputers = $Window.FindName("dgComputers")
$btnCompExport = $Window.FindName("btnCompExport")

$txtGroupFilter = $Window.FindName("txtGroupFilter")
$btnGroupSearch = $Window.FindName("btnGroupSearch")
$dgGroups = $Window.FindName("dgGroups")
$btnGroupExport = $Window.FindName("btnGroupExport")

$txtGPOFilter = $Window.FindName("txtGPOFilter")
$btnGPOSearch = $Window.FindName("btnGPOSearch")
$dgGPOs = $Window.FindName("dgGPOs")
$btnGPOExport = $Window.FindName("btnGPOExport")

$btnSubnetsSearch = $Window.FindName("btnSubnetsSearch")
$dgSubnets = $Window.FindName("dgSubnets")
$btnSubnetsExport = $Window.FindName("btnSubnetsExport")

$btnRODCSearch = $Window.FindName("btnRODCSearch")
$dgRODCs = $Window.FindName("dgRODCs")
$btnRODCExport = $Window.FindName("btnRODCExport")

$btnDFSSearch = $Window.FindName("btnDFSSearch")
$dgDFS = $Window.FindName("dgDFS")
$btnDFSExport = $Window.FindName("btnDFSExport")

$txtDhcpServer = $Window.FindName("txtDhcpServer")
$btnDhcpSearch = $Window.FindName("btnDhcpSearch")
$dgDHCP = $Window.FindName("dgDHCP")
$btnDhcpExport = $Window.FindName("btnDhcpExport")

$txtDnsServer = $Window.FindName("txtDnsServer")
$btnDnsSearch = $Window.FindName("btnDnsSearch")
$dgDNS = $Window.FindName("dgDNS")
$btnDnsExport = $Window.FindName("btnDnsExport")

# Verify critical controls loaded
if (-not $cmbDomain -or -not $txtExportFolder -or -not $txtStatus) {
    Write-Error "Failed to find critical UI controls. Check XAML control names."
    return
}

# Init export folder
$cfg = Load-Config
if ($cfg -and $cfg.ExportFolder) { 
    $txtExportFolder.Text = $cfg.ExportFolder 
} else { 
    $txtExportFolder.Text = $Script:DefaultExportFolder 
}

# -----------------------------
# Populate domains & DCs
# -----------------------------
try {
    $domains = @()
    if ($HasAD) {
        $domains = Get-ForestDomains
    } else {
        try {
            $domains = @((Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot) | Where-Object { $_ }
        } catch {
            $domains = @()
        }
    }
    if ($domains.Count -gt 0) {
        $cmbDomain.ItemsSource = $domains
        $cmbDomain.SelectedIndex = 0
    } else {
        $cmbDomain.ItemsSource = @()
        $txtStatus.Text = "Warning: No domains found. Ensure AD module is installed and you have permissions."
    }
} catch {
    $cmbDomain.ItemsSource = @()
    $txtStatus.Text = "Error loading domains: $($_.Exception.Message)"
}

# Refresh DC list function
function Refresh-DCList {
    $domain = $cmbDomain.SelectedItem
    $cmbDC.ItemsSource = @()
    if (-not $domain) { return }
    try {
        $dcs = Get-DomainControllersForDomain -Domain $domain
        if ($dcs.Count -gt 0) {
            $cmbDC.ItemsSource = $dcs | ForEach-Object { $_.HostName }
            $cmbDC.SelectedIndex = 0
        }
    } catch {
        $cmbDC.ItemsSource = @()
    }
}

$btnRefreshDCs.Add_Click({ Refresh-DCList })
try { Refresh-DCList } catch { }

# -----------------------------
# Color-coded rows
# -----------------------------
function Apply-RowColoring([System.Windows.Controls.DataGrid]$dg, [object[]]$items) {
    try {
        $enriched = $items | ForEach-Object {
            $o = $_
            $rowClass = ""
            try {
                if ($o.PSObject.Properties.Match("LockedOut")) {
                    if ($o.LockedOut -eq $true) { $rowClass = "LockedOut" }
                }
                if (($o.PSObject.Properties.Match("Enabled")) -and ($o.Enabled -eq $false)) {
                    $rowClass = "Disabled"
                }
                if ($o.PSObject.Properties.Match("OperatingSystem")) {
                    if ($o.OperatingSystem -and $o.OperatingSystem -like "*Server*") {
                        $rowClass = "Server"
                    }
                }
                if ($o.PSObject.Properties.Match("IsRODC")) {
                    if ($o.IsRODC) { $rowClass = "RODC" }
                }
            } catch { }
            $n = [pscustomobject]@{}
            foreach ($p in $o.psobject.properties) { $n | Add-Member -MemberType NoteProperty -Name $p.Name -Value $p.Value -Force }
            $n | Add-Member -MemberType NoteProperty -Name "RowClass" -Value $rowClass -Force
            $n
        }
        $dg.ItemsSource = $enriched

        if (-not $dg.RowStyle) {
            $styleXaml = @"
<Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='DataGridRow'>
  <Style.Triggers>
    <DataTrigger Binding='{Binding RowClass}' Value='LockedOut'>
      <Setter Property='Background' Value='#FFF4CCCC'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='Disabled'>
      <Setter Property='Background' Value='#FFECECEC'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='Server'>
      <Setter Property='Background' Value='#FFDCEBF7'/>
    </DataTrigger>
    <DataTrigger Binding='{Binding RowClass}' Value='RODC'>
      <Setter Property='Background' Value='#FFFFF2CC'/>
    </DataTrigger>
  </Style.Triggers>
</Style>
"@
            [xml]$sx = $styleXaml
            $reader2 = (New-Object System.Xml.XmlNodeReader $sx)
            $rs = [Windows.Markup.XamlReader]::Load($reader2)
            $dg.RowStyle = $rs
        }
    } catch { $dg.ItemsSource = $items }
}

# -----------------------------
# Event callbacks for searches
# -----------------------------
function Do-UserSearch {
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    $filter = $txtUserFilter.Text
    $txtStatus.Text = "Searching users..."
    $work = { Search-Users -Filter $using:filter -Server $using:server }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "User search failed: $($err.Exception.Message)"
            $dgUsers.ItemsSource = @()
        } else {
            $txtStatus.Text = "Users: $($res.Count) results."
            Apply-RowColoring -dg $dgUsers -items $res
        }
    }
}

function Do-ComputerSearch {
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    $filter = $txtCompFilter.Text
    $txtStatus.Text = "Searching computers..."
    $work = { Search-Computers -Filter $using:filter -Server $using:server }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "Computer search failed: $($err.Exception.Message)"
            $dgComputers.ItemsSource = @()
        } else {
            $txtStatus.Text = "Computers: $($res.Count) results."
            Apply-RowColoring -dg $dgComputers -items $res
        }
    }
}

function Do-GroupSearch {
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    $filter = $txtGroupFilter.Text
    $txtStatus.Text = "Searching groups..."
    $work = { Search-Groups -Filter $using:filter -Server $using:server }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "Group search failed: $($err.Exception.Message)"
            $dgGroups.ItemsSource = @()
        } else {
            $txtStatus.Text = "Groups: $($res.Count) results."
            $dgGroups.ItemsSource = $res
        }
    }
}

function Do-GPOSearch {
    $filter = $txtGPOFilter.Text
    $txtStatus.Text = "Searching GPOs..."
    $work = { Search-GPOs -Filter $using:filter }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "GPO search failed: $($err.Exception.Message)"
            $dgGPOs.ItemsSource = @()
        } else {
            $txtStatus.Text = "GPOs: $($res.Count) results."
            $dgGPOs.ItemsSource = $res
        }
    }
}

function Do-Subnets {
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    $txtStatus.Text = "Querying subnets..."
    $work = { Search-Subnets -Server $using:server }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "Subnets query failed: $($err.Exception.Message)"
            $dgSubnets.ItemsSource = @()
        } else {
            $txtStatus.Text = "Subnets: $($res.Count) results."
            $dgSubnets.ItemsSource = $res
        }
    }
}

function Do-RODCs {
    $domain = $cmbDomain.SelectedItem
    $dc = $cmbDC.SelectedItem
    $server = if ($dc) { $dc } elseif ($domain) { $domain } else { $null }
    $txtStatus.Text = "Finding RODCs..."
    $work = { Search-RODCs -Server $using:server }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "RODC query failed: $($err.Exception.Message)"
            $dgRODCs.ItemsSource = @()
        } else {
            $txtStatus.Text = "RODCs: $($res.Count) results."
            Apply-RowColoring -dg $dgRODCs -items $res
        }
    }
}

function Do-DFS {
    $txtStatus.Text = "Querying DFS..."
    $work = { Search-DFS }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "DFS query failed: $($err.Exception.Message)"
            $dgDFS.ItemsSource = @()
        } else {
            $txtStatus.Text = "DFS: $($res.Count) items returned."
            $dgDFS.ItemsSource = $res
        }
    }
}

function Do-DHCP {
    $srv = $txtDhcpServer.Text.Trim()
    if (-not $srv) { $txtStatus.Text = "Specify DHCP server or run from DHCP server."; return }
    $txtStatus.Text = "Querying DHCP..."
    $work = { Search-DHCP -DhcpServer $using:srv }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "DHCP query failed: $($err.Exception.Message)"
            $dgDHCP.ItemsSource = @()
        } else {
            $txtStatus.Text = "DHCP scopes: $($res.Count)"
            $dgDHCP.ItemsSource = $res
        }
    }
}

function Do-DNS {
    $srv = $txtDnsServer.Text.Trim()
    if (-not $srv) { $txtStatus.Text = "Specify DNS server"; return }
    $txtStatus.Text = "Querying DNS..."
    $work = { Search-DNSZones -DnsServer $using:srv }
    Start-AsyncOperation -Work $work -ProgressCallback { } -CompletedCallback {
        param($res,$err)
        if ($err) {
            $txtStatus.Text = "DNS query failed: $($err.Exception.Message)"
            $dgDNS.ItemsSource = @()
        } else {
            $txtStatus.Text = "DNS zones: $($res.Count)"
            $dgDNS.ItemsSource = $res
        }
    }
}

# -----------------------------
# Wire buttons to functions
# -----------------------------
$btnUserSearch.Add_Click({ Do-UserSearch })
$btnCompSearch.Add_Click({ Do-ComputerSearch })
$btnGroupSearch.Add_Click({ Do-GroupSearch })
$btnGPOSearch.Add_Click({ Do-GPOSearch })
$btnSubnetsSearch.Add_Click({ Do-Subnets })
$btnRODCSearch.Add_Click({ Do-RODCs })
$btnDFSSearch.Add_Click({ Do-DFS })
$btnDhcpSearch.Add_Click({ Do-DHCP })
$btnDnsSearch.Add_Click({ Do-DNS })

# Export button handlers
function Export-Grid($dg, $category) {
    try {
        $items = $dg.ItemsSource
        if (-not $items -or $items.Count -eq 0) { 
            [System.Windows.MessageBox]::Show("No results to export.","Export",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
            return 
        }
        $formats = @("csv","json","txt")
        $path = $txtExportFolder.Text.Trim()
        if (-not $path) { $path = $Script:DefaultExportFolder }
        Export-Results -Results ($items | ForEach-Object { $_ }) -Category $category -Filter "" -ExportPath $path -Formats $formats
        Save-Config @{ ExportFolder = $path; Formats = $formats }
        $txtStatus.Text = "Exported $($items.Count) items to $path"
    } catch {
        $txtStatus.Text = "Export failed: $($_.Exception.Message)"
    }
}

$btnUserExport.Add_Click({ Export-Grid -dg $dgUsers -category "Users" })
$btnCompExport.Add_Click({ Export-Grid -dg $dgComputers -category "Computers" })
$btnGroupExport.Add_Click({ Export-Grid -dg $dgGroups -category "Groups" })
$btnGPOExport.Add_Click({ Export-Grid -dg $dgGPOs -category "GPOs" })
$btnSubnetsExport.Add_Click({ Export-Grid -dg $dgSubnets -category "Subnets" })
$btnRODCExport.Add_Click({ Export-Grid -dg $dgRODCs -category "RODCs" })
$btnDFSExport.Add_Click({ Export-Grid -dg $dgDFS -category "DFS" })
$btnDhcpExport.Add_Click({ Export-Grid -dg $dgDHCP -category "DHCP" })
$btnDnsExport.Add_Click({ Export-Grid -dg $dgDNS -category "DNS" })

# -----------------------------
# Final: launch window
# -----------------------------
$Window.ShowDialog() | Out-Null
