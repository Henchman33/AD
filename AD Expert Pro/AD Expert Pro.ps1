<#
.SYNOPSIS
  AD Expert Pro - Advanced Multi-Tab WPF Active Directory Management Tool
.DESCRIPTION
  A professional, modern WPF-based Active Directory tool featuring:
    - Dark/Light theme toggle
    - Users, Computers, Servers, Groups, Security Groups, GPOs, Subnets,
      Trusts, DFS, Domain Controllers, T0/T1 Accounts, MSAs, Service Accounts,
      DHCP, DNS tabs
    - Connect to any domain or DC with alternate credentials
    - Subdomain credential management
    - RDP launcher
    - File/Folder copy to remote servers (with credential support)
    - RunspacePool async searches
    - Color-coded DataGrid rows
    - Export (CSV/JSON/TXT)
    - Headless / Scheduled mode support
.NOTES
  Requires PowerShell 5.1 + RSAT (ActiveDirectory, GroupPolicy, DhcpServer, DnsServer, Dfsn)
#>

param(
    [switch]$ScheduledMode,
    [string]$Preset       = "",
    [string]$ExportFolderArg = "",
    [string]$Formats      = "csv"
)

# ─────────────────────────────────────────────
#  ASSEMBLIES
# ─────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms

# ─────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────
$Script:AppName    = "AD Expert Pro"
$Script:ConfigFile = Join-Path $env:USERPROFILE "ADExpertPro.config.json"
$Script:IsDark     = $true   # default theme

$ExportConfig = @{
    BaseExportPath    = Join-Path $env:USERPROFILE "Desktop"
    FolderNamePattern = "ADReport_{1}_{0}"
    TimeStampFormat   = "yyyyMMdd_HHmmss"
    DefaultReportName = "ADExport"
}

function Get-ExportFolder {
    param([string]$ReportName = $ExportConfig.DefaultReportName)
    $ts     = (Get-Date).ToString($ExportConfig.TimeStampFormat)
    $folder = $ExportConfig.FolderNamePattern -f $ts,$ReportName
    $full   = Join-Path $ExportConfig.BaseExportPath $folder
    if (!(Test-Path $full)) { New-Item -Path $full -ItemType Directory -Force | Out-Null }
    return $full
}

$Script:DefaultExportFolder = if ($ExportFolderArg -and $ExportFolderArg.Trim()) { $ExportFolderArg } else { Get-ExportFolder -ReportName "Initial" }
if (!(Test-Path $Script:DefaultExportFolder)) { New-Item -Path $Script:DefaultExportFolder -ItemType Directory -Force | Out-Null }

# ─────────────────────────────────────────────
#  CONFIG HELPERS
# ─────────────────────────────────────────────
function Save-Config { param($cfg) try { $cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $Script:ConfigFile -Encoding UTF8 } catch { Write-Warning "Save-Config: $_" } }
function Load-Config { if (Test-Path $Script:ConfigFile) { try { Get-Content -Path $Script:ConfigFile -Raw | ConvertFrom-Json } catch { $null } } else { $null } }

# ─────────────────────────────────────────────
#  MODULE CHECKS
# ─────────────────────────────────────────────
function Ensure-ModuleLoaded { param([string]$Name) if (Get-Module -ListAvailable -Name $Name) { Import-Module $Name -ErrorAction SilentlyContinue; return $true } else { return $false } }
$HasAD   = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO  = Ensure-ModuleLoaded -Name GroupPolicy
$HasDFS  = Ensure-ModuleLoaded -Name Dfsn
$HasDHCP = Ensure-ModuleLoaded -Name DhcpServer
$HasDNS  = Ensure-ModuleLoaded -Name DnsServer

# ─────────────────────────────────────────────
#  RUNSPACE POOL
# ─────────────────────────────────────────────
$minT = 1; $maxT = 8
$runspacePool = [runspacefactory]::CreateRunspacePool($minT,$maxT)
$runspacePool.ThreadOptions = "ReuseThread"
$runspacePool.Open()

function Invoke-Async {
    param([ScriptBlock]$ScriptBlock,[Parameter(Mandatory)][ScriptBlock]$CompletedCallback)
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $runspacePool
    $ps.AddScript($ScriptBlock) | Out-Null
    $ar = $ps.BeginInvoke()
    [System.Threading.ThreadPool]::QueueUserWorkItem({
        param($ps,$ar,$cb)
        try   { $out = $ps.EndInvoke($ar); & $cb $out $null }
        catch { & $cb $null $_ }
        finally { $ps.Dispose() }
    }, @($ps,$ar,$CompletedCallback)) | Out-Null
}

# ─────────────────────────────────────────────
#  EXPORT / UTILITY
# ─────────────────────────────────────────────
function SafeFileName { param([string]$n) if (-not $n){ $n="results" }; return ($n -replace '[^\w\-\._ ]','_').Trim() }

function Export-Results {
    param([object[]]$Results,[string]$Category,[string]$Filter,[string]$ExportPath,[string[]]$Formats)
    if (!(Test-Path $ExportPath)){ New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
    $ts   = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $base = SafeFileName("$Category`_$Filter`_$ts")
    foreach ($fmt in $Formats) {
        switch ($fmt.ToLower()) {
            "csv"  { $Results | Export-Csv  -Path (Join-Path $ExportPath "$base.csv")  -NoTypeInformation -Force }
            "json" { $Results | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $ExportPath "$base.json") -Encoding UTF8 }
            "txt"  { $Results | Out-String | Set-Content -Path (Join-Path $ExportPath "$base.txt") -Encoding UTF8 }
            default{ $Results | Out-String | Set-Content -Path (Join-Path $ExportPath "$base.txt") -Encoding UTF8 }
        }
    }
}

# ─────────────────────────────────────────────
#  AD DISCOVERY
# ─────────────────────────────────────────────
function Get-ForestDomains {
    if (-not $HasAD){ return @() }
    try { (Get-ADForest -ErrorAction Stop).Domains } catch { @() }
}
function Get-DomainControllers {
    param([string]$Domain,[System.Management.Automation.PSCredential]$Credential=$null)
    if (-not $HasAD){ return @() }
    try {
        $p = @{ Filter='*'; Server=$Domain; ErrorAction='Stop' }
        if ($Credential){ $p.Credential = $Credential }
        Get-ADDomainController @p
    } catch { @() }
}

# ─────────────────────────────────────────────
#  CREDENTIAL STORE (per-domain)
# ─────────────────────────────────────────────
$Script:CredStore = @{}   # key = domain, value = PSCredential

# ─────────────────────────────────────────────
#  AD SEARCH FUNCTIONS
# ─────────────────────────────────────────────
function SearchUsers {
    param([string]$Filter,[string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $props = @("Name","sAMAccountName","DistinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","userPrincipalName","mail","Description","Department","Title")
    $p = @{ Properties=$props; Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    if ($Filter -and $Filter.Trim()) {
        if ($Filter -match '^\(|\=|\&|\|') { $p.LDAPFilter = $Filter; Get-ADUser @p }
        else { $f=$Filter; Get-ADUser -Filter "Name -like '$f' -or sAMAccountName -like '$f' -or mail -like '$f' -or userPrincipalName -like '$f'" @p }
    } else { Get-ADUser -Filter * @p }
     Select-Object Name,sAMAccountName,Enabled,LockedOut,LastLogonDate,Department,Title,Description,DistinguishedName
}

function SearchComputers {
    param([string]$Filter,[string]$Server,[System.Management.Automation.PSCredential]$Cred=$null,[switch]$ServersOnly)
    $props = @("Name","OperatingSystem","OperatingSystemVersion","DistinguishedName","whenCreated","LastLogonDate","IPv4Address","Description")
    $p = @{ Properties=$props; Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    if ($ServersOnly) {
        $res = Get-ADComputer -Filter "OperatingSystem -like '*Server*'" @p
    } elseif ($Filter -and $Filter.Trim()) {
        if ($Filter -match '^\(|\=|\&|\|') { $res = Get-ADComputer -LDAPFilter $Filter @p }
        else { $f=$Filter; $res = Get-ADComputer -Filter "Name -like '$f' -or OperatingSystem -like '$f'" @p }
    } else { $res = Get-ADComputer -Filter * @p }
    $res | Select-Object Name,OperatingSystem,OperatingSystemVersion,IPv4Address,LastLogonDate,Description,DistinguishedName
}

function SearchGroups {
    param([string]$Filter,[string]$Server,[System.Management.Automation.PSCredential]$Cred=$null,[switch]$SecurityOnly)
    $p = @{ Properties=@("member","GroupScope","GroupCategory","Description","whenCreated","ManagedBy"); Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    if ($SecurityOnly) {
        $res = Get-ADGroup -Filter "GroupCategory -eq 'Security'" @p
    } elseif ($Filter -and $Filter.Trim()) {
        $res = Get-ADGroup -Filter "Name -like '$Filter'" @p
    } else { $res = Get-ADGroup -Filter * @p }
    $res | Select-Object Name,GroupScope,GroupCategory,@{n='MemberCount';e={$_.member.Count}},Description,ManagedBy,@{n='Members';e={($_.member | Select-Object -First 5) -join '; '}}
}

function SearchGPOs {
    param([string]$Filter,[string]$Domain,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ All=$true; ErrorAction='SilentlyContinue' }
    if ($Domain){ $p.Domain = $Domain }
    if ($Cred){ $p.Server = $Domain }
    $gpos = Get-GPO @p
    if ($Filter -and $Filter.Trim()){ $gpos = $gpos | Where-Object { $_.DisplayName -like "*$Filter*" } }
    $gpos | Select-Object DisplayName,GpoStatus,Id,Owner,CreationTime,ModificationTime,@{n='LinksTo';e={'(use Get-GPOReport for link info)'}}
}

function SearchSubnets {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Server=$Server; ErrorAction='Stop' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        $cn   = (Get-ADRootDSE @p).configurationNamingContext
        $base = "CN=Subnets,CN=Sites,$cn"
        $sp   = @{ SearchBase=$base; Filter='*'; Properties=@("name","location","siteObject","description"); Server=$Server; ErrorAction='SilentlyContinue' }
        if ($Cred){ $sp.Credential = $Cred }
        Get-ADObject @sp | Select-Object Name,@{n='Location';e={$_.location}},@{n='Description';e={$_.description}},@{n='SiteObject';e={$_.siteObject}},DistinguishedName
    } catch { @() }
}

function SearchTrusts {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Identity=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        Get-ADTrust -Filter * -Server $Server -ErrorAction SilentlyContinue |
            Select-Object Name,Direction,TrustType,DisallowTransivity,IntraForest,SelectiveAuthentication,Source,Target
    } catch { @() }
}

function SearchRODCs {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Filter='*'; Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        (Get-ADDomainController @p) | Where-Object { $_.IsReadOnly } |
            Select-Object HostName,Site,OperatingSystem,IPv4Address,@{n='IsRODC';e={$true}},@{n='Enabled';e={$true}}
    } catch { @() }
}

function SearchAllDCs {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Filter='*'; Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        Get-ADDomainController @p | Select-Object HostName,Site,OperatingSystem,IPv4Address,
            @{n='IsRODC';e={$_.IsReadOnly}},@{n='IsGC';e={$_.IsGlobalCatalog}},Forest,Domain,Enabled
    } catch { @() }
}

# T0 = DA/EA/schema admin tier. T1 = server/infra admins. Simple heuristic by group membership description.
function SearchTierAccounts {
    param([string]$Server,[int]$Tier,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        # Tier-0: members of highly-privileged groups
        $t0Groups = @("Domain Admins","Enterprise Admins","Schema Admins","Administrators","Group Policy Creator Owners")
        # Tier-1: common infra admin groups
        $t1Groups = @("Server Operators","Backup Operators","Account Operators","Network Configuration Operators","Print Operators","Remote Desktop Users")

        $targetGroups = if ($Tier -eq 0) { $t0Groups } else { $t1Groups }
        $members = [System.Collections.ArrayList]@()
        foreach ($grpName in $targetGroups) {
            try {
                $gp = @{ Identity=$grpName; Server=$Server; ErrorAction='SilentlyContinue' }
                if ($Cred){ $gp.Credential = $Cred }
                $grp = Get-ADGroup @gp -Properties member
                if ($grp -and $grp.member) {
                    foreach ($dn in $grp.member) {
                        try {
                            $up = @{ Identity=$dn; Properties=@("Name","sAMAccountName","Enabled","LockedOut","LastLogonDate","Description"); Server=$Server; ErrorAction='SilentlyContinue' }
                            if ($Cred){ $up.Credential = $Cred }
                            $u = Get-ADUser @up
                            if ($u) {
                                [void]$members.Add([PSCustomObject]@{
                                    Name            = $u.Name
                                    sAMAccountName  = $u.sAMAccountName
                                    Enabled         = $u.Enabled
                                    LockedOut       = $u.LockedOut
                                    LastLogonDate   = $u.LastLogonDate
                                    Description     = $u.Description
                                    TierGroup       = $grpName
                                    Tier            = "T$Tier"
                                })
                            }
                        } catch {}
                    }
                }
            } catch {}
        }
        return $members | Sort-Object sAMAccountName -Unique
    } catch { @() }
}

function SearchMSAs {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ Filter='ObjectClass -eq "msDS-GroupManagedServiceAccount" -or ObjectClass -eq "msDS-ManagedServiceAccount"'; Properties=@("Name","sAMAccountName","DistinguishedName","whenCreated","Description","msDS-HostServiceAccount","Enabled"); Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        Get-ADObject @p | Select-Object Name,sAMAccountName,Enabled,@{n='Type';e={if ($_.ObjectClass -eq "msDS-GroupManagedServiceAccount"){"gMSA"}else{"sMSA"}}},
            whenCreated,Description,@{n='HostedBy';e={$_."msDS-HostServiceAccount" -join '; '}},DistinguishedName
    } catch { @() }
}

function SearchServiceAccounts {
    param([string]$Filter,[string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $props = @("Name","sAMAccountName","DistinguishedName","Enabled","LastLogonDate","Description","ServicePrincipalNames","PasswordNeverExpires","PasswordLastSet")
    $p = @{ Properties=$props; Server=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    # Heuristic: accounts with SPNs or named svc*/service* or description containing service
    try {
        $res = if ($Filter -and $Filter.Trim()) {
            Get-ADUser -Filter "Name -like '$Filter' -or sAMAccountName -like '$Filter'" @p
        } else {
            Get-ADUser -Filter "ServicePrincipalNames -like '*' -or sAMAccountName -like 'svc*' -or sAMAccountName -like 'sa_*' -or sAMAccountName -like 'srv*'" @p
        }
        $res | Select-Object Name,sAMAccountName,Enabled,PasswordNeverExpires,PasswordLastSet,LastLogonDate,
            @{n='SPNCount';e={$_.ServicePrincipalNames.Count}},@{n='SPNs';e={($_.ServicePrincipalNames | Select-Object -First 3) -join '; '}},Description,DistinguishedName
    } catch { @() }
}

function SearchDFS {
    param([System.Management.Automation.PSCredential]$Cred=$null)
    try {
        $p = @{ ErrorAction='SilentlyContinue' }
        $roots = Get-DfsnRoot @p
        if (-not $roots){ return @() }
        $roots | Select-Object Path,State,Type,TimeToLive,Description
    } catch { @() }
}

function SearchDHCP {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ ComputerName=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        $scopes = Get-DhcpServerv4Scope @p
        if (-not $scopes){ return @() }
        $scopes | ForEach-Object {
            $stats = $null
            try { $stats = Get-DhcpServerv4ScopeStatistics -ScopeId $_.ScopeId -ComputerName $Server -ErrorAction SilentlyContinue } catch {}
            [PSCustomObject]@{
                ScopeId    = $_.ScopeId
                Name       = $_.Name
                StartRange = $_.StartRange
                EndRange   = $_.EndRange
                SubnetMask = $_.SubnetMask
                State      = $_.State
                InUse      = if ($stats){ $stats.InUse } else { "N/A" }
                Free       = if ($stats){ $stats.Free } else { "N/A" }
                PercentInUse = if ($stats){ "$([math]::Round($stats.PercentageInUse,1))%" } else { "N/A" }
            }
        }
    } catch { @() }
}

function SearchDNS {
    param([string]$Server,[System.Management.Automation.PSCredential]$Cred=$null)
    $p = @{ ComputerName=$Server; ErrorAction='SilentlyContinue' }
    if ($Cred){ $p.Credential = $Cred }
    try {
        $zones = Get-DnsServerZone @p
        if (-not $zones){ return @() }
        $zones | Select-Object ZoneName,ZoneType,IsDsIntegrated,IsReverseLookupZone,DynamicUpdate,ReplicationScope,
            @{n='RecordCount';e={ try{(Get-DnsServerResourceRecord -ZoneName $_.ZoneName -ComputerName $Server -EA SilentlyContinue).Count}catch{"N/A"}}}
    } catch { @() }
}

# ─────────────────────────────────────────────
#  THEME DEFINITIONS
# ─────────────────────────────────────────────
# Dark: deep slate/charcoal + indigo/purple accents + yellow highlights
# Light: clean off-white + blue/purple accents
$Script:Themes = @{
    Dark = @{
        WindowBg          = "#1E1F2E"
        PanelBg           = "#252638"
        TabBg             = "#1A1B2A"
        TabSelected       = "#6C63FF"
        TabHover          = "#2D2E45"
        ControlBg         = "#2D2E45"
        ControlBorder     = "#4A4B6A"
        ControlFocus      = "#6C63FF"
        ButtonBg          = "#6C63FF"
        ButtonHover       = "#7C73FF"
        ButtonFg          = "#FFFFFF"
        DangerBg          = "#E53E3E"
        SuccessBg         = "#38A169"
        AccentYellow      = "#F6E05E"
        AccentCyan        = "#76E4F7"
        TextPrimary       = "#E8E9F3"
        TextSecondary     = "#9898B5"
        TextMuted         = "#5A5B7A"
        GridHeaderBg      = "#1A1B2A"
        GridAltRow        = "#232436"
        GridBorder        = "#3A3B5A"
        StatusBg          = "#141520"
        SeparatorColor    = "#3A3B5A"
        HeaderGradStart   = "#6C63FF"
        HeaderGradEnd     = "#9F7AEA"
        RowLockedOut      = "#4A1515"
        RowDisabled       = "#2A2A3A"
        RowServer         = "#1A2A40"
        RowRODC           = "#2A2A15"
        RowT0             = "#3A1530"
        RowT1             = "#1A2E2A"
        ScrollBg          = "#1E1F2E"
        ScrollThumb       = "#4A4B6A"
        SectionLabel      = "#6C63FF"
    }
    Light = @{
        WindowBg          = "#F0F2FA"
        PanelBg           = "#FFFFFF"
        TabBg             = "#E8EAF6"
        TabSelected       = "#5C54E8"
        TabHover          = "#D8DAF0"
        ControlBg         = "#FFFFFF"
        ControlBorder     = "#C5C9E8"
        ControlFocus      = "#5C54E8"
        ButtonBg          = "#5C54E8"
        ButtonHover       = "#6C64F8"
        ButtonFg          = "#FFFFFF"
        DangerBg          = "#C53030"
        SuccessBg         = "#276749"
        AccentYellow      = "#D69E2E"
        AccentCyan        = "#2B6CB0"
        TextPrimary       = "#1A1B2E"
        TextSecondary     = "#4A4B6A"
        TextMuted         = "#888AAA"
        GridHeaderBg      = "#E8EAF6"
        GridAltRow        = "#F8F9FF"
        GridBorder        = "#D0D3EC"
        StatusBg          = "#E0E3F8"
        SeparatorColor    = "#D0D3EC"
        HeaderGradStart   = "#5C54E8"
        HeaderGradEnd     = "#7C6FE8"
        RowLockedOut      = "#FED7D7"
        RowDisabled       = "#EDF2F7"
        RowServer         = "#EBF8FF"
        RowRODC           = "#FEFCBF"
        RowT0             = "#FED7E2"
        RowT1             = "#C6F6D5"
        ScrollBg          = "#F0F2FA"
        ScrollThumb       = "#C5C9E8"
        SectionLabel      = "#5C54E8"
    }
}

# ─────────────────────────────────────────────
#  XAML TEMPLATE
# ─────────────────────────────────────────────
function Get-XAML {
    param([hashtable]$T)  # T = theme colors
    return @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='$($Script:AppName)'
    Height='840' Width='1400'
    WindowStartupLocation='CenterScreen'
    Background='$($T.WindowBg)'
    FontFamily='Segoe UI'
    FontSize='13'>

  <Window.Resources>
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
                <Setter TargetName='bd' Property='Opacity' Value='0.85'/>
              </Trigger>
              <Trigger Property='IsEnabled' Value='False'>
                <Setter TargetName='bd' Property='Opacity' Value='0.45'/>
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

    <Style x:Key='ModernTextBox' TargetType='TextBox'>
      <Setter Property='Background' Value='$($T.ControlBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
      <Setter Property='CaretBrush' Value='$($T.TextPrimary)'/>
      <Setter Property='SelectionBrush' Value='$($T.ButtonBg)'/>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='TextBox'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='6'>
              <ScrollViewer Margin='0' x:Name='PART_ContentHost' VerticalAlignment='Center'/>
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

    <Style x:Key='ModernPasswordBox' TargetType='PasswordBox'>
      <Setter Property='Background' Value='$($T.ControlBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
      <Setter Property='CaretBrush' Value='$($T.TextPrimary)'/>
      <Setter Property='Template'>
        <Setter.Value>
          <ControlTemplate TargetType='PasswordBox'>
            <Border x:Name='bd' Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='{TemplateBinding BorderThickness}' CornerRadius='6'>
              <ScrollViewer Margin='0' x:Name='PART_ContentHost' VerticalAlignment='Center'/>
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

    <Style x:Key='ModernComboBox' TargetType='ComboBox'>
      <Setter Property='Background' Value='$($T.ControlBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.ControlBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='Padding' Value='8,6'/>
    </Style>

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

    <Style x:Key='ModernDataGrid' TargetType='DataGrid'>
      <Setter Property='Background' Value='$($T.PanelBg)'/>
      <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
      <Setter Property='BorderBrush' Value='$($T.GridBorder)'/>
      <Setter Property='BorderThickness' Value='1'/>
      <Setter Property='GridLinesVisibility' Value='Horizontal'/>
      <Setter Property='HorizontalGridLinesBrush' Value='$($T.GridBorder)'/>
      <Setter Property='RowBackground' Value='$($T.PanelBg)'/>
      <Setter Property='AlternatingRowBackground' Value='$($T.GridAltRow)'/>
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
            <Setter Property='SeparatorBrush' Value='$($T.GridBorder)'/>
          </Style>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key='ModernCheckBox' TargetType='CheckBox'>
      <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='VerticalAlignment' Value='Center'/>
      <Setter Property='Margin' Value='4,0'/>
    </Style>

    <Style x:Key='TabLabel' TargetType='TextBlock'>
      <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
      <Setter Property='FontSize' Value='12'/>
      <Setter Property='FontWeight' Value='SemiBold'/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='56'/>   <!-- Header bar -->
      <RowDefinition Height='Auto'/> <!-- Connection bar -->
      <RowDefinition Height='*'/>    <!-- Main tabs -->
      <RowDefinition Height='30'/>   <!-- Status bar -->
    </Grid.RowDefinitions>

    <!-- ═══ HEADER BAR ═══ -->
    <Border Grid.Row='0' Background='$($T.PanelBg)' BorderBrush='$($T.SeparatorColor)' BorderThickness='0,0,0,1'>
      <Grid Margin='16,0'>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='*'/>
          <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>

        <!-- Logo / Title -->
        <StackPanel Orientation='Horizontal' VerticalAlignment='Center' Grid.Column='0'>
          <Border Width='8' Height='28' Background='$($T.ButtonBg)' CornerRadius='4' Margin='0,0,10,0'/>
          <TextBlock Text='AD Expert Pro' FontSize='18' FontWeight='Bold' Foreground='$($T.TextPrimary)' VerticalAlignment='Center'/>
          <Border Background='$($T.ButtonBg)' CornerRadius='4' Margin='10,0,0,0' Padding='6,2'>
            <TextBlock Text='RSAT' FontSize='10' FontWeight='Bold' Foreground='#FFFFFF'/>
          </Border>
        </StackPanel>

        <!-- Quick Actions -->
        <StackPanel Orientation='Horizontal' HorizontalAlignment='Right' VerticalAlignment='Center' Grid.Column='2' Margin='0,0,8,0'>
          <Button x:Name='btnRDP' Style='{StaticResource ModernButton}' Margin='0,0,8,0' Background='$($T.AccentCyan)' Foreground='$($T.WindowBg)'>
            <StackPanel Orientation='Horizontal'>
              <TextBlock Text='&#xE8D5;' FontFamily='Segoe MDL2 Assets' Margin='0,0,6,0' VerticalAlignment='Center'/>
              <TextBlock Text='RDP Connect' VerticalAlignment='Center'/>
            </StackPanel>
          </Button>
          <Button x:Name='btnFileCopy' Style='{StaticResource ModernButton}' Margin='0,0,8,0' Background='$($T.SuccessBg)'>
            <StackPanel Orientation='Horizontal'>
              <TextBlock Text='&#xE8C8;' FontFamily='Segoe MDL2 Assets' Margin='0,0,6,0' VerticalAlignment='Center'/>
              <TextBlock Text='Copy Files' VerticalAlignment='Center'/>
            </StackPanel>
          </Button>
          <Button x:Name='btnToggleTheme' Style='{StaticResource ModernButton}' Background='$($T.TabHover)' Foreground='$($T.TextPrimary)' Margin='0,0,8,0' ToolTip='Toggle Dark/Light Theme'>
            <TextBlock x:Name='txtThemeIcon' Text='&#xE793;' FontFamily='Segoe MDL2 Assets' FontSize='15'/>
          </Button>
          <Button x:Name='btnNewExportFolder' Style='{StaticResource ModernButton}' Background='$($T.TabHover)' Foreground='$($T.AccentYellow)' ToolTip='Create new timestamped export folder'>
            <StackPanel Orientation='Horizontal'>
              <TextBlock Text='&#xE7C3;' FontFamily='Segoe MDL2 Assets' Margin='0,0,6,0' VerticalAlignment='Center'/>
              <TextBlock Text='New Folder' VerticalAlignment='Center'/>
            </StackPanel>
          </Button>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ═══ CONNECTION BAR ═══ -->
    <Border Grid.Row='1' Background='$($T.PanelBg)' Margin='0,4,0,0' Padding='12,8'>
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='Auto'/> <!-- Domain -->
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/> <!-- DC -->
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/> <!-- Manual DC -->
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/> <!-- Cred section -->
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='Auto'/>
          <ColumnDefinition Width='*'/>    <!-- Export path -->
          <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column='0' Margin='0,0,6,0'>
          <TextBlock Text='DOMAIN' Style='{StaticResource SectionLabel}'/>
          <ComboBox x:Name='cmbDomain' Width='220' Style='{StaticResource ModernComboBox}'/>
        </StackPanel>

        <StackPanel Grid.Column='1' Margin='0,0,6,0'>
          <TextBlock Text=' ' Style='{StaticResource SectionLabel}'/>
          <Button x:Name='btnRefreshDCs' Style='{StaticResource ModernButton}' Content='Refresh DCs' Width='100' Background='$($T.TabHover)' Foreground='$($T.TextPrimary)'/>
        </StackPanel>

        <StackPanel Grid.Column='2' Margin='0,0,6,0'>
          <TextBlock Text='DOMAIN CONTROLLER' Style='{StaticResource SectionLabel}'/>
          <ComboBox x:Name='cmbDC' Width='240' Style='{StaticResource ModernComboBox}'/>
        </StackPanel>

        <StackPanel Grid.Column='3' Margin='0,0,6,0'>
          <TextBlock Text='OR MANUAL DC/SERVER' Style='{StaticResource SectionLabel}'/>
          <TextBox x:Name='txtManualServer' Width='200' Style='{StaticResource ModernTextBox}' ToolTip='Type any DC or server hostname to connect directly'/>
        </StackPanel>

        <StackPanel Grid.Column='4' Margin='0,0,14,0'>
          <TextBlock Text=' ' Style='{StaticResource SectionLabel}'/>
          <Button x:Name='btnConnect' Style='{StaticResource ModernButton}' Content='Connect' Width='90'/>
        </StackPanel>

        <!-- Separator -->
        <Border Grid.Column='5' Width='1' Background='$($T.SeparatorColor)' Margin='4,4,12,4'/>

        <!-- Credentials -->
        <StackPanel Grid.Column='6' Margin='0,0,6,0'>
          <TextBlock Text='CREDENTIALS (optional)' Style='{StaticResource SectionLabel}'/>
          <TextBox x:Name='txtCredUser' Width='160' Style='{StaticResource ModernTextBox}' ToolTip='DOMAIN\Username or UPN'/>
        </StackPanel>
        <StackPanel Grid.Column='7' Margin='0,0,6,0'>
          <TextBlock Text='PASSWORD' Style='{StaticResource SectionLabel}'/>
          <PasswordBox x:Name='pwdCredPass' Width='130' Style='{StaticResource ModernPasswordBox}'/>
        </StackPanel>
        <StackPanel Grid.Column='8' Margin='0,0,14,0'>
          <TextBlock Text=' ' Style='{StaticResource SectionLabel}'/>
          <Button x:Name='btnSetCred' Style='{StaticResource ModernButton}' Content='Set Cred' Width='90' Background='$($T.AccentYellow)' Foreground='$($T.WindowBg)'/>
        </StackPanel>

        <!-- Separator -->
        <Border Grid.Column='9' Width='1' Background='$($T.SeparatorColor)' Margin='4,4,12,4'/>

        <!-- Export path -->
        <StackPanel Grid.Column='10' Margin='0,0,6,0'>
          <TextBlock Text='EXPORT FOLDER' Style='{StaticResource SectionLabel}'/>
          <TextBox x:Name='txtExportFolder' Width='260' Style='{StaticResource ModernTextBox}'/>
        </StackPanel>

        <StackPanel Grid.Column='11' Margin='0,0,6,0'>
          <TextBlock Text=' ' Style='{StaticResource SectionLabel}'/>
          <Button x:Name='btnBrowseExport' Style='{StaticResource ModernButton}' Content='Browse' Width='70' Background='$($T.TabHover)' Foreground='$($T.TextPrimary)'/>
        </StackPanel>

        <!-- Active cred display -->
        <StackPanel Grid.Column='12' Margin='12,0,0,0' VerticalAlignment='Center'>
          <TextBlock Text='ACTIVE CREDENTIAL' Style='{StaticResource SectionLabel}'/>
          <TextBlock x:Name='txtActiveCred' Text='(using current Windows session)' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center'/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ═══ MAIN TABCONTROL ═══ -->
    <TabControl Grid.Row='2' x:Name='tabMain' Margin='0,6,0,0'
                Background='$($T.TabBg)'
                BorderBrush='$($T.GridBorder)' BorderThickness='0,1,0,0'
                TabStripPlacement='Left'>

      <TabControl.Resources>
        <Style TargetType='TabItem'>
          <Setter Property='Background' Value='Transparent'/>
          <Setter Property='Foreground' Value='$($T.TextSecondary)'/>
          <Setter Property='BorderThickness' Value='0'/>
          <Setter Property='Padding' Value='14,10'/>
          <Setter Property='FontSize' Value='12'/>
          <Setter Property='FontWeight' Value='SemiBold'/>
          <Setter Property='MinWidth' Value='140'/>
          <Setter Property='Template'>
            <Setter.Value>
              <ControlTemplate TargetType='TabItem'>
                <Border x:Name='bd' Background='{TemplateBinding Background}' BorderThickness='3,0,0,0' BorderBrush='Transparent' Padding='{TemplateBinding Padding}' Margin='0,1'>
                  <ContentPresenter ContentSource='Header' HorizontalAlignment='Left' VerticalAlignment='Center'/>
                </Border>
                <ControlTemplate.Triggers>
                  <Trigger Property='IsSelected' Value='True'>
                    <Setter TargetName='bd' Property='Background' Value='$($T.PanelBg)'/>
                    <Setter TargetName='bd' Property='BorderBrush' Value='$($T.TabSelected)'/>
                    <Setter Property='Foreground' Value='$($T.TextPrimary)'/>
                  </Trigger>
                  <Trigger Property='IsMouseOver' Value='True'>
                    <Setter TargetName='bd' Property='Background' Value='$($T.TabHover)'/>
                  </Trigger>
                </ControlTemplate.Triggers>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Style>
      </TabControl.Resources>

      <!-- ── USERS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE77B;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Users'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtUserFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0' ToolTip='Name, sAMAccountName, email, or LDAP filter like (objectClass=user)'/>
            <Button x:Name='btnUserSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search' Margin='0,0,6,0'/>
            <Button x:Name='btnUserExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <CheckBox x:Name='chkUserAllDomains' Style='{StaticResource ModernCheckBox}' Content='All domains in forest' Margin='8,0,0,0'/>
            <CheckBox x:Name='chkUserEnabledOnly' Style='{StaticResource ModernCheckBox}' Content='Enabled only' Margin='8,0,0,0'/>
            <TextBlock x:Name='txtUserCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgUsers' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── COMPUTERS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE7EF;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Computers'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtCompFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0' ToolTip='Computer name or OS filter'/>
            <Button x:Name='btnCompSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search' Margin='0,0,6,0'/>
            <Button x:Name='btnCompExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <CheckBox x:Name='chkCompAllDomains' Style='{StaticResource ModernCheckBox}' Content='All domains' Margin='8,0,0,0'/>
            <TextBlock x:Name='txtCompCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgComputers' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── SERVERS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xECCB;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Servers'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtServerFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0' ToolTip='Server name filter'/>
            <Button x:Name='btnServerSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search Servers' Margin='0,0,6,0'/>
            <Button x:Name='btnServerExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtServerCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgServers' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── GROUPS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE902;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Groups'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtGroupFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0'/>
            <Button x:Name='btnGroupSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search' Margin='0,0,6,0'/>
            <Button x:Name='btnGroupExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <CheckBox x:Name='chkGroupAllDomains' Style='{StaticResource ModernCheckBox}' Content='All domains' Margin='8,0,0,0'/>
            <TextBlock x:Name='txtGroupCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgGroups' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── SECURITY GROUPS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE72E;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Security Groups'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtSecGrpFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0'/>
            <Button x:Name='btnSecGrpSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search Security Groups' Margin='0,0,6,0'/>
            <Button x:Name='btnSecGrpExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtSecGrpCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgSecGrps' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── GPOs ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE9E9;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='GPOs'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtGPOFilter' Style='{StaticResource ModernTextBox}' Width='380' Margin='0,0,8,0'/>
            <Button x:Name='btnGPOSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search GPOs' Margin='0,0,6,0'/>
            <Button x:Name='btnGPOExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtGPOCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgGPOs' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── SUBNETS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE968;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Subnets'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <Button x:Name='btnSubnetsSearch' Style='{StaticResource ModernButton}' Content='&#xE72C; Refresh Subnets' Margin='0,0,6,0'/>
            <Button x:Name='btnSubnetsExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtSubnetCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgSubnets' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── TRUSTS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE8F4;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Trusts'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <Button x:Name='btnTrustsSearch' Style='{StaticResource ModernButton}' Content='&#xE72C; Enumerate Trusts' Margin='0,0,6,0'/>
            <Button x:Name='btnTrustsExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtTrustCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgTrusts' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── DOMAIN CONTROLLERS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xECCB;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Domain Controllers'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <Button x:Name='btnDCSearch' Style='{StaticResource ModernButton}' Content='&#xE72C; List All DCs' Margin='0,0,6,0'/>
            <Button x:Name='btnRODCSearch' Style='{StaticResource ModernButton}' Content='RODCs Only' Margin='0,0,6,0' Background='$($T.AccentYellow)' Foreground='$($T.WindowBg)'/>
            <Button x:Name='btnDCExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtDCCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgDCs' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── T0 ACCOUNTS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE902;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='T0 Accounts'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <Border Grid.Row='0' Background='$($T.RowT0)' CornerRadius='6' Padding='10,6' Margin='0,0,0,8'>
            <TextBlock Foreground='$($T.TextPrimary)' FontSize='12'>
              <Run FontWeight='Bold' Text='Tier-0 (T0): '/>
              <Run Text='Domain Admins, Enterprise Admins, Schema Admins, Administrators, Group Policy Creator Owners. Highest privilege accounts.'/>
            </TextBlock>
          </Border>
          <WrapPanel Grid.Row='1' Margin='0,0,0,8'>
            <Button x:Name='btnT0Search' Style='{StaticResource ModernButton}' Content='&#xE72C; Find T0 Accounts' Margin='0,0,6,0'/>
            <Button x:Name='btnT0Export' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtT0Count' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgT0' Grid.Row='2' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── T1 ACCOUNTS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE902;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='T1 Accounts'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <Border Grid.Row='0' Background='$($T.RowT1)' CornerRadius='6' Padding='10,6' Margin='0,0,0,8'>
            <TextBlock Foreground='$($T.TextPrimary)' FontSize='12'>
              <Run FontWeight='Bold' Text='Tier-1 (T1): '/>
              <Run Text='Server Operators, Backup Operators, Account Operators, Network Configuration Operators. Server/infrastructure admins.'/>
            </TextBlock>
          </Border>
          <WrapPanel Grid.Row='1' Margin='0,0,0,8'>
            <Button x:Name='btnT1Search' Style='{StaticResource ModernButton}' Content='&#xE72C; Find T1 Accounts' Margin='0,0,6,0'/>
            <Button x:Name='btnT1Export' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtT1Count' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgT1' Grid.Row='2' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── MSAs ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE7EE;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='MSAs'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <Button x:Name='btnMSASearch' Style='{StaticResource ModernButton}' Content='&#xE72C; Find MSAs / gMSAs' Margin='0,0,6,0'/>
            <Button x:Name='btnMSAExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtMSACount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgMSAs' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── SERVICE ACCOUNTS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE90F;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='Service Accounts'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBox x:Name='txtSvcAcctFilter' Style='{StaticResource ModernTextBox}' Width='300' Margin='0,0,8,0' ToolTip='Filter svc/service accounts - leave blank for all SPNs + svc* prefix accounts'/>
            <Button x:Name='btnSvcAcctSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Search' Margin='0,0,6,0'/>
            <Button x:Name='btnSvcAcctExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtSvcAcctCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgSvcAccts' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── DFS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xEC50;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='DFS'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <Button x:Name='btnDFSSearch' Style='{StaticResource ModernButton}' Content='&#xE72C; Query DFS Roots' Margin='0,0,6,0'/>
            <Button x:Name='btnDFSExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtDFSCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgDFS' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── DHCP ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE968;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='DHCP'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBlock Text='DHCP Server:' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,6,0'/>
            <TextBox x:Name='txtDhcpServer' Style='{StaticResource ModernTextBox}' Width='220' Margin='0,0,8,0' ToolTip='Hostname or IP of DHCP server (not necessarily a DC)'/>
            <Button x:Name='btnDhcpSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Query DHCP Scopes' Margin='0,0,6,0'/>
            <Button x:Name='btnDhcpExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtDHCPCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgDHCP' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

      <!-- ── DNS ── -->
      <TabItem>
        <TabItem.Header>
          <StackPanel Orientation='Horizontal'>
            <TextBlock Text='&#xE774;' FontFamily='Segoe MDL2 Assets' Margin='0,0,8,0'/>
            <TextBlock Text='DNS'/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin='12' Background='$($T.PanelBg)'>
          <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/></Grid.RowDefinitions>
          <WrapPanel Grid.Row='0' Margin='0,0,0,8'>
            <TextBlock Text='DNS Server (DC):' Style='{StaticResource SectionLabel}' VerticalAlignment='Center' Margin='0,0,6,0'/>
            <TextBox x:Name='txtDnsServer' Style='{StaticResource ModernTextBox}' Width='220' Margin='0,0,8,0' ToolTip='DNS server (typically a DC) hostname or IP'/>
            <Button x:Name='btnDnsSearch' Style='{StaticResource ModernButton}' Content='&#xE721; Query DNS Zones' Margin='0,0,6,0'/>
            <Button x:Name='btnDnsExport' Style='{StaticResource ModernButton}' Content='&#xE74E; Export' Margin='0,0,6,0' Background='$($T.SuccessBg)'/>
            <TextBlock x:Name='txtDNSCount' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center' Margin='12,0,0,0'/>
          </WrapPanel>
          <DataGrid x:Name='dgDNS' Grid.Row='1' Style='{StaticResource ModernDataGrid}' AutoGenerateColumns='True' CanUserAddRows='False' IsReadOnly='True'/>
        </Grid>
      </TabItem>

    </TabControl>

    <!-- ═══ STATUS BAR ═══ -->
    <Border Grid.Row='3' Background='$($T.StatusBg)' BorderBrush='$($T.SeparatorColor)' BorderThickness='0,1,0,0'>
      <Grid Margin='12,0'>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='*'/>
          <ColumnDefinition Width='Auto'/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name='txtStatus' Foreground='$($T.TextSecondary)' FontSize='11' VerticalAlignment='Center' TextTrimming='CharacterEllipsis' Text='Ready — Select a domain and click Search or Refresh on any tab.'/>
        <StackPanel Orientation='Horizontal' Grid.Column='1' VerticalAlignment='Center'>
          <TextBlock x:Name='txtConnectedTo' Foreground='$($T.AccentYellow)' FontSize='11' FontWeight='SemiBold' VerticalAlignment='Center' Margin='0,0,12,0'/>
          <TextBlock Text='AD Expert Pro v3.0' Foreground='$($T.TextMuted)' FontSize='11' VerticalAlignment='Center'/>
        </StackPanel>
      </Grid>
    </Border>

  </Grid>
</Window>
"@
}

# ─────────────────────────────────────────────
#  LOAD WINDOW
# ─────────────────────────────────────────────
function Initialize-Window {
    param([hashtable]$T)
    $xamlString = Get-XAML -T $T
    [xml]$xamlXml = $xamlString
    $reader = New-Object System.Xml.XmlNodeReader $xamlXml
    try { [Windows.Markup.XamlReader]::Load($reader) }
    catch { Write-Error "XAML load failed: $_"; throw }
}

$T      = if ($Script:IsDark) { $Script:Themes.Dark } else { $Script:Themes.Light }
$Window = Initialize-Window -T $T

function Get-Ctrl([string]$n) { $Window.FindName($n) }

# ─────────────────────────────────────────────
#  BIND ALL CONTROLS
# ─────────────────────────────────────────────
$cmbDomain        = Get-Ctrl "cmbDomain"
$cmbDC            = Get-Ctrl "cmbDC"
$btnRefreshDCs    = Get-Ctrl "btnRefreshDCs"
$txtManualServer  = Get-Ctrl "txtManualServer"
$btnConnect       = Get-Ctrl "btnConnect"
$txtCredUser      = Get-Ctrl "txtCredUser"
$pwdCredPass      = Get-Ctrl "pwdCredPass"
$btnSetCred       = Get-Ctrl "btnSetCred"
$txtActiveCred    = Get-Ctrl "txtActiveCred"
$txtExportFolder  = Get-Ctrl "txtExportFolder"
$btnNewExportFolder = Get-Ctrl "btnNewExportFolder"
$btnBrowseExport  = Get-Ctrl "btnBrowseExport"
$txtStatus        = Get-Ctrl "txtStatus"
$txtConnectedTo   = Get-Ctrl "txtConnectedTo"
$btnToggleTheme   = Get-Ctrl "btnToggleTheme"
$txtThemeIcon     = Get-Ctrl "txtThemeIcon"
$btnRDP           = Get-Ctrl "btnRDP"
$btnFileCopy      = Get-Ctrl "btnFileCopy"

# Grids and controls per tab
$txtUserFilter  = Get-Ctrl "txtUserFilter"
$btnUserSearch  = Get-Ctrl "btnUserSearch"
$btnUserExport  = Get-Ctrl "btnUserExport"
$dgUsers        = Get-Ctrl "dgUsers"
$chkUserAllDomains = Get-Ctrl "chkUserAllDomains"
$chkUserEnabledOnly = Get-Ctrl "chkUserEnabledOnly"
$txtUserCount   = Get-Ctrl "txtUserCount"

$txtCompFilter  = Get-Ctrl "txtCompFilter"
$btnCompSearch  = Get-Ctrl "btnCompSearch"
$btnCompExport  = Get-Ctrl "btnCompExport"
$dgComputers    = Get-Ctrl "dgComputers"
$chkCompAllDomains = Get-Ctrl "chkCompAllDomains"
$txtCompCount   = Get-Ctrl "txtCompCount"

$txtServerFilter = Get-Ctrl "txtServerFilter"
$btnServerSearch = Get-Ctrl "btnServerSearch"
$btnServerExport = Get-Ctrl "btnServerExport"
$dgServers       = Get-Ctrl "dgServers"
$txtServerCount  = Get-Ctrl "txtServerCount"

$txtGroupFilter  = Get-Ctrl "txtGroupFilter"
$btnGroupSearch  = Get-Ctrl "btnGroupSearch"
$btnGroupExport  = Get-Ctrl "btnGroupExport"
$dgGroups        = Get-Ctrl "dgGroups"
$chkGroupAllDomains = Get-Ctrl "chkGroupAllDomains"
$txtGroupCount   = Get-Ctrl "txtGroupCount"

$txtSecGrpFilter = Get-Ctrl "txtSecGrpFilter"
$btnSecGrpSearch = Get-Ctrl "btnSecGrpSearch"
$btnSecGrpExport = Get-Ctrl "btnSecGrpExport"
$dgSecGrps       = Get-Ctrl "dgSecGrps"
$txtSecGrpCount  = Get-Ctrl "txtSecGrpCount"

$txtGPOFilter   = Get-Ctrl "txtGPOFilter"
$btnGPOSearch   = Get-Ctrl "btnGPOSearch"
$btnGPOExport   = Get-Ctrl "btnGPOExport"
$dgGPOs         = Get-Ctrl "dgGPOs"
$txtGPOCount    = Get-Ctrl "txtGPOCount"

$btnSubnetsSearch = Get-Ctrl "btnSubnetsSearch"
$btnSubnetsExport = Get-Ctrl "btnSubnetsExport"
$dgSubnets        = Get-Ctrl "dgSubnets"
$txtSubnetCount   = Get-Ctrl "txtSubnetCount"

$btnTrustsSearch = Get-Ctrl "btnTrustsSearch"
$btnTrustsExport = Get-Ctrl "btnTrustsExport"
$dgTrusts        = Get-Ctrl "dgTrusts"
$txtTrustCount   = Get-Ctrl "txtTrustCount"

$btnDCSearch     = Get-Ctrl "btnDCSearch"
$btnRODCSearch   = Get-Ctrl "btnRODCSearch"
$btnDCExport     = Get-Ctrl "btnDCExport"
$dgDCs           = Get-Ctrl "dgDCs"
$txtDCCount      = Get-Ctrl "txtDCCount"

$btnT0Search = Get-Ctrl "btnT0Search"; $btnT0Export = Get-Ctrl "btnT0Export"; $dgT0 = Get-Ctrl "dgT0"; $txtT0Count = Get-Ctrl "txtT0Count"
$btnT1Search = Get-Ctrl "btnT1Search"; $btnT1Export = Get-Ctrl "btnT1Export"; $dgT1 = Get-Ctrl "dgT1"; $txtT1Count = Get-Ctrl "txtT1Count"
$btnMSASearch = Get-Ctrl "btnMSASearch"; $btnMSAExport = Get-Ctrl "btnMSAExport"; $dgMSAs = Get-Ctrl "dgMSAs"; $txtMSACount = Get-Ctrl "txtMSACount"
$txtSvcAcctFilter = Get-Ctrl "txtSvcAcctFilter"; $btnSvcAcctSearch = Get-Ctrl "btnSvcAcctSearch"; $btnSvcAcctExport = Get-Ctrl "btnSvcAcctExport"; $dgSvcAccts = Get-Ctrl "dgSvcAccts"; $txtSvcAcctCount = Get-Ctrl "txtSvcAcctCount"

$btnDFSSearch = Get-Ctrl "btnDFSSearch"; $btnDFSExport = Get-Ctrl "btnDFSExport"; $dgDFS = Get-Ctrl "dgDFS"; $txtDFSCount = Get-Ctrl "txtDFSCount"
$txtDhcpServer = Get-Ctrl "txtDhcpServer"; $btnDhcpSearch = Get-Ctrl "btnDhcpSearch"; $btnDhcpExport = Get-Ctrl "btnDhcpExport"; $dgDHCP = Get-Ctrl "dgDHCP"; $txtDHCPCount = Get-Ctrl "txtDHCPCount"
$txtDnsServer  = Get-Ctrl "txtDnsServer";  $btnDnsSearch  = Get-Ctrl "btnDnsSearch";  $btnDnsExport  = Get-Ctrl "btnDnsExport";  $dgDNS  = Get-Ctrl "dgDNS";  $txtDNSCount  = Get-Ctrl "txtDNSCount"

# ─────────────────────────────────────────────
#  HELPERS - GET ACTIVE SERVER / CRED
# ─────────────────────────────────────────────
function Get-ActiveServer {
    $manual = $txtManualServer.Text.Trim()
    if ($manual){ return $manual }
    $dc     = $cmbDC.SelectedItem
    if ($dc){ return $dc }
    $domain = $cmbDomain.SelectedItem
    return $domain
}

function Get-ActiveCred {
    $domain = $cmbDomain.SelectedItem
    if ($Script:CredStore.ContainsKey($domain)){ return $Script:CredStore[$domain] }
    return $null
}

function Set-StatusText { param([string]$msg) $Window.Dispatcher.Invoke([action]{ $txtStatus.Text = $msg }) }
function Set-CountText  { param($ctrl,[int]$n,[string]$label="results") $ctrl.Text = "$n $label" }

# ─────────────────────────────────────────────
#  INIT CONFIG + DOMAINS
# ─────────────────────────────────────────────
$cfg = Load-Config
if ($cfg -and $cfg.ExportFolder){ $txtExportFolder.Text = $cfg.ExportFolder } else { $txtExportFolder.Text = $Script:DefaultExportFolder }
if ($cfg -and $cfg.Theme){ $Script:IsDark = ($cfg.Theme -eq "Dark") }

try {
    if ($HasAD) { $domains = Get-ForestDomains } else { $domains = @() }
    $cmbDomain.ItemsSource = $domains
    if ($domains.Count -gt 0){ $cmbDomain.SelectedIndex = 0 }
} catch { $cmbDomain.ItemsSource = @() }

function Refresh-DCs {
    $domain = $cmbDomain.SelectedItem
    $cmbDC.ItemsSource = @()
    if (-not $domain){ return }
    $txtStatus.Text = "Refreshing DC list for $domain..."
    $cred = Get-ActiveCred
    try {
        $dcs = Get-DomainControllers -Domain $domain -Credential $cred
        if ($dcs){ $cmbDC.ItemsSource = $dcs | ForEach-Object { $_.HostName }; $cmbDC.SelectedIndex = 0 }
        $txtStatus.Text = "DC list refreshed: $($dcs.Count) DCs in $domain"
        $txtConnectedTo.Text = "Connected: $domain"
    } catch { $txtStatus.Text = "Unable to enumerate DCs: $($_.Exception.Message)" }
}

# ─────────────────────────────────────────────
#  ROW STYLE (color coding)
# ─────────────────────────────────────────────
function Apply-RowStyle {
    param($dg, [hashtable]$T)
    try {
        $style = New-Object Windows.Style ([Windows.Controls.DataGridRow])
        $style.Setters.Add((New-Object Windows.Setter ([Windows.Controls.Control]::BackgroundProperty), [Windows.Media.Brushes]::Transparent))

        $lockedTrigger = New-Object Windows.DataTrigger
        $lockedTrigger.Binding = New-Object Windows.Data.Binding "LockedOut"
        $lockedTrigger.Value = $true
        $lockedTrigger.Setters.Add((New-Object Windows.Setter ([Windows.Controls.Control]::BackgroundProperty), (New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($T.RowLockedOut)))))
        $style.Triggers.Add($lockedTrigger)

        $disabledTrigger = New-Object Windows.DataTrigger
        $disabledTrigger.Binding = New-Object Windows.Data.Binding "Enabled"
        $disabledTrigger.Value = $false
        $disabledTrigger.Setters.Add((New-Object Windows.Setter ([Windows.Controls.Control]::BackgroundProperty), (New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($T.RowDisabled)))))
        $style.Triggers.Add($disabledTrigger)

        $rodcTrigger = New-Object Windows.DataTrigger
        $rodcTrigger.Binding = New-Object Windows.Data.Binding "IsRODC"
        $rodcTrigger.Value = $true
        $rodcTrigger.Setters.Add((New-Object Windows.Setter ([Windows.Controls.Control]::BackgroundProperty), (New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($T.RowRODC)))))
        $style.Triggers.Add($rodcTrigger)

        $dg.RowStyle = $style
    } catch { Write-Warning "Apply-RowStyle: $_" }
}

# ─────────────────────────────────────────────
#  GENERIC ASYNC DISPATCHER
# ─────────────────────────────────────────────
function Run-AsyncSearch {
    param([string]$StatusMsg, [ScriptBlock]$Work, [ScriptBlock]$OnComplete, $CountCtrl, $Grid, [string]$CountLabel="results")
    $txtStatus.Text = $StatusMsg
    Invoke-Async -ScriptBlock $Work -CompletedCallback {
        param($result,$err)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($err){
                $txtStatus.Text = "Error: $($err.Exception.Message)"
                if ($Grid){ $Grid.ItemsSource = @() }
            } else {
                if ($Grid){ $Grid.ItemsSource = $result }
                if ($CountCtrl){ $CountCtrl.Text = "$($result.Count) $CountLabel" }
                $txtStatus.Text = "$($result.Count) $CountLabel returned."
            }
        }))
    }
}

# ─────────────────────────────────────────────
#  EXPORT GRID HELPER
# ─────────────────────────────────────────────
function ExportGrid {
    param($dg,[string]$category)
    try {
        $items = $dg.ItemsSource
        if (-not $items -or $items.Count -eq 0){
            [System.Windows.MessageBox]::Show("No data to export.","Export",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        $fmts = $Formats.Split(',') | ForEach-Object { $_.Trim() }
        if (-not $fmts){ $fmts = @('csv') }
        $path = $txtExportFolder.Text.Trim()
        if (-not $path){ $path = Get-ExportFolder -ReportName $category }
        Export-Results -Results ($items | ForEach-Object { $_ }) -Category $category -Filter "" -ExportPath $path -Formats $fmts
        Save-Config @{ ExportFolder=$path; Formats=$fmts; Theme=if($Script:IsDark){"Dark"}else{"Light"} }
        $txtStatus.Text = "Exported $($items.Count) items to $path"
    } catch {
        $txtStatus.Text = "Export error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)","Export Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

# ─────────────────────────────────────────────
#  WIRE EVENTS
# ─────────────────────────────────────────────

# Connection bar
$btnRefreshDCs.Add_Click({ Refresh-DCs })

$btnConnect.Add_Click({
    $manual = $txtManualServer.Text.Trim()
    if ($manual){
        $txtConnectedTo.Text = "Connected: $manual"
        $txtStatus.Text = "Manually set target server to: $manual"
    } else {
        Refresh-DCs
    }
})

$btnSetCred.Add_Click({
    $user = $txtCredUser.Text.Trim()
    $pass = $pwdCredPass.SecurePassword
    $domain = $cmbDomain.SelectedItem
    if (-not $user){ $txtStatus.Text = "Enter a username to set credentials."; return }
    try {
        $cred = New-Object System.Management.Automation.PSCredential($user, $pass)
        if ($domain){ $Script:CredStore[$domain] = $cred }
        $Script:CredStore["__global__"] = $cred
        $txtActiveCred.Text = "Active: $user"
        $txtStatus.Text = "Credentials set for: $user"
        Refresh-DCs
    } catch { $txtStatus.Text = "Credential error: $($_.Exception.Message)" }
})

$btnBrowseExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtExportFolder.Text
    $dlg.Description  = "Select Export Folder"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){ $txtExportFolder.Text = $dlg.SelectedPath }
})

$btnNewExportFolder.Add_Click({
    $f = Get-ExportFolder -ReportName "ADReport"
    $txtExportFolder.Text = $f
    $txtStatus.Text = "New export folder: $f"
})

$cmbDomain.Add_SelectionChanged({ Refresh-DCs })

# ─────────────────────────────────────────────
#  RDP DIALOG
# ─────────────────────────────────────────────
$btnRDP.Add_Click({
    $dlg = New-Object System.Windows.Window
    $dlg.Title = "RDP Connection"
    $dlg.Width = 420; $dlg.Height = 260
    $dlg.WindowStartupLocation = "CenterOwner"
    $dlg.Owner = $Window
    $dlg.Background = [Windows.Media.Brushes]::Transparent
    $T = if ($Script:IsDark) { $Script:Themes.Dark } else { $Script:Themes.Light }
    $dlgXaml = @"
<Grid xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' Background='$($T.PanelBg)'>
  <StackPanel Margin='24'>
    <TextBlock Text='RDP Remote Desktop Connect' FontSize='16' FontWeight='Bold' Foreground='$($T.TextPrimary)' Margin='0,0,0,16'/>
    <TextBlock Text='Server / Hostname:' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <TextBox x:Name='rdpHost' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,10'/>
    <TextBlock Text='Username (optional, e.g. DOMAIN\user):' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <TextBox x:Name='rdpUser' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,16'/>
    <StackPanel Orientation='Horizontal' HorizontalAlignment='Right'>
      <Button x:Name='rdpConnect' Content='Launch RDP' Background='$($T.ButtonBg)' Foreground='White' BorderThickness='0' Padding='16,8' Margin='0,0,8,0' Cursor='Hand'/>
      <Button x:Name='rdpCancel' Content='Cancel' Background='$($T.TabHover)' Foreground='$($T.TextPrimary)' BorderThickness='0' Padding='16,8' Cursor='Hand'/>
    </StackPanel>
  </StackPanel>
</Grid>
"@
    [xml]$dlgXmlDoc = $dlgXaml
    $dlgContent = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $dlgXmlDoc))
    $dlg.Content = $dlgContent

    $rdpHost = $dlgContent.FindName("rdpHost")
    $rdpUser = $dlgContent.FindName("rdpUser")
    $rdpConnect = $dlgContent.FindName("rdpConnect")
    $rdpCancel  = $dlgContent.FindName("rdpCancel")

    # Pre-fill with active server
    $rdpHost.Text = Get-ActiveServer

    $rdpConnect.Add_Click({
        $host_ = $rdpHost.Text.Trim()
        if (-not $host_){ [System.Windows.MessageBox]::Show("Enter a hostname."); return }
        $args_ = "/v:$host_"
        if ($rdpUser.Text.Trim()){ $args_ += " /u:$($rdpUser.Text.Trim())" }
        try { Start-Process "mstsc.exe" -ArgumentList $args_ } catch { [System.Windows.MessageBox]::Show("Failed to launch RDP: $($_.Exception.Message)") }
        $dlg.Close()
    })
    $rdpCancel.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
})

# ─────────────────────────────────────────────
#  FILE COPY DIALOG
# ─────────────────────────────────────────────
$btnFileCopy.Add_Click({
    $T = if ($Script:IsDark) { $Script:Themes.Dark } else { $Script:Themes.Light }
    $dlg = New-Object System.Windows.Window
    $dlg.Title = "Copy Files to Server"; $dlg.Width = 540; $dlg.Height = 360
    $dlg.WindowStartupLocation = "CenterOwner"; $dlg.Owner = $Window
    $dlgXaml = @"
<Grid xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' Background='$($T.PanelBg)'>
  <StackPanel Margin='24'>
    <TextBlock Text='Copy Files / Folder to Remote Server' FontSize='16' FontWeight='Bold' Foreground='$($T.TextPrimary)' Margin='0,0,0,16'/>
    <TextBlock Text='Source (local file or folder path):' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <TextBox x:Name='cpySrc' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,8'/>
    <TextBlock Text='Destination (UNC path, e.g. \\server\share\folder):' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <TextBox x:Name='cpyDst' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,8'/>
    <TextBlock Text='Alternate Username (optional, DOMAIN\user):' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <TextBox x:Name='cpyUser' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,8'/>
    <TextBlock Text='Password:' Foreground='$($T.TextSecondary)' FontSize='12' Margin='0,0,0,4'/>
    <PasswordBox x:Name='cpyPass' Background='$($T.ControlBg)' Foreground='$($T.TextPrimary)' BorderBrush='$($T.ControlBorder)' Padding='8,6' Margin='0,0,0,16'/>
    <TextBlock x:Name='cpyStatus' Foreground='$($T.AccentYellow)' FontSize='11' Margin='0,0,0,8'/>
    <StackPanel Orientation='Horizontal' HorizontalAlignment='Right'>
      <Button x:Name='cpyCopy' Content='&#xE8C8; Copy Now' Background='$($T.SuccessBg)' Foreground='White' BorderThickness='0' Padding='16,8' Margin='0,0,8,0' Cursor='Hand'/>
      <Button x:Name='cpyCancel' Content='Close' Background='$($T.TabHover)' Foreground='$($T.TextPrimary)' BorderThickness='0' Padding='16,8' Cursor='Hand'/>
    </StackPanel>
  </StackPanel>
</Grid>
"@
    [xml]$dlgXmlDoc = $dlgXaml
    $dlgContent = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $dlgXmlDoc))
    $dlg.Content = $dlgContent
    $cpySrc    = $dlgContent.FindName("cpySrc")
    $cpyDst    = $dlgContent.FindName("cpyDst")
    $cpyUser   = $dlgContent.FindName("cpyUser")
    $cpyPass   = $dlgContent.FindName("cpyPass")
    $cpyStatus = $dlgContent.FindName("cpyStatus")
    $cpyCopy   = $dlgContent.FindName("cpyCopy")
    $cpyCancel = $dlgContent.FindName("cpyCancel")

    $cpyCopy.Add_Click({
        $src  = $cpySrc.Text.Trim()
        $dst  = $cpyDst.Text.Trim()
        $user = $cpyUser.Text.Trim()
        $pass_ = $cpyPass.SecurePassword
        if (-not $src -or -not $dst){ $cpyStatus.Text = "Source and destination required."; return }
        try {
            if ($user){
                $cred = New-Object System.Management.Automation.PSCredential($user, $pass_)
                # Map a temp drive for the UNC path then copy
                $driveName = "ADETempDrive"
                $uncFolder = Split-Path $dst -Parent
                if (!(Test-Path "${driveName}:")){ New-PSDrive -Name $driveName -PSProvider FileSystem -Root $uncFolder -Credential $cred -ErrorAction Stop | Out-Null }
                Copy-Item -Path $src -Destination $dst -Recurse -Force -ErrorAction Stop
                Remove-PSDrive -Name $driveName -ErrorAction SilentlyContinue
            } else {
                Copy-Item -Path $src -Destination $dst -Recurse -Force -ErrorAction Stop
            }
            $cpyStatus.Text = "Copy succeeded!"
            $txtStatus.Text = "File copy completed: $src -> $dst"
        } catch {
            $cpyStatus.Text = "Error: $($_.Exception.Message)"
        }
    })
    $cpyCancel.Add_Click({ $dlg.Close() })
    $dlg.ShowDialog() | Out-Null
})

# ─────────────────────────────────────────────
#  THEME TOGGLE
# ─────────────────────────────────────────────
$btnToggleTheme.Add_Click({
    $Script:IsDark = -not $Script:IsDark
    Save-Config @{ ExportFolder=$txtExportFolder.Text; Theme=if($Script:IsDark){"Dark"}else{"Light"} }
    [System.Windows.MessageBox]::Show("Theme changed! Restart the tool to apply the new theme.","Theme Toggle",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
})

# ─────────────────────────────────────────────
#  SEARCH BUTTON WIRES
# ─────────────────────────────────────────────
$btnUserSearch.Add_Click({
    $server = Get-ActiveServer; $cred = Get-ActiveCred; $filter = $txtUserFilter.Text; $allD = $chkUserAllDomains.IsChecked; $enabledOnly = $chkUserEnabledOnly.IsChecked
    $txtStatus.Text = "Searching users..."
    if ($allD) {
        $domains_ = Get-ForestDomains
        $agg = [System.Collections.ArrayList]@(); $done = 0
        foreach ($d in $domains_) {
            $target = $d
            $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchUsers -Filter '$filter' -Server '$target'")
            Invoke-Async -ScriptBlock $sb -CompletedCallback {
                param($r,$e)
                [void]($Window.Dispatcher.Invoke([action]{
                    $done++
                    if ($r){ foreach ($o in $r){ [void]$agg.Add($o) } }
                    $dgUsers.ItemsSource = $agg; $txtUserCount.Text = "$($agg.Count) results"
                    if ($done -eq $domains_.Count){ $txtStatus.Text = "Users: $($agg.Count) across all domains." }
                }))
            }
        }
    } else {
        $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchUsers -Filter '$filter' -Server '$server'")
        Invoke-Async -ScriptBlock $sb -CompletedCallback {
            param($r,$e)
            [void]($Window.Dispatcher.Invoke([action]{
                if ($e){ $txtStatus.Text = "User search error: $($e.Exception.Message)"; $dgUsers.ItemsSource = @(); return }
                $out = if ($enabledOnly){ $r | Where-Object { $_.Enabled -eq $true } } else { $r }
                $dgUsers.ItemsSource = $out; $txtUserCount.Text = "$($out.Count) results"
                Apply-RowStyle -dg $dgUsers -T $T
                $txtStatus.Text = "Users: $($out.Count) results."
            }))
        }
    }
})

$btnCompSearch.Add_Click({
    $server = Get-ActiveServer; $filter = $txtCompFilter.Text
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchComputers -Filter '$filter' -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Computer search error: $($e.Exception.Message)"; $dgComputers.ItemsSource = @(); return }
            $dgComputers.ItemsSource = $r; $txtCompCount.Text = "$($r.Count) results"
            Apply-RowStyle -dg $dgComputers -T $T
            $txtStatus.Text = "Computers: $($r.Count) results."
        }))
    }
})

$btnServerSearch.Add_Click({
    $server = Get-ActiveServer; $filter = $txtServerFilter.Text
    $filterSafe = if ($filter) { "*$filter*" } else { "*" }
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchComputers -Server '$server' -ServersOnly")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Server search error: $($e.Exception.Message)"; $dgServers.ItemsSource = @(); return }
            $out = if ($filter){ $r | Where-Object { $_.Name -like "*$filter*" } } else { $r }
            $dgServers.ItemsSource = $out; $txtServerCount.Text = "$($out.Count) servers"
            $txtStatus.Text = "Servers: $($out.Count) results."
        }))
    }
})

$btnGroupSearch.Add_Click({
    $server = Get-ActiveServer; $filter = $txtGroupFilter.Text
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchGroups -Filter '$filter' -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Group search error: $($e.Exception.Message)"; $dgGroups.ItemsSource = @(); return }
            $dgGroups.ItemsSource = $r; $txtGroupCount.Text = "$($r.Count) groups"
            $txtStatus.Text = "Groups: $($r.Count) results."
        }))
    }
})

$btnSecGrpSearch.Add_Click({
    $server = Get-ActiveServer; $filter = $txtSecGrpFilter.Text
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchGroups -Filter '$filter' -Server '$server' -SecurityOnly")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Security group search error: $($e.Exception.Message)"; $dgSecGrps.ItemsSource = @(); return }
            $dgSecGrps.ItemsSource = $r; $txtSecGrpCount.Text = "$($r.Count) security groups"
            $txtStatus.Text = "Security Groups: $($r.Count) results."
        }))
    }
})

$btnGPOSearch.Add_Click({
    $domain = $cmbDomain.SelectedItem; $filter = $txtGPOFilter.Text
    $sb = [scriptblock]::Create("Import-Module GroupPolicy -EA SilentlyContinue; SearchGPOs -Filter '$filter' -Domain '$domain'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "GPO search error: $($e.Exception.Message)"; $dgGPOs.ItemsSource = @(); return }
            $dgGPOs.ItemsSource = $r; $txtGPOCount.Text = "$($r.Count) GPOs"
            $txtStatus.Text = "GPOs: $($r.Count) results."
        }))
    }
})

$btnSubnetsSearch.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchSubnets -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Subnet error: $($e.Exception.Message)"; $dgSubnets.ItemsSource = @(); return }
            $dgSubnets.ItemsSource = $r; $txtSubnetCount.Text = "$($r.Count) subnets"
            $txtStatus.Text = "Subnets: $($r.Count) results."
        }))
    }
})

$btnTrustsSearch.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchTrusts -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Trust error: $($e.Exception.Message)"; $dgTrusts.ItemsSource = @(); return }
            $dgTrusts.ItemsSource = $r; $txtTrustCount.Text = "$($r.Count) trusts"
            $txtStatus.Text = "Trusts: $($r.Count) results."
        }))
    }
})

$btnDCSearch.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchAllDCs -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "DC error: $($e.Exception.Message)"; $dgDCs.ItemsSource = @(); return }
            $dgDCs.ItemsSource = $r; $txtDCCount.Text = "$($r.Count) DCs"
            Apply-RowStyle -dg $dgDCs -T $T
            $txtStatus.Text = "Domain Controllers: $($r.Count) results."
        }))
    }
})

$btnRODCSearch.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchRODCs -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "RODC error: $($e.Exception.Message)"; $dgDCs.ItemsSource = @(); return }
            $dgDCs.ItemsSource = $r; $txtDCCount.Text = "$($r.Count) RODCs"
            $txtStatus.Text = "RODCs: $($r.Count) results."
        }))
    }
})

$btnT0Search.Add_Click({
    $server = Get-ActiveServer; $tier = 0
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchTierAccounts -Server '$server' -Tier 0")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "T0 error: $($e.Exception.Message)"; $dgT0.ItemsSource = @(); return }
            $dgT0.ItemsSource = $r; $txtT0Count.Text = "$($r.Count) T0 accounts"
            Apply-RowStyle -dg $dgT0 -T $T
            $txtStatus.Text = "T0 Accounts: $($r.Count) results."
        }))
    }
})

$btnT1Search.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchTierAccounts -Server '$server' -Tier 1")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "T1 error: $($e.Exception.Message)"; $dgT1.ItemsSource = @(); return }
            $dgT1.ItemsSource = $r; $txtT1Count.Text = "$($r.Count) T1 accounts"
            Apply-RowStyle -dg $dgT1 -T $T
            $txtStatus.Text = "T1 Accounts: $($r.Count) results."
        }))
    }
})

$btnMSASearch.Add_Click({
    $server = Get-ActiveServer
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchMSAs -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "MSA error: $($e.Exception.Message)"; $dgMSAs.ItemsSource = @(); return }
            $dgMSAs.ItemsSource = $r; $txtMSACount.Text = "$($r.Count) MSAs"
            $txtStatus.Text = "MSAs/gMSAs: $($r.Count) results."
        }))
    }
})

$btnSvcAcctSearch.Add_Click({
    $server = Get-ActiveServer; $filter = $txtSvcAcctFilter.Text
    $sb = [scriptblock]::Create("Import-Module ActiveDirectory -EA SilentlyContinue; SearchServiceAccounts -Filter '$filter' -Server '$server'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "Svc account error: $($e.Exception.Message)"; $dgSvcAccts.ItemsSource = @(); return }
            $dgSvcAccts.ItemsSource = $r; $txtSvcAcctCount.Text = "$($r.Count) accounts"
            $txtStatus.Text = "Service Accounts: $($r.Count) results."
        }))
    }
})

$btnDFSSearch.Add_Click({
    $sb = [scriptblock]::Create("Import-Module Dfsn -EA SilentlyContinue; SearchDFS")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "DFS error: $($e.Exception.Message)"; $dgDFS.ItemsSource = @(); return }
            $dgDFS.ItemsSource = $r; $txtDFSCount.Text = "$($r.Count) roots"
            $txtStatus.Text = "DFS Roots: $($r.Count) results."
        }))
    }
})

$btnDhcpSearch.Add_Click({
    $dhcpSrv = $txtDhcpServer.Text.Trim()
    if (-not $dhcpSrv){ $txtStatus.Text = "Enter a DHCP server hostname/IP first."; return }
    $sb = [scriptblock]::Create("Import-Module DhcpServer -EA SilentlyContinue; SearchDHCP -Server '$dhcpSrv'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "DHCP error: $($e.Exception.Message)"; $dgDHCP.ItemsSource = @(); return }
            $dgDHCP.ItemsSource = $r; $txtDHCPCount.Text = "$($r.Count) scopes"
            $txtStatus.Text = "DHCP Scopes: $($r.Count) results."
        }))
    }
})

$btnDnsSearch.Add_Click({
    $dnsSrv = $txtDnsServer.Text.Trim()
    if (-not $dnsSrv){ $dnsSrv = Get-ActiveServer }
    $sb = [scriptblock]::Create("Import-Module DnsServer -EA SilentlyContinue; SearchDNS -Server '$dnsSrv'")
    Invoke-Async -ScriptBlock $sb -CompletedCallback {
        param($r,$e)
        [void]($Window.Dispatcher.Invoke([action]{
            if ($e){ $txtStatus.Text = "DNS error: $($e.Exception.Message)"; $dgDNS.ItemsSource = @(); return }
            $dgDNS.ItemsSource = $r; $txtDNSCount.Text = "$($r.Count) zones"
            $txtStatus.Text = "DNS Zones: $($r.Count) results."
        }))
    }
})

# ─────────────────────────────────────────────
#  EXPORT WIRES
# ─────────────────────────────────────────────
$btnUserExport.Add_Click({ ExportGrid -dg $dgUsers -category "Users" })
$btnCompExport.Add_Click({ ExportGrid -dg $dgComputers -category "Computers" })
$btnServerExport.Add_Click({ ExportGrid -dg $dgServers -category "Servers" })
$btnGroupExport.Add_Click({ ExportGrid -dg $dgGroups -category "Groups" })
$btnSecGrpExport.Add_Click({ ExportGrid -dg $dgSecGrps -category "SecurityGroups" })
$btnGPOExport.Add_Click({ ExportGrid -dg $dgGPOs -category "GPOs" })
$btnSubnetsExport.Add_Click({ ExportGrid -dg $dgSubnets -category "Subnets" })
$btnTrustsExport.Add_Click({ ExportGrid -dg $dgTrusts -category "Trusts" })
$btnDCExport.Add_Click({ ExportGrid -dg $dgDCs -category "DomainControllers" })
$btnT0Export.Add_Click({ ExportGrid -dg $dgT0 -category "T0_Accounts" })
$btnT1Export.Add_Click({ ExportGrid -dg $dgT1 -category "T1_Accounts" })
$btnMSAExport.Add_Click({ ExportGrid -dg $dgMSAs -category "MSAs" })
$btnSvcAcctExport.Add_Click({ ExportGrid -dg $dgSvcAccts -category "ServiceAccounts" })
$btnDFSExport.Add_Click({ ExportGrid -dg $dgDFS -category "DFS" })
$btnDhcpExport.Add_Click({ ExportGrid -dg $dgDHCP -category "DHCP" })
$btnDnsExport.Add_Click({ ExportGrid -dg $dgDNS -category "DNS" })

# ─────────────────────────────────────────────
#  INITIAL DC LOAD
# ─────────────────────────────────────────────
try { Refresh-DCs } catch {}

# ─────────────────────────────────────────────
#  SCHEDULED / HEADLESS MODE
# ─────────────────────────────────────────────
if ($ScheduledMode) {
    switch ($Preset) {
        "LockedOutUsers" {
            try {
                Import-Module ActiveDirectory -EA SilentlyContinue
                $locked = Search-ADAccount -LockedOut -UsersOnly -EA Stop | ForEach-Object {
                    Get-ADUser $_.SamAccountName -Properties LockedOut,LastLogonDate,Enabled,DistinguishedName -EA SilentlyContinue
                } | Select-Object Name,sAMAccountName,Enabled,LockedOut,LastLogonDate,DistinguishedName
                $fmts = $Formats.Split(',') | ForEach-Object { $_.Trim() }
                $path  = Get-ExportFolder -ReportName $Preset
                Export-Results -Results $locked -Category $Preset -Filter "" -ExportPath $path -Formats $fmts
                Write-Host "Exported $($locked.Count) locked-out users to $path"
            } catch { Write-Error "Preset $Preset failed: $($_.Exception.Message)" }
        }
        "T0Accounts" {
            try {
                Import-Module ActiveDirectory -EA SilentlyContinue
                $domain_ = (Get-ADDomain -EA Stop).DNSRoot
                $t0 = SearchTierAccounts -Server $domain_ -Tier 0
                $fmts = $Formats.Split(',') | ForEach-Object { $_.Trim() }
                $path  = Get-ExportFolder -ReportName $Preset
                Export-Results -Results $t0 -Category $Preset -Filter "" -ExportPath $path -Formats $fmts
                Write-Host "Exported $($t0.Count) T0 accounts to $path"
            } catch { Write-Error "Preset $Preset failed: $($_.Exception.Message)" }
        }
        default { Write-Host "No preset named '$Preset'. Available: LockedOutUsers, T0Accounts" }
    }
    $runspacePool.Close(); $runspacePool.Dispose()
    return
}

# ─────────────────────────────────────────────
#  SHOW WINDOW
# ─────────────────────────────────────────────
$Window.Add_Closed({
    try { $runspacePool.Close(); $runspacePool.Dispose() } catch {}
})

$Window.ShowDialog() | Out-Null
try { $runspacePool.Close(); $runspacePool.Dispose() } catch {}
