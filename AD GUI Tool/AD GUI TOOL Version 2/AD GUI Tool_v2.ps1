<#
.SYNOPSIS
  AD GUI Tool_v2.ps1 - Enterprise WPF AD Search & Reporting Console
  Author: Stephen McKee - Systems Administrator 2

.DESCRIPTION
  WPF GUI to query Active Directory (Users, Computers, OUs, GPOs, Groups, Subnets, Locked-out users, Service Accounts, Firewall GPO rules, etc.)
  Includes presets, save/load of settings, optional credential storage (protected), export to many formats, and basic charting.

.NOTES
  Run elevated for event-log lockout-origin queries. Requires ActiveDirectory module. For GPO features GroupPolicy module is required.
#>

# -----------------------------
# Prereqs & assemblies
# -----------------------------
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check for charting assembly
$ChartingAvailable = $false
try {
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction Stop
    $ChartingAvailable = $true
} catch {
    Write-Warning "System.Windows.Forms.DataVisualization assembly not found. Charting features will be disabled."
}

# -----------------------------
# Configuration / Helpers
# -----------------------------
$Global:AppFolder = "C:\Temp\ADSearchTool"
$Global:ExportFolder = Join-Path $Global:AppFolder "Export"
$Global:ConfigFile = Join-Path $Global:AppFolder "config.json"
If (!(Test-Path $Global:AppFolder)) { New-Item -Path $Global:AppFolder -ItemType Directory -Force | Out-Null }
If (!(Test-Path $Global:ExportFolder)) { New-Item -Path $Global:ExportFolder -ItemType Directory -Force | Out-Null }

function Ensure-ModuleLoaded {
    param([string]$Name)
    if (Get-Module -ListAvailable -Name $Name) {
        Import-Module $Name -ErrorAction SilentlyContinue
        return $true
    } else {
        return $false
    }
}

$HasAD = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO = Ensure-ModuleLoaded -Name GroupPolicy

# DPAPI wrappers for credential storage
function Protect-Credential {
    param([PSCredential]$Credential)
    if (-not $Credential) { return $null }
    $plain = ($Credential.UserName + "`n" + ($Credential.GetNetworkCredential().Password))
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($plain)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    return [System.Convert]::ToBase64String($protected)
}
function Unprotect-Credential {
    param([string]$ProtectedString)
    if (-not $ProtectedString) { return $null }
    try {
        $bytes = [System.Convert]::FromBase64String($ProtectedString)
        $un = [System.Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $plain = [System.Text.Encoding]::UTF8.GetString($un)
        $parts = $plain -split "`n",2
        $username = $parts[0]
        $password = if ($parts.Count -ge 2) { $parts[1] } else { "" }
        return New-Object System.Management.Automation.PSCredential ($username,(ConvertTo-SecureString $password -AsPlainText -Force))
    } catch { return $null }
}

# Save / Load config
function Save-Config {
    param($config)
    $json = $config | ConvertTo-Json -Depth 6
    $json | Out-File -FilePath $Global:ConfigFile -Encoding UTF8
}
function Load-Config {
    if (Test-Path $Global:ConfigFile) {
        try { return Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json } catch { return $null }
    } else { return $null }
}

# Export helper (header + formats)
function Export-Results {
    param(
        [Parameter(Mandatory=$true)] [array]$Results,
        [Parameter(Mandatory=$true)] [string]$Category,
        [Parameter(Mandatory=$true)] [string]$Filter,
        [Parameter(Mandatory=$true)] [string]$ExportPath,
        [Parameter(Mandatory=$true)] [string[]]$Formats
    )

    if (!(Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $baseName = "{0}_{1}" -f (($Category -replace '[^\w\-_\. ]','_').Trim(), $timestamp)
    $total = $Results.Count
    $header = @"
Active Directory Search - Export
Search Type : $Category
Filter Used : $Filter
Export Time : $(Get-Date)
Total Found : $total
"@

    if ($total -eq 0) {
        # Create small "no results" placeholder files for selected formats
        foreach ($fmt in $Formats) {
            switch ($fmt.ToLower()) {
                "csv" {
                    $file = Join-Path $ExportPath ("$baseName.csv")
                    "# $header" | Out-File -FilePath $file -Encoding UTF8
                }
                "xml" {
                    $file = Join-Path $ExportPath ("$baseName.xml")
                    "<results><summary>$( [System.Security.SecurityElement]::Escape($header) )</summary></results>" | Out-File -FilePath $file -Encoding UTF8
                }
                "html" {
                    $file = Join-Path $ExportPath ("$baseName.html")
                    "<html><body><pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre><h3>No results</h3></body></html>" | Out-File -FilePath $file -Encoding UTF8
                }
                "txt" {
                    $file = Join-Path $ExportPath ("$baseName.txt")
                    $header | Out-File -FilePath $file -Encoding UTF8
                }
                "excel" {
                    $file = Join-Path $ExportPath ("$baseName.csv")
                    "# $header" | Out-File -FilePath $file -Encoding UTF8
                }
                "pdf" {
                    $file = Join-Path $ExportPath ("$baseName.html")
                    "<html><body><pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre><h3>No results</h3></body></html>" | Out-File -FilePath $file -Encoding UTF8
                }
                "docx" {
                    $file = Join-Path $ExportPath ("$baseName.txt")
                    $header | Out-File -FilePath $file -Encoding UTF8
                }
            }
        }
        [System.Windows.MessageBox]::Show("No results found. Placeholder export files created in $ExportPath.","Export - No Results",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    foreach ($fmt in $Formats) {
        switch ($fmt.ToLower()) {
            "csv" {
                $file = Join-Path $ExportPath ("$baseName.csv")
                $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                $csv = $Results | ConvertTo-Csv -NoTypeInformation
                $headerLines + $csv | Out-File -FilePath $file -Encoding UTF8
            }
            "xml" {
                $file = Join-Path $ExportPath ("$baseName.xml")
                $xmlComment = "<!-- " + ($header -replace '--','- -') + " -->`n"
                $Results | Export-Clixml -Path $file
                $content = Get-Content -Path $file -Raw
                $xmlComment + $content | Out-File -FilePath $file -Encoding UTF8
            }
            "html" {
                $file = Join-Path $ExportPath ("$baseName.html")
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $file -Encoding UTF8
            }
            "txt" {
                $file = Join-Path $ExportPath ("$baseName.txt")
                $header | Out-File -FilePath $file -Encoding UTF8
                $Results | Out-String | Out-File -FilePath $file -Append -Encoding UTF8
            }
            "excel" {
                $file = Join-Path $ExportPath ("$baseName.xlsx")
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    try {
                        $Results | Export-Excel -Path $file -WorksheetName "Results" -AutoSize -Title ("AD Search Results - " + $Category)
                    } catch {
                        # fallback
                        $csvFile = Join-Path $ExportPath ("$baseName.csv")
                        $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                        $csv = $Results | ConvertTo-Csv -NoTypeInformation
                        $headerLines + $csv | Out-File -FilePath $csvFile -Encoding UTF8
                    }
                } else {
                    $csvFile = Join-Path $ExportPath ("$baseName.csv")
                    $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                    $csv = $Results | ConvertTo-Csv -NoTypeInformation
                    $headerLines + $csv | Out-File -FilePath $csvFile -Encoding UTF8
                }
            }
            "pdf" {
                $htmlFile = Join-Path $ExportPath ("$baseName.html")
                $pdfFile  = Join-Path $ExportPath ("$baseName.pdf")
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $htmlFile -Encoding UTF8
                $wk = (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue).Path
                if ($wk) {
                    & $wk $htmlFile $pdfFile
                } else {
                    # user converts manually
                }
            }
            "docx" {
                $htmlFile = Join-Path $ExportPath ("$baseName.html")
                $docxFile = Join-Path $ExportPath ("$baseName.docx")
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$([System.Web.HttpUtility]::HtmlEncode($header))</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $htmlFile -Encoding UTF8
                try {
                    $word = New-Object -ComObject Word.Application -ErrorAction Stop
                    $doc = $word.Documents.Add($htmlFile)
                    $doc.SaveAs([ref] $docxFile, [ref]16)
                    $doc.Close()
                    $word.Quit()
                } catch {
                    # save HTML anyway
                }
            }
        }
    }

    [System.Windows.MessageBox]::Show("Export complete. $total results exported to $ExportPath.","Export Complete",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
}

# -----------------------------
# Search functions (core)
# -----------------------------
# Ensure AD module loaded where needed
if (-not $HasAD) {
    [System.Windows.MessageBox]::Show("ActiveDirectory module not found. Many functions will be unavailable.","Missing Module",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) | Out-Null
} else {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
}

function Search-Users {
    param([string]$filter)
    $props = @("Name","sAMAccountName","distinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","memberOf","userPrincipalName")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADUser -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADUser -Filter { Name -like $f -or sAMAccountName -like $f -or mail -like $f -or userPrincipalName -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    return $res | Select-Object @{n='Type';e={'User'}}, Name,sAMAccountName,distinguishedName,Enabled,LockedOut,LastLogonDate,whenCreated,userPrincipalName,@{n='MemberOf';e={$_.memberOf -join '; '}}
}

function Search-Computers {
    param([string]$filter)
    $props = @("Name","OperatingSystem","OperatingSystemVersion","distinguishedName","whenCreated","lastLogonDate")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADComputer -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADComputer -Filter { Name -like $f -or OperatingSystem -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    return $res | Select-Object @{n='Type';e={'Computer'}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}}
}

function Search-OUs {
    param([string]$filter)
    $res = Get-ADOrganizationalUnit -Filter { Name -like $filter } -Properties distinguishedName,whenCreated -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'OU'}}, Name,distinguishedName,whenCreated
}

function Search-GPOs {
    param([string]$filter)
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        throw "GroupPolicy module not available."
    }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All | Where-Object { $_.DisplayName -like $filter }
    $out = foreach ($g in $gpos) {
        $links = (Get-GPOLink -Guid $g.Id).LinksTo | ForEach-Object { $_.Scope } -join "; "
        [pscustomobject]@{
            Type = "GPO"
            Name = $g.DisplayName
            Id = $g.Id
            Owner = $g.Owner
            CreationTime = $g.CreationTime
            ModificationTime = $g.ModificationTime
            Links = $links
        }
    }
    return $out
}

function Search-Groups {
    param([string]$filter)
    $props = @("Name","distinguishedName","GroupCategory","GroupScope","member")
    $res = Get-ADGroup -Filter { Name -like $filter } -Properties $props -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'Group'}}, Name,GroupScope,GroupCategory,distinguishedName,@{n='Members';e={$_.member -join '; '}}
}

function Search-ServiceAccounts {
    param([string]$filter)
    $res = Get-ADUser -Filter { servicePrincipalName -like $filter -or sAMAccountName -like $filter } -Properties servicePrincipalName,description,distinguishedName -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'ServiceAccount'}}, Name,sAMAccountName,servicePrincipalName,distinguishedName,description
}

function Search-ServersOrWorkstations {
    param([string]$filter,[switch]$Servers)
    $osFilter = if ($Servers) { "*Server*" } else { "*Windows*" }
    $res = Get-ADComputer -Filter { OperatingSystem -like $osFilter -and Name -like $filter } -Properties OperatingSystem,OperatingSystemVersion,distinguishedName,lastLogonDate -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={if ($Servers) {'Server'}else{'Workstation'}}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.lastLogonDate}}
}

function Search-LockedOutUsers {
    param([switch]$ResolveOrigin)
    $locked = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue |
             Get-ADUser -Properties LockedOut,LastLogonDate,whenCreated,sAMAccountName,distinguishedName -ErrorAction SilentlyContinue
    $out = @()
    foreach ($u in $locked) {
        $record = [pscustomobject]@{
            Type = "LockedUser"
            Name = $u.Name
            sAMAccountName = $u.sAMAccountName
            DistinguishedName = $u.DistinguishedName
            LockedOut = $u.LockedOut
            LastLogon = $u.LastLogonDate
        }
        if ($ResolveOrigin) {
            $dcs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
            $origin = $null
            foreach ($dc in $dcs) {
                try {
                    $query = @"
<QueryList>
  <Query Id='0' Path='Security'>
    <Select Path='Security'>
      *[System[(EventID=4740)]] and *[EventData[Data and (Data='$($u.sAMAccountName)')]]
    </Select>
  </Query>
</QueryList>
"@
                    $events = Get-WinEvent -ComputerName $dc.HostName -FilterXml $query -MaxEvents 1 -ErrorAction SilentlyContinue
                    if ($events -and $events.Count -gt 0) {
                        $ev = $events[0]
                        $data = [xml]$ev.ToXml()
                        $td = $data.Event.EventData.Data
                        $caller = ($td | Where-Object { $_.Name -eq "CallerComputerName" }).'#text'
                        $origin = @{
                            DomainController = $dc.HostName
                            CallerComputer   = $caller
                            EventTime        = $ev.TimeCreated
                        }
                        break
                    }
                } catch { }
            }
            if ($origin) {
                $originValue = ($origin | Out-String).Trim()
            } else {
                $originValue = "Origin not found or insufficient permissions"
            }
            $record | Add-Member -MemberType NoteProperty -Name "LockoutOrigin" -Value $originValue -Force
        }
        $out += $record
    }
    return $out
}

function Search-Subnets {
    try {
        $configNaming = (Get-ADRootDSE).configurationNamingContext
        $base = "CN=Subnets,CN=Sites,$configNaming"
        $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -ErrorAction SilentlyContinue
        return $subnets | Select-Object @{n='Type';e={'ADSubnet'}}, name,@{n='Location';e={$_.location}},@{n='DistinguishedName';e={$_.DistinguishedName}}
    } catch {
        return @()
    }
}

function Search-GPOFirewallRules {
    param([string]$filter)
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) { throw "GroupPolicy module unavailable." }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All
    $matches = @()
    foreach ($g in $gpos) {
        $xml = Get-GPOReport -Guid $g.Id -ReportType Xml
        [xml]$gxml = $xml
        $policies = @()
        if ($gxml.GPO.Computer.ExtensionData.Extension.Policy) { $policies = $gxml.GPO.Computer.ExtensionData.Extension.Policy }
        foreach ($p in $policies) {
            if ($p.Name -like "*Firewall*" -or $p.Setting -like "*Firewall*" -or $p.Name -like $filter -or $p.Setting -like $filter) {
                $matches += [pscustomobject]@{
                    Type     = "GPOFirewall"
                    GPOName  = $g.DisplayName
                    Policy   = $p.Name
                    Setting  = $p.Setting
                    Links    = ((Get-GPOLink -Guid $g.Id).LinksTo | ForEach-Object { $_.Scope }) -join "; "
                }
            }
        }
    }
    return $matches
}

# -----------------------------
# Presets definitions
# -----------------------------
$Presets = @(
    [pscustomobject]@{Name="Disabled Accounts"; Category="User"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { Enabled -eq $false } -Properties Enabled,whenCreated | Select Name,sAMAccountName,distinguishedName,Enabled,whenCreated } },
    [pscustomobject]@{Name="Locked-Out Users"; Category="Locked-out Users (basic)"; Filter="*"; ScriptBlock={ Search-LockedOutUsers } },
    [pscustomobject]@{Name="Service Accounts (SPN)"; Category="Service Accounts"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { servicePrincipalName -like '*' } -Properties servicePrincipalName | Select Name,sAMAccountName,servicePrincipalName,distinguishedName } },
    [pscustomobject]@{Name="Password Never Expires"; Category="User"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { PasswordNeverExpires -eq $true } -Properties PasswordNeverExpires | Select Name,sAMAccountName,distinguishedName,PasswordNeverExpires } },
    [pscustomobject]@{Name="Domain Admins Members"; Category="Security Group"; Filter="Domain Admins"; ScriptBlock={ Get-ADGroupMember -Identity "Domain Admins" -Recursive | Select Name,sAMAccountName,distinguishedName,@{n='Group';e={'Domain Admins'}} } },
    [pscustomobject]@{Name="Recently Created Accounts (30d)"; Category="User"; Filter="*"; ScriptBlock={ $since = (Get-Date).AddDays(-30); Get-ADUser -Filter { whenCreated -ge $since } -Properties whenCreated | Select Name,sAMAccountName,distinguishedName,whenCreated } },
    [pscustomobject]@{Name="Inactive Computers (90d)"; Category="Computer"; Filter="*"; ScriptBlock={ $cut = (Get-Date).AddDays(-90); Get-ADComputer -Filter * -Properties LastLogonDate | Where-Object { $_.LastLogonDate -lt $cut -or -not $_.LastLogonDate } | Select Name,OperatingSystem,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}} } },
    [pscustomobject]@{Name="Domain Controllers"; Category="Servers (by OS or group)"; Filter="*Server*"; ScriptBlock={ Get-ADDomainController -Filter * | Select HostName,Site,OperatingSystem } },
    [pscustomobject]@{Name="GPOs with Log on as a service"; Category="GPO"; Filter="*"; ScriptBlock={ param($f) ; if (-not (Get-Module -ListAvailable -Name GroupPolicy)) { throw "GroupPolicy module missing" } ; Get-GPO -All | Where-Object { $_.DisplayName -like $f } | ForEach-Object { $links=(Get-GPOLink -Guid $_.Id).LinksTo | ForEach-Object { $_.Scope } -join '; '; [pscustomobject]@{Name=$_.DisplayName;Id=$_.Id;Links=$links} } } }
)

# -----------------------------
# WPF XAML layout - Tabbed UI
# -----------------------------
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:winForms="clr-namespace:System.Windows.Forms.Integration;assembly=WindowsFormsIntegration"
        xmlns:wf="clr-namespace:System.Windows.Forms.DataVisualization.Charting;assembly=System.Windows.Forms.DataVisualization"
        Title="AD Search Tool - Enterprise v2" Height="760" Width="1200" WindowStartupLocation="CenterScreen">
  <Grid Margin="8">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,6">
      <Label Content="Search Category:" VerticalAlignment="Center"/>
      <ComboBox x:Name="cmbCategory" Width="300" Margin="8,0,0,0"/>
      <Label Content="Filter:" VerticalAlignment="Center" Margin="12,0,0,0"/>
      <TextBox x:Name="txtFilter" Width="420" Margin="8,0,0,0"/>
      <Button x:Name="btnRun" Content="Run Search" Width="110" Margin="12,0,0,0"/>
      <Button x:Name="btnPresets" Content="Presets" Width="90" Margin="6,0,0,0"/>
      <Button x:Name="btnClear" Content="Clear" Width="70" Margin="6,0,0,0"/>
    </StackPanel>

    <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,0,0,6">
      <CheckBox x:Name="chkResolveSIDs" Content="Resolve SIDs" IsChecked="True" Margin="0,0,12,0"/>
      <CheckBox x:Name="chkIncludeDisabled" Content="Include disabled accounts" IsChecked="True" Margin="0,0,12,0"/>
      <CheckBox x:Name="chkEventLookup" Content="Lockout origin lookup (slow)" IsChecked="False" Margin="0,0,12,0"/>
      <Label Content="Export folder:" VerticalAlignment="Center" Margin="6,0,0,0"/>
      <TextBox x:Name="txtExportFolder" Width="360" Margin="6,0,0,0"/>
      <Button x:Name="btnOpenExport" Content="Open" Width="70" Margin="6,0,0,0"/>
      <Button x:Name="btnExport" Content="Export" Width="90" Margin="12,0,0,0"/>
    </StackPanel>

    <!-- Tab control -->
    <TabControl x:Name="tabControl" Grid.Row="2">
      <TabItem Header="Results">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="200"/>
          </Grid.RowDefinitions>
          <DataGrid x:Name="dgResults" Grid.Row="0" AutoGenerateColumns="True" IsReadOnly="True"/>
          <winForms:WindowsFormsHost Grid.Row="1" x:Name="hostChart">
            <wf:Chart x:Name="chart" />
          </winForms:WindowsFormsHost>
        </Grid>
      </TabItem>

      <TabItem Header="Presets">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal" Margin="6">
            <Button x:Name="btnRunPreset" Content="Run Selected Preset" Width="150" Margin="0,0,8,0"/>
            <Button x:Name="btnRefreshPresets" Content="Refresh Presets" Width="120"/>
          </StackPanel>
          <ListBox x:Name="lstPresets" Grid.Row="1" Margin="6"/>
        </Grid>
      </TabItem>

      <TabItem Header="Settings">
        <StackPanel Margin="8">
          <Label Content="Remember last settings in config file:"/>
          <WrapPanel>
            <Button x:Name="btnSaveSettings" Content="Save Settings" Width="120" Margin="4"/>
            <Button x:Name="btnLoadSettings" Content="Load Settings" Width="120" Margin="4"/>
            <Button x:Name="btnClearSettings" Content="Clear Settings" Width="120" Margin="4"/>
          </WrapPanel>
          <Separator Margin="6"/>
          <Label Content="Optional: Store alternate credentials (protected)"/>
          <StackPanel Orientation="Horizontal">
            <Label Content="Username:" VerticalAlignment="Center"/>
            <TextBox x:Name="txtAltUser" Width="240" Margin="6"/>
            <Label Content="Password:" VerticalAlignment="Center" Margin="6,0,0,0"/>
            <PasswordBox x:Name="txtAltPass" Width="240" Margin="6"/>
            <Button x:Name="btnSaveCred" Content="Save Creds (DPAPI)" Width="140" Margin="6"/>
            <Button x:Name="btnClearCred" Content="Clear Creds" Width="100" Margin="6"/>
          </StackPanel>
        </StackPanel>
      </TabItem>

    </TabControl>

    <StatusBar Grid.Row="3" VerticalAlignment="Bottom" Height="26">
      <StatusBarItem>
        <TextBlock x:Name="txtStatus" Text="Ready."/>
      </StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
'@

# Prepare XAML - load namespaces and helpers
try {
    [xml]$xamlObj = $Xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlObj
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Error "Failed to load XAML: $_"
    return
}

# Helper to find named elements
function Get-Element($name) { return $window.FindName($name) }

# Map WPF elements
$cmbCategory = Get-Element "cmbCategory"
$txtFilter = Get-Element "txtFilter"
$btnRun = Get-Element "btnRun"
$btnPresets = Get-Element "btnPresets"
$btnClear = Get-Element "btnClear"
$chkResolveSIDs = Get-Element "chkResolveSIDs"
$chkIncludeDisabled = Get-Element "chkIncludeDisabled"
$chkEventLookup = Get-Element "chkEventLookup"
$txtExportFolder = Get-Element "txtExportFolder"
$btnOpenExport = Get-Element "btnOpenExport"
$btnExport = Get-Element "btnExport"
$dgResults = Get-Element "dgResults"
$hostChart = Get-Element "hostChart"
$lstPresets = Get-Element "lstPresets"
$btnRunPreset = Get-Element "btnRunPreset"
$btnRefreshPresets = Get-Element "btnRefreshPresets"
$btnSaveSettings = Get-Element "btnSaveSettings"
$btnLoadSettings = Get-Element "btnLoadSettings"
$btnClearSettings = Get-Element "btnClearSettings"
$txtAltUser = Get-Element "txtAltUser"
$txtAltPass = Get-Element "txtAltPass"
$btnSaveCred = Get-Element "btnSaveCred"
$btnClearCred = Get-Element "btnClearCred"
$txtStatus = Get-Element "txtStatus"
$tabControl = Get-Element "tabControl"

# Set defaults
$cmbCategory.ItemsSource = @("User","Computer","OU","GPO","Security Group","Service Accounts","Servers (by OS or group)","Workstations (by OS or group)","Locked-out Users (basic)","Locked-out Users (with origin/event lookup)","Subnets (AD Sites & Services)","Firewall (GPO firewall rules)","All: Users+Computers")
$cmbCategory.SelectedIndex = 0
$txtFilter.Text = "*"
$txtExportFolder.Text = $Global:ExportFolder

# Populate presets list
function Refresh-Presets {
    $lstPresets.Items.Clear()
    foreach ($p in $Presets) {
        $lstPresets.Items.Add($p.Name) | Out-Null
    }
}
Refresh-Presets

# Create chart only if assembly is available
if ($ChartingAvailable) {
    try {
        $chartHost = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
        $chartHost.Width = 1000
        $chartHost.Height = 200
        $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "Main"
        $chartHost.ChartAreas.Add($area)
        $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Series1"
        $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
        $chartHost.Series.Add($series)
        $hostChart.Child = $chartHost
    } catch {
        Write-Warning "Chart creation failed. Chart area will be empty."
    }
} else {
    # If charting not available, hide the chart area
    if ($hostChart -ne $null) {
        $hostChart.Visibility = "Collapsed"
    }
}

# Load config if exists
$config = Load-Config
if ($config) {
    if ($config.LastFilter) { $txtFilter.Text = $config.LastFilter }
    if ($config.LastCategory) { $cmbCategory.SelectedItem = $config.LastCategory }
    if ($config.ExportFolder) { $txtExportFolder.Text = $config.ExportFolder }
    if ($config.CredProtected) {
        $cred = Unprotect-Credential -ProtectedString $config.CredProtected
        if ($cred) {
            $txtAltUser.Text = $cred.UserName
            # do not show password
        }
    }
}

# Utility to present results in WPF DataGrid (convert PSObject array to DataTable / ItemsSource)
function Show-Results {
    param([array]$results)
    if (-not $results) {
        $dgResults.ItemsSource = $null
        $txtStatus.Text = "No results"
        return
    }
    $dgResults.ItemsSource = $results
    $txtStatus.Text = "Found $($results.Count) item(s)."
    
    # Update chart (example: OS distribution for computer results)
    if ($ChartingAvailable -and $hostChart.Child -ne $null) {
        try {
            $chartHost = $hostChart.Child
            $chartHost.Series["Series1"].Points.Clear()
            if ($results -and $results[0].PSObject.Properties.Name -contains "OperatingSystem") {
                $groups = $results | Group-Object -Property OperatingSystem | Sort-Object Count -Descending
                foreach ($g in $groups) {
                    $pt = $chartHost.Series["Series1"].Points.Add($g.Count)
                    $pt.AxisLabel = $g.Name
                }
            }
        } catch {
            # Silently fail if chart update fails
        }
    }
}

# Main search dispatcher
function Run-Search {
    param([string]$Category, [string]$Filter, [switch]$ResolveOrigin)
    $txtStatus.Text = "Running search..."
    $window.Dispatcher.Invoke([action]{},[System.Windows.Threading.DispatcherPriority]::Background)
    $results = @()
    try {
        switch ($Category) {
            "User" { $results = Search-Users -filter $Filter }
            "Computer" { $results = Search-Computers -filter $Filter }
            "OU" { $results = Search-OUs -filter $Filter }
            "GPO" {
                if (-not $HasGPO) { [System.Windows.MessageBox]::Show("GroupPolicy module not available.","Missing Module"); $results = @() } else { $results = Search-GPOs -filter $Filter }
            }
            "Security Group" { $results = Search-Groups -filter $Filter }
            "Service Accounts" { $results = Search-ServiceAccounts -filter $Filter }
            "Servers (by OS or group)" { $results = Search-ServersOrWorkstations -filter $Filter -Servers }
            "Workstations (by OS or group)" { $results = Search-ServersOrWorkstations -filter $Filter }
            "Locked-out Users (basic)" { $results = Search-LockedOutUsers -ResolveOrigin:$false }
            "Locked-out Users (with origin/event lookup)" { $results = Search-LockedOutUsers -ResolveOrigin:$ResolveOrigin }
            "Subnets (AD Sites & Services)" { $results = Search-Subnets }
            "Firewall (GPO firewall rules)" { if (-not $HasGPO) { [System.Windows.MessageBox]::Show("GroupPolicy module not available.","Missing Module"); $results=@() } else { $results = Search-GPOFirewallRules -filter $Filter } }
            "All: Users+Computers" {
                $results = @()
                $results += (Search-Users -filter $Filter)
                $results += (Search-Computers -filter $Filter)
            }
            default { $results = @() }
        }
    } catch {
        [System.Windows.MessageBox]::Show("Search error: $($_.Exception.Message)","Error")
        $txtStatus.Text = "Error during search."
        return $null
    }
    return $results
}

# Buttons events
$btnRun.Add_Click({
    $category = $cmbCategory.SelectedItem
    $filter = $txtFilter.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filter)) { $filter = "*" }
    $resolve = $chkEventLookup.IsChecked
    $res = Run-Search -Category $category -Filter $filter -ResolveOrigin:$resolve
    Show-Results -results $res
})

$btnClear.Add_Click({
    $txtFilter.Text = "*"
    $dgResults.ItemsSource = $null
    $txtStatus.Text = "Cleared."
    if ($ChartingAvailable -and $hostChart.Child -ne $null) {
        $chartHost = $hostChart.Child
        $chartHost.Series["Series1"].Points.Clear()
    }
})

$btnPresets.Add_Click({
    $tabControl.SelectedIndex = 1
})

$btnRefreshPresets.Add_Click({
    Refresh-Presets
})

$btnRunPreset.Add_Click({
    $sel = $lstPresets.SelectedItem
    if (-not $sel) { [System.Windows.MessageBox]::Show("Select a preset first.","Presets") ; return }
    $preset = $Presets | Where-Object { $_.Name -eq $sel } | Select-Object -First 1
    if (-not $preset) { [System.Windows.MessageBox]::Show("Preset not found.","Presets") ; return }
    try {
        # Some presets use ScriptBlock with param($f)
        $sb = $preset.ScriptBlock
        if ($sb.Parameters.Count -gt 0) {
            $res = & $sb $txtFilter.Text
        } else {
            $res = & $sb
        }
        # convert results to array
        $arr = @()
        foreach ($r in $res) { $arr += $r }
        Show-Results -results $arr
    } catch {
        [System.Windows.MessageBox]::Show("Preset execution error: $($_.Exception.Message)","Presets")
    }
})

$btnOpenExport.Add_Click({
    $path = $txtExportFolder.Text
    if (-not (Test-Path $path)) { New-Item -Path $path -ItemType Directory -Force | Out-Null }
    Start-Process -FilePath $path
})

# Export button
$btnExport.Add_Click({
    $items = $dgResults.ItemsSource
    if (-not $items) { [System.Windows.MessageBox]::Show("No results to export.","Export") ; return }
    $formats = @("csv","xml","html","txt")  # default set
    # show small selection dialog - for brevity use messagebox choice; in production make a multi-select dialog
    $choices = [System.Windows.MessageBox]::Show("Export in all formats (CSV, XML, HTML, TXT, Excel, PDF, DOCX)?`nChoose Yes (all) or No (CSV/HTML only).","Export Options",[System.Windows.MessageBoxButton]::YesNoCancel)
    if ($choices -eq [System.Windows.MessageBoxResult]::Cancel) { return }
    if ($choices -eq [System.Windows.MessageBoxResult]::Yes) {
        $formats = @("csv","xml","html","txt","excel","pdf","docx")
    } else {
        $formats = @("csv","html")
    }
    # Convert ItemsSource (IList) to PSObject array
    $arr = @()
    foreach ($i in $items) { $arr += $i }
    Export-Results -Results $arr -Category $cmbCategory.SelectedItem -Filter $txtFilter.Text -ExportPath $txtExportFolder.Text -Formats $formats
})

# Settings Save/Load
$btnSaveSettings.Add_Click({
    $cfg = [ordered]@{}
    $cfg.LastCategory = $cmbCategory.SelectedItem
    $cfg.LastFilter = $txtFilter.Text
    $cfg.ExportFolder = $txtExportFolder.Text
    if ($txtAltUser.Text -and $txtAltPass.Password) {
        $cred = New-Object System.Management.Automation.PSCredential ($txtAltUser.Text, (ConvertTo-SecureString $txtAltPass.Password -AsPlainText -Force))
        $cfg.CredProtected = Protect-Credential -Credential $cred
    }
    Save-Config -config $cfg
    [System.Windows.MessageBox]::Show("Settings saved to $Global:ConfigFile","Settings")
})

$btnLoadSettings.Add_Click({
    $cfg = Load-Config
    if (-not $cfg) { [System.Windows.MessageBox]::Show("No saved settings found.","Settings") ; return }
    if ($cfg.LastCategory) { $cmbCategory.SelectedItem = $cfg.LastCategory }
    if ($cfg.LastFilter) { $txtFilter.Text = $cfg.LastFilter }
    if ($cfg.ExportFolder) { $txtExportFolder.Text = $cfg.ExportFolder }
    if ($cfg.CredProtected) {
        $cred = Unprotect-Credential -ProtectedString $cfg.CredProtected
        if ($cred) { $txtAltUser.Text = $cred.UserName }
    }
    [System.Windows.MessageBox]::Show("Settings loaded.","Settings")
})

$btnClearSettings.Add_Click({
    if (Test-Path $Global:ConfigFile) { Remove-Item $Global:ConfigFile -Force }
    [System.Windows.MessageBox]::Show("Settings cleared.","Settings")
})

$btnSaveCred.Add_Click({
    if (-not $txtAltUser.Text -or -not $txtAltPass.Password) { [System.Windows.MessageBox]::Show("Enter username and password to save.","Credentials") ; return }
    $cred = New-Object System.Management.Automation.PSCredential ($txtAltUser.Text, (ConvertTo-SecureString $txtAltPass.Password -AsPlainText -Force))
    $cfg = Load-Config
    if (-not $cfg) { $cfg = [ordered]@{} }
    $cfg.CredProtected = Protect-Credential -Credential $cred
    Save-Config -config $cfg
    [System.Windows.MessageBox]::Show("Credentials saved (protected to current user profile).","Credentials")
})

$btnClearCred.Add_Click({
    $cfg = Load-Config
    if ($cfg -and $cfg.CredProtected) { 
        $cfg.PSObject.Properties.Remove('CredProtected') 
        Save-Config -config $cfg 
        $txtAltUser.Text = ""
        $txtAltPass.Password = ""
        [System.Windows.MessageBox]::Show("Stored credentials cleared.","Credentials") 
    }
    else { [System.Windows.MessageBox]::Show("No stored credentials found.","Credentials") }
})

# Keyboard shortcuts
$window.Add_KeyDown({
    param($sender,$e)
    if ($e.Key -eq 'F5') { 
        # Trigger the Run Search button click
        $btnRun.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
    }
})

# Show window
$null = $window.ShowDialog()
