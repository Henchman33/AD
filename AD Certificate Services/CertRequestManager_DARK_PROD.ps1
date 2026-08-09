#Requires -Version 5.1
<#
.SYNOPSIS
    Certificate Request Manager - WPF GUI for building and submitting internal AD CS
    certificate requests (CSR -> SAN.inf -> policy CSR -> issued .cer/.p7b).

.DESCRIPTION
    Designed to run directly on the Issuing CA/PKI server. Given a customer- or
    self-generated .csr file, this tool:
      1. Creates a per-server working folder under a configurable base path
      2. Builds a correct SAN.inf (SAN extension policy file) from entered SAN entries
      3. Runs "certreq -policy" to merge the SAN extension into the CSR
      4. Runs "certreq -submit" against the issuing CA with the chosen template
      5. Leaves the resulting .csr / .cer / .p7b files in the working folder

    Certificate templates are pulled live from the CA (Get-CATemplate) so the
    exact internal template Name (not the display name) is always used with
    certreq -attrib "CertificateTemplate:<Name>".

.NOTES
    Run elevated, on the Issuing CA, as an account with Manage CA / Issue and
    Manage Certificates permission (or at least Request Certificates permission
    plus auto-approval on the template) so certreq -submit succeeds without a
    manual approval step. If a request lands Pending, use the "Retrieve Pending"
    action once it has been approved in the Certification Authority console.
#>

[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName Microsoft.VisualBasic   # for folder browser InputBox fallback

# ------------------------------------------------------------------------------------
# Config persistence (remembers CA config string + base save path between runs)
# ------------------------------------------------------------------------------------
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::CurrentDirectory }
$script:ConfigPath = Join-Path $script:ScriptRoot 'CertRequestManager.config.json'

function Get-ToolConfig {
    $default = [ordered]@{
        CAConfig       = 'USRNOPPKICA01.MYIGT.COM\MYIGT-ISSUING-CA'
        BasePath       = 'C:\PKI\CertRequests'
        DomainSuffixes = @('igtsap.ad.igt.com', 'ad.igt.com', 'is.ad.igt.com', 'myigt.com')
    }
    if (Test-Path $script:ConfigPath) {
        try {
            $loaded = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($key in $default.Keys) {
                if ($loaded.PSObject.Properties.Name -contains $key -and $loaded.$key) {
                    $default[$key] = $loaded.$key
                }
            }
        } catch { }
    }
    return $default
}

function Save-ToolConfig {
    param($Config)
    try {
        $Config | ConvertTo-Json -Depth 3 | Set-Content -Path $script:ConfigPath -Encoding UTF8
    } catch { }
}

$script:Config = Get-ToolConfig

# ------------------------------------------------------------------------------------
# XAML
# ------------------------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Certificate Request Manager - Issuing CA" Height="980" Width="1220"
        WindowStartupLocation="CenterScreen" Background="#1E1E2E" FontFamily="Segoe UI" FontSize="14">
    <Window.Resources>
        <SolidColorBrush x:Key="Panel" Color="#181825"/>
        <SolidColorBrush x:Key="Border" Color="#313244"/>
        <SolidColorBrush x:Key="TextMain" Color="#CDD6F4"/>
        <SolidColorBrush x:Key="TextSub" Color="#A6ADC8"/>
        <SolidColorBrush x:Key="Accent" Color="#89B4FA"/>
        <SolidColorBrush x:Key="Good" Color="#A6E3A1"/>
        <SolidColorBrush x:Key="Warn" Color="#F9E2AF"/>
        <SolidColorBrush x:Key="Bad" Color="#F38BA8"/>

        <Style TargetType="Label">
            <Setter Property="Foreground" Value="{StaticResource TextSub}"/>
            <Setter Property="Padding" Value="0,5,0,3"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource Panel}"/>
            <Setter Property="Foreground" Value="{StaticResource TextMain}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="CaretBrush" Value="{StaticResource TextMain}"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{StaticResource Panel}"/>
            <Setter Property="Foreground" Value="#11111B"/>
            <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="15"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="{StaticResource TextMain}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,0,10,0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="{StaticResource Accent}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{StaticResource Panel}"/>
            <Setter Property="Foreground" Value="{StaticResource TextMain}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Border}"/>
            <Setter Property="RowBackground" Value="{StaticResource Panel}"/>
            <Setter Property="AlternatingRowBackground" Value="#1E1E2E"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource Border}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="RowHeight" Value="30"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Certificate Request Manager" FontSize="18" FontWeight="Bold"
                   Foreground="{StaticResource TextMain}" Margin="0,0,0,12"/>

        <!-- Request Details -->
        <GroupBox Grid.Row="1" Header="Request Details" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,10,0">
                    <Label Content="Server / Service Name (e.g. SERVER01)"/>
                    <TextBox x:Name="TxtServerName"/>
                </StackPanel>
                <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,10,0">
                    <Label Content="Domain Suffix for FQDN"/>
                    <ComboBox x:Name="CmbDomainSuffix" IsEditable="True"/>
                </StackPanel>
                <StackPanel Grid.Row="0" Grid.Column="2">
                    <Label Content="Resulting FQDN"/>
                    <TextBox x:Name="TxtFqdn" IsReadOnly="True" Background="#11111B"/>
                </StackPanel>

                <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,8,10,0">
                    <Label Content="Certificate Template (Get-CATemplate 'Name')"/>
                    <ComboBox x:Name="CmbTemplate" IsEditable="True"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="1" Margin="0,8,10,0">
                    <Label Content="Issuing CA -config string"/>
                    <TextBox x:Name="TxtCAConfig"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,8,0,0">
                    <Label Content="Base Working Folder (on this server)"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="TxtBasePath" Grid.Column="0"/>
                        <Button x:Name="BtnBrowseBase" Grid.Column="1" Content="..." Margin="6,0,0,0" Padding="8,4"/>
                    </Grid>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- CSR Source -->
        <GroupBox Grid.Row="2" Header="CSR Source" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row="0" Grid.Column="0">
                    <Label Content="CSR File Path (customer/self-generated .csr)"/>
                    <TextBox x:Name="TxtCsrPath"/>
                </StackPanel>
                <Button x:Name="BtnBrowseCsr" Grid.Row="0" Grid.Column="1" Content="Browse .csr" Margin="8,20,0,0" Padding="10,6"/>

                <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,8,0,0">
                    <Label Content="Or paste raw CSR text here (used instead of the file above if not empty)"/>
                    <TextBox x:Name="TxtCsrPaste" Height="80" AcceptsReturn="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="13"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- SAN entries -->
        <GroupBox Grid.Row="3" Header="Subject Alternative Names (SAN.inf)" Padding="10">
            <StackPanel>
                <DataGrid x:Name="GridSan" Height="160" AutoGenerateColumns="False" CanUserAddRows="False">
                    <DataGrid.Columns>
                        <DataGridComboBoxColumn Header="Type" Width="120" x:Name="ColSanType"/>
                        <DataGridTextColumn Header="Value" Width="*" Binding="{Binding Value}"/>
                    </DataGrid.Columns>
                </DataGrid>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                    <Button x:Name="BtnAddSan" Content="+ Add SAN Row"/>
                    <Button x:Name="BtnRemoveSan" Content="Remove Selected"/>
                    <Button x:Name="BtnAddFqdnSan" Content="Add FQDN as DNS SAN"/>
                </StackPanel>
            </StackPanel>
        </GroupBox>

        <!-- Actions + Log -->
        <Grid Grid.Row="4">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                <Button x:Name="BtnStep1" Content="1. Create Folder + SAN.inf" Padding="12,8"/>
                <Button x:Name="BtnStep2" Content="2. certreq -policy" Padding="12,8"/>
                <Button x:Name="BtnStep3" Content="3. certreq -submit (Issue)" Padding="12,8"/>
                <Button x:Name="BtnRetrievePending" Content="Retrieve Pending Request" Padding="12,8"/>
                <Button x:Name="BtnRefreshTemplates" Content="Refresh Templates from CA" Padding="12,8"/>
                <Button x:Name="BtnOpenFolder" Content="Open Working Folder" Padding="12,8"/>
            </StackPanel>

            <GroupBox Grid.Row="1" Header="Output Log" Padding="8">
                <TextBox x:Name="TxtLog" IsReadOnly="True" FontFamily="Consolas" FontSize="13"
                         VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                         TextWrapping="NoWrap" Background="#11111B"/>
            </GroupBox>
        </Grid>

        <StatusBar Grid.Row="5" Background="#181825" Margin="0,10,0,0">
            <StatusBarItem>
                <TextBlock x:Name="TxtStatus" Text="Ready." Foreground="#CDD6F4" FontSize="14"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# ------------------------------------------------------------------------------------
# Named element lookup
# ------------------------------------------------------------------------------------
$ctrl = @{}
foreach ($name in @(
        'TxtServerName','CmbDomainSuffix','TxtFqdn','CmbTemplate','TxtCAConfig','TxtBasePath','BtnBrowseBase',
        'TxtCsrPath','BtnBrowseCsr','TxtCsrPaste','GridSan','BtnAddSan','BtnRemoveSan','BtnAddFqdnSan',
        'BtnStep1','BtnStep2','BtnStep3','BtnRetrievePending','BtnRefreshTemplates','BtnOpenFolder',
        'TxtLog','TxtStatus')) {
    $ctrl[$name] = $window.FindName($name)
}

# ------------------------------------------------------------------------------------
# Logging helpers
# ------------------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $ctrl.TxtLog.AppendText("[$stamp] [$Level] $Message`r`n")
    $ctrl.TxtLog.ScrollToEnd()
}

function Set-Status {
    param([string]$Message, [string]$Color = '#CDD6F4')
    $ctrl.TxtStatus.Text = $Message
    $ctrl.TxtStatus.Foreground = [Windows.Media.BrushConverter]::new().ConvertFromString($Color)
}

# ------------------------------------------------------------------------------------
# SAN grid data source (bindable list of PSObjects: Type, Value)
# ------------------------------------------------------------------------------------
$script:SanEntries = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
$ctrl.GridSan.ItemsSource = $script:SanEntries

# Populate the Type combo column with DNS / IPAddress choices
$typeCol = $ctrl.GridSan.Columns[0]
$typeCol.ItemsSource = @('dns', 'ipaddress')
$typeCol.SelectedItemBinding = New-Object System.Windows.Data.Binding 'Type'

function Add-SanRow {
    param([string]$Type = 'dns', [string]$Value = '')
    $script:SanEntries.Add([pscustomobject]@{ Type = $Type; Value = $Value })
}

# ------------------------------------------------------------------------------------
# Populate static combo boxes from config
# ------------------------------------------------------------------------------------
foreach ($suffix in $script:Config.DomainSuffixes) { [void]$ctrl.CmbDomainSuffix.Items.Add($suffix) }
$ctrl.CmbDomainSuffix.SelectedIndex = 0
$ctrl.TxtCAConfig.Text = $script:Config.CAConfig
$ctrl.TxtBasePath.Text = $script:Config.BasePath

# Fallback template list (used until "Refresh Templates from CA" is run / if CA module unavailable)
$fallbackTemplates = @(
    'WorkstationAuthenticationOnly','SCCM DP Certificate','Intune User Code Signing','Application Code Signing',
    'MYIGT macOS Client Authentication','ConfigMgrClientCertificateNoCDP','ConfigMgrClientCertificateSupplyFQDN',
    'WinWiFi Device Cert','Intune ConfigMgr Client Certificate','HorizonWebServerTemplate','ConfigMgrClientCertificate',
    'Workstation Auth SCCM Distribution Point','MYIGT Workstation Authentication','MYIGT User Authentication',
    'MYIGT SCOM Certificate Template','IGTCMG Web Server Template','CodeSigningInternal','Web Server SCCM Site System',
    'sbX Signing Certificate','sbX Licensing Service','PxGrid-ISE','MYIGT Printer Authentication',
    'MYIGT Intune Workstation Authentication','MYIGT Intune User Authentication','MyIGT RAS and IAS Server',
    'ITALY SURFACE','IGTPLC - ConfigMgr Web Server Certificate','IGTPLC - CMG Cloud Distribution Point',
    'IDC checkpoint','Domain Controller Authentication (Kerberos)','Administrator','WEbServerIGTInternal'
)
foreach ($t in $fallbackTemplates) { [void]$ctrl.CmbTemplate.Items.Add($t) }
$ctrl.CmbTemplate.Text = 'WEbServerIGTInternal'

Write-Log "Loaded fallback template list (display names may not match certreq's internal Name). Click 'Refresh Templates from CA' to pull exact Name values via Get-CATemplate."

# Seed one blank SAN row
Add-SanRow -Type 'dns' -Value ''

# ------------------------------------------------------------------------------------
# FQDN auto-update
# ------------------------------------------------------------------------------------
function Update-Fqdn {
    $srv = $ctrl.TxtServerName.Text.Trim()
    $suf = $ctrl.CmbDomainSuffix.Text.Trim()
    if ($srv -and $suf) {
        $ctrl.TxtFqdn.Text = "$($srv.ToLower()).$($suf.ToLower())"
    } else {
        $ctrl.TxtFqdn.Text = ''
    }
}
$ctrl.TxtServerName.Add_TextChanged({ Update-Fqdn })
$ctrl.CmbDomainSuffix.Add_TextChanged({ Update-Fqdn })
$ctrl.CmbDomainSuffix.Add_SelectionChanged({ Update-Fqdn })

# ------------------------------------------------------------------------------------
# Working-folder path resolution
# ------------------------------------------------------------------------------------
function Get-WorkingFolder {
    $srv = $ctrl.TxtServerName.Text.Trim()
    if (-not $srv) { throw 'Enter a Server / Service Name first.' }
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    if ($srv.IndexOfAny($invalid) -ge 0) { throw 'Server / Service Name contains invalid characters for a folder name.' }
    $base = $ctrl.TxtBasePath.Text.Trim()
    if (-not $base) { throw 'Set a Base Working Folder first.' }
    return (Join-Path $base $srv)
}

# ------------------------------------------------------------------------------------
# Browse buttons
# ------------------------------------------------------------------------------------
$ctrl.BtnBrowseCsr.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSR files (*.csr)|*.csr|All files (*.*)|*.*'
    $dlg.Title = 'Select CSR file'
    if ($dlg.ShowDialog()) {
        $ctrl.TxtCsrPath.Text = $dlg.FileName
        Write-Log "Selected CSR file: $($dlg.FileName)"
    }
})

$ctrl.BtnBrowseBase.Add_Click({
    $folder = [Microsoft.VisualBasic.Interaction]::InputBox(
        'Base working folder path (all per-server request folders are created under this path):',
        'Base Working Folder', $ctrl.TxtBasePath.Text)
    if ($folder) { $ctrl.TxtBasePath.Text = $folder }
})

$ctrl.BtnOpenFolder.Add_Click({
    try {
        $folder = Get-WorkingFolder
        if (Test-Path $folder) {
            Start-Process explorer.exe $folder
        } else {
            Write-Log "Folder does not exist yet: $folder" 'WARN'
        }
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
    }
})

# ------------------------------------------------------------------------------------
# SAN row buttons
# ------------------------------------------------------------------------------------
$ctrl.BtnAddSan.Add_Click({ Add-SanRow -Type 'dns' -Value '' })

$ctrl.BtnRemoveSan.Add_Click({
    $selected = $ctrl.GridSan.SelectedItem
    if ($selected) { [void]$script:SanEntries.Remove($selected) }
})

$ctrl.BtnAddFqdnSan.Add_Click({
    Update-Fqdn
    $fqdn = $ctrl.TxtFqdn.Text.Trim()
    if (-not $fqdn) { Write-Log 'Enter Server Name and Domain Suffix first.' 'WARN'; return }
    if (-not ($script:SanEntries | Where-Object { $_.Type -eq 'dns' -and $_.Value -eq $fqdn })) {
        Add-SanRow -Type 'dns' -Value $fqdn
        Write-Log "Added FQDN as DNS SAN: $fqdn"
    }
})

# ------------------------------------------------------------------------------------
# Refresh templates from the live CA (Get-CATemplate = exact internal Name)
# ------------------------------------------------------------------------------------
$ctrl.BtnRefreshTemplates.Add_Click({
    try {
        if (-not (Get-Command Get-CATemplate -ErrorAction SilentlyContinue)) {
            Import-Module PSPKI -ErrorAction Stop
        }
        $templates = Get-CATemplate | Select-Object -ExpandProperty Templates -ErrorAction SilentlyContinue
        if (-not $templates) { $templates = Get-CATemplate | Sort-Object Name }
        $current = $ctrl.CmbTemplate.Text
        $ctrl.CmbTemplate.Items.Clear()
        foreach ($t in ($templates | Sort-Object Name)) { [void]$ctrl.CmbTemplate.Items.Add($t.Name) }
        $ctrl.CmbTemplate.Text = $current
        Write-Log "Loaded $($ctrl.CmbTemplate.Items.Count) templates directly from the CA." 'INFO'
        Set-Status 'Templates refreshed from CA.' '#A6E3A1'
    } catch {
        Write-Log "Could not query the CA directly (needs to run on/against the CA with the ADCS or PSPKI module): $($_.Exception.Message)" 'WARN'
        Write-Log "Falling back to certutil -CATemplates parsing..." 'INFO'
        try {
            $raw = & certutil -CATemplates 2>&1
            $names = @()
            foreach ($line in $raw) {
                if ($line -match '^([^:]+):') { $names += $Matches[1].Trim() }
            }
            if ($names) {
                $current = $ctrl.CmbTemplate.Text
                $ctrl.CmbTemplate.Items.Clear()
                foreach ($n in ($names | Sort-Object -Unique)) { [void]$ctrl.CmbTemplate.Items.Add($n) }
                $ctrl.CmbTemplate.Text = $current
                Write-Log "Loaded $($names.Count) templates via certutil -CATemplates." 'INFO'
            } else {
                Write-Log 'certutil -CATemplates returned no parsable template names.' 'ERROR'
            }
        } catch {
            Write-Log "certutil fallback also failed: $($_.Exception.Message)" 'ERROR'
        }
    }
})

# ------------------------------------------------------------------------------------
# Run an external command, streaming stdout/stderr into the log. Returns exit code.
# ------------------------------------------------------------------------------------
function Invoke-LoggedCommand {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [string]$WorkingDirectory
    )
    # Build a properly quoted argument string (ProcessStartInfo.ArgumentList is not
    # reliably present under Windows PowerShell 5.1 / .NET Framework, so avoid it).
    $quotedArgs = $ArgumentList | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }
    $argString = $quotedArgs -join ' '

    Write-Log "> $FilePath $argString"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $argString
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($stdout) { $stdout -split "`r?`n" | Where-Object { $_ } | ForEach-Object { Write-Log $_ } }
    if ($stderr) { $stderr -split "`r?`n" | Where-Object { $_ } | ForEach-Object { Write-Log $_ 'ERROR' } }

    return [pscustomobject]@{ ExitCode = $proc.ExitCode; StdOut = $stdout; StdErr = $stderr }
}

# ------------------------------------------------------------------------------------
# Step 1: Create working folder, drop in the CSR, write SAN.inf
# ------------------------------------------------------------------------------------
$ctrl.BtnStep1.Add_Click({
    try {
        Update-Fqdn
        $srv  = $ctrl.TxtServerName.Text.Trim()
        $fqdn = $ctrl.TxtFqdn.Text.Trim()
        if (-not $srv)  { throw 'Enter a Server / Service Name.' }
        if (-not $fqdn) { throw 'Select a Domain Suffix so the FQDN can be built.' }

        $sanRows = $script:SanEntries | Where-Object { $_.Value -and $_.Value.Trim() }
        if (-not $sanRows) { throw 'Add at least one SAN entry (or click "Add FQDN as DNS SAN").' }

        $folder = Get-WorkingFolder
        if (-not (Test-Path $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
            Write-Log "Created working folder: $folder"
        } else {
            Write-Log "Working folder already exists: $folder" 'WARN'
        }

        # --- Place the CSR ---
        $csrDest = Join-Path $folder "$srv.csr"
        $pasted = $ctrl.TxtCsrPaste.Text.Trim()
        if ($pasted) {
            Set-Content -Path $csrDest -Value $pasted -Encoding ASCII
            Write-Log "Wrote pasted CSR text to: $csrDest"
        } elseif ($ctrl.TxtCsrPath.Text.Trim()) {
            $src = $ctrl.TxtCsrPath.Text.Trim()
            if (-not (Test-Path $src)) { throw "CSR file not found: $src" }
            Copy-Item -Path $src -Destination $csrDest -Force
            Write-Log "Copied CSR from $src to $csrDest"
        } else {
            throw 'Provide a CSR either by browsing to a file or pasting the CSR text.'
        }

        # --- Build SAN.inf ---
        $sanLines = foreach ($row in $sanRows) { "$($row.Type)=$($row.Value.Trim())" }
        $sanText = ($sanLines -join '&')

        $infLines = @(
            '[Version]'
            'Signature="$Windows NT$"'
            ''
            '[Extensions]'
            '2.5.29.17 = "{text}"'
        )
        foreach ($line in $sanLines) { $infLines += "_continue_ = `"$line&`"" }

        $infPath = Join-Path $folder 'SAN.inf'
        Set-Content -Path $infPath -Value $infLines -Encoding ASCII
        Write-Log "Wrote SAN.inf with $($sanRows.Count) SAN entries to: $infPath"
        Write-Log "SAN string: $sanText"

        Set-Status "Step 1 complete: $folder" '#A6E3A1'
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Set-Status 'Step 1 failed - see log.' '#F38BA8'
    }
})

# ------------------------------------------------------------------------------------
# Step 2: certreq -policy  (merge SAN.inf into the CSR -> policy CSR)
# ------------------------------------------------------------------------------------
$ctrl.BtnStep2.Add_Click({
    try {
        $srv  = $ctrl.TxtServerName.Text.Trim()
        $fqdn = $ctrl.TxtFqdn.Text.Trim()
        $caConfig = $ctrl.TxtCAConfig.Text.Trim()
        if (-not $srv -or -not $fqdn) { throw 'Server Name / FQDN not set - run Step 1 first.' }
        if (-not $caConfig) { throw 'Set the Issuing CA -config string.' }

        $folder = Get-WorkingFolder
        $csrIn  = "$srv.csr"
        $infFile = 'SAN.inf'
        $csrOut = "$fqdn.csr"

        foreach ($f in @($csrIn, $infFile)) {
            if (-not (Test-Path (Join-Path $folder $f))) { throw "Missing $f in $folder - run Step 1 first." }
        }

        $result = Invoke-LoggedCommand -FilePath 'certreq.exe' `
            -ArgumentList @('-policy', '-config', $caConfig, $csrIn, $infFile, $csrOut) `
            -WorkingDirectory $folder

        if ($result.ExitCode -eq 0 -and (Test-Path (Join-Path $folder $csrOut))) {
            Write-Log "Policy CSR created: $csrOut" 'INFO'
            Set-Status "Step 2 complete: $csrOut" '#A6E3A1'
        } else {
            throw "certreq -policy exited with code $($result.ExitCode). Check the log above for details."
        }
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Set-Status 'Step 2 failed - see log.' '#F38BA8'
    }
})

# ------------------------------------------------------------------------------------
# Step 3: certreq -submit  (issue the cert against the chosen template)
# ------------------------------------------------------------------------------------
$ctrl.BtnStep3.Add_Click({
    try {
        $srv  = $ctrl.TxtServerName.Text.Trim()
        $fqdn = $ctrl.TxtFqdn.Text.Trim()
        $caConfig = $ctrl.TxtCAConfig.Text.Trim()
        $template = $ctrl.CmbTemplate.Text.Trim()
        if (-not $fqdn) { throw 'FQDN not set - run Step 1 first.' }
        if (-not $caConfig) { throw 'Set the Issuing CA -config string.' }
        if (-not $template) { throw 'Choose or type a Certificate Template Name.' }
        if ($template -match '\s') {
            Write-Log "Template '$template' contains spaces - certreq needs the internal Name (no spaces). Click 'Refresh Templates from CA' and re-select if this fails." 'WARN'
        }

        $folder = Get-WorkingFolder
        $csrIn = "$fqdn.csr"
        if (-not (Test-Path (Join-Path $folder $csrIn))) { throw "Missing $csrIn in $folder - run Step 2 first." }

        $cerOut = "$fqdn.cer"
        $p7bOut = "$fqdn.p7b"

        $result = Invoke-LoggedCommand -FilePath 'certreq.exe' `
            -ArgumentList @('-submit', '-attrib', "CertificateTemplate:$template", '-config', $caConfig, $csrIn, $cerOut, $p7bOut) `
            -WorkingDirectory $folder

        $combined = "$($result.StdOut)`n$($result.StdErr)"
        if ($combined -match 'Certificate Request Processor: The request submission failed') {
            throw 'Request submission failed - verify the template Name and CA -config string.'
        } elseif ($combined -match 'RequestId "(\d+)"') {
            $reqId = $Matches[1]
            if ($combined -match 'Certificate retrieved') {
                Write-Log "Certificate issued. Request ID: $reqId. Files: $cerOut, $p7bOut" 'INFO'
                Set-Status "Certificate issued (Request ID $reqId)." '#A6E3A1'
            } else {
                Write-Log "Request submitted but is PENDING approval on the CA. Request ID: $reqId" 'WARN'
                Write-Log "Approve it in the Certification Authority console, then use 'Retrieve Pending Request'." 'WARN'
                Set-Status "Pending CA approval - Request ID $reqId" '#F9E2AF'
            }
        } else {
            Write-Log 'Could not confirm issuance from certreq output - review the log above.' 'WARN'
        }
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Set-Status 'Step 3 failed - see log.' '#F38BA8'
    }
})

# ------------------------------------------------------------------------------------
# Retrieve a previously pending request once approved
# ------------------------------------------------------------------------------------
$ctrl.BtnRetrievePending.Add_Click({
    try {
        $reqId = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the pending Request ID to retrieve:', 'Retrieve Pending Request', '')
        if (-not $reqId) { return }
        $caConfig = $ctrl.TxtCAConfig.Text.Trim()
        $fqdn = $ctrl.TxtFqdn.Text.Trim()
        $folder = Get-WorkingFolder
        $cerOut = "$fqdn.cer"

        $result = Invoke-LoggedCommand -FilePath 'certreq.exe' `
            -ArgumentList @('-retrieve', '-config', $caConfig, $reqId, $cerOut) `
            -WorkingDirectory $folder

        if ($result.ExitCode -eq 0 -and (Test-Path (Join-Path $folder $cerOut))) {
            Write-Log "Retrieved issued certificate: $cerOut" 'INFO'
            Set-Status 'Pending request retrieved.' '#A6E3A1'
        } else {
            Write-Log 'Request may still be pending or was denied - check the Certification Authority console.' 'WARN'
        }
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
    }
})

# ------------------------------------------------------------------------------------
# Persist config on close
# ------------------------------------------------------------------------------------
$window.Add_Closing({
    $script:Config.CAConfig = $ctrl.TxtCAConfig.Text.Trim()
    $script:Config.BasePath = $ctrl.TxtBasePath.Text.Trim()
    Save-ToolConfig -Config $script:Config
})

Write-Log 'Certificate Request Manager ready. Run elevated on the Issuing CA for certreq -submit to succeed without manual approval.'
[void]$window.ShowDialog()
