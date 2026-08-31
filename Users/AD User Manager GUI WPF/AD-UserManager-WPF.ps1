#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Modern WPF AD User Manager – Light theme, Blue tabs, Credential support
.DESCRIPTION
    Create / Manage (multi-select bulk) / Duplicate users.
    Connect to any domain/child domain with alternate credentials.
    Searchable OU TreeView, progress bar, CSV export.
.NOTES
    Run: powershell.exe -STA -File .\AD-UserManager-WPF.ps1
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

#region ===== Business Logic =====

function Get-ADDomainContext {
    param([string]$DomainFQDN, [PSCredential]$Credential)
    try {
        $p = @{}
        if (-not [string]::IsNullOrWhiteSpace($DomainFQDN)) { $p.Identity = $DomainFQDN }
        if ($Credential) { $p.Credential = $Credential }
        return Get-ADDomain @p -ErrorAction Stop
    } catch {
        throw "Failed to resolve domain '$DomainFQDN': $($_.Exception.Message)"
    }
}

function Get-OUTree {
    param([string]$SearchBase, [string]$Server, [PSCredential]$Credential)
    $p = @{ Filter = '*'; SearchBase = $SearchBase; Properties = 'DistinguishedName','Name' }
    if ($Server)     { $p.Server     = $Server }
    if ($Credential) { $p.Credential = $Credential }
    Get-ADOrganizationalUnit @p | Sort-Object DistinguishedName
}

function Get-SecurityGroups {
    param([string]$SearchBase, [string]$Server, [PSCredential]$Credential)
    $p = @{ Filter = "GroupCategory -eq 'Security'"; SearchBase = $SearchBase; Properties = 'Name','DistinguishedName','SamAccountName' }
    if ($Server)     { $p.Server     = $Server }
    if ($Credential) { $p.Credential = $Credential }
    Get-ADGroup @p | Sort-Object Name
}

function New-ADUserFromForm {
    param([hashtable]$UserData, [string]$Server, [PSCredential]$Credential)
    $params = @{
        Name                  = $UserData.DisplayName
        GivenName             = $UserData.GivenName
        Surname               = $UserData.Surname
        SamAccountName        = $UserData.SamAccountName
        UserPrincipalName     = $UserData.UPN
        DisplayName           = $UserData.DisplayName
        Description           = $UserData.Description
        Path                  = $UserData.OU
        AccountPassword       = (ConvertTo-SecureString $UserData.Password -AsPlainText -Force)
        Enabled               = $true
        ChangePasswordAtLogon = $UserData.ChangePasswordAtLogon
        PasswordNeverExpires  = $UserData.PasswordNeverExpires
        ErrorAction           = 'Stop'
    }
    if ($Server)     { $params.Server     = $Server }
    if ($Credential) { $params.Credential = $Credential }
    $user = New-ADUser @params -PassThru
    foreach ($g in @($UserData.Groups)) {
        $ap = @{ Identity = $g; Members = $user.SamAccountName; ErrorAction = 'Stop' }
        if ($Server)     { $ap.Server     = $Server }
        if ($Credential) { $ap.Credential = $Credential }
        Add-ADGroupMember @ap
    }
    return $user
}

function Copy-ADUserAsTemplate {
    param([string]$SourceSam, [hashtable]$NewUserData, [string]$Server, [PSCredential]$Credential)
    $gp = @{ Identity = $SourceSam; Properties = '*'; ErrorAction = 'Stop' }
    if ($Server)     { $gp.Server     = $Server }
    if ($Credential) { $gp.Credential = $Credential }
    $source = Get-ADUser @gp

    $params = @{
        Name                  = $NewUserData.DisplayName
        GivenName             = $NewUserData.GivenName
        Surname               = $NewUserData.Surname
        SamAccountName        = $NewUserData.SamAccountName
        UserPrincipalName     = $NewUserData.UPN
        DisplayName           = $NewUserData.DisplayName
        Description           = if ($NewUserData.Description) { $NewUserData.Description } else { $source.Description }
        Path                  = $NewUserData.OU
        AccountPassword       = (ConvertTo-SecureString $NewUserData.Password -AsPlainText -Force)
        Enabled               = $true
        ChangePasswordAtLogon = $NewUserData.ChangePasswordAtLogon
        PasswordNeverExpires  = $NewUserData.PasswordNeverExpires
        ErrorAction           = 'Stop'
    }
    if ($source.Department) { $params.Department = $source.Department }
    if ($source.Title)      { $params.Title      = $source.Title }
    if ($source.Office)     { $params.Office     = $source.Office }
    if ($source.Company)    { $params.Company    = $source.Company }
    if ($source.Manager)    { $params.Manager    = $source.Manager }
    if ($Server)     { $params.Server     = $Server }
    if ($Credential) { $params.Credential = $Credential }

    $newUser = New-ADUser @params -PassThru

    $mp = @{ Identity = $source.DistinguishedName; ErrorAction = 'SilentlyContinue' }
    if ($Server)     { $mp.Server     = $Server }
    if ($Credential) { $mp.Credential = $Credential }
    $groups = Get-ADPrincipalGroupMembership @mp | Where-Object { $_.SamAccountName -ne 'Domain Users' }
    foreach ($g in $groups) {
        try {
            $ap = @{ Identity = $g; Members = $newUser.SamAccountName; ErrorAction = 'Stop' }
            if ($Server)     { $ap.Server     = $Server }
            if ($Credential) { $ap.Credential = $Credential }
            Add-ADGroupMember @ap
        } catch { Write-Warning "Group $($g.Name): $($_.Exception.Message)" }
    }
    return $newUser
}

function Write-ADLog {
    param([string]$Action, [string]$Target, [string]$Result, [string]$Details = '')
    $logDir = Join-Path $env:USERPROFILE 'AD-UserManager-Logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir ("ADUserMgr_{0:yyyy-MM-dd}.csv" -f (Get-Date))
    [PSCustomObject]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Operator  = $env:USERNAME
        Computer  = $env:COMPUTERNAME
        Action    = $Action
        Target    = $Target
        Result    = $Result
        Details   = $Details
    } | Export-Csv $logFile -Append -NoTypeInformation -Encoding UTF8
}

function Build-OUTreeViewItems {
    param([array]$OUs, [string]$DomainDN)
    $root = New-Object System.Windows.Controls.TreeViewItem
    $root.Header = "(Domain Root)"
    $root.Tag = $DomainDN
    $root.IsExpanded = $true
    $nodeMap = @{ $DomainDN = $root }

    foreach ($ou in ($OUs | Sort-Object { ($_.DistinguishedName -split ',').Count })) {
        $dn = $ou.DistinguishedName
        $parentDN = ($dn -split ',', 2)[1]
        $item = New-Object System.Windows.Controls.TreeViewItem
        $item.Header = $ou.Name
        $item.Tag = $dn
        $item.IsExpanded = $false
        $nodeMap[$dn] = $item
        if ($nodeMap.ContainsKey($parentDN)) {
            [void]$nodeMap[$parentDN].Items.Add($item)
        } else {
            [void]$root.Items.Add($item)
        }
    }
    return $root
}

function Filter-OUTree {
    param($TreeView, [string]$Filter, $AllRoot)
    $TreeView.Items.Clear()
    if ([string]::IsNullOrWhiteSpace($Filter)) {
        [void]$TreeView.Items.Add($AllRoot)
        return
    }
    $filterLower = $Filter.ToLower()
    function Copy-MatchingNodes($sourceItem) {
        $match = $sourceItem.Header.ToString().ToLower().Contains($filterLower)
        $newItem = $null
        $childMatches = @()
        foreach ($child in $sourceItem.Items) {
            $copied = Copy-MatchingNodes $child
            if ($copied) { $childMatches += $copied }
        }
        if ($match -or $childMatches.Count -gt 0) {
            $newItem = New-Object System.Windows.Controls.TreeViewItem
            $newItem.Header = $sourceItem.Header
            $newItem.Tag = $sourceItem.Tag
            $newItem.IsExpanded = $true
            foreach ($c in $childMatches) { [void]$newItem.Items.Add($c) }
        }
        return $newItem
    }
    $filtered = Copy-MatchingNodes $AllRoot
    if ($filtered) { [void]$TreeView.Items.Add($filtered) }
}

#endregion

#region ===== XAML – White background + Dark-blue tabs with white text =====

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User Manager" Height="780" Width="1040"
        MinHeight="650" MinWidth="920"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13"
        Background="White" Foreground="#1A1A1A">

  <Window.Resources>
    <!-- Light professional palette -->
    <SolidColorBrush x:Key="BgWindow"     Color="White"/>
    <SolidColorBrush x:Key="BgPanel"      Color="#F7F9FC"/>
    <SolidColorBrush x:Key="BgInput"      Color="White"/>
    <SolidColorBrush x:Key="BorderSoft"   Color="#D0D7DE"/>
    <SolidColorBrush x:Key="Accent"       Color="#1B4F72"/>
    <SolidColorBrush x:Key="AccentHover"  Color="#2874A6"/>
    <SolidColorBrush x:Key="Success"      Color="#1E8449"/>
    <SolidColorBrush x:Key="TextPrimary"  Color="#1A1A1A"/>
    <SolidColorBrush x:Key="TextMuted"    Color="#5D6D7E"/>

    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="Background" Value="{StaticResource BgPanel}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Padding" Value="10"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
      <Setter Property="CaretBrush" Value="{StaticResource TextPrimary}"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="2"/>
    </Style>
    <Style TargetType="PasswordBox">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
      <Setter Property="CaretBrush" Value="{StaticResource TextPrimary}"/>
      <Setter Property="Padding" Value="6,4"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
      <Setter Property="Padding" Value="4"/>
      <Setter Property="Margin" Value="2"/>
    </Style>
    <Style TargetType="ListBox">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
    </Style>
    <Style TargetType="ListView">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
    </Style>
    <Style TargetType="TreeView">
      <Setter Property="Background" Value="{StaticResource BgInput}"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,7"/>
      <Setter Property="Margin" Value="4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="#E8EEF4"/>
      <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSoft}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style x:Key="PrimaryBtn" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="Margin" Value="4"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style x:Key="SuccessBtn" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource Success}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="Margin" Value="4"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="8"/>
      <Setter Property="Background" Value="#E5E8EB"/>
      <Setter Property="Foreground" Value="{StaticResource Accent}"/>
    </Style>

    <!-- Darker blue tabs with white lettering -->
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="18,10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="Bd" Background="#1B4F72" BorderBrush="#154360" BorderThickness="1,1,1,0"
                    CornerRadius="4,4,0,0" Margin="2,0,2,0" Padding="{TemplateBinding Padding}">
              <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#2874A6"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="#1B4F72"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#2874A6"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <DockPanel>
    <!-- Domain + Credential bar -->
    <Border DockPanel.Dock="Top" Background="#F0F4F8" BorderBrush="#D0D7DE" BorderThickness="0,0,0,1" Padding="14,12">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <TextBlock Text="Domain / Child Domain:" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <ComboBox x:Name="cmbDomain" Grid.Column="1" IsEditable="True" Height="30" Margin="0,0,10,0"/>
        <Button x:Name="btnCredentials" Grid.Column="2" Content="Credentials…" Margin="0,0,6,0"/>
        <Button x:Name="btnLoadDomain" Grid.Column="3" Content="Load / Connect" Style="{StaticResource PrimaryBtn}"/>
        <TextBlock x:Name="lblCurrent" Grid.Column="4" Text="Current: (none)" VerticalAlignment="Center"
                   Margin="12,0,0,0" Foreground="#1B4F72" FontWeight="SemiBold"/>

        <TextBlock x:Name="lblCredStatus" Grid.Row="1" Grid.ColumnSpan="5"
                   Text="Using current Windows credentials" Foreground="#5D6D7E" FontSize="12" Margin="0,8,0,0"/>
      </Grid>
    </Border>

    <!-- Progress + Status -->
    <Border DockPanel.Dock="Bottom" Background="#F0F4F8" BorderBrush="#D0D7DE" BorderThickness="0,1,0,0" Padding="12,8">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ProgressBar x:Name="progressBar" Grid.Row="0" Height="6" Margin="0,0,0,6" Visibility="Collapsed"/>
        <TextBlock x:Name="txtStatus" Grid.Row="1" Text="Ready – load a domain to begin"/>
      </Grid>
    </Border>

    <!-- Tabs -->
    <TabControl x:Name="tabs" Margin="10" Background="Transparent">

      <!-- CREATE USER -->
      <TabItem Header="Create User">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid Margin="4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="1.1*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <GroupBox Header="Identity" Grid.Column="0" Grid.Row="0">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="130"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Given Name *" Grid.Row="0" VerticalAlignment="Center"/>
                <TextBox x:Name="txtGiven" Grid.Column="1" Grid.Row="0"/>
                <TextBlock Text="Surname *" Grid.Row="1" VerticalAlignment="Center"/>
                <TextBox x:Name="txtSur" Grid.Column="1" Grid.Row="1"/>
                <TextBlock Text="SamAccountName *" Grid.Row="2" VerticalAlignment="Center"/>
                <TextBox x:Name="txtSam" Grid.Column="1" Grid.Row="2"/>
                <TextBlock Text="UPN *" Grid.Row="3" VerticalAlignment="Center"/>
                <TextBox x:Name="txtUPN" Grid.Column="1" Grid.Row="3"/>
                <TextBlock Text="Display Name" Grid.Row="4" VerticalAlignment="Center"/>
                <TextBox x:Name="txtDisp" Grid.Column="1" Grid.Row="4"/>
                <TextBlock Text="Description" Grid.Row="5" VerticalAlignment="Center"/>
                <TextBox x:Name="txtDesc" Grid.Column="1" Grid.Row="5"/>
              </Grid>
            </GroupBox>

            <GroupBox Header="Password &amp; Options" Grid.Column="0" Grid.Row="1">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="130"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Text="Password *" VerticalAlignment="Center"/>
                  <PasswordBox x:Name="txtPwd" Grid.Column="1" Height="30"/>
                </Grid>
                <CheckBox x:Name="chkNeverExpire" Content="Password never expires"/>
                <CheckBox x:Name="chkChangeAtLogon" Content="User must change password at next logon" IsChecked="True"/>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="Organizational Unit *" Grid.Column="1" Grid.Row="0" Grid.RowSpan="2">
              <DockPanel>
                <TextBox x:Name="txtOUFilter" DockPanel.Dock="Top" Height="28" Margin="0,0,0,8"/>
                <TextBlock DockPanel.Dock="Bottom" Text="Select an OU in the tree. Filter box searches names."
                           Foreground="#5D6D7E" Margin="0,6,0,0" TextWrapping="Wrap" FontSize="11"/>
                <TreeView x:Name="tvOU" Height="260"/>
              </DockPanel>
            </GroupBox>

            <GroupBox Header="Initial Security Group Membership (Ctrl+Click multi-select)" Grid.Column="1" Grid.Row="2">
              <ListBox x:Name="lstGroups" SelectionMode="Multiple" Height="160"
                       ScrollViewer.VerticalScrollBarVisibility="Auto"/>
            </GroupBox>

            <Button x:Name="btnCreate" Content="Create User" Style="{StaticResource PrimaryBtn}"
                    Grid.Column="0" Grid.Row="3" HorizontalAlignment="Left" Margin="6,14,6,6"/>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- MANAGE USERS -->
      <TabItem Header="Manage Users">
        <Grid Margin="6">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="170"/>
          </Grid.RowDefinitions>

          <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Search (Name / SAM / UPN):" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox x:Name="txtSearch" Width="260" Height="28"/>
            <Button x:Name="btnSearch" Content="Search" Style="{StaticResource PrimaryBtn}"/>
            <Button x:Name="btnExport" Content="Export Results"/>
            <TextBlock x:Name="lblSelCount" Text="" VerticalAlignment="Center" Margin="16,0,0,0" Foreground="#1B4F72"/>
          </StackPanel>

          <ListView x:Name="lvUsers" Grid.Row="1" SelectionMode="Extended"
                    ScrollViewer.VerticalScrollBarVisibility="Auto">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="SAM" Width="140" DisplayMemberBinding="{Binding SamAccountName}"/>
                <GridViewColumn Header="Display Name" Width="180" DisplayMemberBinding="{Binding DisplayName}"/>
                <GridViewColumn Header="Enabled" Width="70" DisplayMemberBinding="{Binding Enabled}"/>
                <GridViewColumn Header="OU" Width="280" DisplayMemberBinding="{Binding OU}"/>
                <GridViewColumn Header="UPN" Width="180" DisplayMemberBinding="{Binding UserPrincipalName}"/>
              </GridView>
            </ListView.View>
          </ListView>

          <WrapPanel Grid.Row="2" Margin="0,10,0,6">
            <Button x:Name="btnBulkEnable" Content="Enable Selected"/>
            <Button x:Name="btnBulkDisable" Content="Disable Selected"/>
            <Button x:Name="btnBulkResetPwd" Content="Reset Password"/>
            <Button x:Name="btnBulkMoveOU" Content="Change OU"/>
            <Button x:Name="btnBulkAddGroup" Content="Add Group(s)"/>
            <Button x:Name="btnRefreshMembers" Content="Refresh Member Of"/>
          </WrapPanel>

          <GroupBox Header="Member Of (first selected user)" Grid.Row="3">
            <ListBox x:Name="lstMemberOf" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
          </GroupBox>
        </Grid>
      </TabItem>

      <!-- DUPLICATE -->
      <TabItem Header="Duplicate User (Template)">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="6">
            <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
              <TextBlock Text="Source user (SAM or UPN):" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBox x:Name="txtSrc" Width="260" Height="28"/>
              <Button x:Name="btnLoadSrc" Content="Load Template" Style="{StaticResource PrimaryBtn}"/>
            </StackPanel>
            <TextBlock x:Name="lblSrcInfo" Text="No template loaded." Foreground="#1E8449" Margin="0,0,0,12" TextWrapping="Wrap"/>

            <GroupBox Header="New User Details">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Column="0" Grid.Row="0" Margin="0,4">
                  <TextBlock Text="Given Name *"/>
                  <TextBox x:Name="txtDGiven"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Grid.Row="0" Margin="8,4,0,0">
                  <TextBlock Text="Surname *"/>
                  <TextBox x:Name="txtDSur"/>
                </StackPanel>
                <StackPanel Grid.Column="0" Grid.Row="1" Margin="0,4">
                  <TextBlock Text="SamAccountName *"/>
                  <TextBox x:Name="txtDSam"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Grid.Row="1" Margin="8,4,0,0">
                  <TextBlock Text="UPN *"/>
                  <TextBox x:Name="txtDUPN"/>
                </StackPanel>
                <StackPanel Grid.Column="0" Grid.Row="2" Margin="0,4">
                  <TextBlock Text="Display Name"/>
                  <TextBox x:Name="txtDDisp"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Grid.Row="2" Margin="8,4,0,0">
                  <TextBlock Text="Password *"/>
                  <PasswordBox x:Name="txtDPwd" Height="30"/>
                </StackPanel>
                <StackPanel Grid.Column="0" Grid.Row="3" Grid.ColumnSpan="2" Margin="0,8,0,0">
                  <TextBlock Text="Target OU *  (filter + select in tree)"/>
                  <TextBox x:Name="txtDOUFilter" Height="28" Margin="0,4,0,6"/>
                  <TreeView x:Name="tvDOU" Height="160"/>
                </StackPanel>
                <StackPanel Grid.Column="0" Grid.Row="4" Margin="0,8,0,0">
                  <CheckBox x:Name="chkDNever" Content="Password never expires"/>
                  <CheckBox x:Name="chkDChange" Content="Must change at next logon" IsChecked="True"/>
                </StackPanel>
                <StackPanel Grid.Column="0" Grid.Row="5" Grid.ColumnSpan="2" Margin="0,10,0,0">
                  <TextBlock Text="Groups that will be copied from template:"/>
                  <ListBox x:Name="lstDGroups" Height="80" Margin="0,4,0,0"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <Button x:Name="btnDuplicate" Content="Create from Template" Style="{StaticResource SuccessBtn}"
                    HorizontalAlignment="Left" Margin="0,14,0,0" IsEnabled="False"/>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
    </TabControl>
  </DockPanel>
</Window>
"@

#endregion

#region ===== Load Window & Controls =====

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

function Get-Control($name) { $window.FindName($name) }

$cmbDomain        = Get-Control 'cmbDomain'
$btnCredentials   = Get-Control 'btnCredentials'
$btnLoadDomain    = Get-Control 'btnLoadDomain'
$lblCurrent       = Get-Control 'lblCurrent'
$lblCredStatus    = Get-Control 'lblCredStatus'
$txtStatus        = Get-Control 'txtStatus'
$progressBar      = Get-Control 'progressBar'

$txtGiven         = Get-Control 'txtGiven'
$txtSur           = Get-Control 'txtSur'
$txtSam           = Get-Control 'txtSam'
$txtUPN           = Get-Control 'txtUPN'
$txtDisp          = Get-Control 'txtDisp'
$txtDesc          = Get-Control 'txtDesc'
$txtPwd           = Get-Control 'txtPwd'
$chkNeverExpire   = Get-Control 'chkNeverExpire'
$chkChangeAtLogon = Get-Control 'chkChangeAtLogon'
$txtOUFilter      = Get-Control 'txtOUFilter'
$tvOU             = Get-Control 'tvOU'
$lstGroups        = Get-Control 'lstGroups'
$btnCreate        = Get-Control 'btnCreate'

$txtSearch        = Get-Control 'txtSearch'
$btnSearch        = Get-Control 'btnSearch'
$btnExport        = Get-Control 'btnExport'
$lblSelCount      = Get-Control 'lblSelCount'
$lvUsers          = Get-Control 'lvUsers'
$btnBulkEnable    = Get-Control 'btnBulkEnable'
$btnBulkDisable   = Get-Control 'btnBulkDisable'
$btnBulkResetPwd  = Get-Control 'btnBulkResetPwd'
$btnBulkMoveOU    = Get-Control 'btnBulkMoveOU'
$btnBulkAddGroup  = Get-Control 'btnBulkAddGroup'
$btnRefreshMembers= Get-Control 'btnRefreshMembers'
$lstMemberOf      = Get-Control 'lstMemberOf'

$txtSrc           = Get-Control 'txtSrc'
$btnLoadSrc       = Get-Control 'btnLoadSrc'
$lblSrcInfo       = Get-Control 'lblSrcInfo'
$txtDGiven        = Get-Control 'txtDGiven'
$txtDSur          = Get-Control 'txtDSur'
$txtDSam          = Get-Control 'txtDSam'
$txtDUPN          = Get-Control 'txtDUPN'
$txtDDisp         = Get-Control 'txtDDisp'
$txtDPwd          = Get-Control 'txtDPwd'
$txtDOUFilter     = Get-Control 'txtDOUFilter'
$tvDOU            = Get-Control 'tvDOU'
$chkDNever        = Get-Control 'chkDNever'
$chkDChange       = Get-Control 'chkDChange'
$lstDGroups       = Get-Control 'lstDGroups'
$btnDuplicate     = Get-Control 'btnDuplicate'

#endregion

#region ===== Shared State =====
$script:CurrentDomain = $null
$script:CurrentServer = $null
$script:Credential    = $null
$script:OUList        = @()
$script:GroupList     = @()
$script:TemplateUser  = $null
$script:TemplateGroups= @()
$script:OURoot        = $null
$script:DOURoot       = $null
$script:SelectedOUDN  = $null
$script:SelectedDOUDN = $null
#endregion

#region ===== Progress helpers =====
function Show-Progress {
    param([int]$Value = 0, [int]$Maximum = 100, [string]$Status = '')
    $progressBar.Visibility = "Visible"
    $progressBar.Maximum = $Maximum
    $progressBar.Value = $Value
    if ($Status) { $txtStatus.Text = $Status }
    $window.Dispatcher.Invoke([Action]{}, "Background")
}
function Hide-Progress { $progressBar.Visibility = "Collapsed"; $progressBar.Value = 0 }
#endregion

#region ===== Event Handlers =====

# Auto-fill
$txtGiven.Add_TextChanged({
    if ($txtSur.Text -and $txtGiven.Text) {
        $txtDisp.Text = "$($txtGiven.Text) $($txtSur.Text)"
        if (-not $txtSam.Tag) {
            $txtSam.Text = ("{0}.{1}" -f $txtGiven.Text.Substring(0,[Math]::Min(1,$txtGiven.Text.Length)), $txtSur.Text).ToLower() -replace '[^a-z0-9.]',''
        }
    }
})
$txtSur.Add_TextChanged({
    if ($txtGiven.Text -and $txtSur.Text) {
        $txtDisp.Text = "$($txtGiven.Text) $($txtSur.Text)"
        if (-not $txtSam.Tag) {
            $txtSam.Text = ("{0}.{1}" -f $txtGiven.Text.Substring(0,[Math]::Min(1,$txtGiven.Text.Length)), $txtSur.Text).ToLower() -replace '[^a-z0-9.]',''
        }
    }
})
$txtSam.Add_TextChanged({ $txtSam.Tag = $true })

# OU filter + selection
$txtOUFilter.Add_TextChanged({
    if ($script:OURoot) { Filter-OUTree -TreeView $tvOU -Filter $txtOUFilter.Text -AllRoot $script:OURoot }
})
$txtDOUFilter.Add_TextChanged({
    if ($script:DOURoot) { Filter-OUTree -TreeView $tvDOU -Filter $txtDOUFilter.Text -AllRoot $script:DOURoot }
})
$tvOU.Add_SelectedItemChanged({
    if ($tvOU.SelectedItem -and $tvOU.SelectedItem.Tag) { $script:SelectedOUDN = $tvOU.SelectedItem.Tag }
})
$tvDOU.Add_SelectedItemChanged({
    if ($tvDOU.SelectedItem -and $tvDOU.SelectedItem.Tag) { $script:SelectedDOUDN = $tvDOU.SelectedItem.Tag }
})

# Credentials button
$btnCredentials.Add_Click({
    $cred = Get-Credential -Message "Enter credentials for the target domain (leave blank / Cancel to use current Windows credentials)"
    if ($cred) {
        $script:Credential = $cred
        $lblCredStatus.Text = "Using credentials: $($cred.UserName)"
        $lblCredStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
    } else {
        $script:Credential = $null
        $lblCredStatus.Text = "Using current Windows credentials"
        $lblCredStatus.Foreground = [System.Windows.Media.Brushes]::Gray
    }
})

# Domain load (with optional credential)
$window.Add_Loaded({
    try {
        $forestParams = @{}
        if ($script:Credential) { $forestParams.Credential = $script:Credential }
        $forest = Get-ADForest @forestParams
        $cmbDomain.Items.Clear()
        foreach ($d in $forest.Domains) { [void]$cmbDomain.Items.Add($d) }
        $curr = (Get-ADDomain).DNSRoot
        if (-not $cmbDomain.Items.Contains($curr)) { [void]$cmbDomain.Items.Add($curr) }
        $cmbDomain.Text = $curr
        $txtStatus.Text = "Forest domains loaded. Select domain and click Load / Connect."
    } catch {
        $txtStatus.Text = "Could not enumerate forest. Type a domain FQDN manually and supply credentials if needed."
    }
})

$btnLoadDomain.Add_Click({
    $domainFqdn = $cmbDomain.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domainFqdn)) {
        [System.Windows.MessageBox]::Show("Enter or select a domain FQDN.", "Input required")
        return
    }
    $txtStatus.Text = "Connecting to $domainFqdn ..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $script:CurrentDomain = Get-ADDomainContext -DomainFQDN $domainFqdn -Credential $script:Credential
        $script:CurrentServer = $null

        $script:OUList = @(Get-OUTree -SearchBase $script:CurrentDomain.DistinguishedName -Server $script:CurrentServer -Credential $script:Credential)
        $script:OURoot  = Build-OUTreeViewItems -OUs $script:OUList -DomainDN $script:CurrentDomain.DistinguishedName
        $script:DOURoot = Build-OUTreeViewItems -OUs $script:OUList -DomainDN $script:CurrentDomain.DistinguishedName

        $tvOU.Items.Clear();  [void]$tvOU.Items.Add($script:OURoot)
        $tvDOU.Items.Clear(); [void]$tvDOU.Items.Add($script:DOURoot)
        $script:SelectedOUDN  = $script:CurrentDomain.DistinguishedName
        $script:SelectedDOUDN = $script:CurrentDomain.DistinguishedName

        $script:GroupList = @(Get-SecurityGroups -SearchBase $script:CurrentDomain.DistinguishedName -Server $script:CurrentServer -Credential $script:Credential)
        $lstGroups.Items.Clear()
        foreach ($g in $script:GroupList) { [void]$lstGroups.Items.Add($g.Name) }

        $lblCurrent.Text = "Current: $($script:CurrentDomain.DNSRoot)"
        $txtStatus.Text = "Connected to $($script:CurrentDomain.DNSRoot) — $($script:OUList.Count) OUs, $($script:GroupList.Count) security groups."
        Write-ADLog -Action "ConnectDomain" -Target $domainFqdn -Result "Success" -Details $(if ($script:Credential) { "Cred=$($script:Credential.UserName)" } else { "Current user" })
    } catch {
        $txtStatus.Text = "Connection failed."
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Domain connection error")
        Write-ADLog -Action "ConnectDomain" -Target $domainFqdn -Result "Failed" -Details $_.Exception.Message
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Create User
$btnCreate.Add_Click({
    if (-not $script:CurrentDomain) { [System.Windows.MessageBox]::Show("Load a domain first."); return }
    if ([string]::IsNullOrWhiteSpace($txtGiven.Text) -or [string]::IsNullOrWhiteSpace($txtSur.Text) -or
        [string]::IsNullOrWhiteSpace($txtSam.Text) -or [string]::IsNullOrWhiteSpace($txtUPN.Text) -or
        [string]::IsNullOrWhiteSpace($txtPwd.Password) -or -not $script:SelectedOUDN) {
        [System.Windows.MessageBox]::Show("Fill all required fields and select an OU in the tree.")
        return
    }

    $selectedGroups = @()
    foreach ($name in $lstGroups.SelectedItems) {
        $g = $script:GroupList | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($g) { $selectedGroups += $g.DistinguishedName }
    }

    $userData = @{
        GivenName             = $txtGiven.Text.Trim()
        Surname               = $txtSur.Text.Trim()
        SamAccountName        = $txtSam.Text.Trim()
        UPN                   = $txtUPN.Text.Trim()
        DisplayName           = if ($txtDisp.Text) { $txtDisp.Text.Trim() } else { "$($txtGiven.Text) $($txtSur.Text)" }
        Description           = $txtDesc.Text.Trim()
        OU                    = $script:SelectedOUDN
        Password              = $txtPwd.Password
        PasswordNeverExpires  = [bool]$chkNeverExpire.IsChecked
        ChangePasswordAtLogon = [bool]$chkChangeAtLogon.IsChecked
        Groups                = $selectedGroups
    }

    $r = [System.Windows.MessageBox]::Show("Create user '$($userData.SamAccountName)' in:`n$($userData.OU)`n`nContinue?", "Confirm Create", "YesNo", "Question")
    if ($r -ne "Yes") { return }

    $txtStatus.Text = "Creating user..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $newUser = New-ADUserFromForm -UserData $userData -Server $script:CurrentServer -Credential $script:Credential
        $txtStatus.Text = "Created: $($newUser.SamAccountName)"
        [System.Windows.MessageBox]::Show("User '$($newUser.SamAccountName)' created successfully.", "Success")
        Write-ADLog -Action "CreateUser" -Target $newUser.SamAccountName -Result "Success" -Details "OU=$($userData.OU)"
        $txtPwd.Password = ""
    } catch {
        $txtStatus.Text = "Create failed."
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Create User Error")
        Write-ADLog -Action "CreateUser" -Target $userData.SamAccountName -Result "Failed" -Details $_.Exception.Message
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Search
$btnSearch.Add_Click({
    if (-not $script:CurrentDomain) { [System.Windows.MessageBox]::Show("Load a domain first."); return }
    $term = $txtSearch.Text.Trim()
    if (-not $term) { [System.Windows.MessageBox]::Show("Enter a search term."); return }

    $txtStatus.Text = "Searching..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $lvUsers.Items.Clear()
    $lstMemberOf.Items.Clear()
    try {
        $filter = "Name -like '*$term*' -or SamAccountName -like '*$term*' -or UserPrincipalName -like '*$term*'"
        $p = @{
            Filter     = $filter
            SearchBase = $script:CurrentDomain.DistinguishedName
            Properties = 'Enabled','DistinguishedName','UserPrincipalName','DisplayName'
        }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        $users = Get-ADUser @p | Select-Object -First 400

        foreach ($u in $users) {
            $ou = ($u.DistinguishedName -split ',', 2)[1]
            $item = [PSCustomObject]@{
                SamAccountName    = $u.SamAccountName
                DisplayName       = $u.DisplayName
                Enabled           = $u.Enabled
                OU                = $ou
                UserPrincipalName = $u.UserPrincipalName
                DN                = $u.DistinguishedName
                ADUser            = $u
            }
            [void]$lvUsers.Items.Add($item)
        }
        $txtStatus.Text = "Found $($users.Count) user(s)."
        $lblSelCount.Text = ""
    } catch {
        $txtStatus.Text = "Search failed."
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Search Error")
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

# Export
$btnExport.Add_Click({
    if ($lvUsers.Items.Count -eq 0) { [System.Windows.MessageBox]::Show("No results to export."); return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv"
    $dlg.FileName = "ADUsers_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($dlg.ShowDialog()) {
        $lvUsers.Items | Select-Object SamAccountName, DisplayName, Enabled, OU, UserPrincipalName |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $txtStatus.Text = "Exported $($lvUsers.Items.Count) rows to $($dlg.FileName)"
        [System.Windows.MessageBox]::Show("Exported successfully.", "Export")
    }
})

$lvUsers.Add_SelectionChanged({
    $count = $lvUsers.SelectedItems.Count
    $lblSelCount.Text = if ($count -gt 0) { "$count selected" } else { "" }
    $lstMemberOf.Items.Clear()
    if ($count -eq 0) { return }
    $first = $lvUsers.SelectedItems[0]
    try {
        $p = @{ Identity = $first.DN; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        $groups = Get-ADPrincipalGroupMembership @p | Sort-Object Name
        foreach ($g in $groups) { [void]$lstMemberOf.Items.Add($g.Name) }
    } catch {
        [void]$lstMemberOf.Items.Add("(error loading groups)")
    }
})

# Bulk action helper
function Invoke-BulkAction {
    param([array]$Items, [string]$ConfirmMsg, [string]$ActionName, [scriptblock]$Action)
    if ($Items.Count -eq 0) { return }
    $r = [System.Windows.MessageBox]::Show($ConfirmMsg, "Confirm", "YesNo", "Question")
    if ($r -ne "Yes") { return }

    $total = $Items.Count
    $ok = 0
    Show-Progress -Value 0 -Maximum $total -Status "$ActionName 0 / $total ..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $i = 0
    foreach ($item in $Items) {
        $i++
        try {
            & $Action $item
            $ok++
            Write-ADLog -Action $ActionName -Target $item.SamAccountName -Result "Success"
        } catch {
            Write-ADLog -Action $ActionName -Target $item.SamAccountName -Result "Failed" -Details $_.Exception.Message
        }
        Show-Progress -Value $i -Maximum $total -Status "$ActionName $i / $total ..."
    }
    Hide-Progress
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $txtStatus.Text = "$ActionName completed: $ok of $total succeeded."
    $lvUsers.Items.Refresh()
}

$btnBulkEnable.Add_Click({
    Invoke-BulkAction -Items @($lvUsers.SelectedItems) -ConfirmMsg "Enable $($lvUsers.SelectedItems.Count) account(s)?" -ActionName "EnableAccount" -Action {
        param($item)
        $p = @{ Identity = $item.SamAccountName; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        Enable-ADAccount @p
        $item.Enabled = $true
    }
})

$btnBulkDisable.Add_Click({
    Invoke-BulkAction -Items @($lvUsers.SelectedItems) -ConfirmMsg "Disable $($lvUsers.SelectedItems.Count) account(s)?" -ActionName "DisableAccount" -Action {
        param($item)
        $p = @{ Identity = $item.SamAccountName; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        Disable-ADAccount @p
        $item.Enabled = $false
    }
})

$btnBulkResetPwd.Add_Click({
    $sel = @($lvUsers.SelectedItems)
    if ($sel.Count -eq 0) { return }

    $pwdWin = New-Object System.Windows.Window
    $pwdWin.Title = "Reset Password – $($sel.Count) user(s)"
    $pwdWin.Width = 400; $pwdWin.Height = 210
    $pwdWin.WindowStartupLocation = "CenterOwner"
    $pwdWin.Background = "White"
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = "16"
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = "New password for all selected users:"
    $tb.Margin = "0,0,0,8"
    $pb = New-Object System.Windows.Controls.PasswordBox
    $pb.Height = 30
    $chk = New-Object System.Windows.Controls.CheckBox
    $chk.Content = "User must change password at next logon"
    $chk.IsChecked = $true
    $chk.Margin = "0,10,0,0"
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Right"
    $btnPanel.Margin = "0,16,0,0"
    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "Reset"; $btnOk.Width = 90; $btnOk.Margin = "0,0,8,0"
    $btnOk.Background = "#1B4F72"; $btnOk.Foreground = "White"
    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"; $btnCancel.Width = 90
    $btnPanel.Children.Add($btnOk); $btnPanel.Children.Add($btnCancel)
    $sp.Children.Add($tb); $sp.Children.Add($pb); $sp.Children.Add($chk); $sp.Children.Add($btnPanel)
    $pwdWin.Content = $sp

    $script:PwdResult = $null
    $btnOk.Add_Click({
        if ($pb.Password) {
            $script:PwdResult = @{ Password = $pb.Password; MustChange = [bool]$chk.IsChecked }
            $pwdWin.DialogResult = $true
            $pwdWin.Close()
        }
    })
    $btnCancel.Add_Click({ $pwdWin.Close() })
    $pwdWin.Owner = $window
    $null = $pwdWin.ShowDialog()
    if (-not $script:PwdResult) { return }

    $sec = ConvertTo-SecureString $script:PwdResult.Password -AsPlainText -Force
    Invoke-BulkAction -Items $sel -ConfirmMsg "Apply the new password to $($sel.Count) account(s)?" -ActionName "ResetPassword" -Action {
        param($item)
        $p = @{ Identity = $item.SamAccountName; NewPassword = $sec; Reset = $true; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        Set-ADAccountPassword @p
        if ($script:PwdResult.MustChange) {
            $su = @{ Identity = $item.SamAccountName; ChangePasswordAtLogon = $true; ErrorAction = 'SilentlyContinue' }
            if ($script:CurrentServer) { $su.Server = $script:CurrentServer }
            if ($script:Credential)    { $su.Credential = $script:Credential }
            Set-ADUser @su
        }
    }
})

$btnBulkMoveOU.Add_Click({
    $sel = @($lvUsers.SelectedItems)
    if ($sel.Count -eq 0) { return }

    $moveWin = New-Object System.Windows.Window
    $moveWin.Title = "Move $($sel.Count) user(s) to OU"
    $moveWin.Width = 480; $moveWin.Height = 420
    $moveWin.WindowStartupLocation = "CenterOwner"
    $moveWin.Background = "White"
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = "14"
    $filterBox = New-Object System.Windows.Controls.TextBox
    $filterBox.Height = 28
    $filterBox.Margin = "0,0,0,8"
    $tv = New-Object System.Windows.Controls.TreeView
    $tv.Height = 280
    if ($script:OURoot) {
        $rootCopy = Build-OUTreeViewItems -OUs $script:OUList -DomainDN $script:CurrentDomain.DistinguishedName
        [void]$tv.Items.Add($rootCopy)
    }
    $filterBox.Add_TextChanged({
        if ($script:OURoot) { Filter-OUTree -TreeView $tv -Filter $filterBox.Text -AllRoot $script:OURoot }
    })
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Right"
    $btnPanel.Margin = "0,12,0,0"
    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "Move"; $btnOk.Width = 90; $btnOk.Margin = "0,0,8,0"
    $btnOk.Background = "#1B4F72"; $btnOk.Foreground = "White"
    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"; $btnCancel.Width = 90
    $btnPanel.Children.Add($btnOk); $btnPanel.Children.Add($btnCancel)
    $sp.Children.Add($filterBox); $sp.Children.Add($tv); $sp.Children.Add($btnPanel)
    $moveWin.Content = $sp

    $script:MoveDN = $null
    $btnOk.Add_Click({
        if ($tv.SelectedItem -and $tv.SelectedItem.Tag) {
            $script:MoveDN = $tv.SelectedItem.Tag
            $moveWin.DialogResult = $true
            $moveWin.Close()
        } else {
            [System.Windows.MessageBox]::Show("Select an OU in the tree.")
        }
    })
    $btnCancel.Add_Click({ $moveWin.Close() })
    $moveWin.Owner = $window
    $null = $moveWin.ShowDialog()
    if (-not $script:MoveDN) { return }

    Invoke-BulkAction -Items $sel -ConfirmMsg "Move $($sel.Count) user(s) to the selected OU?" -ActionName "MoveOU" -Action {
        param($item)
        $p = @{ Identity = $item.DN; TargetPath = $script:MoveDN; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        Move-ADObject @p
    }
    $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$btnBulkAddGroup.Add_Click({
    $sel = @($lvUsers.SelectedItems)
    if ($sel.Count -eq 0) { return }

    $grpWin = New-Object System.Windows.Window
    $grpWin.Title = "Add groups to $($sel.Count) user(s)"
    $grpWin.Width = 420; $grpWin.Height = 440
    $grpWin.WindowStartupLocation = "CenterOwner"
    $grpWin.Background = "White"
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = "12"
    $lb = New-Object System.Windows.Controls.ListBox
    $lb.SelectionMode = "Multiple"
    $lb.Height = 310
    foreach ($g in $script:GroupList) { [void]$lb.Items.Add($g.Name) }
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Right"
    $btnPanel.Margin = "0,12,0,0"
    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "Add Selected"; $btnOk.Padding = "12,7"; $btnOk.Margin = "0,0,8,0"
    $btnOk.Background = "#1B4F72"; $btnOk.Foreground = "White"
    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"; $btnCancel.Padding = "12,7"
    $btnPanel.Children.Add($btnOk); $btnPanel.Children.Add($btnCancel)
    $sp.Children.Add($lb); $sp.Children.Add($btnPanel)
    $grpWin.Content = $sp

    $script:SelectedGroupNames = @()
    $btnOk.Add_Click({
        $script:SelectedGroupNames = @($lb.SelectedItems)
        $grpWin.DialogResult = $true
        $grpWin.Close()
    })
    $btnCancel.Add_Click({ $grpWin.Close() })
    $grpWin.Owner = $window
    $null = $grpWin.ShowDialog()
    if ($script:SelectedGroupNames.Count -eq 0) { return }

    $totalOps = $sel.Count * $script:SelectedGroupNames.Count
    $done = 0
    Show-Progress -Value 0 -Maximum $totalOps -Status "Adding groups..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    foreach ($item in $sel) {
        foreach ($gname in $script:SelectedGroupNames) {
            $g = $script:GroupList | Where-Object { $_.Name -eq $gname } | Select-Object -First 1
            if ($g) {
                try {
                    $p = @{ Identity = $g; Members = $item.SamAccountName; ErrorAction = 'Stop' }
                    if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
                    if ($script:Credential)    { $p.Credential = $script:Credential }
                    Add-ADGroupMember @p
                    Write-ADLog -Action "AddGroup" -Target $item.SamAccountName -Result "Success" -Details $gname
                } catch {
                    Write-ADLog -Action "AddGroup" -Target $item.SamAccountName -Result "Failed" -Details $_.Exception.Message
                }
            }
            $done++
            Show-Progress -Value $done -Maximum $totalOps
        }
    }
    Hide-Progress
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $txtStatus.Text = "Group membership updates finished."
})

$btnRefreshMembers.Add_Click({
    if ($lvUsers.SelectedItems.Count -gt 0) {
        $lvUsers.RaiseEvent((New-Object System.Windows.Controls.SelectionChangedEventArgs(
            [System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
            [System.Collections.ArrayList]@(),
            [System.Collections.ArrayList]@($lvUsers.SelectedItems)
        )))
    }
})

# Duplicate
$btnLoadSrc.Add_Click({
    if (-not $script:CurrentDomain) { [System.Windows.MessageBox]::Show("Load a domain first."); return }
    $src = $txtSrc.Text.Trim()
    if (-not $src) { return }
    $txtStatus.Text = "Loading template..."
    try {
        $p = @{ Identity = $src; Properties = '*'; ErrorAction = 'Stop' }
        if ($script:CurrentServer) { $p.Server = $script:CurrentServer }
        if ($script:Credential)    { $p.Credential = $script:Credential }
        $script:TemplateUser = Get-ADUser @p

        $mp = @{ Identity = $script:TemplateUser.DistinguishedName; ErrorAction = 'SilentlyContinue' }
        if ($script:CurrentServer) { $mp.Server = $script:CurrentServer }
        if ($script:Credential)    { $mp.Credential = $script:Credential }
        $script:TemplateGroups = @(Get-ADPrincipalGroupMembership @mp |
            Where-Object { $_.SamAccountName -ne 'Domain Users' } | Sort-Object Name)

        $lblSrcInfo.Text = "Template: $($script:TemplateUser.DisplayName) ($($script:TemplateUser.SamAccountName)) — $($script:TemplateGroups.Count) groups will be copied."
        $lstDGroups.Items.Clear()
        foreach ($g in $script:TemplateGroups) { [void]$lstDGroups.Items.Add($g.Name) }
        $btnDuplicate.IsEnabled = $true
        $txtStatus.Text = "Template loaded."
    } catch {
        $lblSrcInfo.Text = "Failed to load template."
        $btnDuplicate.IsEnabled = $false
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Load Template Error")
    }
})

$btnDuplicate.Add_Click({
    if (-not $script:TemplateUser) { return }
    if ([string]::IsNullOrWhiteSpace($txtDGiven.Text) -or [string]::IsNullOrWhiteSpace($txtDSur.Text) -or
        [string]::IsNullOrWhiteSpace($txtDSam.Text) -or [string]::IsNullOrWhiteSpace($txtDUPN.Text) -or
        [string]::IsNullOrWhiteSpace($txtDPwd.Password) -or -not $script:SelectedDOUDN) {
        [System.Windows.MessageBox]::Show("Fill all required fields and select a target OU in the tree.")
        return
    }

    $newData = @{
        GivenName             = $txtDGiven.Text.Trim()
        Surname               = $txtDSur.Text.Trim()
        SamAccountName        = $txtDSam.Text.Trim()
        UPN                   = $txtDUPN.Text.Trim()
        DisplayName           = if ($txtDDisp.Text) { $txtDDisp.Text.Trim() } else { "$($txtDGiven.Text) $($txtDSur.Text)" }
        Description           = $null
        OU                    = $script:SelectedDOUDN
        Password              = $txtDPwd.Password
        PasswordNeverExpires  = [bool]$chkDNever.IsChecked
        ChangePasswordAtLogon = [bool]$chkDChange.IsChecked
    }

    $r = [System.Windows.MessageBox]::Show(
        "Create '$($newData.SamAccountName)' by copying from '$($script:TemplateUser.SamAccountName)' (including $($script:TemplateGroups.Count) groups)?",
        "Confirm Duplicate", "YesNo", "Question")
    if ($r -ne "Yes") { return }

    $txtStatus.Text = "Creating from template..."
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $newUser = Copy-ADUserAsTemplate -SourceSam $script:TemplateUser.SamAccountName -NewUserData $newData -Server $script:CurrentServer -Credential $script:Credential
        $txtStatus.Text = "Created from template: $($newUser.SamAccountName)"
        [System.Windows.MessageBox]::Show("User '$($newUser.SamAccountName)' created successfully from template.", "Success")
        Write-ADLog -Action "DuplicateUser" -Target $newUser.SamAccountName -Result "Success" -Details "Source=$($script:TemplateUser.SamAccountName)"
        $txtDPwd.Password = ""
    } catch {
        $txtStatus.Text = "Duplicate failed."
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Duplicate Error")
        Write-ADLog -Action "DuplicateUser" -Target $newData.SamAccountName -Result "Failed" -Details $_.Exception.Message
    } finally {
        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
})

#endregion

[void]$window.ShowDialog()
