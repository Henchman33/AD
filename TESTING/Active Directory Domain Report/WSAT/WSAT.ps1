<#
================================================================================
    WindowsServerAudit - Windows Server Audit Tool (Redesigned)
    Version: 0.4.0
    
    Description:
    Comprehensive server audit tool with modern WPF GUI.
    Provides detailed analysis of system configuration, security, services,
    roles, and replication with export functions (CSV, HTML).
    
    Design: Sidebar menu (left) with audit categories + output area (right)
    
================================================================================
#>

[xml]$Global:XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="WindowsServerAudit v0.4.0 - Windows Server Audit Tool"
    Width="1850"
    Height="1000"
    Background="#F8FAFC"
    FontFamily="Segoe UI"
    ResizeMode="CanResizeWithGrip"
    WindowStartupLocation="CenterScreen">

    <!--  Window Resources for Modern Styling  -->
    <Window.Resources>
        <!--  Modern Card Style  -->
        <Style x:Key="ModernCard" TargetType="Border">
            <Setter Property="Background" Value="White" />
            <Setter Property="CornerRadius" Value="6" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#E5E7EB" />
        </Style>

        <!--  Primary Button Style  -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#6366F1" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="24,14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="FontSize" Value="14" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            CornerRadius="4">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5B5BD6" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#4F46E5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Category Header Style  -->
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Foreground" Value="#1F2937" />
            <Setter Property="Margin" Value="0,8,0,4" />
        </Style>

        <!--  Sidebar Menu Button Style  -->
        <Style x:Key="SidebarButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="12,8" />
            <Setter Property="FontWeight" Value="Normal" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="HorizontalAlignment" Value="Stretch" />
            <Setter Property="HorizontalContentAlignment" Value="Left" />
            <Setter Property="Margin" Value="0,1" />
            <Setter Property="Foreground" Value="#374151" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Padding="{TemplateBinding Padding}"
                            Background="{TemplateBinding Background}"
                            CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F3F4F6" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E5E7EB" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!--  Active Sidebar Button Style  -->
        <Style
            x:Key="SidebarButtonActive"
            BasedOn="{StaticResource SidebarButton}"
            TargetType="Button">
            <Setter Property="Background" Value="#EEF2FF" />
            <Setter Property="Foreground" Value="#6366F1" />
            <Setter Property="FontWeight" Value="Medium" />
        </Style>

        <!--  Expandable Section Style  -->
        <Style x:Key="ExpanderStyle" TargetType="Expander">
            <Setter Property="IsExpanded" Value="True" />
            <Setter Property="Margin" Value="0,2,0,6" />
            <Setter Property="Background" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="#E5E7EB" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="Foreground" Value="#374151" />
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="50" />
            <!--  Header  -->
            <RowDefinition Height="*" />
            <!--  Content Area  -->
            <RowDefinition Height="25" />
            <!--  Footer  -->
        </Grid.RowDefinitions>

        <!--  Header (Grid.Row="0")  -->
        <Border
            Grid.Row="0"
            Background="#374151"
            BorderBrush="#E5E7EB"
            BorderThickness="0,0,0,1">
            <Grid Margin="20,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <!--  App Title  -->
                <StackPanel
                    Grid.Column="0"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <TextBlock
                        FontSize="20"
                        FontWeight="SemiBold"
                        Foreground="#F9FAFB"
                        Text="⚙️ WindowsServerAudit v0.4.0 - Windows Server Audit" />
                </StackPanel>

                <!--  Status Bar  -->
                <StackPanel
                    Grid.Column="1"
                    VerticalAlignment="Center"
                    Orientation="Vertical"
                    Background="#FF4C5B73"
                    Height="49"
                    Margin="1475,0,-20,0"
                    Grid.ColumnSpan="2">
                    <TextBlock Text="Results" FontSize="14" Foreground="#FFC1FFDB" HorizontalAlignment="Center"/>
                    <TextBlock x:Name="TotalResultCountText" Text="0" FontSize="20" FontWeight="Bold" 
                               Foreground="#FFEAEAEA" HorizontalAlignment="Center" Margin="0,2,0,0" Width="60" TextAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>

        <!--  Content Area (Grid.Row="1")  -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320" />
                <!--  Sidebar  -->
                <ColumnDefinition Width="24" />
                <!--  Spacing  -->
                <ColumnDefinition Width="*" />
                <!--  Main Content  -->
            </Grid.ColumnDefinitions>

            <!--  Modern Sidebar with Categorized Reports  -->
            <ScrollViewer
                Grid.Column="0"
                Margin="0,0,0,10"
                HorizontalScrollBarVisibility="Disabled"
                VerticalScrollBarVisibility="Auto">
                <Border
                    Padding="16,16"
                    Background="#F8FAFC"
                    BorderBrush="#E5E7EB"
                    BorderThickness="1"
                    CornerRadius="6">
                    <StackPanel>
                        <!--  Sidebar Header  -->
                        <Grid Margin="0,0,0,16">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto" />
                                <RowDefinition Height="Auto" />
                            </Grid.RowDefinitions>
                            <TextBlock
                                Grid.Row="0"
                                Margin="0,0,0,8"
                                FontSize="16"
                                FontWeight="Bold"
                                Foreground="#374151"
                                Text="📊 Audit Reports" />
                        </Grid>

                        <!--  System Information  -->
                        <Expander
                            Header="🖥️ System Information"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonSystemInfo" Content="System Overview" Style="{StaticResource SidebarButton}" ToolTip="Retrieve system information" />
                                <Button x:Name="ButtonOSInfo" Content="OS Details" Style="{StaticResource SidebarButton}" ToolTip="Get operating system details" />
                                <Button x:Name="ButtonHardwareInfo" Content="Hardware Summary" Style="{StaticResource SidebarButton}" ToolTip="Display hardware information" />
                                <Button x:Name="ButtonCPUInfo" Content="CPU Details" Style="{StaticResource SidebarButton}" ToolTip="CPU specifications" />
                                <Button x:Name="ButtonMemoryInfo" Content="Memory Details" Style="{StaticResource SidebarButton}" ToolTip="Memory information" />
                                <Button x:Name="ButtonStorageInfo" Content="Storage Summary" Style="{StaticResource SidebarButton}" ToolTip="Disk and volume information" />
                            </StackPanel>
                        </Expander>

                        <!--  Network Configuration  -->
                        <Expander
                            Header="🌐 Network Configuration"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonNetConfig" Content="IP Configuration" Style="{StaticResource SidebarButton}" ToolTip="Network IP settings" />
                                <Button x:Name="ButtonNetAdapters" Content="Network Adapters" Style="{StaticResource SidebarButton}" ToolTip="Network adapter status" />
                                <Button x:Name="ButtonTCPConnections" Content="Active Connections" Style="{StaticResource SidebarButton}" ToolTip="Listen ports and connections" />
                                <Button x:Name="ButtonFirewallRules" Content="Firewall Rules" Style="{StaticResource SidebarButton}" ToolTip="Active firewall rules" />
                            </StackPanel>
                        </Expander>

                        <!--  Services & Tasks  -->
                        <Expander
                            Header="⚙️ Services &amp; Tasks"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonAutomaticServices" Content="Automatic Services" Style="{StaticResource SidebarButton}" ToolTip="Services set to automatic" />
                                <Button x:Name="ButtonRunningServices" Content="Running Services" Style="{StaticResource SidebarButton}" ToolTip="Currently running services" />
                                <Button x:Name="ButtonScheduledTasks" Content="Scheduled Tasks" Style="{StaticResource SidebarButton}" ToolTip="Ready scheduled tasks" />
                            </StackPanel>
                        </Expander>

                        <!--  Roles &amp; Features  -->
                        <Expander
                            Header="📦 Roles &amp; Features"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonInstalledFeatures" Content="Installed Features" Style="{StaticResource SidebarButton}" ToolTip="Windows features in use" />
                                <Button x:Name="ButtonInstalledPrograms" Content="Installed Programs" Style="{StaticResource SidebarButton}" ToolTip="Software inventory" />
                                <Button x:Name="ButtonWindowsUpdates" Content="Recent Updates" Style="{StaticResource SidebarButton}" ToolTip="Latest Windows patches" />
                            </StackPanel>
                        </Expander>

                        <!--  IIS  -->
                        <Expander
                            Header="🌐 IIS Web Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonIISWebsites" Content="Websites" Style="{StaticResource SidebarButton}" ToolTip="IIS websites" />
                                <Button x:Name="ButtonIISAppPools" Content="Application Pools" Style="{StaticResource SidebarButton}" ToolTip="App pool configuration" />
                                <Button x:Name="ButtonIISBindings" Content="SSL Bindings" Style="{StaticResource SidebarButton}" ToolTip="SSL certificates and bindings" />
                            </StackPanel>
                        </Expander>

                        <!--  RDS / WTS  -->
                        <Expander
                            Header="🖥️ Remote Desktop Services"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonRDSCollections" Content="Session Collections" Style="{StaticResource SidebarButton}" ToolTip="RDS session collections" />
                                <Button x:Name="ButtonRDSSessionHosts" Content="Session Hosts" Style="{StaticResource SidebarButton}" ToolTip="RDS session hosts" />
                                <Button x:Name="ButtonRDSLicensing" Content="RDS Licensing" Style="{StaticResource SidebarButton}" ToolTip="RDS licensing configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  DFS  -->
                        <Expander
                            Header="🔀 Distributed File System"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDFSNamespaces" Content="DFS Namespaces" Style="{StaticResource SidebarButton}" ToolTip="DFS namespace configuration" />
                                <Button x:Name="ButtonDFSReplication" Content="Replication Groups" Style="{StaticResource SidebarButton}" ToolTip="DFS replication groups" />
                            </StackPanel>
                        </Expander>

                        <!--  Print Server  -->
                        <Expander
                            Header="🖨️ Print Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonPrinters" Content="Printers" Style="{StaticResource SidebarButton}" ToolTip="Print server printers" />
                                <Button x:Name="ButtonPrinterDrivers" Content="Printer Drivers" Style="{StaticResource SidebarButton}" ToolTip="Installed printer drivers" />
                            </StackPanel>
                        </Expander>

                        <!--  WSUS  -->
                        <Expander
                            Header="🔄 WSUS (Update Services)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonWSUSConfig" Content="WSUS Configuration" Style="{StaticResource SidebarButton}" ToolTip="WSUS server configuration" />
                                <Button x:Name="ButtonWSUSGroups" Content="Computer Groups" Style="{StaticResource SidebarButton}" ToolTip="WSUS computer target groups" />
                                <Button x:Name="ButtonWSUSUpdates" Content="Available Updates" Style="{StaticResource SidebarButton}" ToolTip="Available updates in WSUS" />
                            </StackPanel>
                        </Expander>

                        <!--  Hyper-V  -->
                        <Expander
                            Header="🔧 Hyper-V Virtualization"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonHyperVVMs" Content="Virtual Machines" Style="{StaticResource SidebarButton}" ToolTip="Hyper-V virtual machines" />
                                <Button x:Name="ButtonHyperVSwitches" Content="Virtual Switches" Style="{StaticResource SidebarButton}" ToolTip="Hyper-V virtual network switches" />
                                <Button x:Name="ButtonHyperVSnapshots" Content="Snapshots" Style="{StaticResource SidebarButton}" ToolTip="VM snapshots" />
                            </StackPanel>
                        </Expander>

                        <!--  NRAS / NPS  -->
                        <Expander
                            Header="🔐 Network Access Protection"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonNPASConfig" Content="NRAS/NPS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Network Policy Server configuration" />
                                <Button x:Name="ButtonNASClients" Content="RADIUS NAS Clients" Style="{StaticResource SidebarButton}" ToolTip="RADIUS network access servers" />
                            </StackPanel>
                        </Expander>

                        <!--  KMS (Volume Activation)  -->
                        <Expander
                            Header="🔑 KMS (Volume Activation)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonKMSConfig" Content="KMS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Key Management Service configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  WDS (Windows Deployment)  -->
                        <Expander
                            Header="📦 WDS (Deployment Services)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonWDSConfig" Content="WDS Configuration" Style="{StaticResource SidebarButton}" ToolTip="Windows Deployment Services configuration" />
                                <Button x:Name="ButtonWDSBootImages" Content="Boot Images" Style="{StaticResource SidebarButton}" ToolTip="WDS boot images" />
                                <Button x:Name="ButtonWDSInstallImages" Content="Install Images" Style="{StaticResource SidebarButton}" ToolTip="WDS install images" />
                            </StackPanel>
                        </Expander>

                        <!--  File Services &amp; SMB  -->
                        <Expander
                            Header="📁 File Services (SMB)"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonFileShares" Content="File Shares" Style="{StaticResource SidebarButton}" ToolTip="SMB file shares" />
                                <Button x:Name="ButtonSharePermissions" Content="Share Permissions" Style="{StaticResource SidebarButton}" ToolTip="Share ACLs and permissions" />
                                <Button x:Name="ButtonFileQuotas" Content="File Quotas (FSRM)" Style="{StaticResource SidebarButton}" ToolTip="File Server Resource Manager quotas" />
                                <Button x:Name="ButtonShadowCopies" Content="Shadow Copies" Style="{StaticResource SidebarButton}" ToolTip="Volume shadow copies / snapshots" />
                            </StackPanel>
                        </Expander>

                        <!--  Advanced Active Directory  -->
                        <Expander
                            Header="🌳 AD Advanced Info"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonADDCExtended" Content="Domain Controllers (Extended)" Style="{StaticResource SidebarButton}" ToolTip="Detailed DC information" />
                                <Button x:Name="ButtonADFunctionalLevels" Content="Functional Levels" Style="{StaticResource SidebarButton}" ToolTip="Domain and forest functional levels" />
                                <Button x:Name="ButtonADSites" Content="Sites &amp; Subnets" Style="{StaticResource SidebarButton}" ToolTip="AD sites and subnets configuration" />
                                <Button x:Name="ButtonADGPO" Content="Group Policy Summary" Style="{StaticResource SidebarButton}" ToolTip="GPO overview and statistics" />
                                <Button x:Name="ButtonADClustering" Content="Failover Clustering" Style="{StaticResource SidebarButton}" ToolTip="Failover cluster configuration" />
                            </StackPanel>
                        </Expander>

                        <!--  Event Logs  -->
                        <Expander
                            Header="📋 Event Logs"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonSystemEvents" Content="System Events (24h)" Style="{StaticResource SidebarButton}" ToolTip="System event log" />
                                <Button x:Name="ButtonAppEvents" Content="Application Events (24h)" Style="{StaticResource SidebarButton}" ToolTip="Application event log" />
                                <Button x:Name="ButtonSecurityEvents" Content="Security Events (100)" Style="{StaticResource SidebarButton}" ToolTip="Security event log" />
                                <Button x:Name="ButtonFailedLogons" Content="Failed Logon Attempts" Style="{StaticResource SidebarButton}" ToolTip="Event ID 4625" />
                                <Button x:Name="ButtonAccountLockouts" Content="Account Lockouts" Style="{StaticResource SidebarButton}" ToolTip="Event ID 4740" />
                            </StackPanel>
                        </Expander>

                        <!--  Security  -->
                        <Expander
                            Header="🔒 Security &amp; Users"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonLocalUsers" Content="Local Users" Style="{StaticResource SidebarButton}" ToolTip="Local user accounts" />
                                <Button x:Name="ButtonLocalGroups" Content="Local Groups" Style="{StaticResource SidebarButton}" ToolTip="Local group membership" />
                                <Button x:Name="ButtonPrivilegeAudit" Content="Privilege Use Audit" Style="{StaticResource SidebarButton}" ToolTip="Event IDs 4672/4673/4674" />
                            </StackPanel>
                        </Expander>

                        <!--  Active Directory  -->
                        <Expander
                            Header="🌳 Active Directory"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonADDC" Content="Domain Controllers" Style="{StaticResource SidebarButton}" ToolTip="AD domain controller status" />
                                <Button x:Name="ButtonADDomain" Content="Domain Info" Style="{StaticResource SidebarButton}" ToolTip="AD domain properties" />
                                <Button x:Name="ButtonADForest" Content="Forest Info" Style="{StaticResource SidebarButton}" ToolTip="AD forest properties" />
                                <Button x:Name="ButtonADOUs" Content="Organizational Units" Style="{StaticResource SidebarButton}" ToolTip="AD OUs" />
                                <Button x:Name="ButtonADAdmins" Content="Domain Admins" Style="{StaticResource SidebarButton}" ToolTip="Domain admin members" />
                                <Button x:Name="ButtonADComputers" Content="Computer Accounts" Style="{StaticResource SidebarButton}" ToolTip="AD computer objects" />
                                <Button x:Name="ButtonADReplStatus" Content="Replication Status" Style="{StaticResource SidebarButton}" ToolTip="AD replication status" />
                                <Button x:Name="ButtonADTrusts" Content="Trust Relationships" Style="{StaticResource SidebarButton}" ToolTip="AD domain trusts" />
                            </StackPanel>
                        </Expander>

                        <!--  DNS  -->
                        <Expander
                            Header="🔗 DNS Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDNSConfig" Content="DNS Configuration" Style="{StaticResource SidebarButton}" ToolTip="DNS server settings" />
                                <Button x:Name="ButtonDNSZones" Content="DNS Zones" Style="{StaticResource SidebarButton}" ToolTip="DNS zones" />
                                <Button x:Name="ButtonDNSForwarders" Content="DNS Forwarders" Style="{StaticResource SidebarButton}" ToolTip="DNS forwarders" />
                                <Button x:Name="ButtonDNSCache" Content="DNS Cache" Style="{StaticResource SidebarButton}" ToolTip="DNS cache entries" />
                            </StackPanel>
                        </Expander>

                        <!--  DHCP  -->
                        <Expander
                            Header="📡 DHCP Server"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button x:Name="ButtonDHCPConfig" Content="DHCP Configuration" Style="{StaticResource SidebarButton}" ToolTip="DHCP server settings" />
                                <Button x:Name="ButtonDHCPv4Scopes" Content="IPv4 Scopes" Style="{StaticResource SidebarButton}" ToolTip="DHCP IPv4 scopes" />
                                <Button x:Name="ButtonDHCPv6Scopes" Content="IPv6 Scopes" Style="{StaticResource SidebarButton}" ToolTip="DHCP IPv6 scopes" />
                                <Button x:Name="ButtonDHCPReservations" Content="DHCP Reservations" Style="{StaticResource SidebarButton}" ToolTip="DHCP reservations" />
                            </StackPanel>
                        </Expander>

                        <!--  Export Options  -->
                        <StackPanel Margin="0,20,0,0">
                            <Button
                                x:Name="ButtonExportCSV"
                                Content="📥 Export to CSV"
                                Style="{StaticResource PrimaryButton}"
                                Margin="0,0,0,8" />
                            <Button
                                x:Name="ButtonClearOutput"
                                Content="🗑️ Clear Output"
                                Style="{StaticResource PrimaryButton}" />
                        </StackPanel>
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!--  Output Area  -->
            <Border
                Grid.Column="2"
                Padding="20"
                Background="White"
                BorderBrush="#E5E7EB"
                BorderThickness="1"
                CornerRadius="6">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>

                    <TextBlock
                        Grid.Row="0"
                        Margin="0,0,0,12"
                        FontSize="14"
                        FontWeight="SemiBold"
                        Foreground="#374151"
                        Text="📊 Audit Results" />

                    <DataGrid
                        x:Name="DataGridResults"
                        Grid.Row="1"
                        AlternatingRowBackground="#F8FAFC"
                        AutoGenerateColumns="True"
                        Background="White"
                        BorderBrush="#E5E7EB"
                        BorderThickness="1"
                        CanUserReorderColumns="True"
                        CanUserResizeColumns="True"
                        CanUserSortColumns="True"
                        GridLinesVisibility="Horizontal"
                        HeadersVisibility="Column"
                        IsReadOnly="True"
                        RowBackground="White">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#F3F4F6" />
                                <Setter Property="Foreground" Value="#374151" />
                                <Setter Property="FontWeight" Value="SemiBold" />
                                <Setter Property="BorderBrush" Value="#E5E7EB" />
                                <Setter Property="BorderThickness" Value="0,0,1,1" />
                                <Setter Property="Padding" Value="12,8" />
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="Padding" Value="12,6" />
                                <Setter Property="BorderThickness" Value="0" />
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="#EEF2FF" />
                                        <Setter Property="Foreground" Value="#6366F1" />
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.CellStyle>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>

        <!--  Footer (Grid.Row="2")  -->
        <Border
            Grid.Row="2"
            Background="#E5E7EB"
            Padding="12,4">
            <TextBlock
                x:Name="StatusBarText"
                FontSize="11"
                Foreground="#6B7280"
                Text="Ready" />
        </Border>
    </Grid>
</Window>
"@

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Create WPF window
$reader = New-Object System.Xml.XmlNodeReader $Global:XAML
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get control references
$DataGridResults = $window.FindName("DataGridResults")
$StatusBarText = $window.FindName("StatusBarText")
$TotalResultCountText = $window.FindName("TotalResultCountText")

# Variables for UI state
$script:ResultsCount = 0
$script:AllResults = @()
$script:CurrentButton = $null

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

function Update-Output {
    param([PSObject[]]$Data)
    try {
        # Convert all inputs to a consistent array format
        $processedData = @()
        
        if ($null -eq $Data -or $Data.Count -eq 0) {
            $script:ResultsCount = 0
        } else {
            # Handle different input types
            foreach ($item in $Data) {
                if ($null -ne $item) {
                    # If single object (not in an array)
                    if ($item -is [System.Collections.IEnumerable] -and $item -isnot [string] -and $item -isnot [System.Collections.Specialized.OrderedDictionary]) {
                        foreach ($subitem in $item) {
                            $processedData += $subitem
                        }
                    } else {
                        $processedData += $item
                    }
                }
            }
            
            # Create ObservableCollection
            $collection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
            foreach ($item in $processedData) {
                $collection.Add($item)
            }
            $script:ResultsCount = $collection.Count
            
            # UI update on Dispatcher thread
            $window.Dispatcher.Invoke({
                try {
                    $DataGridResults.ItemsSource = $null
                    
                    # Auto-generate columns based on data
                    $DataGridResults.AutoGenerateColumns = $true
                    
                    $DataGridResults.ItemsSource = $collection
                    $TotalResultCountText.Text = $script:ResultsCount.ToString()
                    
                    # Force refresh
                    $DataGridResults.Items.Refresh()
                } catch {
                    Write-Host "UI-Update error: $_" -ForegroundColor Red
                }
            }, "Normal")
            
            return
        }
        
        # Fallback if data is empty
        $window.Dispatcher.Invoke({
            $DataGridResults.ItemsSource = $null
            $TotalResultCountText.Text = "0"
        }, "Normal")
        
    } catch {
        Write-Host "Update-Output error: $_" -ForegroundColor Red
        $window.Dispatcher.Invoke({
            try {
                $DataGridResults.ItemsSource = $null
                $TotalResultCountText.Text = "0"
            } catch {}
        }, "Normal")
    }
}

function Clear-Output {
    $DataGridResults.ItemsSource = @()
    $script:ResultsCount = 0
    $TotalResultCountText.Text = "0"
}

function Update-Status {
    param([string]$Status)
    $StatusBarText.Text = $Status
    $window.Dispatcher.Invoke([Action]{}, "Background")
}

# ============================================================================
# TABLE FORMATTING FUNCTIONS - REMOVED (Using DataGrid Now)
# ============================================================================

# All outputs now go directly through DataGrid using PSObjects
# The old Format-* functions are no longer needed

# ============================================================================
# AUDIT FUNCTIONS
# ============================================================================

function Get-SystemInformation {
    try {
        $info = Get-ComputerInfo -ErrorAction Stop
        $data = @(
            @{"Attribute" = "Computer Name"; "Value" = $info.CsComputerName},
            @{"Attribute" = "Domain"; "Value" = $info.CsDomain},
            @{"Attribute" = "Operating System"; "Value" = $info.OsName},
            @{"Attribute" = "Install Date"; "Value" = $info.OsInstallDate},
            @{"Attribute" = "Last Boot"; "Value" = $info.OsLastBootUpTime}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-OSDetails {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $data = @(
            @{"Property" = "Caption"; "Value" = $os.Caption},
            @{"Property" = "Version"; "Value" = $os.Version},
            @{"Property" = "Build"; "Value" = $os.BuildNumber},
            @{"Property" = "Total Memory"; "Value" = "$('{0:N0}' -f ($os.TotalVisibleMemorySize / 1024)) MB"},
            @{"Property" = "Free Memory"; "Value" = "$('{0:N0}' -f ($os.FreePhysicalMemory / 1024)) MB"},
            @{"Property" = "Last Boot"; "Value" = $os.LastBootUpTime}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-HardwareSummary {
    try {
        $hw = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $data = @(
            @{"Property" = "Manufacturer"; "Value" = $hw.Manufacturer},
            @{"Property" = "Model"; "Value" = $hw.Model},
            @{"Property" = "Processors"; "Value" = $hw.NumberOfProcessors},
            @{"Property" = "Logical Cores"; "Value" = $hw.NumberOfLogicalProcessors},
            @{"Property" = "RAM (GB)"; "Value" = "$('{0:N2}' -f ($hw.TotalPhysicalMemory / 1GB))"},
            @{"Property" = "System Type"; "Value" = $hw.SystemType}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-CPUDetails {
    try {
        $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $data = @(
            @{"Property" = "Name"; "Value" = $cpu.Name},
            @{"Property" = "Cores"; "Value" = $cpu.NumberOfCores},
            @{"Property" = "Logical Processors"; "Value" = $cpu.NumberOfLogicalProcessors},
            @{"Property" = "Speed (GHz)"; "Value" = "$('{0:N2}' -f ($cpu.MaxClockSpeed / 1000))"},
            @{"Property" = "Architecture"; "Value" = $cpu.Architecture},
            @{"Property" = "Cache (KB)"; "Value" = $cpu.L3CacheSize}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-MemoryDetails {
    try {
        $mem = Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop
        $data = @()
        $data += @{"Property" = "Number of Modules"; "Value" = $mem.Count}
        foreach ($m in $mem) {
            $data += @{
                "Property" = $m.PartNumber
                "Value" = "$('{0:N2}' -f ($m.Capacity / 1GB)) GB @ $($m.Speed) MHz"
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-StorageSummary {
    try {
        $disks = Get-CimInstance Win32_LogicalDisk -ErrorAction Stop | Where-Object DriveType -eq 3
        $data = @()
        foreach ($disk in $disks) {
            if ($disk.Size -gt 0) {
                $used = '{0:N2}' -f (($disk.Size - $disk.FreeSpace) / 1GB)
                $total = '{0:N2}' -f ($disk.Size / 1GB)
                $percent = [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 2)
                $freePercent = [math]::Round(100 - $percent, 2)
            } else {
                $used = "0"
                $total = "0"
                $percent = "0"
                $freePercent = "100"
            }
            $data += @{
                "Drive" = $disk.Name
                "Total(GB)" = $total
                "Used(GB)" = $used
                "Free(%)" = $freePercent
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-NetworkConfiguration {
    try {
        $config = Get-NetIPConfiguration -ErrorAction Stop
        $data = @()
        foreach ($cfg in $config) {
            $data += @{
                "Interface" = $cfg.InterfaceAlias
                "IPv4" = ($cfg.IPv4Address.IPAddress -join ', ')
                "Gateway" = ($cfg.IPv4DefaultGateway.NextHopAddress -join ', ')
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-NetworkAdapters {
    try {
        $adapters = Get-NetAdapter -ErrorAction Stop
        $data = @()
        foreach ($adapter in $adapters) {
            $speedGbps = '{0:N2}' -f ($adapter.Speed / 1000000000)
            $data += @{
                "Name" = $adapter.Name
                "Status" = $adapter.Status
                "Speed(Gbps)" = $speedGbps
                "MAC" = $adapter.MacAddress
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ActiveConnections {
    try {
        $connections = Get-NetTCPConnection -State Listen -ErrorAction Stop | Select-Object LocalAddress, LocalPort, OwningProcess -First 50
        $data = @()
        foreach ($conn in $connections) {
            $process = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
            $data += @{
                "IP:Port" = "$($conn.LocalAddress):$($conn.LocalPort)"
                "PID" = $conn.OwningProcess
                "Process" = $process.ProcessName
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-FirewallRules {
    try {
        $rules = Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' } | Select-Object DisplayName, Direction, Action -First 50
        if ($null -eq $rules) {
            return @([PSCustomObject]@{"Information" = "No enabled firewall rules found"})
        }
        $data = @()
        foreach ($rule in $rules) {
            $name = $rule.DisplayName.Substring(0, [Math]::Min(40, $rule.DisplayName.Length))
            $data += @{
                "Name" = $name
                "Direction" = $rule.Direction
                "Action" = $rule.Action
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-AutomaticServices {
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name | Select-Object -First 50
        return $services | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Status" = $_.Status; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-RunningServices {
    try {
        $services = Get-Service -ErrorAction Stop | Where-Object Status -eq 'Running' | Sort-Object Name | Select-Object -First 50
        return $services | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ScheduledTasks {
    try {
        $tasks = Get-ScheduledTask -ErrorAction Stop | Where-Object State -eq 'Ready' | Select-Object -First 50
        return $tasks | ForEach-Object { [PSCustomObject]@{"Path" = $_.TaskPath; "Name" = $_.TaskName; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-InstalledFeatures {
    try {
        $features = Get-WindowsFeature -ErrorAction Stop | Where-Object Installed -eq $true | Select-Object Name, DisplayName
        return $features | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Display" = $_.DisplayName} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-InstalledPrograms {
    try {
        $programs = Get-CimInstance Win32_Product -ErrorAction Stop | Select-Object Name, Version, Vendor | Sort-Object Name | Select-Object -First 100
        return $programs | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Version" = $_.Version; "Vendor" = $_.Vendor} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WindowsUpdates {
    try {
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 20
        return $updates | ForEach-Object { [PSCustomObject]@{"KB" = $_.HotFixID; "InstalledOn" = $_.InstalledOn} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-SystemEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No system events found"})
        }
        return $events | Select-Object -First 50 | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
                "Message" = $_.Message.Substring(0, [Math]::Min(50, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ApplicationEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No application events found"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-SecurityEvents {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'} -MaxEvents 100 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No security events found"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "ID" = $_.Id
                "Level" = $_.LevelDisplayName
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-FailedLogons {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No failed logon attempts found"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "Message" = $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-AccountLockouts {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4740} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No account lockouts found"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "Message" = $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-LocalUsers {
    try {
        $users = Get-LocalUser -ErrorAction Stop
        if ($null -eq $users) {
            return @([PSCustomObject]@{"Information" = "No local users found"})
        }
        $data = @()
        foreach ($user in $users) {
            $data += @{
                "Name" = $user.Name
                "Enabled" = if($user.Enabled) {"Yes"} else {"No"}
                "LastLogon" = $user.LastLogon
                "PasswordRequired" = if($user.PasswordRequired) {"Yes"} else {"No"}
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-LocalGroups {
    try {
        $groups = Get-LocalGroup -ErrorAction Stop
        if ($null -eq $groups) {
            return @([PSCustomObject]@{"Information" = "No local groups found"})
        }
        return $groups | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-PrivilegeAudit {
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4672,4673,4674} -MaxEvents 50 -ErrorAction Stop
        if ($null -eq $events) {
            return @([PSCustomObject]@{"Information" = "No privilege audit events found"})
        }
        return $events | ForEach-Object { 
            [PSCustomObject]@{
                "Time" = $_.TimeCreated
                "ID" = $_.Id
                "Message" = $_.Message.Substring(0, [Math]::Min(60, $_.Message.Length))
            } 
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# AD Functions (old-style still converted to PSObjects)
function Get-ADDomainControllers {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        return $dcs | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Site" = $_.Site; "IP" = $_.IPv4Address; "GC" = $_.IsGlobalCatalog} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADDomainInfo {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $data = @(
            @{"Property" = "Name"; "Value" = $domain.Name},
            @{"Property" = "NetBIOS"; "Value" = $domain.NetBIOSName},
            @{"Property" = "Mode"; "Value" = $domain.DomainMode},
            @{"Property" = "PDC"; "Value" = $domain.PDCEmulator}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DNSZones {
    try {
        $zones = Get-DnsServerZone -ErrorAction Stop | Select-Object -First 50
        return $zones | ForEach-Object { [PSCustomObject]@{"Zone" = $_.ZoneName; "Type" = $_.ZoneType; "DS" = $_.IsDsIntegrated} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-FileShares {
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name, Path, Description, ShareType
        return $shares | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Path" = $_.Path; "Type" = $_.ShareType} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-IISWebsites {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $sites = Get-Website -ErrorAction Stop | Select-Object Name, State, PhysicalPath
        return $sites | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Path" = $_.PhysicalPath} }
    } catch {
        return @([PSCustomObject]@{"Error" = "IIS: " + $_.Exception.Message})
    }
}

function Get-RDSCollections {
    try {
        $collections = Get-RDSessionCollection -ErrorAction Stop | Select-Object CollectionName, CollectionDescription
        return $collections | ForEach-Object { [PSCustomObject]@{"Name" = $_.CollectionName; "Description" = $_.CollectionDescription} }
    } catch {
        return @([PSCustomObject]@{"Error" = "RDS: " + $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS
# ============================================================================

function Get-ADForestInfo {
    try {
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"Property" = "Name"; "Value" = $forest.Name},
            @{"Property" = "Mode"; "Value" = $forest.ForestMode},
            @{"Property" = "Domains"; "Value" = ($forest.Domains | Measure-Object).Count},
            @{"Property" = "Sites"; "Value" = ($forest.Sites | Measure-Object).Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADOUs {
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -ErrorAction Stop | Select-Object Name, DistinguishedName | Sort-Object Name | Select-Object -First 50
        return $ous | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "DN" = $_.DistinguishedName} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADDomainAdmins {
    try {
        $admins = Get-ADGroupMember -Identity 'Domain Admins' -ErrorAction Stop | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled -ErrorAction SilentlyContinue | Select-Object Name, SamAccountName, Enabled, LastLogonDate
        return $admins | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Account" = $_.SamAccountName; "Enabled" = $_.Enabled; "LastLogon" = $_.LastLogonDate} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADComputers {
    try {
        $computers = Get-ADComputer -Filter * -Properties OperatingSystem, LastLogonDate -ErrorAction Stop | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Sort-Object LastLogonDate -Descending | Select-Object -First 50
        return $computers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "OS" = $_.OperatingSystem; "LastLogon" = $_.LastLogonDate; "Enabled" = $_.Enabled} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADReplicationStatus {
    try {
        $output = repadmin /replsummary 2>&1
        $data = @()
        
        # Parse repadmin output
        foreach ($line in $output) {
            if (-not [string]::IsNullOrWhiteSpace($line) -and $line -match '\S') {
                $data += @{
                    "Information" = $line.Trim()
                }
            }
        }
        
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No replication information available"})
        }
        
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADTrusts {
    try {
        $trusts = Get-ADTrust -Filter * -ErrorAction Stop | Select-Object Name, Direction, TrustType
        return $trusts | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Direction" = $_.Direction; "Type" = $_.TrustType} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# DNS FUNCTIONS
# ============================================================================

function Get-DNSConfiguration {
    try {
        # Try Get-DnsServerSetting first (better than Get-DnsServer)
        $dnsSettings = Get-DnsServerSetting -ErrorAction SilentlyContinue
        
        if ($null -ne $dnsSettings) {
            $data = @(
                @{"Setting" = "Listen Addresses"; "Value" = ($dnsSettings.ListeningIPAddress -join ", ")},
                @{"Setting" = "All Zones Writeable"; "Value" = if($dnsSettings.AllowZoneEditing) {"Yes"} else {"No"}},
                @{"Setting" = "DNSSEC Enabled"; "Value" = if($dnsSettings.EnableDnsSec) {"Yes"} else {"No"}},
                @{"Setting" = "Log Queries"; "Value" = if($dnsSettings.LogQueries) {"Yes"} else {"No"}},
                @{"Setting" = "Write to Log"; "Value" = if($dnsSettings.WriteToLog) {"Yes"} else {"No"}}
            )
            return $data | ForEach-Object { [PSCustomObject]$_ }
        } else {
            # Fallback: Get-DnsServer (older Windows versions)
            $dns = Get-DnsServer -ErrorAction Stop
            $data = @(
                @{"Setting" = "Computer"; "Value" = $dns.ComputerName},
                @{"Setting" = "Version"; "Value" = $dns.Version}
            )
            return $data | ForEach-Object { [PSCustomObject]$_ }
        }
    } catch {
        return @([PSCustomObject]@{"Error" = "DNS not available or no access: $($_.Exception.Message)"})
    }
}

function Get-DNSForwarders {
    try {
        $fwd = Get-DnsServerForwarder -ErrorAction Stop
        if ($null -ne $fwd.IPAddress) {
            return $fwd.IPAddress | ForEach-Object { [PSCustomObject]@{"Forwarder" = $_} }
        } else {
            return @([PSCustomObject]@{"Forwarder" = "No forwarders configured"})
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DNSCache {
    try {
        $cache = Get-DnsServerCache -ErrorAction Stop
        return @([PSCustomObject]@{"CacheSize" = $cache.CacheSize; "MaxTTL" = $cache.MaxTTL; "MaxNegativeTTL" = $cache.MaxNegativeTTL})
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# DHCP FUNCTIONS
# ============================================================================

function Get-DHCPConfiguration {
    try {
        $dhcp = Get-DhcpServerInDC -ErrorAction Stop | Select-Object -First 10
        return $dhcp | ForEach-Object { [PSCustomObject]@{"Server" = $_.ToString()} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPv4Scopes {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop | Select-Object ScopeId, Name, StartRange, EndRange, State | Select-Object -First 50
        return $scopes | ForEach-Object { [PSCustomObject]@{"ScopeID" = $_.ScopeId; "Name" = $_.Name; "Start" = $_.StartRange; "End" = $_.EndRange; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPv6Scopes {
    try {
        $scopes = Get-DhcpServerv6Scope -ErrorAction Stop | Select-Object Prefix, Name, State | Select-Object -First 50
        return $scopes | ForEach-Object { [PSCustomObject]@{"Prefix" = $_.Prefix; "Name" = $_.Name; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPReservations {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        $reservations = @()
        foreach ($scope in $scopes) {
            $scopeRes = Get-DhcpServerv4Reservation -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue | Select-Object -First 20
            $reservations += $scopeRes
        }
        return $reservations | ForEach-Object { [PSCustomObject]@{"ScopeID" = $_.ScopeId; "IP" = $_.IPAddress; "Name" = $_.Name; "Type" = $_.ReservationType} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# IIS FUNCTIONS
# ============================================================================

function Get-IISAppPools {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $pools = Get-ChildItem IIS:\AppPools -ErrorAction Stop | Select-Object Name, State, ManagedRuntimeVersion
        return $pools | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Runtime" = $_.ManagedRuntimeVersion} }
    } catch {
        return @([PSCustomObject]@{"Error" = "IIS: " + $_.Exception.Message})
    }
}

function Get-IISBindings {
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $sites = Get-ChildItem IIS:\Sites -ErrorAction Stop
        $bindings = @()
        foreach ($site in $sites) {
            $site.Bindings.Collection | ForEach-Object { $bindings += $_ }
        }
        return $bindings | Select-Object -First 50 | ForEach-Object { [PSCustomObject]@{"Protocol" = $_.Protocol; "IP" = $_.BindingInformation} }
    } catch {
        return @([PSCustomObject]@{"Error" = "IIS: " + $_.Exception.Message})
    }
}

# ============================================================================
# RDS FUNCTIONS
# ============================================================================

function Get-RDSSessionHosts {
    try {
        $hosts = Get-RDSessionHost -ErrorAction Stop | Select-Object SessionHost, NewConnectionAllowed | Select-Object -First 50
        return $hosts | ForEach-Object { [PSCustomObject]@{"Host" = $_.SessionHost; "NewConnections" = $_.NewConnectionAllowed} }
    } catch {
        return @([PSCustomObject]@{"Error" = "RDS: " + $_.Exception.Message})
    }
}

function Get-RDSActiveLicensing {
    try {
        $licensing = Get-RDLicenseConfiguration -ErrorAction Stop
        return @([PSCustomObject]@{"Mode" = $licensing.Mode; "Server" = $licensing.LicenseServer; "IssuedLicenses" = $licensing.IssuedLicenses})
    } catch {
        return @([PSCustomObject]@{"Error" = "RDS: " + $_.Exception.Message})
    }
}

# ============================================================================
# DFS FUNCTIONS
# ============================================================================

function Get-DFSNamespaces {
    try {
        $namespaces = Get-DfsnRoot -ErrorAction Stop | Select-Object Path, Type, State | Select-Object -First 50
        return $namespaces | ForEach-Object { [PSCustomObject]@{"Path" = $_.Path; "Type" = $_.Type; "State" = $_.State} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DFSReplicationGroups {
    try {
        $groups = Get-DfsReplicationGroup -ErrorAction Stop | Select-Object GroupName, State, Description | Select-Object -First 50
        return $groups | ForEach-Object { [PSCustomObject]@{"Name" = $_.GroupName; "State" = $_.State; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# PRINT SERVER FUNCTIONS
# ============================================================================

function Get-PrintServers {
    try {
        $printers = Get-Printer -ErrorAction Stop | Select-Object Name, DriverName, Shared, Published | Select-Object -First 50
        return $printers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Driver" = $_.DriverName; "Shared" = $_.Shared; "Published" = $_.Published} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-PrinterDrivers {
    try {
        $drivers = Get-PrinterDriver -ErrorAction Stop | Select-Object Name, Manufacturer, DriverVersion | Select-Object -First 50
        return $drivers | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Vendor" = $_.Manufacturer; "Version" = $_.DriverVersion} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# WSUS FUNCTIONS
# ============================================================================

function Get-WSUSConfiguration {
    try {
        $wsus = Get-WsusServer -ErrorAction Stop | Select-Object Name, PortNumber, ServerProtocolVersion
        return @([PSCustomObject]@{"Name" = $wsus.Name; "Port" = $wsus.PortNumber; "Protocol" = $wsus.ServerProtocolVersion})
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WSUSComputerTargetGroups {
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $groups = $server | Get-WsusComputerTargetGroup -ErrorAction Stop | Select-Object Name | Select-Object -First 50
        return $groups | ForEach-Object { [PSCustomObject]@{"Group" = $_.Name} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WSUSUpdates {
    try {
        $server = Get-WsusServer -ErrorAction Stop
        $updates = $server.GetUpdates() | Select-Object Title, Classification, ApprovedCount | Select-Object -First 50
        return $updates | ForEach-Object { [PSCustomObject]@{"Title" = $_.Title; "Classification" = $_.Classification; "Approved" = $_.ApprovedCount} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# HYPER-V FUNCTIONS
# ============================================================================

function Get-HyperVVirtualMachines {
    try {
        $vms = Get-VM -ErrorAction Stop | Select-Object Name, State, MemoryAssigned, ProcessorCount | Select-Object -First 50
        return $vms | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "RAM(GB)" = [Math]::Round($_.MemoryAssigned / 1GB, 2); "CPUs" = $_.ProcessorCount} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-HyperVSwitches {
    try {
        $switches = Get-VMSwitch -ErrorAction Stop | Select-Object Name, SwitchType, NetAdapterInterfaceDescription | Select-Object -First 50
        return $switches | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Type" = $_.SwitchType; "Adapter" = $_.NetAdapterInterfaceDescription} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-HyperVSnapshots {
    try {
        $snapshots = Get-VMSnapshot -ErrorAction Stop | Select-Object VMName, Name, CreationTime | Select-Object -First 50
        return $snapshots | ForEach-Object { [PSCustomObject]@{"VM" = $_.VMName; "Snapshot" = $_.Name; "Created" = $_.CreationTime} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# NRAS / NPS FUNCTIONS
# ============================================================================

function Get-NPASConfiguration {
    try {
        $nasclients = Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address | Select-Object -First 50
        return $nasclients | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "IP" = $_.Address} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-NASClients {
    try {
        $nasclients = Get-NpsRadiusClient -ErrorAction Stop | Select-Object Name, Address | Select-Object -First 50
        return $nasclients | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "IP" = $_.Address} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# KMS FUNCTIONS
# ============================================================================

function Get-KMSConfiguration {
    try {
        $kms = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction Stop
        return @([PSCustomObject]@{"ServiceRunning" = $kms.IsS_Running; "Version" = $kms.Version; "VLActivationInterval" = $kms.VLActivationInterval})
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# WDS FUNCTIONS
# ============================================================================

function Get-WDSConfiguration {
    try {
        $wdsReg = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\WDSServer\Providers\WDSPXE" -ErrorAction SilentlyContinue
        if ($wdsReg) {
            return @([PSCustomObject]@{"Status" = "WDS service exists"; "Version" = $wdsReg.Version})
        } else {
            return @([PSCustomObject]@{"Status" = "WDS service not found"})
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WDSBootImages {
    try {
        $wdsPath = "C:\RemoteInstall\Boot" # Standard WDS Path
        if (Test-Path $wdsPath) {
            $images = Get-ChildItem -Path $wdsPath -Filter "*.wim" -ErrorAction SilentlyContinue | Select-Object -First 50
            return $images | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Size(MB)" = [Math]::Round($_.Length / 1MB, 2)} }
        } else {
            return @([PSCustomObject]@{"Info" = "WDS Boot Images path not found"})
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WDSInstallImages {
    try {
        $wdsPath = "C:\RemoteInstall\Images" # Standard WDS Path
        if (Test-Path $wdsPath) {
            $images = Get-ChildItem -Path $wdsPath -Filter "*.wim" -ErrorAction SilentlyContinue | Select-Object -First 50
            return $images | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Size(MB)" = [Math]::Round($_.Length / 1MB, 2)} }
        } else {
            return @([PSCustomObject]@{"Info" = "WDS Install Images path not found"})
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# FILE SERVICES & SMB FUNCTIONS
# ============================================================================

function Get-FileSharePermissions {
    try {
        $shares = Get-SmbShare -ErrorAction Stop | Select-Object Name | Select-Object -First 20
        $perms = @()
        foreach ($share in $shares) {
            $sharePerm = Get-SmbShareAccess -Name $share.Name -ErrorAction SilentlyContinue
            $perms += $sharePerm
        }
        return $perms | Select-Object -First 100 | ForEach-Object { [PSCustomObject]@{"Share" = $_.Name; "Account" = $_.AccountName; "Access" = $_.AccessRight} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-FileServerQuotas {
    try {
        $quotas = Get-FsrmQuota -ErrorAction Stop | Select-Object Path, Size, SoftLimit | Select-Object -First 50
        return $quotas | ForEach-Object { [PSCustomObject]@{"Path" = $_.Path; "Size(MB)" = [Math]::Round($_.Size / 1MB, 2); "SoftLimit" = $_.SoftLimit} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ShadowCopies {
    try {
        $shadows = Get-CimInstance -ClassName Win32_ShadowCopy -ErrorAction Stop | Select-Object VolumeName, InstallDate, ID | Select-Object -First 50
        return $shadows | ForEach-Object { [PSCustomObject]@{"Volume" = $_.VolumeName; "Date" = $_.InstallDate; "ID" = $_.ID.Substring(0, [Math]::Min(20, $_.ID.Length))} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS
# ============================================================================

function Get-ADDomainControllerExtended {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop
        if ($null -eq $dcs) {
            return @([PSCustomObject]@{"Information" = "No domain controllers found"})
        }
        $data = @()
        foreach ($dc in $dcs) {
            $data += @{
                "Name" = $dc.Name
                "Site" = $dc.Site
                "OS" = $dc.OperatingSystem
                "GC" = if($dc.IsGlobalCatalog) {"Yes"} else {"No"}
                "RODC" = if($dc.IsReadOnly) {"Yes"} else {"No"}
                "IPv4" = $dc.IPv4Address
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADDomainFunctionalLevel {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop
        return @(
            [PSCustomObject]@{"Property" = "Domain Name"; "Value" = $domain.Name},
            [PSCustomObject]@{"Property" = "Domain Mode"; "Value" = $domain.DomainMode},
            [PSCustomObject]@{"Property" = "Forest Name"; "Value" = $forest.Name},
            [PSCustomObject]@{"Property" = "Forest Mode"; "Value" = $forest.ForestMode}
        )
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADSiteConfiguration {
    try {
        $sites = Get-ADReplicationSite -ErrorAction Stop | Select-Object Name, Description | Select-Object -First 50
        return $sites | ForEach-Object { [PSCustomObject]@{"Site" = $_.Name; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADGroupPolicySummary {
    try {
        $gpos = Get-GPO -All -ErrorAction Stop | Measure-Object
        $topGPOs = Get-GPO -All -ErrorAction Stop | Select-Object DisplayName, CreationTime | Sort-Object CreationTime -Descending | Select-Object -First 30
        return $topGPOs | ForEach-Object { [PSCustomObject]@{"Name" = $_.DisplayName; "Created" = $_.CreationTime} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADClusterInformation {
    try {
        $cluster = Get-Cluster -ErrorAction Stop
        $nodes = Get-ClusterNode -ErrorAction Stop | Select-Object Name, State | Select-Object -First 50
        return $nodes | ForEach-Object { [PSCustomObject]@{"Node" = $_.Name; "State" = $_.State; "Cluster" = $cluster.Name} }
    } catch {
        return @([PSCustomObject]@{"Error" = "Failover Clustering: " + $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED AD FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-ADUserAccounts {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"FSMO_Role" = "PDC Emulator"; "Owner" = $domain.PDCEmulator},
            @{"FSMO_Role" = "RID Master"; "Owner" = $domain.RIDMaster},
            @{"FSMO_Role" = "Infrastructure Master"; "Owner" = $domain.InfrastructureMaster},
            @{"FSMO_Role" = "Schema Master"; "Owner" = $forest.SchemaMaster},
            @{"FSMO_Role" = "Domain Naming Master"; "Owner" = $forest.DomainNamingMaster}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADGroupAccounts {
    try {
        $forest = Get-ADForest -ErrorAction Stop
        $data = @(
            @{"Property" = "Schema Version"; "Value" = $forest.SchemaVersion},
            @{"Property" = "Forest Mode"; "Value" = $forest.ForestMode},
            @{"Property" = "Exchange Version"; "Value" = $forest.ExchangeVersion},
            @{"Property" = "Domains Count"; "Value" = ($forest.Domains | Measure-Object).Count},
            @{"Property" = "Sites Count"; "Value" = ($forest.Sites | Measure-Object).Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADServiceAccounts {
    try {
        $features = Get-ADOptionalFeature -Filter * -ErrorAction Stop | Select-Object Name, @{Name='Status';Expression={if($_.EnabledScopes.Count -gt 0) {"Enabled"} else {"Disabled"}}} | Sort-Object Name | Select-Object -First 50
        if ($null -eq $features) {
            return @([PSCustomObject]@{"Information" = "No optional features found"})
        }
        return $features | ForEach-Object { [PSCustomObject]@{"Feature" = $_.Name; "Status" = $_.Status} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADComputerAccounts {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly
        if ($null -eq $dcs) {
            return @([PSCustomObject]@{"Information" = "No DCs found"})
        }
        return $dcs | ForEach-Object { [PSCustomObject]@{"DC" = $_.Name; "Site" = $_.Site; "IP" = $_.IPv4Address; "OS" = $_.OperatingSystem; "GC" = $_.IsGlobalCatalog; "RODC" = $_.IsReadOnly} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADPasswordPolicy {
    try {
        $trusts = Get-ADTrust -Filter * -ErrorAction Stop | Select-Object Name, Direction, TrustType, TrustAttributes | Sort-Object Name | Select-Object -First 50
        if ($null -eq $trusts) {
            return @([PSCustomObject]@{"Information" = "No trust relationships found"})
        }
        return $trusts | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Direction" = $_.Direction; "Type" = $_.TrustType} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADReplicationSites {
    try {
        $gpos = Get-GPO -All -ErrorAction Stop | Select-Object DisplayName, CreationTime, ModificationTime | Sort-Object ModificationTime -Descending | Select-Object -First 50
        if ($null -eq $gpos) {
            return @([PSCustomObject]@{"Information" = "No GPOs found"})
        }
        return $gpos | ForEach-Object { [PSCustomObject]@{"Name" = $_.DisplayName; "Created" = $_.CreationTime; "Modified" = $_.ModificationTime} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADReplicationSubnets {
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop
        $data = @()
        foreach ($dc in $dcs) {
            try {
                $services = @("NTDS", "DFSR", "DNS", "KDC") | ForEach-Object { 
                    $svc = Get-Service -Name $_ -ComputerName $dc.HostName -ErrorAction SilentlyContinue
                    if ($null -ne $svc) { "$_`:$($svc.Status.ToString())" }
                }
                $data += @{
                    "DC" = $dc.Name
                    "Reachable" = "Yes"
                    "Services" = ($services -join ", ")
                }
            } catch {
                $data += @{
                    "DC" = $dc.Name
                    "Reachable" = "No"
                    "Services" = "Unable to check"
                }
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No DC health data available"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ADReplicationSiteLinks {
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $dcSites = Get-ADDomainController -Filter * -ErrorAction Stop | Group-Object Site
        $data = @()
        foreach ($siteGroup in $dcSites) {
            $data += @{
                "Site" = $siteGroup.Name
                "DC_Count" = ($siteGroup.Group | Measure-Object).Count
                "DCs" = ($siteGroup.Group.Name -join ", ")
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No site information available"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED DNS FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-DNSResourceRecords {
    try {
        $zones = Get-DnsServerZone -ErrorAction Stop | Select-Object -ExpandProperty ZoneName
        $records = @()
        foreach ($zone in $zones | Select-Object -First 10) {
            $zoneRecords = Get-DnsServerResourceRecord -ZoneName $zone -ErrorAction SilentlyContinue | Select-Object -First 30
            $records += $zoneRecords | Select-Object -Property Name, RecordType, @{Name='TTL';Expression={$_.TTL}}, @{Name='Zone';Expression={$zone}} | Select-Object -First 50
        }
        if ($null -eq $records -or $records.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No DNS resource records found"})
        }
        return $records | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "Type" = $_.RecordType; "TTL" = $_.TTL; "Zone" = $_.Zone} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DNSServerStatistics {
    try {
        $stats = Get-DnsServerStatistics -ErrorAction Stop
        if ($null -eq $stats) {
            return @([PSCustomObject]@{"Information" = "No DNS statistics available"})
        }
        return @(
            [PSCustomObject]@{"Statistic" = "Queries (Total)"; "Value" = $stats.Stats.TotalQueries},
            [PSCustomObject]@{"Statistic" = "Responses (Total)"; "Value" = $stats.Stats.TotalResponses},
            [PSCustomObject]@{"Statistic" = "Cache Hits"; "Value" = $stats.Stats.CacheHits},
            [PSCustomObject]@{"Statistic" = "Cache Misses"; "Value" = $stats.Stats.CacheMisses}
        )
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DNSServerScavengingSettings {
    try {
        $scavenging = Get-DnsServerScavenging -ErrorAction Stop
        if ($null -eq $scavenging) {
            return @([PSCustomObject]@{"Information" = "No scavenging settings found"})
        }
        return @(
            [PSCustomObject]@{"Setting" = "ScavengingInterval"; "Value" = $scavenging.ScavengingInterval},
            [PSCustomObject]@{"Setting" = "RefreshInterval"; "Value" = $scavenging.RefreshInterval},
            [PSCustomObject]@{"Setting" = "NoRefreshInterval"; "Value" = $scavenging.NoRefreshInterval},
            [PSCustomObject]@{"Setting" = "ScavengingState"; "Value" = $scavenging.ScavengingState}
        )
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DNSSecSettings {
    try {
        $dnssec = Get-DnsServerDnsSec -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $dnssec) {
            return @([PSCustomObject]@{"Information" = "No DNSSEC settings found"})
        }
        return $dnssec | ForEach-Object { [PSCustomObject]@{"Zone" = $_.ZoneName; "DNSSEC" = $_.DnsSecState; "KSKCount" = $_.KSKCount; "ZSKCount" = $_.ZSKCount} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# EXTENDED DHCP FUNCTIONS (ADDITIONAL)
# ============================================================================

function Get-DHCPServerInformation {
    try {
        $servers = Get-DhcpServerInDC -ErrorAction Stop
        if ($null -eq $servers) {
            return @([PSCustomObject]@{"Information" = "No DHCP servers found in AD"})
        }
        return $servers | ForEach-Object { [PSCustomObject]@{"Server" = $_.ToString()} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPServerLeases {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        $leases = @()
        foreach ($scope in $scopes | Select-Object -First 5) {
            $scopeLeases = Get-DhcpServerv4Lease -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue | Select-Object -First 50
            $leases += $scopeLeases
        }
        if ($null -eq $leases -or $leases.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No active DHCP leases found"})
        }
        return $leases | ForEach-Object { [PSCustomObject]@{"IP" = $_.IPAddress; "Scope" = $_.ScopeId; "Client" = $_.HostName; "LeaseExpires" = $_.LeaseExpiryTime} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPServerOptions {
    try {
        $options = Get-DhcpServerv4OptionValue -ErrorAction Stop | Select-Object OptionId, Name, Value | Sort-Object OptionId | Select-Object -First 50
        if ($null -eq $options) {
            return @([PSCustomObject]@{"Information" = "No DHCP server options found"})
        }
        return $options | ForEach-Object { [PSCustomObject]@{"OptionID" = $_.OptionId; "Name" = $_.Name; "Value" = $_.Value} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPv4ScopeStatistics {
    try {
        $scopes = Get-DhcpServerv4Scope -ErrorAction Stop
        if ($null -eq $scopes) {
            return @([PSCustomObject]@{"Information" = "No DHCP scopes found"})
        }
        return $scopes | ForEach-Object { 
            $stats = Get-DhcpServerv4ScopeStatistics -ScopeId $_.ScopeId -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                "Scope" = $_.ScopeId
                "Name" = $_.Name
                "AddressesInUse" = $stats.AddressesInUse
                "AddressesAvailable" = $stats.AddressesAvailable
                "PercentageInUse" = $stats.PercentageInUse
            }
        }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DHCPServerAuthorization {
    try {
        $auth = Get-DhcpServerv4Failover -ErrorAction Stop | Select-Object Name, State, Mode, ServerRole | Select-Object -First 50
        if ($null -eq $auth) {
            return @([PSCustomObject]@{"Information" = "No DHCP failover configuration found"})
        }
        return $auth | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "State" = $_.State; "Mode" = $_.Mode; "Role" = $_.ServerRole} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# HTML EXPORT FUNCTIONS
# ============================================================================

function Export-AuditResultsToHtml {
    param(
        [string]$Title,
        [PSObject[]]$Data,
        [string]$Category
    )
    
    try {
        if ($null -eq $Data -or $Data.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No data to export!", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $false
        }
        
        # HTML Header with modern styling
        $htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WindowsServerAudit - $Title</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            padding: 20px;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
        }
        .header p {
            font-size: 14px;
            opacity: 0.9;
        }
        .content {
            padding: 30px;
        }
        .audit-info {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        .audit-info p {
            margin: 5px 0;
            font-size: 14px;
            line-height: 1.6;
        }
        .audit-info strong {
            color: #667eea;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        thead {
            background: #f3f4f6;
        }
        th {
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #374151;
            border-bottom: 2px solid #667eea;
        }
        td {
            padding: 12px;
            border-bottom: 1px solid #e5e7eb;
        }
        tr:nth-child(even) {
            background: #f9fafb;
        }
        tr:hover {
            background: #eff6ff;
        }
        .footer {
            background: #f3f4f6;
            padding: 20px 30px;
            text-align: center;
            font-size: 12px;
            color: #6b7280;
            border-top: 1px solid #e5e7eb;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>⚙️ WindowsServerAudit - Audit Report</h1>
            <p>Windows Server Audit Tool v0.4.0</p>
        </div>
        <div class="content">
            <div class="audit-info">
                <p><strong>Category:</strong> $Category</p>
                <p><strong>Check:</strong> $Title</p>
                <p><strong>Server:</strong> $(hostname)</p>
                <p><strong>Domain:</strong> $([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName)</p>
                <p><strong>Created:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
                <p><strong>Results:</strong> $($Data.Count) entries</p>
            </div>
"@

        # Convert data to HTML table
        $htmlTable = $Data | ConvertTo-Html -Fragment -As Table
        
        # HTML Footer
        $htmlFooter = @"
            $htmlTable
        </div>
        <div class="footer">
            <p>WindowsServerAudit © 2025 | Generated: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@

        # Create filename with timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeTitle = ($Title -replace '[^a-zA-Z0-9äöüÄÖÜß_]', '_').Substring(0, [Math]::Min(30, $Title.Length))
        $fileName = "AuditReport_${safeTitle}_${timestamp}.html"
        
        # Show SaveFileDialog
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*"
        $dialog.FileName = $fileName
        $dialog.DefaultExt = "html"
        $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $fullContent = $htmlHeader + "`r`n" + $htmlTable + "`r`n" + $htmlFooter
            $fullContent | Out-File -FilePath $dialog.FileName -Encoding UTF8 -Force
            
            [System.Windows.MessageBox]::Show("HTML export successful!`n`nFile: $($dialog.FileName)", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            # Optionally open HTML file in default browser
            try {
                Start-Process $dialog.FileName
            } catch { }
            
            return $true
        }
        
        return $false
    } catch {
        [System.Windows.MessageBox]::Show("HTML export error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# ============================================================================
# EVENT HANDLERS
# ============================================================================

# System Information Buttons
$window.FindName("ButtonSystemInfo").Add_Click({
    Update-Status "Loading System Information..."
    Update-Output (Get-SystemInformation)
    Update-Status "Done"
})

$window.FindName("ButtonOSInfo").Add_Click({
    Update-Status "Loading OS Details..."
    Update-Output (Get-OSDetails)
    Update-Status "Done"
})

$window.FindName("ButtonHardwareInfo").Add_Click({
    Update-Status "Loading Hardware Info..."
    Update-Output (Get-HardwareSummary)
    Update-Status "Done"
})

$window.FindName("ButtonCPUInfo").Add_Click({
    Update-Status "Loading CPU Details..."
    Update-Output (Get-CPUDetails)
    Update-Status "Done"
})

$window.FindName("ButtonMemoryInfo").Add_Click({
    Update-Status "Loading Memory Details..."
    Update-Output (Get-MemoryDetails)
    Update-Status "Done"
})

$window.FindName("ButtonStorageInfo").Add_Click({
    Update-Status "Loading Storage Info..."
    Update-Output (Get-StorageSummary)
    Update-Status "Done"
})

# Network Buttons
$window.FindName("ButtonNetConfig").Add_Click({
    Update-Status "Loading Network IP Configuration..."
    Update-Output (Get-NetworkConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonNetAdapters").Add_Click({
    Update-Status "Loading Network Adapters..."
    Update-Output (Get-NetworkAdapters)
    Update-Status "Done"
})

$window.FindName("ButtonTCPConnections").Add_Click({
    Update-Status "Loading Active Connections..."
    Update-Output (Get-ActiveConnections)
    Update-Status "Done"
})

$window.FindName("ButtonFirewallRules").Add_Click({
    Update-Status "Loading Firewall Rules..."
    Update-Output (Get-FirewallRules)
    Update-Status "Done"
})

# Services Buttons
$window.FindName("ButtonAutomaticServices").Add_Click({
    Update-Status "Loading Automatic Services..."
    Update-Output (Get-AutomaticServices)
    Update-Status "Done"
})

$window.FindName("ButtonRunningServices").Add_Click({
    Update-Status "Loading Running Services..."
    Update-Output (Get-RunningServices)
    Update-Status "Done"
})

$window.FindName("ButtonScheduledTasks").Add_Click({
    Update-Status "Loading Scheduled Tasks..."
    Update-Output (Get-ScheduledTasks)
    Update-Status "Done"
})

# Roles & Features Buttons
$window.FindName("ButtonInstalledFeatures").Add_Click({
    Update-Status "Loading Installed Features..."
    Update-Output (Get-InstalledFeatures)
    Update-Status "Done"
})

$window.FindName("ButtonInstalledPrograms").Add_Click({
    Update-Status "Loading Installed Programs..."
    Update-Output (Get-InstalledPrograms)
    Update-Status "Done"
})

$window.FindName("ButtonWindowsUpdates").Add_Click({
    Update-Status "Loading Windows Updates..."
    Update-Output (Get-WindowsUpdates)
    Update-Status "Done"
})

# Event Logs Buttons
$window.FindName("ButtonSystemEvents").Add_Click({
    Update-Status "Loading System Events..."
    Update-Output (Get-SystemEvents)
    Update-Status "Done"
})

$window.FindName("ButtonAppEvents").Add_Click({
    Update-Status "Loading Application Events..."
    Update-Output (Get-ApplicationEvents)
    Update-Status "Done"
})

$window.FindName("ButtonSecurityEvents").Add_Click({
    Update-Status "Loading Security Events..."
    Update-Output (Get-SecurityEvents)
    Update-Status "Done"
})

$window.FindName("ButtonFailedLogons").Add_Click({
    Update-Status "Loading Failed Logons..."
    Update-Output (Get-FailedLogons)
    Update-Status "Done"
})

$window.FindName("ButtonAccountLockouts").Add_Click({
    Update-Status "Loading Account Lockouts..."
    Update-Output (Get-AccountLockouts)
    Update-Status "Done"
})

# Security & Users Buttons
$window.FindName("ButtonLocalUsers").Add_Click({
    Update-Status "Loading Local Users..."
    Update-Output (Get-LocalUsers)
    Update-Status "Done"
})

$window.FindName("ButtonLocalGroups").Add_Click({
    Update-Status "Loading Local Groups..."
    Update-Output (Get-LocalGroups)
    Update-Status "Done"
})

$window.FindName("ButtonPrivilegeAudit").Add_Click({
    Update-Status "Loading Privilege Audit..."
    Update-Output (Get-PrivilegeAudit)
    Update-Status "Done"
})

# Active Directory Buttons
$window.FindName("ButtonADDC").Add_Click({
    Update-Status "Loading AD Domain Controllers..."
    Update-Output (Get-ADDomainControllers)
    Update-Status "Done"
})

$window.FindName("ButtonADDomain").Add_Click({
    Update-Status "Loading AD Domain Info..."
    Update-Output (Get-ADDomainInfo)
    Update-Status "Done"
})

$window.FindName("ButtonADForest").Add_Click({
    Update-Status "Loading AD Forest Info..."
    Update-Output (Get-ADForestInfo)
    Update-Status "Done"
})

$window.FindName("ButtonADOUs").Add_Click({
    Update-Status "Loading AD OUs..."
    Update-Output (Get-ADOUs)
    Update-Status "Done"
})

$window.FindName("ButtonADAdmins").Add_Click({
    Update-Status "Loading AD Admins..."
    Update-Output (Get-ADDomainAdmins)
    Update-Status "Done"
})

$window.FindName("ButtonADComputers").Add_Click({
    Update-Status "Loading AD Computers..."
    Update-Output (Get-ADComputers)
    Update-Status "Done"
})

$window.FindName("ButtonADReplStatus").Add_Click({
    Update-Status "Loading AD Replication Status..."
    Update-Output (Get-ADReplicationStatus)
    Update-Status "Done"
})

$window.FindName("ButtonADTrusts").Add_Click({
    Update-Status "Loading AD Trusts..."
    Update-Output (Get-ADTrusts)
    Update-Status "Done"
})

$window.FindName("ButtonADDCExtended").Add_Click({
    Update-Status "Loading AD DC Extended..."
    Update-Output (Get-ADDomainControllerExtended)
    Update-Status "Done"
})

$window.FindName("ButtonADFunctionalLevels").Add_Click({
    Update-Status "Loading AD Functional Levels..."
    Update-Output (Get-ADDomainFunctionalLevel)
    Update-Status "Done"
})

$window.FindName("ButtonADSites").Add_Click({
    Update-Status "Loading AD Sites..."
    Update-Output (Get-ADSiteConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonADGPO").Add_Click({
    Update-Status "Loading AD GPOs..."
    Update-Output (Get-ADGroupPolicySummary)
    Update-Status "Done"
})

$window.FindName("ButtonADClustering").Add_Click({
    Update-Status "Loading AD Clustering..."
    Update-Output (Get-ADClusterInformation)
    Update-Status "Done"
})

# DNS Buttons
$window.FindName("ButtonDNSConfig").Add_Click({
    Update-Status "Loading DNS Configuration..."
    Update-Output (Get-DNSConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonDNSZones").Add_Click({
    Update-Status "Loading DNS Zones..."
    Update-Output (Get-DNSZones)
    Update-Status "Done"
})

$window.FindName("ButtonDNSForwarders").Add_Click({
    Update-Status "Loading DNS Forwarders..."
    Update-Output (Get-DNSForwarders)
    Update-Status "Done"
})

$window.FindName("ButtonDNSCache").Add_Click({
    Update-Status "Loading DNS Cache..."
    Update-Output (Get-DNSCache)
    Update-Status "Done"
})

# DHCP Buttons
$window.FindName("ButtonDHCPConfig").Add_Click({
    Update-Status "Loading DHCP Configuration..."
    Update-Output (Get-DHCPConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonDHCPv4Scopes").Add_Click({
    Update-Status "Loading DHCP IPv4 Scopes..."
    Update-Output (Get-DHCPv4Scopes)
    Update-Status "Done"
})

$window.FindName("ButtonDHCPv6Scopes").Add_Click({
    Update-Status "Loading DHCP IPv6 Scopes..."
    Update-Output (Get-DHCPv6Scopes)
    Update-Status "Done"
})

$window.FindName("ButtonDHCPReservations").Add_Click({
    Update-Status "Loading DHCP Reservations..."
    Update-Output (Get-DHCPReservations)
    Update-Status "Done"
})

# IIS Buttons
$window.FindName("ButtonIISWebsites").Add_Click({
    Update-Status "Loading IIS Websites..."
    Update-Output (Get-IISWebsites)
    Update-Status "Done"
})

$window.FindName("ButtonIISAppPools").Add_Click({
    Update-Status "Loading IIS App Pools..."
    Update-Output (Get-IISAppPools)
    Update-Status "Done"
})

$window.FindName("ButtonIISBindings").Add_Click({
    Update-Status "Loading IIS Bindings..."
    Update-Output (Get-IISBindings)
    Update-Status "Done"
})

# RDS Buttons
$window.FindName("ButtonRDSCollections").Add_Click({
    Update-Status "Loading RDS Collections..."
    Update-Output (Get-RDSCollections)
    Update-Status "Done"
})

$window.FindName("ButtonRDSSessionHosts").Add_Click({
    Update-Status "Loading RDS Session Hosts..."
    Update-Output (Get-RDSSessionHosts)
    Update-Status "Done"
})

$window.FindName("ButtonRDSLicensing").Add_Click({
    Update-Status "Loading RDS Licensing..."
    Update-Output (Get-RDSActiveLicensing)
    Update-Status "Done"
})

# DFS Buttons
$window.FindName("ButtonDFSNamespaces").Add_Click({
    Update-Status "Loading DFS Namespaces..."
    Update-Output (Get-DFSNamespaces)
    Update-Status "Done"
})

$window.FindName("ButtonDFSReplication").Add_Click({
    Update-Status "Loading DFS Replication Groups..."
    Update-Output (Get-DFSReplicationGroups)
    Update-Status "Done"
})

# Print Server Buttons
$window.FindName("ButtonPrinters").Add_Click({
    Update-Status "Loading Printers..."
    Update-Output (Get-PrintServers)
    Update-Status "Done"
})

$window.FindName("ButtonPrinterDrivers").Add_Click({
    Update-Status "Loading Printer Drivers..."
    Update-Output (Get-PrinterDrivers)
    Update-Status "Done"
})

# WSUS Buttons
$window.FindName("ButtonWSUSConfig").Add_Click({
    Update-Status "Loading WSUS Configuration..."
    Update-Output (Get-WSUSConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonWSUSGroups").Add_Click({
    Update-Status "Loading WSUS Computer Target Groups..."
    Update-Output (Get-WSUSComputerTargetGroups)
    Update-Status "Done"
})

$window.FindName("ButtonWSUSUpdates").Add_Click({
    Update-Status "Loading WSUS Updates..."
    Update-Output (Get-WSUSUpdates)
    Update-Status "Done"
})

# Hyper-V Buttons
$window.FindName("ButtonHyperVVMs").Add_Click({
    Update-Status "Loading Hyper-V VMs..."
    Update-Output (Get-HyperVVirtualMachines)
    Update-Status "Done"
})

$window.FindName("ButtonHyperVSwitches").Add_Click({
    Update-Status "Loading Hyper-V Switches..."
    Update-Output (Get-HyperVSwitches)
    Update-Status "Done"
})

$window.FindName("ButtonHyperVSnapshots").Add_Click({
    Update-Status "Loading Hyper-V Snapshots..."
    Update-Output (Get-HyperVSnapshots)
    Update-Status "Done"
})

# NRAS/NPS Buttons
$window.FindName("ButtonNPASConfig").Add_Click({
    Update-Status "Loading NRAS/NPS Configuration..."
    Update-Output (Get-NPASConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonNASClients").Add_Click({
    Update-Status "Loading NAS Clients..."
    Update-Output (Get-NASClients)
    Update-Status "Done"
})

# KMS Buttons
$window.FindName("ButtonKMSConfig").Add_Click({
    Update-Status "Loading KMS Configuration..."
    Update-Output (Get-KMSConfiguration)
    Update-Status "Done"
})

# WDS Buttons
$window.FindName("ButtonWDSConfig").Add_Click({
    Update-Status "Loading WDS Configuration..."
    Update-Output (Get-WDSConfiguration)
    Update-Status "Done"
})

$window.FindName("ButtonWDSBootImages").Add_Click({
    Update-Status "Loading WDS Boot Images..."
    Update-Output (Get-WDSBootImages)
    Update-Status "Done"
})

$window.FindName("ButtonWDSInstallImages").Add_Click({
    Update-Status "Loading WDS Install Images..."
    Update-Output (Get-WDSInstallImages)
    Update-Status "Done"
})

# File Services Buttons
$window.FindName("ButtonFileShares").Add_Click({
    Update-Status "Loading File Shares..."
    Update-Output (Get-FileShares)
    Update-Status "Done"
})

$window.FindName("ButtonSharePermissions").Add_Click({
    Update-Status "Loading Share Permissions..."
    Update-Output (Get-FileSharePermissions)
    Update-Status "Done"
})

$window.FindName("ButtonFileQuotas").Add_Click({
    Update-Status "Loading File Quotas..."
    Update-Output (Get-FileServerQuotas)
    Update-Status "Done"
})

$window.FindName("ButtonShadowCopies").Add_Click({
    Update-Status "Loading Shadow Copies..."
    Update-Output (Get-ShadowCopies)
    Update-Status "Done"
})

# Export & Clear Buttons
$window.FindName("ButtonExportCSV").Add_Click({
    try {
        if ($null -eq $DataGridResults.ItemsSource -or $DataGridResults.ItemsSource.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No data to export!", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Display export options
        $result = [System.Windows.MessageBox]::Show("Choose export format:`n`nYes = CSV Export`nNo = HTML Export", "Export Format", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            # CSV Export
            $dialog = New-Object System.Windows.Forms.SaveFileDialog
            $dialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            $dialog.DefaultExt = "csv"
            $dialog.FileName = "WindowsServerAudit_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $dialog.FileName
                
                # Export data to CSV
                $DataGridResults.ItemsSource | Export-Csv -Path $exportPath -Encoding UTF8 -NoTypeInformation -Force
                
                Update-Status "CSV export successful: $exportPath"
                [System.Windows.MessageBox]::Show("CSV export completed successfully!`n`nFile: $exportPath", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        } else {
            # HTML Export
            $title = $StatusBarText.Text
            if ([string]::IsNullOrWhiteSpace($title)) { $title = "Audit Results" }
            
            Export-AuditResultsToHtml -Title $title -Data $DataGridResults.ItemsSource -Category "Audit"
        }
    } catch {
        Update-Status "Export error: $_"
        [System.Windows.MessageBox]::Show("Error during export: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$window.FindName("ButtonClearOutput").Add_Click({
    Clear-Output
    Update-Status "Output cleared"
})

# ============================================================================
# MAIN
# ============================================================================

$window.ShowDialog()

# ============================================================================
# EXTENDED CERTIFICATE FUNCTIONS
# ============================================================================

function Get-InstalledCertificates {
    try {
        $certs = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $certs) {
            return @([PSCustomObject]@{"Information" = "No certificates found"})
        }
        $data = @()
        foreach ($cert in $certs) {
            $data += @{
                "Thumbprint" = $cert.Thumbprint.Substring(0, 16)
                "Subject" = $cert.Subject
                "Issuer" = $cert.Issuer
                "ValidFrom" = $cert.NotBefore
                "ValidTo" = $cert.NotAfter
                "DaysLeft" = ($cert.NotAfter - (Get-Date)).Days
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-CertificateAuthorities {
    try {
        $cas = Get-ChildItem -Path Cert:\LocalMachine\Root -ErrorAction Stop | Select-Object -First 50
        if ($null -eq $cas) {
            return @([PSCustomObject]@{"Information" = "No CAs found"})
        }
        $data = @()
        foreach ($ca in $cas) {
            $data += @{
                "CA-Name" = $ca.Subject
                "Thumbprint" = $ca.Thumbprint.Substring(0, 16)
                "ValidFrom" = $ca.NotBefore
                "ValidTo" = $ca.NotAfter
                "Issuer" = $ca.Issuer
            }
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ExpiringCertificates {
    try {
        $certs = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction Stop
        $expiring = @()
        foreach ($cert in $certs) {
            $daysLeft = ($cert.NotAfter - (Get-Date)).Days
            if ($daysLeft -le 90 -and $daysLeft -ge 0) {
                $expiring += @{
                    "Subject" = $cert.Subject
                    "DaysLeft" = $daysLeft
                    "ExpiresOn" = $cert.NotAfter
                    "Status" = if ($daysLeft -le 30) {"CRITICAL"} else {"WARNING"}
                }
            }
        }
        if ($expiring.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "No expiring certificates in the next 90 days"})
        }
        return $expiring | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

# ============================================================================
# MICROSOFT SERVER ROLES FUNCTIONS
# ============================================================================

function Get-InstalledRoles {
    try {
        $roles = Get-WindowsFeature -ErrorAction Stop | Where-Object { $_.Installed -eq $true } | Select-Object Name, DisplayName, FeatureType
        if ($null -eq $roles) {
            return @([PSCustomObject]@{"Information" = "No roles/features installed"})
        }
        return $roles | ForEach-Object { [PSCustomObject]@{"Name" = $_.Name; "DisplayName" = $_.DisplayName; "Type" = $_.FeatureType} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-SQLServerInfo {
    try {
        $sqlServices = Get-Service -Name MSSQLSERVER -ErrorAction SilentlyContinue
        if ($null -eq $sqlServices) {
            return @([PSCustomObject]@{"Information" = "SQL Server not installed"})
        }
        
        $data = @(
            @{"Service" = "MSSQLSERVER"; "Status" = $sqlServices.Status},
            @{"Service" = "SQL Server Agent"; "Status" = (Get-Service -Name SQLSERVERAGENT -ErrorAction SilentlyContinue).Status},
            @{"Service" = "SQL Browser"; "Status" = (Get-Service -Name SQLBrowser -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ExchangeServerInfo {
    try {
        $exService = Get-Service -Name MSExchangeServiceHost -ErrorAction SilentlyContinue
        if ($null -eq $exService) {
            return @([PSCustomObject]@{"Information" = "Exchange Server not installed"})
        }
        
        $data = @(
            @{"Component" = "Exchange Service Host"; "Status" = $exService.Status},
            @{"Component" = "Information Store"; "Status" = (Get-Service -Name MSExchangeIS -ErrorAction SilentlyContinue).Status},
            @{"Component" = "Transport"; "Status" = (Get-Service -Name MSExchangeTransport -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-SharePointInfo {
    try {
        $spService = Get-Service -Name SPAdminV4 -ErrorAction SilentlyContinue
        if ($null -eq $spService) {
            return @([PSCustomObject]@{"Information" = "SharePoint not installed"})
        }
        
        $data = @(
            @{"Service" = "SP Admin"; "Status" = $spService.Status},
            @{"Service" = "SP Timer"; "Status" = (Get-Service -Name SPTimerV4 -ErrorAction SilentlyContinue).Status}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-DomainControllerHealth {
    try {
        $dcHealth = dcdiag /v 2>&1 | Select-Object -First 50
        if ($null -eq $dcHealth) {
            return @([PSCustomObject]@{"Information" = "No DC health information available"})
        }
        $data = @()
        foreach ($line in $dcHealth) {
            if ($line -match "passed|failed") {
                $data += @{"Test" = $line}
            }
        }
        if ($data.Count -eq 0) {
            return @([PSCustomObject]@{"Information" = "DC health check performed"})
        }
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-HyperVInfo {
    try {
        $hvService = Get-Service -Name vmms -ErrorAction SilentlyContinue
        if ($null -eq $hvService) {
            return @([PSCustomObject]@{"Information" = "Hyper-V not installed"})
        }
        
        $vms = Get-VM -ErrorAction SilentlyContinue | Measure-Object
        $data = @(
            @{"Component" = "Hyper-V Service"; "Status" = $hvService.Status},
            @{"Component" = "VMs installed"; "Status" = $vms.Count}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-ServerUpdates {
    try {
        $updates = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 20
        if ($null -eq $updates) {
            return @([PSCustomObject]@{"Information" = "No updates installed"})
        }
        return $updates | ForEach-Object { [PSCustomObject]@{"KB" = $_.HotFixID; "InstalledOn" = $_.InstalledOn; "Description" = $_.Description} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-NetworkSecurity {
    try {
        $firewallEnabled = Get-NetFirewallProfile -All -ErrorAction Stop | Select-Object Name, Enabled
        if ($null -eq $firewallEnabled) {
            return @([PSCustomObject]@{"Information" = "Firewall status not available"})
        }
        return $firewallEnabled | ForEach-Object { [PSCustomObject]@{"Profile" = $_.Name; "Enabled" = if($_.Enabled) {"Yes"} else {"No"}} }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-WindowsDefender {
    try {
        $defender = Get-Service -Name WinDefend -ErrorAction SilentlyContinue
        $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        
        if ($null -eq $defender) {
            return @([PSCustomObject]@{"Information" = "Windows Defender not available"})
        }
        
        $data = @(
            @{"Component" = "Service Status"; "Status" = $defender.Status},
            @{"Component" = "Real-time Protection"; "Status" = if($defenderStatus.RealTimeProtectionEnabled) {"Enabled"} else {"Disabled"}},
            @{"Component" = "Definitions Update"; "Status" = $defenderStatus.AntivirusSignatureLastUpdated}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-UserAccountControl {
    try {
        $uac = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue
        $data = @(
            @{"Setting" = "User Account Control"; "Status" = if($uac.EnableLUA -eq 1) {"Activated"} else {"Deactivated"}}
        )
        return $data | ForEach-Object { [PSCustomObject]$_ }
    } catch {
        return @([PSCustomObject]@{"Error" = $_.Exception.Message})
    }
}

function Get-SystemBackupConfig {
    try {
        $backup = Get-WBBackupSet -ErrorAction SilentlyContinue | Select-Object -First 5
        if ($null -eq $backup) {
            return @([PSCustomObject]@{"Information" = "No backups configured or Windows Backup is not active"})
        }
        return $backup | ForEach-Object { [PSCustomObject]@{"BackupSetId" = $_.BackupSetId; "BackupTime" = $_.BackupTime; "Items" = ($_.Items.Count)} }
    } catch {
        return @([PSCustomObject]@{"Information" = "Windows Backup not available - using other backup solution"})
    }
}
