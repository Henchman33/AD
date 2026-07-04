<#
================================================================================
    AD Reporting Tool - Active Directory Reporting Tool
    Version 1.2 PROD
    
    Beschreibung:
    Umfassendes Reporting-Tool für Active Directory mit moderner WPF-GUI.
    Bietet detaillierte Analysen von Benutzern, Gruppen, Computern, OUs,
    GPOs und weiteren AD-Objekten mit Export-Funktionen (CSV, HTML).
    
================================================================================

VORAUSSETZUNGEN:
    - Windows PowerShell 5.1
    - Active Directory PowerShell Module (RSAT-AD-PowerShell)

================================================================================
#>

[xml]$Global:XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="AD Reporting Tool Version 1.2 PRO Series"
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

        <!--  Modern Button Style  -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#D1D5DB" />
            <Setter Property="Padding" Value="16,10" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border
                            x:Name="border"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="4">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F3F4F6" />
                                <Setter TargetName="border" Property="BorderBrush" Value="#6366F1" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E5E7EB" />
                                <Setter TargetName="border" Property="BorderBrush" Value="#4F46E5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
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

        <!--  Expandable Section Style - Simplified  -->
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
                        Text="AD Reporting Tool Version 1.2 PRO Series" />
                </StackPanel>

                <!--  Results Counter  -->
                <StackPanel
                    Grid.Column="1"
                    VerticalAlignment="Center"
                    Orientation="Vertical" Background="#FF4C5B73" Height="49" Margin="1475,0,-20,0" Grid.ColumnSpan="2">
                    <TextBlock Text="Total Results" FontSize="14" Foreground="#FFC1FFDB" HorizontalAlignment="Center"/>
                    <TextBlock x:Name="TotalResultCountText" Text="0" FontSize="20" FontWeight="Bold" 
                               Foreground="#FFEAEAEA" HorizontalAlignment="Center" Margin="0,2,0,0" Width="60" TextAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>

        <!--  Content Area (Grid.Row="1")  -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320" />
                <!--  Enhanced Sidebar  -->
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
                                Text="📊 Quick Reports" />
                        </Grid>

                        <!--  Domain Overview Reports  -->
                        <Expander
                            Header="📊 Domain Overview"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickAllUsers"
                                    Content="All Users"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Retrieve all user accounts in the domain" />
                                <Button
                                    x:Name="ButtonQuickGroups"
                                    Content="All Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Retrieve all groups in the domain" />
                                <Button
                                    x:Name="ButtonQuickComputers"
                                    Content="All Computers"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Retrieve all computer accounts" />
                                <Button
                                    x:Name="ButtonQuickOSSummary"
                                    Content="Operating System Summary"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Summary of OS versions in use" />
                                <Button
                                    x:Name="ButtonQuickOUHierarchy"
                                    Content="OU Hierarchy"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show Organizational Unit hierarchy" />
                                <Button
                                    x:Name="ButtonQuickSitesSubnets"
                                    Content="Sites and Subnets"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Display sites and subnets" />
                            </StackPanel>
                        </Expander>

                        <!--  Statistics and Reports  -->
                        <Expander
                            Header="📈 Statistics and Reports"
                            IsExpanded="True"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickSecurityDashboard"
                                    Content="Security Assessment Dashboard"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Overall AD security assessment with recommendations" />
                                <Button
                                    x:Name="ButtonQuickCompromiseIndicators"
                                    Content="AD Compromise Indicators"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Signs of compromise and suspicious activities" />
                                <Button
                                    x:Name="ButtonQuickDepartmentStats"
                                    Content="Department Statistics"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Generate department statistics" />
                                <Button
                                    x:Name="ButtonQuickDepartmentSecurity"
                                    Content="Department Security"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze department security posture" />
                                <Button
                                    x:Name="ButtonQuickOrgChart"
                                    Content="Organization Chart"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Export organizational hierarchy and structure" />
                            </StackPanel>
                        </Expander>

                        <!--  User Account Management  -->
                        <Expander
                            Header="👤 User Accounts"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickDisabledUsers"
                                    Content="Disabled Users"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find all disabled user accounts" />
                                <Button
                                    x:Name="ButtonQuickInactiveUsers"
                                    Content="Inactive Users (90+ days)"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users inactive for more than 90 days" />
                                <Button
                                    x:Name="ButtonQuickLockedUsers"
                                    Content="Locked Users"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show accounts currently locked out" />
                                <Button
                                    x:Name="ButtonQuickNeverLoggedOn"
                                    Content="Never Logged On"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users who have never logged into the system" />
                                <Button
                                    x:Name="ButtonQuickExpiringAccounts"
                                    Content="Expiring Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Accounts approaching expiration date" />
                                <Button
                                    x:Name="ButtonQuickRecentlyCreatedUsers"
                                    Content="Recently Created (30d)"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="New user accounts created within 30 days" />
                                <Button
                                    x:Name="ButtonQuickRecentlyDeletedUsers"
                                    Content="Recently Deleted"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Recently deleted user accounts" />
                                <Button
                                    x:Name="ButtonQuickRecentlyModifiedUsers"
                                    Content="Recently Modified"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="User accounts with recent changes" />
                                <Button
                                    x:Name="ButtonQuickUsersByDepartment"
                                    Content="Users by Department"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Group users by their department" />
                                <Button
                                    x:Name="ButtonQuickUsersByManager"
                                    Content="Users by Manager"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show users organized by manager" />
                                <Button
                                    x:Name="ButtonQuickUsersWithoutManager"
                                    Content="Users without Manager"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with no assigned manager" />
                                <Button
                                    x:Name="ButtonQuickUsersMissingRequiredAttributes"
                                    Content="Missing Required Attributes"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with incomplete profile information" />
                                <Button
                                    x:Name="ButtonQuickGuestAccountStatus"
                                    Content="Guest Account Status"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check status of guest accounts" />
                                <Button
                                    x:Name="ButtonQuickLoginTimeAnalysis"
                                    Content="Login Time Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze unusual login times and patterns" />
                                <Button
                                    x:Name="ButtonQuickAccountCreationTimeline"
                                    Content="Account Creation Timeline"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Timeline of account creations with creator info" />
                                <Button
                                    x:Name="ButtonQuickDeletedObjectsRecovery"
                                    Content="Deleted Objects Recovery"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Recently deleted AD objects and recovery options" />
                                <Button
                                    x:Name="ButtonQuickProxyAddressChanges"
                                    Content="ProxyAddress Changes"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Track changes to ProxyMailAddress attributes" />
                            </StackPanel>
                        </Expander>

                        <!--  Password Security  -->
                        <Expander
                            Header="🔐 User Security"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickNeverExpire"
                                    Content="Password Never Expires"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with non-expiring passwords" />
                                <Button
                                    x:Name="ButtonQuickPasswordExpiringSoon"
                                    Content="Password Expiring (7d)"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Passwords expiring within 7 days" />
                                <Button
                                    x:Name="ButtonQuickExpiredPasswords"
                                    Content="Expired Passwords"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with expired passwords" />
                                <Button
                                    x:Name="ButtonQuickStalePasswords"
                                    Content="Stale Passwords"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with very old passwords" />
                                <Button
                                    x:Name="ButtonQuickNeverChangingPasswords"
                                    Content="Never Changing Passwords"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users who never changed their password" />
                                <Button
                                    x:Name="ButtonQuickReversibleEncryption"
                                    Content="Reversible Encryption"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users with reversible password encryption" />
                                <Button
                                    x:Name="ButtonQuickWeakPasswordPolicy"
                                    Content="Weak Password Policies"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze weak password policies" />
                                <Button
                                    x:Name="ButtonQuickPasswordPolicySummary"
                                    Content="Password Policy Summary"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Summary of password policies" />
                                <Button
                                    x:Name="ButtonQuickAccountLockoutPolicies"
                                    Content="Account Lockout Policies"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Review account lockout policies" />
                                <Button
                                    x:Name="ButtonQuickFineGrainedPasswordPolicies"
                                    Content="Fine-Grained Password Policies"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show fine-grained password policies" />
                            </StackPanel>
                        </Expander>

                        <!--  Group Management  -->
                        <Expander
                            Header="👥 Group Management"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickEmptyGroups"
                                    Content="Empty Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups with no members" />
                                <Button
                                    x:Name="ButtonQuickSecurityGroups"
                                    Content="Security Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show only security groups" />
                                <Button
                                    x:Name="ButtonQuickDistributionGroups"
                                    Content="Distribution Lists"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show only distribution groups" />
                                <Button
                                    x:Name="ButtonQuickGroupsByTypeScope"
                                    Content="Groups by Type/Scope"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Categorize groups by type and scope" />
                                <Button
                                    x:Name="ButtonQuickMailEnabledGroups"
                                    Content="Mail Enabled Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups enabled for email" />
                                <Button
                                    x:Name="ButtonQuickDynamicDistGroups"
                                    Content="Dynamic Distribution Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Dynamic distribution groups" />
                                <Button
                                    x:Name="ButtonQuickNestedGroups"
                                    Content="Nested Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups containing other groups" />
                                <Button
                                    x:Name="ButtonQuickCircularGroups"
                                    Content="Circular References"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups with circular membership" />
                                <Button
                                    x:Name="ButtonQuickGroupsWithoutOwners"
                                    Content="Groups without Owners"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups with no assigned owners" />
                                <Button
                                    x:Name="ButtonQuickLargeGroups"
                                    Content="Large Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups with many members" />
                                <Button
                                    x:Name="ButtonQuickRecentlyModifiedGroups"
                                    Content="Recently Modified Groups"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Groups with recent changes" />
                            </StackPanel>
                        </Expander>

                        <!--  Computer Management  -->
                        <Expander
                            Header="💻 Computer Management"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickInactiveComputers"
                                    Content="Inactive Computers (90+ days)"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Computers inactive for more than 90 days" />
                                <Button
                                    x:Name="ButtonQuickComputersNeverLoggedOn"
                                    Content="Computers Never Logged On"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Computer accounts that never logged on" />
                                <Button
                                    x:Name="ButtonQuickComputersByLocation"
                                    Content="Computers by Location"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Group computers by location" />
                                <Button
                                    x:Name="ButtonQuickComputersByOSVersion"
                                    Content="Computers by OS Version"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Group computers by OS version" />
                                <Button
                                    x:Name="ButtonQuickDuplicateComputerNames"
                                    Content="Duplicate Computer Names"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find duplicate computer names" />
                                <Button
                                    x:Name="ButtonQuickBitLockerStatus"
                                    Content="BitLocker Status"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check BitLocker encryption status" />
                            </StackPanel>
                        </Expander>

                        <!--  Privileged Accounts and Security  -->
                        <Expander
                            Header="🛡️ Privileged Accounts"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickAdminUsers"
                                    Content="Administrators"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show all administrative accounts" />
                                <Button
                                    x:Name="ButtonQuickPrivilegedAccounts"
                                    Content="Privileged Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show all privileged accounts" />
                                <Button
                                    x:Name="ButtonQuickServiceAccountsOverview"
                                    Content="Service Accounts Overview"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Overview of all service accounts" />
                                <Button
                                    x:Name="ButtonQuickManagedServiceAccounts"
                                    Content="Managed Service Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show managed service accounts (MSA)" />
                                <Button
                                    x:Name="ButtonQuickServiceAccountsSPN"
                                    Content="Service Accounts with SPN"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Service accounts with Service Principal Names" />
                                <Button
                                    x:Name="ButtonQuickHighPrivServiceAccounts"
                                    Content="High Privilege Service Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Service accounts with high privileges" />
                                <Button
                                    x:Name="ButtonQuickServiceAccountPasswordAge"
                                    Content="Service Account Password Age"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check service account password ages" />
                                <Button
                                    x:Name="ButtonQuickUnusedServiceAccounts"
                                    Content="Unused Service Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Unused or obsolete service accounts" />
                                <Button
                                    x:Name="ButtonQuickRiskyGroupMemberships"
                                    Content="Risky Memberships"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find risky group memberships" />
                                <Button
                                    x:Name="ButtonQuickKerberosDES"
                                    Content="Kerberos DES Users"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Users using weak DES encryption" />
                                <Button
                                    x:Name="ButtonQuickUsersWithSPN"
                                    Content="Users with SPN"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="User accounts with Service Principal Names" />
                                <Button
                                    x:Name="ButtonQuickUsersDuplicateLogonNames"
                                    Content="Duplicate Logon Names"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find duplicate user logon names" />
                                <Button
                                    x:Name="ButtonQuickOrphanedSIDsUsers"
                                    Content="Orphaned SIDs"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="User accounts with orphaned SIDs" />
                            </StackPanel>
                        </Expander>

                        <!--  Security Threats Analysis  -->
                        <Expander
                            Header="🔍 Security Threats"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickKerberoastable"
                                    Content="Kerberoastable Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Accounts vulnerable to Kerberoasting" />
                                <Button
                                    x:Name="ButtonQuickASREPRoastable"
                                    Content="ASREPRoastable Accounts"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Accounts vulnerable to ASREPRoasting" />
                                <Button
                                    x:Name="ButtonQuickHoneyTokens"
                                    Content="Honey Tokens"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Identify honey token accounts" />
                                <Button
                                    x:Name="ButtonQuickPrivilegeEscalation"
                                    Content="Privilege Escalation Paths"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find privilege escalation paths" />
                                <Button
                                    x:Name="ButtonQuickExposedCredentials"
                                    Content="Exposed Credentials"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check for exposed credentials" />
                                <Button
                                    x:Name="ButtonQuickSuspiciousLogons"
                                    Content="Suspicious Logons"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Detect suspicious logon patterns" />
                                <Button
                                    x:Name="ButtonQuickForeignSecurityPrincipals"
                                    Content="Foreign Security Principals"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find foreign security principals" />
                                <Button
                                    x:Name="ButtonQuickSIDHistoryAbuse"
                                    Content="SID History Abuse"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Detect SID history abuse" />
                                <Button
                                    x:Name="ButtonQuickAuthProtocolAnalysis"
                                    Content="Authentication Protocol Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Breakdown of NTLM vs. Kerberos usage" />
                                <Button
                                    x:Name="ButtonQuickFailedAuthPatterns"
                                    Content="Failed Authentication Patterns"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Advanced analysis of failed login attempts" />
                            </StackPanel>
                        </Expander>

                        <!--  Permissions and Access Control  -->
                        <Expander
                            Header="🔑 Permissions and Access"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickACLAnalysis"
                                    Content="ACL Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze Access Control Lists" />
                                <Button
                                    x:Name="ButtonQuickInheritanceBreaks"
                                    Content="Inheritance Breaks"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find ACL inheritance breaks" />
                                <Button
                                    x:Name="ButtonQuickAdminSDHolderObjects"
                                    Content="AdminSDHolder Objects"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show AdminSDHolder protected objects" />
                                <Button
                                    x:Name="ButtonQuickAdvancedDelegation"
                                    Content="Advanced Delegation"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Advanced delegation analysis" />
                                <Button
                                    x:Name="ButtonQuickDelegation"
                                    Content="Delegation Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="General delegation analysis" />
                                <Button
                                    x:Name="ButtonQuickSchemaPermissions"
                                    Content="Schema Permissions"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check schema permissions" />
                                <Button
                                    x:Name="ButtonQuickDCSyncRights"
                                    Content="DCSync Rights Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze DCSync rights" />
                                <Button
                                    x:Name="ButtonQuickSchemaAdmins"
                                    Content="Schema Admin Paths"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show schema admin privilege paths" />
                            </StackPanel>
                        </Expander>

                        <!--  Compliance and Audit  -->
                        <Expander
                            Header="📜 Compliance and Audit"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickDSGVOCompliance"
                                    Content="DSGVO/GDPR Compliance"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="DSGVO compliance report with sensitive data access" />
                            </StackPanel>
                        </Expander>

                        <!--  Policies and Group Policy Objects  -->
                        <Expander
                            Header="📋 Policies and GPOs"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickGPOOverview"
                                    Content="GPO Overview"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="General Group Policy Objects overview" />
                                <Button
                                    x:Name="ButtonQuickUnlinkedGPOs"
                                    Content="Unlinked GPOs"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Find unlinked Group Policy Objects" />
                                <Button
                                    x:Name="ButtonQuickEmptyGPOs"
                                    Content="Empty GPOs"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show empty Group Policy Objects" />
                                <Button
                                    x:Name="ButtonQuickGPOPermissions"
                                    Content="GPO Permissions"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze GPO permissions" />
                                <Button
                                    x:Name="ButtonQuickConditionalAccessPolicies"
                                    Content="Conditional Access Policies"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Report on conditional access configurations" />
                                <Button
                                    x:Name="ButtonQuickGPOConflictAnalysis"
                                    Content="GPO Conflict Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Detect conflicting GPO settings and priorities" />
                                <Button
                                    x:Name="ButtonQuickGPOSecurityBaseline"
                                    Content="GPO Security Baseline"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Compare GPOs against security best practices" />
                            </StackPanel>
                        </Expander>

                        <!--  AD Infrastructure and Health  -->
                        <Expander
                            Header="🏥 AD Infrastructure"
                            IsExpanded="False"
                            Style="{StaticResource ExpanderStyle}">
                            <StackPanel>
                                <Button
                                    x:Name="ButtonQuickFSMORoles"
                                    Content="FSMO Role Holders"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show FSMO role holders" />
                                <Button
                                    x:Name="ButtonQuickDCStatus"
                                    Content="Domain Controller Status"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check domain controller status" />
                                <Button
                                    x:Name="ButtonQuickReplicationStatus"
                                    Content="Replication Status"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check replication status" />
                                <Button
                                    x:Name="ButtonQuickSYSVOLHealth"
                                    Content="SYSVOL Health Check"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze SYSVOL health" />
                                <Button
                                    x:Name="ButtonQuickDNSHealth"
                                    Content="DNS Health Analysis"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check DNS health and configuration" />
                                <Button
                                    x:Name="ButtonQuickBackupStatus"
                                    Content="Backup Readiness"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check backup readiness status" />
                                <Button
                                    x:Name="ButtonQuickTrustRelationships"
                                    Content="Trust Relationships"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Show trust relationships" />
                                <Button
                                    x:Name="ButtonQuickSchemaAnalysis"
                                    Content="Schema Extensions"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze Active Directory schema extensions" />
                                <Button
                                    x:Name="ButtonQuickCertificateAnalysis"
                                    Content="Certificate Security"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Analyze certificate security" />
                                <Button
                                    x:Name="ButtonQuickQuotasLimits"
                                    Content="Quotas and Limits"
                                    Style="{StaticResource SidebarButton}"
                                    ToolTip="Check directory quotas and limits" />
                            </StackPanel>
                        </Expander>
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!--  Main Content Area  -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <!--  Filter Section  -->
                    <RowDefinition Height="20" />
                    <!--  Spacing  -->
                    <RowDefinition Height="*" />
                    <!--  Results  -->
                </Grid.RowDefinitions>

                <!--  Enhanced Filter Section  -->
                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition />
                        <ColumnDefinition Width="25" />
                        <ColumnDefinition Width="Auto" MinWidth="435" />
                        <ColumnDefinition Width="20" />
                        <ColumnDefinition Width="Auto" />
                    </Grid.ColumnDefinitions>

                    <!--  Advanced Filter Card  -->
                    <Border
                        Grid.Column="0"
                        Padding="20,16"
                        Background="#FFFFFF"
                        Style="{StaticResource ModernCard}">
                        <StackPanel>
                            <TextBlock Margin="0,0,0,12" Style="{StaticResource CategoryHeader}" Text="Search with Filter" />

                            <!--  Object Type Selection  -->
                            <StackPanel
                                Margin="0,0,0,12"
                                Background="#F9FAFB"
                                Orientation="Horizontal">
                                <TextBlock
                                    Width="90"
                                    VerticalAlignment="Center"
                                    FontWeight="Medium"
                                    Text="Object Type:" />
                                <RadioButton
                                    x:Name="RadioButtonUser"
                                    Margin="12,0"
                                    VerticalAlignment="Center"
                                    Content="User"
                                    IsChecked="True" />
                                <RadioButton
                                    x:Name="RadioButtonGroup"
                                    Margin="12,0"
                                    VerticalAlignment="Center"
                                    Content="Group" />
                                <RadioButton
                                    x:Name="RadioButtonComputer"
                                    Margin="12,0"
                                    VerticalAlignment="Center"
                                    Content="Computer" />
                                <RadioButton
                                    x:Name="RadioButtonGroupMemberships"
                                    Margin="12,0"
                                    VerticalAlignment="Center"
                                    Content="Memberships" />
                            </StackPanel>

                            <!--  Filter 1  -->
                            <Grid Margin="0,0,0,8" Background="#FFF9F9F9">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80" />
                                    <ColumnDefinition Width="140" />
                                    <ColumnDefinition Width="100" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    FontWeight="Medium"
                                    Text="Filter 1:" />
                                <ComboBox
                                    x:Name="ComboBoxFilterAttribute1"
                                    Grid.Column="1"
                                    Margin="8,0"
                                    VerticalAlignment="Center" />
                                <ComboBox
                                    x:Name="ComboBoxFilterOperator1"
                                    Grid.Column="2"
                                    Margin="8,0"
                                    VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True" />
                                    <ComboBoxItem Content="Equals" />
                                    <ComboBoxItem Content="StartsWith" />
                                    <ComboBoxItem Content="EndsWith" />
                                    <ComboBoxItem Content="NotEqual" />
                                </ComboBox>
                                <TextBox
                                    x:Name="TextBoxFilterValue1"
                                    Grid.Column="3"
                                    Margin="8,0,0,0"
                                    Padding="8,6"
                                    VerticalAlignment="Center" />
                            </Grid>

                            <!--  Logic Selector  -->
                            <StackPanel Margin="0,0,0,8" Orientation="Horizontal">
                                <TextBlock
                                    Width="80"
                                    VerticalAlignment="Center"
                                    FontWeight="Medium"
                                    Text="Logic:" />
                                <RadioButton
                                    x:Name="RadioButtonAnd"
                                    Margin="8,0"
                                    VerticalAlignment="Center"
                                    Content="AND"
                                    IsChecked="True" />
                                <RadioButton
                                    x:Name="RadioButtonOr"
                                    Margin="12,0"
                                    VerticalAlignment="Center"
                                    Content="OR" />
                                <CheckBox
                                    x:Name="CheckBoxUseSecondFilter"
                                    Margin="20,0"
                                    VerticalAlignment="Center"
                                    Content="Use second filter" />
                            </StackPanel>

                            <!--  Filter 2  -->
                            <Grid
                                x:Name="SecondFilterPanel"
                                Background="#FFF9F9F9"
                                IsEnabled="False">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80" />
                                    <ColumnDefinition Width="140" />
                                    <ColumnDefinition Width="100" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    FontWeight="Medium"
                                    Text="Filter 2:" />
                                <ComboBox
                                    x:Name="ComboBoxFilterAttribute2"
                                    Grid.Column="1"
                                    Margin="8,0"
                                    VerticalAlignment="Center" />
                                <ComboBox
                                    x:Name="ComboBoxFilterOperator2"
                                    Grid.Column="2"
                                    Margin="8,0"
                                    VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True" />
                                    <ComboBoxItem Content="Equals" />
                                    <ComboBoxItem Content="StartsWith" />
                                    <ComboBoxItem Content="EndsWith" />
                                    <ComboBoxItem Content="NotEqual" />
                                </ComboBox>
                                <TextBox
                                    x:Name="TextBoxFilterValue2"
                                    Grid.Column="3"
                                    Margin="8,0,0,0"
                                    Padding="8,6"
                                    VerticalAlignment="Center" />
                            </Grid>
                        </StackPanel>
                    </Border>

                    <!--  Attributes Selection Card  -->
                    <Border
                        Grid.Column="2"
                        Padding="20,16"
                        Background="#FFFFFF"
                        BorderBrush="#E5E7EB"
                        Style="{StaticResource ModernCard}">
                        <StackPanel>
                            <Grid Height="25" Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*" />
                                    <ColumnDefinition Width="Auto" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    Margin="0,0,0,4"
                                    Style="{StaticResource CategoryHeader}"
                                    Text="Export Attributes" />
                                <StackPanel Grid.Column="1" Orientation="Horizontal">
                                    <Button
                                        x:Name="ButtonSelectAllAttributes"
                                        Width="40"
                                        Height="25"
                                        Margin="0,0,4,0"
                                        Padding="6,4"
                                        Content="All"
                                        FontSize="10"
                                        Style="{StaticResource ModernButton}"
                                        ToolTip="Select all attributes" />
                                    <Button
                                        x:Name="ButtonSelectNoneAttributes"
                                        Width="40"
                                        Height="25"
                                        Padding="6,4"
                                        Content="None"
                                        FontSize="10"
                                        Style="{StaticResource ModernButton}"
                                        ToolTip="Deselect all attributes" />
                                </StackPanel>
                            </Grid>

                            <!--  Attribute Categories  -->
                            <TabControl
                                x:Name="TabControlAttributes"
                                Height="140"
                                Background="Transparent"
                                BorderThickness="0">
                                <TabItem FontSize="11" Header="Basic">
                                    <ListBox
                                        x:Name="ListBoxBasicAttributes"
                                        Background="White"
                                        BorderThickness="0"
                                        SelectionMode="Multiple">
                                        <ListBoxItem Content="DisplayName" IsSelected="True" />
                                        <ListBoxItem Content="SamAccountName" IsSelected="True" />
                                        <ListBoxItem Content="GivenName" />
                                        <ListBoxItem Content="Surname" />
                                        <ListBoxItem Content="mail" />
                                        <ListBoxItem Content="Department" />
                                        <ListBoxItem Content="Title" />
                                        <ListBoxItem Content="Enabled" IsSelected="True" />
                                    </ListBox>
                                </TabItem>
                                <TabItem FontSize="11" Header="Security">
                                    <ListBox
                                        x:Name="ListBoxSecurityAttributes"
                                        Background="Transparent"
                                        BorderThickness="0"
                                        SelectionMode="Multiple">
                                        <ListBoxItem Content="LastLogonTimestamp" />
                                        <ListBoxItem Content="PasswordExpired" />
                                        <ListBoxItem Content="PasswordLastSet" />
                                        <ListBoxItem Content="AccountExpirationDate" />
                                        <ListBoxItem Content="badPwdCount" />
                                        <ListBoxItem Content="lockoutTime" />
                                        <ListBoxItem Content="UserAccountControl" />
                                        <ListBoxItem Content="memberOf" />
                                    </ListBox>
                                </TabItem>
                                <TabItem FontSize="11" Header="Extended">
                                    <ListBox
                                        x:Name="ListBoxExtendedAttributes"
                                        Background="Transparent"
                                        BorderThickness="0"
                                        SelectionMode="Multiple">
                                        <ListBoxItem Content="whenCreated" />
                                        <ListBoxItem Content="whenChanged" />
                                        <ListBoxItem Content="Manager" />
                                        <ListBoxItem Content="Company" />
                                        <ListBoxItem Content="physicalDeliveryOfficeName" />
                                        <ListBoxItem Content="telephoneNumber" />
                                        <ListBoxItem Content="homeDirectory" />
                                        <ListBoxItem Content="scriptPath" />
                                    </ListBox>
                                </TabItem>
                            </TabControl>
                        </StackPanel>
                    </Border>

                    <!--  Action Buttons  -->
                    <StackPanel
                        Grid.Column="4"
                        MinWidth="180"
                        Margin="0,44,0,63">
                        <Button
                            x:Name="ButtonQueryAD"
                            Width="178"
                            Height="50"
                            Margin="0,0,0,12"
                            Content="SEARCH"
                            FontSize="16"
                            Style="{StaticResource PrimaryButton}" Background="#FF313261" />
                        <StackPanel
                            HorizontalAlignment="Center"
                            VerticalAlignment="Center"
                            Orientation="Horizontal">
                            <Button
                                x:Name="ButtonExportCSV"
                                Width="85"
                                Height="38"
                                Margin="0,0,8,0"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                Background="#FFBDBDBD"
                                Content="CSV"
                                Style="{StaticResource ModernButton}" />
                            <Button
                                x:Name="ButtonExportHTML"
                                Width="85"
                                Height="38"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                Background="#FFBDBDBD"
                                Content="HTML"
                                Style="{StaticResource ModernButton}" />
                        </StackPanel>
                    </StackPanel>

                </Grid>

                <!--  Results Section  -->
                <Border
                    Grid.Row="2"
                    Margin="0,0,0,10"
                    Padding="20,16"
                    Background="#FFFFFF"
                    Style="{StaticResource ModernCard}">
                    <Grid Margin="-10,0,-20,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>

                        <!--  Results Header  -->
                        <Grid Grid.Row="0" Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                            </Grid.ColumnDefinitions>
                            <TextBlock
                                Grid.Column="0"
                                Style="{StaticResource CategoryHeader}"
                                Text="Query Results" />
                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button
                                    x:Name="ButtonRefresh"
                                    Margin="0,0,8,0"
                                    Padding="8,4"
                                    Content="Reset Query Window"
                                    Style="{StaticResource ModernButton}" />
                                <Button
                                    x:Name="ButtonCopy"
                                    Margin="0,0,10,0"
                                    Padding="8,4"
                                    Content="Copy Selected Rows"
                                    Style="{StaticResource ModernButton}" />
                            </StackPanel>
                        </Grid>

                        <!--  Enhanced DataGrid  -->
                        <DataGrid
                            x:Name="DataGridResults"
                            Grid.Row="1"
                            Margin="0,0,10,0"
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
        </Grid>

        <!--  Footer (Grid.Row="2")  -->
        <Border
            Grid.Row="2"
            Background="#374151"
            BorderBrush="#E5E7EB"
            BorderThickness="0,1,0,0">
            <Grid Margin="20,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <TextBlock
                    Grid.Column="0"
                    VerticalAlignment="Center"
                    FontSize="11"
                    Foreground="#E5E7EB"
                    Text="AD Reporting Tool  |  Version: 1.2  |  Last Update: 06.30.2025" />

                <StackPanel
                    Grid.Column="1"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center"
                    Orientation="Horizontal">
                    <Ellipse x:Name="StatusIndicator" Width="16" Height="16" Fill="#10B981" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TextBlockStatus" Text="Ready" FontWeight="Medium" Foreground="#FFFFBBBB" Margin="0,0,16,0" FontSize="16" Height="22"/>
                </StackPanel>

                <TextBlock
                    Grid.Column="2"
                    Width="141"
                    Margin="620,0,0,0"
                    HorizontalAlignment="Left"
                    VerticalAlignment="Center"
                    FontSize="11"
                    Foreground="#E5E7EB"
                    Text="© 2025 Henchman33" />
                <TextBlock
                    Name="linkWebsite"
                    Grid.Column="2"
                    Margin="780,0,0,0"
                    HorizontalAlignment="Left"
                    VerticalAlignment="Center"
                    Cursor="Hand"
                    FontSize="11"
                    Foreground="#93C5FD"
                    Text="https://github.com/Henchman33" Width="160" />
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ===== INITIALISIERUNG =====
# Progress Bars deaktivieren für saubere GUI-Darstellung
$ProgressPreference = 'SilentlyContinue'

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Assembly fÃ¼r WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # FÃ¼r SaveFileDialog

# --- Globale AD-Gruppennamen fÃ¼r Deutsch/Englisch KompatibilitÃ¤t ---
$Global:ADGroupNames = @{
    DomainAdmins = @("Domain Admins", "DomÃ¤nen-Admins")
    EnterpriseAdmins = @("Enterprise Admins", "Organisations-Admins")
    SchemaAdmins = @("Schema Admins", "Schema-Admins")
    Administrators = @("Administrators", "Administratoren")
    AccountOperators = @("Account Operators", "Konten-Operatoren")
    ServerOperators = @("Server Operators", "Server-Operatoren")
    BackupOperators = @("Backup Operators", "Sicherungs-Operatoren")
    PrintOperators = @("Print Operators", "Druck-Operatoren")
    Replicator = @("Replicator", "Replikations-Operator")
    RemoteDesktopUsers = @("Remote Desktop Users", "Remotedesktopbenutzer")
    PowerUsers = @("Power Users", "Hauptbenutzer")
    DomainControllers = @("Domain Controllers", "DomÃ¤nencontroller")
    EnterpriseDomainControllers = @("Enterprise Domain Controllers", "Organisations-DomÃ¤nencontroller")
}

# Hilfsfunktion zum Finden von AD-Gruppen in beiden Sprachen
Function Get-ADGroupByNames {
    param(
        [string[]]$GroupNames,
        [switch]$ReturnAll
    )
    
    $foundGroups = @()
    foreach ($name in $GroupNames) {
        try {
            $group = Get-ADGroup -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue
            if ($group) {
                if ($ReturnAll) {
                    $foundGroups += $group
                } else {
                    return $group
                }
            }
        } catch {
            # Ignore errors for non-existent groups
        }
    }
    
    if ($ReturnAll) {
        return $foundGroups
    }
    return $null
}

# --- Log-Funktion fÃ¼r konsistente Fehlerausgabe ---
Function Write-ADReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory=$false)]
        [switch]$GUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$Terminal
    )
    
    # StandardmÃ¤ÃŸig nur GUI, Terminal deaktiviert für EXE-Konvertierung
    if (-not $GUI -and -not $Terminal) {
        $GUI = $true
        $Terminal = $false  # Deaktiviert für EXE-Version
    }
    
    # Ausgabe in der GUI
    if ($GUI -and $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $Message
    }
    
    # Ausgabe im Terminal (deaktiviert für EXE-Kompilierung)
    # if ($Terminal) {
    #     switch ($Type) {
    #         'Info'    { Write-Host $Message }
    #         'Warning' { Write-Warning $Message }
    #         'Error'   { Write-Error $Message }
    #     }
    # }
}

# --- Debug-Log-Funktion fÃ¼r konsistente Debug-Ausgabe ---
Function Write-DebugLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Category = 'Debug'
    )
    
    # Debug-Ausgabe nur wenn $DebugPreference gesetzt ist
    if ($DebugPreference -ne 'SilentlyContinue') {
        Write-Debug "[$Category] $Message"
    }
}

# --- Funktion zum Abrufen von AD-Daten ---
Function Get-ADReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue,
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute2,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue2,
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator2 = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterLogic = "AND",
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter,
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "User"
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Error: Active Directory module not found." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        # Basis-Eigenschaften hinzufÃ¼gen
        if ('DistinguishedName' -notin $PropertiesToLoad) { $PropertiesToLoad += 'DistinguishedName' }
        if ('ObjectClass' -notin $PropertiesToLoad) { $PropertiesToLoad += 'ObjectClass' }
        $PropertiesToLoad = $PropertiesToLoad | Select-Object -Unique

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("DisplayName", "SamAccountName", "ObjectClass")
        }

        # Filter erstellen
        $Filter = "*"
        
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } else {
            # Erstelle Filter basierend auf den Eingaben
            $FilterPart1 = ""
            $FilterPart2 = ""
            
            # Erster Filter
            if (-not [string]::IsNullOrWhiteSpace($FilterValue) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute $FilterAttribute -Value $FilterValue -Operator $FilterOperator
            } elseif (-not [string]::IsNullOrWhiteSpace($FilterValue) -and [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute "DisplayName" -Value $FilterValue -Operator "Contains"
            }
            
            # Zweiter Filter (wenn aktiviert)
            if (-not [string]::IsNullOrWhiteSpace($FilterValue2) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute2)) {
                $FilterPart2 = Build-FilterString -Attribute $FilterAttribute2 -Value $FilterValue2 -Operator $FilterOperator2
            }
            
            # Kombiniere Filter mit UND/ODER Logik
            if ($FilterPart1 -and $FilterPart2) {
                if ($FilterLogic -eq "AND") {
                    $Filter = "$FilterPart1 -and $FilterPart2"
                } else {
                    $Filter = "$FilterPart1 -or $FilterPart2"
                }
            } elseif ($FilterPart1) {
                $Filter = $FilterPart1
            } elseif ($FilterPart2) {
                $Filter = $FilterPart2
            }
        }

        Write-ADReportLog -Message "Executing AD query with filter: $Filter" -Type Info -Terminal

        # AD-Abfrage basierend auf Objekttyp
        $Results = $null
        switch ($ObjectType) {
            "User" {
                if ($Filter -and $Filter.Trim() -eq "LockedOut -eq `$true") {
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop
                    if ($LockedOutAccounts) {
                        $Results = foreach ($Account in $LockedOutAccounts) {
                            try {
                                Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue
                            } catch {
                                Write-Warning "Could not get details for user $($Account.SamAccountName): $($_.Exception.Message)"
                                $null
                            }
                        }
                        $Results = $Results | Where-Object {$_ -ne $null}
                    }
                } else {
                    # Debug-Ausgabe fÃ¼r Filter
                    Write-DebugLog "Executing Get-ADUser with filter: $Filter" -Category "ADQuery"
                    $Results = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
                }
            }
            "Group" {
                $Results = @(Get-ADGroup -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
            "Computer" {
                $Results = @(Get-ADComputer -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
        }

        if ($Results) {
            $Results = $Results | Select-Object -Property $PropertiesToLoad -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Results
        } else {
            Write-ADReportLog -Message "No objects found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# Hilfsfunktion zum Erstellen von Filter-Strings
Function Build-FilterString {
    param(
        [string]$Attribute,
        [string]$Value,
        [string]$Operator
    )
    
    switch ($Operator) {
        "Contains" { return "$Attribute -like '*$Value*'" }
        "Equals" { return "$Attribute -eq '$Value'" }
        "StartsWith" { return "$Attribute -like '$Value*'" }
        "EndsWith" { return "$Attribute -like '*$Value'" }
        "NotEqual" { return "$Attribute -ne '$Value'" }
        default { return "$Attribute -like '*$Value*'" }
    }
}

# --- Funktion zum Abrufen der Gruppenmitgliedschaften eines Benutzers ---
Function Get-UserGroupMemberships {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        $User = Get-ADUser -Identity $SamAccountName -Properties SamAccountName, Name -ErrorAction Stop | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount' # HinzugefÃ¼gt: Name fÃ¼r UserDisplayName
        if (-not $User) {
            Write-ADReportLog -Message "Benutzer $SamAccountName nicht gefunden." -Type Warning
            return $null
        }
        
        $Groups = Get-ADPrincipalGroupMembership -Identity $User -ErrorAction Stop | 
                  Get-ADGroup -Properties Name, SamAccountName, Description, GroupCategory, GroupScope -ErrorAction SilentlyContinue # HinzugefÃ¼gt: SamAccountName fÃ¼r GroupSamAccountName

        if ($Groups) {
            $GroupMemberships = $Groups | ForEach-Object {
                [PSCustomObject]@{
                    UserDisplayName = $User.Name
                    UserSamAccountName = $User.SamAccountName
                    GroupName = $_.Name
                    GroupSamAccountName = $_.SamAccountName
                    GroupDescription = $_.Description
                    GroupCategory = $_.GroupCategory
                    GroupScope = $_.GroupScope
                }
            }
            return $GroupMemberships
        } else {
            Write-ADReportLog -Message "Keine Gruppenmitgliedschaften fÃ¼r Benutzer $SamAccountName gefunden." -Type Info
            return [System.Collections.ArrayList]::new() # Leeres Array zurÃ¼ckgeben, um Fehler zu vermeiden
        }
    } catch {
        $ErrorMessage = "Fehler beim Abrufen der Gruppenmitgliedschaften fuer $($SamAccountName): $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Gruppendaten ---
Function Get-ADGroupReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        # Debug-Info ins Logfile schreiben
        Write-DebugLog "FÃ¼hre Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus" -Category "ADQuery"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurÃ¼ckgeben
            $Output = $Groups | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No groups found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD group query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Computerdaten ---
Function Get-ADComputerReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        # Debug-Info ins Logfile schreiben
        Write-DebugLog "FÃ¼hre Get-ADComputer mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus" -Category "ADQuery"
        $Computers = Get-ADComputer -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Computers) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            $Output = $Computers | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No computers found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD computer query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von Gruppenmitgliedschaftsberichten ---
Function Get-ADGroupMembershipsReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilterAttribute,

        [Parameter(Mandatory=$true)]
        [string]$FilterValue
    )

    Write-ADReportLog -Message "Fetching group membership data for filter: $FilterAttribute = $FilterValue" -Type Info -Terminal
    $ReportOutput = @()

    try {
        $TargetObjects = Get-ADObject -Filter "$FilterAttribute -like '*$FilterValue*'" -Properties DisplayName, SamAccountName, MemberOf, Member, ObjectClass -ErrorAction SilentlyContinue

        if (-not $TargetObjects) {
            Write-ADReportLog -Message "No objects found for the specified filter '$FilterAttribute -like *$FilterValue*'." -Type Warning
            return $ReportOutput
        }

        # Gruppiere Objekte nach Typ für bessere Strukturierung
        $GroupedObjects = $TargetObjects | Group-Object {
            if ($_.ObjectClass -is [array]) { $_.ObjectClass[-1] } else { $_.ObjectClass }
        }

        foreach ($ObjectGroup in $GroupedObjects) {
            foreach ($TargetObject in $ObjectGroup.Group) {
                $objectClassSimple = if ($TargetObject.ObjectClass -is [array]) {
                    $TargetObject.ObjectClass[-1]
                } else {
                    $TargetObject.ObjectClass
                }
                
                $objDisplayName = if ($TargetObject.DisplayName) { $TargetObject.DisplayName } else { $TargetObject.Name }
                $objSamAccountName = if ($TargetObject.SamAccountName) { $TargetObject.SamAccountName } else { "N/A" }

                Write-ADReportLog -Message "Processing object: $objDisplayName (Type: $objectClassSimple)" -Type Info

                # Header-Eintrag für das aktuelle Objekt
                $ReportOutput += [PSCustomObject]@{
                    ObjectName = "$objDisplayName ($objSamAccountName)"
                    ObjectSAM = ""
                    ObjectType = $objectClassSimple.ToUpper()
                    Relationship = ""
                    RelatedObject = ""
                    RelatedObjectSAM = ""
                    RelatedObjectType = ""
                }

                if ($objectClassSimple -eq 'user' -or $objectClassSimple -eq 'computer') {
                    if ($TargetObject.MemberOf) {
                        foreach ($groupDN in $TargetObject.MemberOf) {
                            try {
                                $groupObject = Get-ADGroup -Identity $groupDN -Properties DisplayName, SamAccountName -ErrorAction Stop
                                $groupDisplayName = if ($groupObject.DisplayName) { $groupObject.DisplayName } else { $groupObject.Name }
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ IST MITGLIED VON"
                                    RelatedObject = $groupDisplayName
                                    RelatedObjectSAM = $groupObject.SamAccountName
                                    RelatedObjectType = "Group"
                                }
                            } catch {
                                Write-ADReportLog -Message "Error resolving group DN '$groupDN' for object '$objDisplayName': $($_.Exception.Message)" -Type Warning
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ IST MITGLIED VON (FEHLER)"
                                    RelatedObject = $groupDN
                                    RelatedObjectSAM = "N/A"
                                    RelatedObjectType = "Group (Fehler)"
                                }
                            }
                        }
                    } else {
                        $ReportOutput += [PSCustomObject]@{
                            ObjectName = ""
                            ObjectSAM = ""
                            ObjectType = ""
                            Relationship = "└─ KEINE GRUPPENMITGLIEDSCHAFTEN"
                            RelatedObject = ""
                            RelatedObjectSAM = ""
                            RelatedObjectType = ""
                        }
                    }
                } elseif ($objectClassSimple -eq 'group') {
                    if ($TargetObject.Member) {
                        foreach ($memberDN in $TargetObject.Member) {
                            try {
                                $memberObject = Get-ADObject -Identity $memberDN -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                                $memberObjectClassSimple = if ($memberObject.ObjectClass -is [array]) {
                                    $memberObject.ObjectClass[-1]
                                } else {
                                    $memberObject.ObjectClass
                                }
                                $memberName = if ($memberObject.DisplayName) { $memberObject.DisplayName } else { $memberObject.Name }
                                $memberSam = if ($memberObject.SamAccountName) { $memberObject.SamAccountName } else { "N/A" }

                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ HAT MITGLIED"
                                    RelatedObject = $memberName
                                    RelatedObjectSAM = $memberSam
                                    RelatedObjectType = $memberObjectClassSimple
                                }
                            } catch {
                                Write-ADReportLog -Message "Error resolving member DN '$memberDN' for group '$objDisplayName': $($_.Exception.Message)" -Type Warning
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ HAT MITGLIED (FEHLER)"
                                    RelatedObject = $memberDN
                                    RelatedObjectSAM = "N/A"
                                    RelatedObjectType = "Unbekannt (Fehler)"
                                }
                            }
                        }
                    } else {
                        $ReportOutput += [PSCustomObject]@{
                            ObjectName = ""
                            ObjectSAM = ""
                            ObjectType = ""
                            Relationship = "└─ KEINE MITGLIEDER"
                            RelatedObject = ""
                            RelatedObjectSAM = ""
                            RelatedObjectType = ""
                        }
                    }
                }

                # Leerzeile nach jedem Objekt für bessere Lesbarkeit
                $ReportOutput += [PSCustomObject]@{
                    ObjectName = ""
                    ObjectSAM = ""
                    ObjectType = ""
                    Relationship = ""
                    RelatedObject = ""
                    RelatedObjectSAM = ""
                    RelatedObjectType = ""
                }
            }
        }
    } catch {
        $ErrorMessage = "Error in Get-ADGroupMembershipsReportData: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
    }
    return $ReportOutput
}

# --- Security Audit Functions ---
Function Get-WeakPasswordPolicyUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with weak password policies..." -Type Info -Terminal
        
        # Properties relevant for comprehensive password policy analysis
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", 
            "PasswordNotRequired", "PasswordLastSet", "LastLogonDate", "AdminCount",
            "CannotChangePassword", "SmartcardLogonRequired", "TrustedForDelegation",
            "DoesNotRequirePreAuth", "UseDESKeyOnly", "AccountExpirationDate",
            "LastBadPasswordAttempt", "BadLogonCount", "LogonCount", "PrimaryGroup",
            "MemberOf", "ServicePrincipalNames", "UserAccountControl", "LockedOut",
            "TrustedToAuthForDelegation", "AllowReversiblePasswordEncryption",
            "whenCreated", "Description", "UserPrincipalName", "DistinguishedName"
        )
        
        # Retrieve Domain Password Policy for comparisons
        $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        $MinPasswordAge = $DomainPasswordPolicy.MinPasswordAge.Days
        $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
        $MinPasswordLength = $DomainPasswordPolicy.MinPasswordLength
        
        Write-ADReportLog -Message "Domain Password Policy - Min Age: $MinPasswordAge days, Max Age: $MaxPasswordAge days, Min Length: $MinPasswordLength chars" -Type Info -Terminal
        
        # Load all users with relevant properties
        $AllUsers = Get-ADUser -Filter * -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        Write-ADReportLog -Message "$($AllUsers.Count) users loaded for analysis..." -Type Info -Terminal
        
        # Define high-risk service account patterns
        $ServiceAccountPatterns = @("svc", "service", "app", "sql", "iis", "web", "backup", "sync", "admin", "sa")
        $TestAccountPatterns = @("test", "temp", "demo", "guest", "anonymous", "trial")
        
        # Enhanced analysis for weak password policies
        $WeakPasswordUsers = foreach ($user in $AllUsers) {
            $issues = @()
            $riskLevel = 0
            $recommendations = @()
            $securityFlags = @()
            
            # 1. Password never expires
            if ($user.PasswordNeverExpires -eq $true) {
                $issues += "Password never expires"
                $riskLevel += 3
                $recommendations += "Enable password expiration"
                $securityFlags += "NO_EXPIRY"
            }
            
            # 2. Password not required
            if ($user.PasswordNotRequired -eq $true) {
                $issues += "Password not required"
                $riskLevel += 5  # Critical risk
                $recommendations += "Enforce password requirement"
                $securityFlags += "NO_PASSWORD_REQ"
            }
            
            # 3. No password set or extremely old password
            if ($user.PasswordLastSet -eq $null) {
                $issues += "No password set"
                $riskLevel += 5
                $recommendations += "Set password immediately"
                $securityFlags += "NO_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 3))) {
                $issues += "Password extremely outdated (>$($MaxPasswordAge * 3) days)"
                $riskLevel += 4
                $recommendations += "Force password reset immediately"
                $securityFlags += "ANCIENT_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 2))) {
                $issues += "Password very old (>$($MaxPasswordAge * 2) days)"
                $riskLevel += 3
                $recommendations += "Schedule password reset"
                $securityFlags += "OLD_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-365)) {
                $issues += "Password older than 1 year"
                $riskLevel += 2
                $recommendations += "Password reset recommended"
                $securityFlags += "STALE_PASSWORD"
            }
            
            # 4. User cannot change password
            if ($user.CannotChangePassword -eq $true) {
                $issues += "Cannot change password"
                $riskLevel += 2
                $recommendations += "Allow password changes (except for service accounts)"
                $securityFlags += "NO_CHANGE_ALLOWED"
            }
            
            # 5. Kerberos Pre-Authentication disabled (ASREPRoast attack possible)
            if ($user.DoesNotRequirePreAuth -eq $true) {
                $issues += "Kerberos Pre-Auth disabled (ASREPRoast vulnerability)"
                $riskLevel += 4
                $recommendations += "Enable Kerberos Pre-Authentication"
                $securityFlags += "ASREP_ROASTABLE"
            }
            
            # 6. Weak encryption (DES)
            if ($user.UseDESKeyOnly -eq $true) {
                $issues += "Uses weak DES encryption"
                $riskLevel += 3
                $recommendations += "Disable DES encryption"
                $securityFlags += "WEAK_ENCRYPTION"
            }
            
            # 7. Reversible password encryption
            if ($user.AllowReversiblePasswordEncryption -eq $true) {
                $issues += "Reversible password encryption enabled"
                $riskLevel += 4
                $recommendations += "Disable reversible password encryption"
                $securityFlags += "REVERSIBLE_ENCRYPTION"
            }
            
            # 8. Smartcard authentication not used for privileged accounts
            if ($user.AdminCount -eq 1 -and $user.SmartcardLogonRequired -eq $false) {
                $issues += "Privileged account without smartcard requirement"
                $riskLevel += 3
                $recommendations += "Enable smartcard authentication for admin accounts"
                $securityFlags += "ADMIN_NO_SMARTCARD"
            }
            
            # 9. Delegation for normal user accounts
            if (($user.TrustedForDelegation -eq $true -or $user.TrustedToAuthForDelegation -eq $true) -and $user.AdminCount -ne 1) {
                $issues += "Delegation enabled for standard user"
                $riskLevel += 3
                $recommendations += "Restrict delegation to service accounts only"
                $securityFlags += "UNEXPECTED_DELEGATION"
            }
            
            # 10. Excessive failed logon attempts
            if ($user.BadLogonCount -gt 10) {
                $issues += "Excessive failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 2
                $recommendations += "Investigate account for potential compromise"
                $securityFlags += "HIGH_FAILED_LOGONS"
            } elseif ($user.BadLogonCount -gt 5) {
                $issues += "Multiple failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 1
                $recommendations += "Monitor account activity"
                $securityFlags += "FAILED_LOGONS"
            }
            
            # 11. Service account without SPN (potentially misconfigured)
            $isServiceAccount = $false
            foreach ($pattern in $ServiceAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isServiceAccount = $true
                    break
                }
            }
            
            if ($isServiceAccount) {
                if (-not $user.ServicePrincipalNames) {
                    $issues += "Service account without SPN"
                    $riskLevel += 1
                    $recommendations += "Configure SPN for service account"
                    $securityFlags += "SERVICE_NO_SPN"
                }
                
                # Service accounts should not be interactive
                if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                    $issues += "Interactive logons detected for service account"
                    $riskLevel += 2
                    $recommendations += "Review service account usage"
                    $securityFlags += "SERVICE_INTERACTIVE"
                }
            }
            
            # 12. Never logged in accounts with high privileges
            if ($user.LogonCount -eq 0 -and $user.AdminCount -eq 1) {
                $issues += "Admin account never used"
                $riskLevel += 3
                $recommendations += "Disable unused admin account"
                $securityFlags += "UNUSED_ADMIN"
            } elseif ($user.LogonCount -eq 0 -and $user.whenCreated -lt (Get-Date).AddDays(-30)) {
                $issues += "Account never used (created >30 days ago)"
                $riskLevel += 1
                $recommendations += "Consider disabling unused account"
                $securityFlags += "NEVER_USED"
            }
            
            # 13. Inactive privileged accounts
            if ($user.AdminCount -eq 1 -and $user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-60)) {
                $issues += "Privileged account inactive for >60 days"
                $riskLevel += 3
                $recommendations += "Review inactive admin account"
                $securityFlags += "INACTIVE_ADMIN"
            } elseif ($user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-180)) {
                $issues += "Account inactive for >180 days"
                $riskLevel += 2
                $recommendations += "Consider disabling inactive account"
                $securityFlags += "LONG_INACTIVE"
            }
            
            # 14. Test/temp accounts without expiration
            $isTestAccount = $false
            foreach ($pattern in $TestAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isTestAccount = $true
                    break
                }
            }
            
            if ($isTestAccount) {
                if ($user.AccountExpirationDate -eq $null) {
                    $issues += "Test/temp account without expiration date"
                    $riskLevel += 2
                    $recommendations += "Set expiration date for temporary accounts"
                    $securityFlags += "TEST_NO_EXPIRY"
                }
                
                if ($user.whenCreated -lt (Get-Date).AddDays(-90)) {
                    $issues += "Old test account (>90 days)"
                    $riskLevel += 1
                    $recommendations += "Review necessity of old test account"
                    $securityFlags += "OLD_TEST_ACCOUNT"
                }
            }
            
            # 15. Locked accounts with admin privileges
            if ($user.LockedOut -eq $true -and $user.AdminCount -eq 1) {
                $issues += "Locked privileged account"
                $riskLevel += 2
                $recommendations += "Investigate locked admin account"
                $securityFlags += "LOCKED_ADMIN"
            }
            
            # 16. Accounts with suspicious creation patterns
            if ($user.whenCreated -gt (Get-Date).AddDays(-7) -and $user.AdminCount -eq 1) {
                $issues += "Recently created admin account"
                $riskLevel += 2
                $recommendations += "Verify legitimacy of new admin account"
                $securityFlags += "NEW_ADMIN"
            }
            
            # 17. Accounts with generic or weak naming
            $weakNames = @("admin", "administrator", "user", "guest", "test", "temp", "service", "default")
            foreach ($weakName in $weakNames) {
                if ($user.SamAccountName -eq $weakName -or $user.SamAccountName -like "$weakName*") {
                    $issues += "Generic/predictable account name"
                    $riskLevel += 1
                    $recommendations += "Use non-predictable account names"
                    $securityFlags += "WEAK_NAMING"
                    break
                }
            }
            
            # Only return users with identified vulnerabilities
            if ($issues.Count -gt 0) {
                # Categorize risk level
                $riskCategory = switch ([int]$riskLevel) {
                    {$_ -ge 10} { [string]"Critical" }
                    {$_ -ge 7} { [string]"High" }
                    {$_ -ge 4} { [string]"Medium" }
                    {$_ -ge 2} { [string]"Low" }
                    default    { [string]"Minimal" }
                }
                

                # Add context information
                $contextInfo = @()
                if ($user.AdminCount -eq 1) { $contextInfo += "Privileged Account" }
                if ($user.Enabled -eq $false) { $contextInfo += "Disabled" }
                if ($user.LockedOut -eq $true) { $contextInfo += "Locked" }
                if ($user.ServicePrincipalNames) { $contextInfo += "Service Account" }
                if ($isServiceAccount) { $contextInfo += "Service Pattern" }
                if ($isTestAccount) { $contextInfo += "Test/Temp Account" }
                if ($user.SamAccountName -match "^(admin|administrator|root|sa|service|svc)") { $contextInfo += "System Account" }
                
                # Calculate compliance status
                $complianceIssues = 0
                if ($user.PasswordNeverExpires) { $complianceIssues++ }
                if ($user.PasswordNotRequired) { $complianceIssues++ }
                if ($user.DoesNotRequirePreAuth) { $complianceIssues++ }
                if ($user.UseDESKeyOnly) { $complianceIssues++ }
                if ($user.AllowReversiblePasswordEncryption) { $complianceIssues++ }
                
                $complianceStatus = if ($complianceIssues -eq 0) { "Compliant" } 
                                   elseif ($complianceIssues -le 2) { "Partially Compliant" } 
                                   else { "Non-Compliant" }
                
                # Calculate password age in days
                $passwordAge = if ($user.PasswordLastSet) { 
                    [math]::Round((New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).TotalDays) 
                } else { "Never Set" }
                
                # Calculate account age
                $accountAge = [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays)
                
                # Determine urgency level
                $urgencyLevel = if ($user.AdminCount -eq 1 -and $riskLevel -ge 7) { "Immediate Action Required" }
                               elseif ($riskLevel -ge 10) { "Critical" }
                               elseif ($riskLevel -ge 7) { "Urgent" }
                               elseif ($riskLevel -ge 4) { "High Priority" }
                               elseif ($riskLevel -ge 2) { "Medium Priority" }
                               else { "Low Priority" }
                
                # Enhanced output object with comprehensive analysis
                [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    LockedOut = $user.LockedOut
                    Context = if ($contextInfo) { $contextInfo -join ", " } else { "Standard User" }
                    PasswordLastSet = $user.PasswordLastSet
                    PasswordAge = $passwordAge
                    AccountCreated = $user.whenCreated
                    AccountAge = $accountAge
                    LastLogonDate = $user.LastLogonDate
                    LogonCount = $user.LogonCount
                    BadLogonCount = $user.BadLogonCount
                    LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                    Vulnerabilities = $issues -join "; "
                    SecurityFlags = $securityFlags -join ", "
                    RiskLevel = $riskCategory
                    RiskScore = $riskLevel
                    ComplianceStatus = $complianceStatus
                    ComplianceIssues = $complianceIssues
                    UrgencyLevel = $urgencyLevel
                    Recommendations = $recommendations -join "; "
                    Description = $user.Description
                    LastAssessment = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    TotalIssuesFound = $issues.Count
                    RequiresImmediateAction = ($urgencyLevel -eq "Immediate Action Required" -or $urgencyLevel -eq "Critical")
                }
            }
        }
        
        # Enhanced statistics for logging
        $totalIssues = $WeakPasswordUsers.Count
        $criticalIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        $adminIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Privileged Account*" }).Count
        $nonCompliant = ($WeakPasswordUsers | Where-Object { $_.ComplianceStatus -eq "Non-Compliant" }).Count
        $immediateAction = ($WeakPasswordUsers | Where-Object { $_.RequiresImmediateAction -eq $true }).Count
        $serviceAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Service*" }).Count
        $testAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Test*" }).Count
        
        Write-ADReportLog -Message "Password Policy Analysis completed:" -Type Info -Terminal
        Write-ADReportLog -Message "  Total: $totalIssues users with vulnerabilities" -Type Info -Terminal
        Write-ADReportLog -Message "  Risk Distribution - Critical: $criticalIssues, High: $highIssues, Medium: $mediumIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Privileged accounts affected: $adminIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Service accounts with issues: $serviceAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Test/temp accounts with issues: $testAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Compliance violations: $nonCompliant" -Type Info -Terminal
        Write-ADReportLog -Message "  Requiring immediate action: $immediateAction" -Type Info -Terminal
        
        return $WeakPasswordUsers
        
    } catch {
        $ErrorMessage = "Error analyzing password policies: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-RiskyGroupMemberships {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing risky group memberships..." -Type Info -Terminal
        
        # Define high-privileged groups with German and English names
        $RiskyGroups = @(
            'Domain Controllers', 'Domänencontroller',
            'Enterprise Admins', 'Organisations-Admins',
            'Domain Admins', 'Domänen-Admins', 
            'Account Operators', 'Konten-Operatoren',
            'Remote Desktop Users', 'Remotedesktopbenutzer',
            'Enterprise Domain Controllers', 'Organisations-Domänencontroller',
            'Schema Admins', 'Schema-Admins',
            'Backup Operators', 'Sicherungsoperatoren',
            'Replicator', 'Replikations-Operator',
            'Server Operators', 'Server-Operatoren',
            'Print Operators', 'Druckoperatoren',
            'Power Users', 'Hauptbenutzer',
            'Administrators', 'Administratoren'
        )

        # Add additional groups from configuration if available
        if ($Global:ADGroupNames) {
            $RiskyGroups += foreach ($groupType in $Global:ADGroupNames.GetEnumerator()) {
                if (-not [string]::IsNullOrEmpty($groupType.Value)) {
                    $groupType.Value
                }
            }
        }
        
        $RiskyUsers = [System.Collections.Generic.List[PSObject]]::new()
        
        # Group analysis with enhanced error handling
        foreach ($groupName in $RiskyGroups) {
            try {
                if ([string]::IsNullOrEmpty($groupName)) {
                    Write-ADReportLog -Message "Skipping empty group name" -Type Warning
                    continue
                }

                # Try to find group using both name formats (German/English)
                $group = Get-ADGroup -Filter "Name -eq '$groupName' -or SamAccountName -eq '$groupName'" -ErrorAction SilentlyContinue
                if (-not $group) {
                    Write-ADReportLog -Message "Group '$groupName' not found" -Type Warning
                    continue
                }
                
                $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue |
                    Where-Object { $_.objectClass -eq "user" } |
                    ForEach-Object {
                        try {
                            Get-ADObject -Identity $_.DistinguishedName -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                        } catch {
                            Write-ADReportLog -Message "Error retrieving user object: $($_.Exception.Message)" -Type Warning
                            $null
                        }
                    } | Where-Object { $_ -ne $null }
                
                foreach ($member in $members) {
                    try {
                        $userDetails = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName, Enabled, LastLogonDate, PasswordLastSet, Department, Description -ErrorAction Stop
                        
                        # Dynamic risk assessment based on group name (normalized to English)
                        $normalizedGroupName = switch -Wildcard ($group.Name) {
                            { $_ -match '(Domain Admins|Domänen-Admins)' } { "Domain Admins" }
                            { $_ -match '(Enterprise Admins|Organisations-Admins)' } { "Enterprise Admins" }
                            { $_ -match '(Schema Admins|Schema-Admins)' } { "Schema Admins" }
                            { $_ -match '(Administrators|Administratoren)' } { "Administrators" }
                            { $_ -match '(Domain Controllers|Domänencontroller)' } { "Domain Controllers" }
                            { $_ -match '(Account Operators|Konten-Operatoren)' } { "Account Operators" }
                            { $_ -match '(Server Operators|Server-Operatoren)' } { "Server Operators" }
                            { $_ -match '(Backup Operators|Sicherungsoperatoren)' } { "Backup Operators" }
                            { $_ -match '(Print Operators|Druckoperatoren)' } { "Print Operators" }
                            { $_ -match '(Power Users|Hauptbenutzer)' } { "Power Users" }
                            { $_ -match '(Remote Desktop Users|Remotedesktopbenutzer)' } { "Remote Desktop Users" }
                            { $_ -match '(Replicator|Replikations-Operator)' } { "Replicator" }
                            default { $group.Name }
                        }
                        
                        # Risk level assessment (output in English for GUI)
                        $riskLevel = switch ($normalizedGroupName) {
                            { $_ -in @("Domain Admins", "Enterprise Admins", "Schema Admins") } { "Critical" }
                            { $_ -in @("Administrators", "Domain Controllers") } { "High" }
                            { $_ -in @("Account Operators", "Server Operators", "Backup Operators") } { "Medium" }
                            { $_ -in @("Print Operators", "Power Users", "Remote Desktop Users", "Replicator") } { "Low" }
                            default { "Medium" }
                        }
                        
                        # Enhanced recommendation generation
                        $recommendation = switch ($true) {
                            (-not $userDetails.Enabled) { 
                                "Remove disabled account from privileged group immediately" 
                            }
                            ($null -eq $userDetails.LastLogonDate) { 
                                "Account never logged on - review necessity and remove if not required" 
                            }
                            ($userDetails.LastLogonDate -lt (Get-Date).AddDays(-90)) { 
                                "Inactive account (>90 days) - review and consider removal" 
                            }
                            ($userDetails.LastLogonDate -lt (Get-Date).AddDays(-30)) { 
                                "Account inactive for 30+ days - verify still needed" 
                            }
                            ($riskLevel -eq "Critical") { 
                                "Critical privileged account - implement enhanced monitoring and MFA" 
                            }
                            ($riskLevel -eq "High") { 
                                "High-risk account - regular access review required" 
                            }
                            default { 
                                "Regular monitoring and periodic access review required" 
                            }
                        }
                        
                        # Account type classification
                        $accountType = switch ($true) {
                            ($userDetails.SamAccountName -match '^(svc|service)') { "Service Account" }
                            ($userDetails.SamAccountName -match '^(admin|adm)') { "Administrative Account" }
                            ($userDetails.SamAccountName -match '^(test|tmp|temp)') { "Test/Temporary Account" }
                            ($userDetails.Description -match '(service|system|automation)') { "System Account" }
                            default { "User Account" }
                        }
                        
                        # Days since last logon calculation
                        $daysSinceLastLogon = if ($userDetails.LastLogonDate) {
                            [math]::Round((New-TimeSpan -Start $userDetails.LastLogonDate -End (Get-Date)).TotalDays)
                        } else {
                            999
                        }
                        
                        # Password age calculation
                        $passwordAge = if ($userDetails.PasswordLastSet) {
                            [math]::Round((New-TimeSpan -Start $userDetails.PasswordLastSet -End (Get-Date)).TotalDays)
                        } else {
                            999
                        }
                        
                        # Security flags
                        $securityFlags = @()
                        if (-not $userDetails.Enabled) { $securityFlags += "Disabled" }
                        if ($daysSinceLastLogon -gt 90) { $securityFlags += "Inactive" }
                        if ($passwordAge -gt 365) { $securityFlags += "Old Password" }
                        if ($null -eq $userDetails.LastLogonDate) { $securityFlags += "Never Logged On" }
                        
                        # Create result object with proper typing
                        $riskUser = [PSCustomObject]@{
                            RiskyGroup = [string]$normalizedGroupName
                            DisplayName = [string]($userDetails.DisplayName -replace '[^\x20-\x7E]', '')
                            SamAccountName = [string]$userDetails.SamAccountName
                            Enabled = [bool]$userDetails.Enabled
                            LastLogonDate = if ($userDetails.LastLogonDate) { $userDetails.LastLogonDate.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                            PasswordLastSet = if ($userDetails.PasswordLastSet) { $userDetails.PasswordLastSet.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                            DaysSinceLastLogon = [int]$daysSinceLastLogon
                            PasswordAge = [int]$passwordAge
                            RiskLevel = [string]$riskLevel
                            AccountType = [string]$accountType
                            SecurityFlags = [string]($securityFlags -join ", ")
                            Recommendation = [string]$recommendation
                        }
                        
                        $RiskyUsers.Add($riskUser)
                    }
                    catch {
                        Write-ADReportLog -Message "Error processing user $($member.SamAccountName): $($_.Exception.Message)" -Type Warning
                    }
                }
            }
            catch {
                Write-ADReportLog -Message "Group analysis failed for '$groupName': $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Deduplication and sorting
        $UniqueRiskyUsers = $RiskyUsers | 
            Sort-Object SamAccountName -Unique |
            Sort-Object @{
                Expression = {
                    switch ($_.RiskLevel) {
                        "Critical" { 1 }
                        "High" { 2 }
                        "Medium" { 3 }
                        "Low" { 4 }
                        default { 5 }
                    }
                }
            }, @{Expression = {-not $_.Enabled}}, DaysSinceLastLogon -Descending
        
        # Enhanced statistics logging
        if ($UniqueRiskyUsers.Count -gt 0) {
            $criticalCount = ($UniqueRiskyUsers | Where-Object { $_.RiskLevel -eq "Critical" }).Count
            $highCount = ($UniqueRiskyUsers | Where-Object { $_.RiskLevel -eq "High" }).Count
            $disabledCount = ($UniqueRiskyUsers | Where-Object { -not $_.Enabled }).Count
            $inactiveCount = ($UniqueRiskyUsers | Where-Object { $_.DaysSinceLastLogon -gt 90 }).Count
            $neverLoggedCount = ($UniqueRiskyUsers | Where-Object { $_.LastLogonDate -eq "Never" }).Count
            
            Write-ADReportLog -Message "Risky group membership analysis completed. Found $($UniqueRiskyUsers.Count) users in privileged groups." -Type Info -Terminal
            Write-ADReportLog -Message "Risk Distribution - Critical: $criticalCount, High: $highCount, Medium: $(($UniqueRiskyUsers | Where-Object { $_.RiskLevel -eq "Medium" }).Count), Low: $(($UniqueRiskyUsers | Where-Object { $_.RiskLevel -eq "Low" }).Count)" -Type Info -Terminal
            Write-ADReportLog -Message "Security Issues - Disabled: $disabledCount, Inactive (>90d): $inactiveCount, Never logged on: $neverLoggedCount" -Type Info -Terminal
        } else {
            Write-ADReportLog -Message "No risky group memberships found." -Type Info -Terminal
        }
        
        return $UniqueRiskyUsers
    }
    catch {
        $ErrorMessage = "Error in risky group membership analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return @()
    }
}

Function Get-PrivilegedAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere Konten mit erhÃ¶hten Rechten..." -Type Info -Terminal
        
        # Eigenschaften fÃ¼r privilegierte Konten
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "AdminCount", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires",
            "ServicePrincipalNames", "TrustedForDelegation", "TrustedToAuthForDelegation"
        )
        
        # Alle Benutzer mit AdminCount = 1 (historisch privilegiert)
        $AdminCountUsers = Get-ADUser -Filter "AdminCount -eq 1" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Service-Konten (Konten mit SPNs)
        $ServiceAccounts = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Konten mit Delegierungsrechten
        $DelegationAccounts = Get-ADUser -Filter "TrustedForDelegation -eq `$true -or TrustedToAuthForDelegation -eq `$true" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Alle privilegierten Konten zusammenfÃ¼hren
        $AllPrivilegedAccounts = @()
        $AllPrivilegedAccounts += $AdminCountUsers
        $AllPrivilegedAccounts += $ServiceAccounts
        $AllPrivilegedAccounts += $DelegationAccounts
        
        # Duplikate entfernen und analysieren
        $UniquePrivilegedAccounts = $AllPrivilegedAccounts | Sort-Object SamAccountName -Unique | ForEach-Object {
            $account = $_
            
            # Risikofaktoren analysieren
            $riskFactors = @()
            $riskLevel = 0
            
            if ($account.AdminCount -eq 1) {
                $riskFactors += "AdminCount set"
                $riskLevel += 2
            }
            
            if ($account.ServicePrincipalNames) {
                $riskFactors += "Service-Account (SPN)"
                $riskLevel += 1
            }
            
            if ($account.TrustedForDelegation) {
                $riskFactors += "Delegation activated"
                $riskLevel += 2
            }
            
            if ($account.TrustedToAuthForDelegation) {
                $riskFactors += "Constrained Delegation"
                $riskLevel += 1
            }
            
            if ($account.PasswordNeverExpires) {
                $riskFactors += "Password never expires"
                $riskLevel += 1
            }
            
            if ($account.Enabled -eq $false) {
                $riskFactors += "Account disabled"
                $riskLevel += 3  # Deactivated privileged accounts are a high risk
            }

            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $riskFactors += "Inactive (>90 days)"
                $riskLevel += 2
            }
            
            # Privilegien-Level bestimmen
            $privilegeLevel = "Standard"
            if ($account.AdminCount -eq 1 -and ($account.TrustedForDelegation -or $account.TrustedToAuthForDelegation)) {
                $privilegeLevel = "Critical"
            } elseif ($account.AdminCount -eq 1) {
                $privilegeLevel = "High"
            } elseif ($account.ServicePrincipalNames -or $account.TrustedForDelegation) {
                $privilegeLevel = "Medium"
            }
            
            # Risikostufe bestimmen
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 5} { "Critical" }
                {$_ -ge 3} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            # Empfehlungen generieren
            $recommendations = @()
            if ($account.Enabled -eq $false -and $account.AdminCount -eq 1) {
                $recommendations += "Reset AdminCount for deactivated account"
            }
            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $recommendations += "Check account usage"
            }
            if ($account.PasswordNeverExpires -and $account.AdminCount -eq 1) {
                $recommendations += "Enable password expiration"
            }
            if ($account.TrustedForDelegation) {
                $recommendations += "Review delegation rights"
            }
            
            [PSCustomObject]@{
                AdminCount = $account.AdminCount
                ServiceAccount = [bool]$account.ServicePrincipalNames
                DisplayName = $account.DisplayName
                SamAccountName = $account.SamAccountName
                Enabled = $account.Enabled
                LastLogonDate = $account.LastLogonDate
                PasswordLastSet = $account.PasswordLastSet
                PrivilegeLevel = $privilegeLevel
                RiskFactors = $riskFactors -join "; "
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Regular review" }
                Delegation = $account.TrustedForDelegation -or $account.TrustedToAuthForDelegation
            }
        }
        
        Write-ADReportLog -Message "$($UniquePrivilegedAccounts.Count) Konten mit erhÃ¶hten Rechten gefunden." -Type Info -Terminal
        return $UniquePrivilegedAccounts
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der privilegierten Konten: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Organisationsstruktur-Reports ---
Function Get-DepartmentStatistics {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing department statistics..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Abteilungsinformationen
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, whenCreated, PasswordLastSet, PasswordNeverExpires, LockedOut, Title -ErrorAction Stop
        
        # Gruppiere nach Abteilung und erstelle leeres Array für Ergebnisse
        $DepartmentStats = [System.Collections.ArrayList]@()
        $DepartmentGroups = $Users | Group-Object { 
            $dept = $_.Department
            if ([string]::IsNullOrWhiteSpace($dept)) { 
                "(No Department)" 
            } else { 
                [string]$dept 
            }
        }
        
        foreach ($dept in $DepartmentGroups) {
            $deptName = [string]$dept.Name
            $deptUsers = $dept.Group
            
            # Statistiken berechnen mit sicherer Typisierung
            $totalUsers = [int]($deptUsers | Measure-Object).Count
            $enabledCount = [int]($deptUsers | Where-Object { $_.Enabled -eq $true } | Measure-Object).Count
            $disabledCount = [int]($deptUsers | Where-Object { $_.Enabled -eq $false } | Measure-Object).Count
            $lockedCount = [int]($deptUsers | Where-Object { $_.LockedOut -eq $true } | Measure-Object).Count
            $neverExpireCount = [int]($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true } | Measure-Object).Count
            $inactiveCount = [int]($deptUsers | Where-Object { 
                $_.LastLogonDate -and 
                $_.LastLogonDate -lt (Get-Date).AddDays(-90) 
            } | Measure-Object).Count
            
            # Durchschnittliches Kontoalter berechnen
            $accountAges = @()
            foreach ($user in $deptUsers) {
                if ($user.whenCreated) {
                    $age = [int]((Get-Date) - $user.whenCreated).Days
                    $accountAges += $age
                }
            }
            
            $avgAccountAge = if ($accountAges.Count -gt 0) {
                [math]::Round(($accountAges | Measure-Object -Average).Average, 1)
            } else {
                0
            }
            
            # Titel zählen
            $uniqueTitles = [int](@($deptUsers | Where-Object { -not [string]::IsNullOrEmpty($_.Title) } | Select-Object -ExpandProperty Title -Unique).Count)
            
            # Sicherheits-Score berechnen
            $securityScore = 100.0
            if ($totalUsers -gt 0) {
                $issueCount = [int]($disabledCount + $lockedCount + $neverExpireCount + $inactiveCount)
                $securityScore = [math]::Round(100 - (($issueCount / $totalUsers) * 100), 1)
            }
            
            # Health Status basierend auf Security Score
            $healthStatus = switch ($securityScore) {
                {$_ -ge 90} { "Excellent" }
                {$_ -ge 80} { "Good" }
                {$_ -ge 70} { "Fair" }
                {$_ -ge 60} { "Poor" }
                default { "Critical" }
            }
            
            # Empfehlungen basierend auf Kennzahlen
            $recommendations = @()
            if ($disabledCount -gt 0) {
                $recommendations += "Review $disabledCount disabled accounts"
            }
            if ($lockedCount -gt 0) {
                $recommendations += "Unlock $lockedCount locked accounts"
            }
            if ($inactiveCount -gt 0) {
                $recommendations += "Review $inactiveCount inactive accounts"
            }
            if ($neverExpireCount -gt 0) {
                $recommendations += "Enable password expiration for $neverExpireCount accounts"
            }
            
            # Füge Statistiken zum Array hinzu
            $null = $DepartmentStats.Add([PSCustomObject]@{
                Department = [string]$deptName
                TotalUsers = [int]$totalUsers
                EnabledUsers = [int]$enabledCount
                DisabledUsers = [int]$disabledCount
                LockedUsers = [int]$lockedCount
                InactiveUsers = [int]$inactiveCount
                PasswordNeverExpires = [int]$neverExpireCount
                AvgAccountAgeDays = [double]$avgAccountAge
                UniqueTitles = [int]$uniqueTitles
                SecurityScore = [double]$securityScore
                HealthStatus = [string]$healthStatus
                Recommendations = [string]($recommendations -join "; ")
            })
        }
        
        Write-ADReportLog -Message "Department statistics analysis completed. $($DepartmentStats.Count) departments found." -Type Info -Terminal
        return $DepartmentStats | Sort-Object SecurityScore -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing department statistics: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DepartmentSecurityRisks {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing security risks by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit erweiterten Eigenschaften - nur standardmäßige AD-Attribute verwenden
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, 
                                                  LockedOut, AdminCount, ServicePrincipalNames, DoesNotRequirePreAuth,
                                                  TrustedForDelegation, AllowReversiblePasswordEncryption, DisplayName,
                                                  SamAccountName, whenCreated, Description -ErrorAction Stop
        
        # Erstelle leeres Array für Ergebnisse
        $DepartmentRisks = [System.Collections.ArrayList]@()
        
        # Gruppiere nach Abteilung - verwende nur das Standard-Department Attribut
        $DepartmentGroups = $Users | Group-Object { 
            $dept = $_.Department
            if ([string]::IsNullOrWhiteSpace($dept)) {
                "(No Department)"
            } else {
                $dept
            }
        }
        
        foreach ($dept in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($dept.Name)) { "(No Department)" } else { [string]$dept.Name }
            $deptUsers = $dept.Group
            $totalUsers = ($deptUsers | Measure-Object).Count
            
            if ($totalUsers -eq 0) { continue }
            
            # Risiko-Metriken berechnen mit sauberer Typisierung
            $adminUsers = [int]($deptUsers | Where-Object { $_.AdminCount -eq 1 } | Measure-Object).Count
            $serviceAccounts = [int]($deptUsers | Where-Object { 
                $null -ne $_.ServicePrincipalNames -and @($_.ServicePrincipalNames).Count -gt 0 
            } | Measure-Object).Count
            
            $kerberoastable = [int]($deptUsers | Where-Object { 
                $null -ne $_.ServicePrincipalNames -and @($_.ServicePrincipalNames).Count -gt 0 -and $_.Enabled -eq $true 
            } | Measure-Object).Count
            
            $asrepRoastable = [int]($deptUsers | Where-Object { $_.DoesNotRequirePreAuth -eq $true } | Measure-Object).Count
            $delegationEnabled = [int]($deptUsers | Where-Object { $_.TrustedForDelegation -eq $true } | Measure-Object).Count
            $reversiblePwd = [int]($deptUsers | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true } | Measure-Object).Count
            $neverExpire = [int]($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true } | Measure-Object).Count
            $oldPasswords = [int]($deptUsers | Where-Object { 
                $_.PasswordLastSet -and $_.PasswordLastSet -lt (Get-Date).AddDays(-180) 
            } | Measure-Object).Count
            
            $inactiveUsers = [int]($deptUsers | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90) 
            } | Measure-Object).Count
            
            $disabledUsers = [int]($deptUsers | Where-Object { $_.Enabled -eq $false } | Measure-Object).Count
            $lockedUsers = [int]($deptUsers | Where-Object { $_.LockedOut -eq $true } | Measure-Object).Count
            
            # Risiko-Score berechnen (gewichtete Bewertung)
            [int]$riskScore = 0
            $riskScore += $adminUsers * 5       # Privilegierte Konten haben höchstes Risiko
            $riskScore += $kerberoastable * 4   # Kerberoastable Konten sind kritisch
            $riskScore += $asrepRoastable * 4   # ASREPRoastable Konten sind kritisch
            $riskScore += $delegationEnabled * 3 # Delegation ist riskant
            $riskScore += $reversiblePwd * 5    # Reversible Passwörter sind kritisch
            $riskScore += [int]($neverExpire * 1.5)     # Passwörter die nie ablaufen
            $riskScore += [int]($oldPasswords * 1.2)    # Alte Passwörter
            $riskScore += [int]($inactiveUsers * 0.8)   # Inaktive Konten
            
            # Normalisierung auf Benutzerzahl für faire Vergleichbarkeit
            $normalizedRiskScore = if ($totalUsers -gt 0) {
                [math]::Round(($riskScore / $totalUsers) * 10, 1)
            } else {
                0
            }
            
            # Risiko-Level bestimmen (String für GUI-Anzeige)
            [string]$riskLevel = switch ($normalizedRiskScore) {
                {$_ -ge 30} { "Critical" }
                {$_ -ge 20} { "High" }
                {$_ -ge 10} { "Medium" }
                {$_ -gt 2} { "Low" }
                default { "Minimal" }
            }
            
            # Sicherheitsempfehlungen generieren
            $recommendations = [System.Collections.ArrayList]@()
            if ($adminUsers -gt 0) {
                $null = $recommendations.Add("Review privileged account usage")
            }
            if ($kerberoastable -gt 0) {
                $null = $recommendations.Add("Audit service accounts with SPNs")
            }
            if ($asrepRoastable -gt 0) {
                $null = $recommendations.Add("Enable Kerberos pre-authentication")
            }
            if ($reversiblePwd -gt 0) {
                $null = $recommendations.Add("Disable reversible password encryption")
            }
            if ($oldPasswords -gt ($totalUsers * 0.3)) {
                $null = $recommendations.Add("Enforce password rotation policy")
            }
            if ($inactiveUsers -gt ($totalUsers * 0.2)) {
                $null = $recommendations.Add("Review and disable inactive accounts")
            }
            
            $recommendationString = if ($recommendations.Count -gt 0) {
                [string]($recommendations -join "; ")
            } else {
                [string]"Continue monitoring security posture"
            }
            
            # Security Impact Assessment
            $securityImpact = [string]"Medium"
            if ($normalizedRiskScore -ge 30) {
                $securityImpact = "Critical - Immediate action required"
            } elseif ($normalizedRiskScore -ge 20) {
                $securityImpact = "High - Priority remediation needed"
            } elseif ($normalizedRiskScore -ge 10) {
                $securityImpact = "Medium - Monitor and improve"
            } else {
                $securityImpact = "Low - Standard monitoring sufficient"
            }
            
            # Füge Risiken zum Array hinzu mit expliziter Typisierung
            $deptRiskObject = [PSCustomObject]@{
                Department = [string]$deptName
                TotalUsers = [int]$totalUsers
                AdminUsers = [int]$adminUsers
                ServiceAccounts = [int]$serviceAccounts
                Kerberoastable = [int]$kerberoastable
                ASREPRoastable = [int]$asrepRoastable
                DelegationEnabled = [int]$delegationEnabled
                ReversiblePasswords = [int]$reversiblePwd
                PasswordNeverExpires = [int]$neverExpire
                OldPasswords = [int]$oldPasswords
                InactiveUsers = [int]$inactiveUsers
                DisabledUsers = [int]$disabledUsers
                LockedUsers = [int]$lockedUsers
                RiskScore = [double]$normalizedRiskScore
                RiskLevel = [string]$riskLevel
                SecurityImpact = [string]$securityImpact
                Recommendations = [string]$recommendationString
            }
            
            $null = $DepartmentRisks.Add($deptRiskObject)
        }
        
        # Statistiken für Logging
        $totalDepartments = $DepartmentRisks.Count
        $criticalDepts = ($DepartmentRisks | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highRiskDepts = ($DepartmentRisks | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumRiskDepts = ($DepartmentRisks | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        
        Write-ADReportLog -Message "Department security risk analysis completed. $totalDepartments departments analyzed." -Type Info -Terminal
        Write-ADReportLog -Message "Risk distribution - Critical: $criticalDepts, High: $highRiskDepts, Medium: $mediumRiskDepts" -Type Info -Terminal
        
        if ($criticalDepts -gt 0) {
            Write-ADReportLog -Message "WARNING: $criticalDepts departments with critical security risks require immediate attention!" -Type Warning -Terminal
        }
        
        return $DepartmentRisks | Sort-Object RiskScore -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing department security risks: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ASREPRoastableAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing ASREPRoastable accounts (pre-authentication disabled)..." -Type Info -Terminal
        
        # Lade alle Benutzer mit deaktivierter Kerberos-Vorauthentifizierung
        $ASREPUsers = Get-ADUser -Filter "DoesNotRequirePreAuth -eq `$true" -Properties DoesNotRequirePreAuth, DisplayName,
                                                                                        SamAccountName, Enabled, PasswordLastSet,
                                                                                        PasswordNeverExpires, LastLogonDate, AdminCount,
                                                                                        Description, whenCreated, Department -ErrorAction Stop
        
        $ASREPRoastableAccounts = [System.Collections.ArrayList]@()
        foreach ($user in $ASREPUsers) {
            # Risiko-Bewertung
            $riskFactors = @()
            [int]$riskLevel = 5  # Basis-Risiko für ASREP Roastable ist hoch
            
            if ($user.AdminCount -eq 1) {
                $riskFactors += "Privileged Account"
                $riskLevel += 3
            }
            
            if ($user.Enabled -eq $true) {
                $riskFactors += "Account Enabled"
                $riskLevel += 1
            }
            
            if ($user.PasswordNeverExpires) {
                $riskFactors += "Password Never Expires"
                $riskLevel += 2
            }
            
            if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $riskFactors += "Recently Active"
                $riskLevel += 1
            }
            
            # Risiko-Level als String für GUI-Anzeige
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 8} { "Critical" }
                {$_ -ge 6} { "High" }
                {$_ -ge 4} { "Medium" }
                default { "Low" }
            }
            
            # Risiko-Faktoren als String zusammenfassen
            $riskFactorsString = if ($riskFactors.Count -gt 0) {
                $riskFactors -join "; "
            } else {
                "Basic ASREP vulnerability"
            }
            
            $null = $ASREPRoastableAccounts.Add([PSCustomObject]@{
                DisplayName = if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { $user.SamAccountName } else { $user.DisplayName }
                SamAccountName = $user.SamAccountName
                Department = if ([string]::IsNullOrWhiteSpace($user.Department)) { "Not Specified" } else { $user.Department }
                Enabled = $user.Enabled
                DoesNotRequirePreAuth = $true
                PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet } else { "Never" }
                PasswordNeverExpires = $user.PasswordNeverExpires
                LastLogonDate = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                AdminAccount = ($user.AdminCount -eq 1)
                AccountAge = if ($user.whenCreated) { 
                    [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays) 
                } else { 
                    0 
                }
                RiskFactors = $riskFactorsString
                RiskLevel = $overallRisk
                RiskScore = $riskLevel
                Remediation = "Enable Kerberos pre-authentication immediately"
                SecurityImpact = "Account vulnerable to ASREPRoast attacks - password hash can be extracted offline"
                Description = if ([string]::IsNullOrWhiteSpace($user.Description)) { "No description available" } else { $user.Description }
            })
        }
        
        # Status-Update für GUI
        try {
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "ASREPRoastable accounts analysis completed. $($ASREPRoastableAccounts.Count) vulnerable accounts found."
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI status: $($_.Exception.Message)" -Type Warning
        }
        
        Write-ADReportLog -Message "ASREPRoastable accounts analysis completed. $($ASREPRoastableAccounts.Count) accounts found." -Type Info -Terminal
        return $ASREPRoastableAccounts | Sort-Object RiskScore -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing ASREPRoastable accounts: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Fehler-Status für GUI
        try {
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI error status: $($_.Exception.Message)" -Type Warning
        }
        
        return @()
    }
}

Function Get-DelegationAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing delegation settings..." -Type Info -Terminal
        
        $DelegatedObjects = @()
        
        # Unconstrained Delegation Analysis (Critical Risk)
        Write-ADReportLog -Message "Scanning for unconstrained delegation..." -Type Info
        $UnconstrainedDelegation = Get-ADObject -Filter {TrustedForDelegation -eq $true -and TrustedToAuthForDelegation -eq $false} `
                                                -Properties Name, ObjectClass, DistinguishedName, whenCreated, whenChanged, Description, ObjectSid -ErrorAction SilentlyContinue
        
        foreach ($obj in $UnconstrainedDelegation) {
            $pathParts = $obj.DistinguishedName -split ','
            $containerPath = ""
            for ($i = 1; $i -lt $pathParts.Count; $i++) {
                $containerPath += ($pathParts[$i] -replace '^(CN=|OU=|DC=)', '') + " → "
            }
            $containerPath = $containerPath.TrimEnd(" → ")
            
            $ageInDays = 0
            if ($obj.whenCreated) {
                $ageInDays = [math]::Round((New-TimeSpan -Start $obj.whenCreated -End (Get-Date)).TotalDays)
            }
            
            $DelegatedObjects += [PSCustomObject]@{
                TreeStructure = "🔴 UNCONSTRAINED DELEGATION"
                ObjectName = [string]$obj.Name
                ObjectType = [string]$obj.ObjectClass
                DelegationType = [string]"Unconstrained Delegation"
                ContainerPath = [string]$containerPath
                AllowedServices = [string]"ALL SERVICES (No Restrictions)"
                RiskLevel = [string]"Critical"
                SecurityImpact = [string]"Can impersonate any user to any service - Golden Ticket equivalent"
                Created = if ($obj.whenCreated) { $obj.whenCreated.ToString("dd.MM.yyyy HH:mm") } else { "Unknown" }
                LastModified = if ($obj.whenChanged) { $obj.whenChanged.ToString("dd.MM.yyyy HH:mm") } else { "Never"  }
                AccountAge = [int]$ageInDays
                Description = if ([string]::IsNullOrWhiteSpace($obj.Description)) { "No description" } else { [string]$obj.Description }
                Recommendation = [string]"IMMEDIATE: Remove unconstrained delegation or convert to constrained delegation"
                ComplianceNote = [string]"Violates security best practices - enables privilege escalation"
            }
        }
        
        # Constrained Delegation Analysis (Medium Risk)
        Write-ADReportLog -Message "Scanning for constrained delegation..." -Type Info
        $ConstrainedDelegation = Get-ADObject -Filter {TrustedToAuthForDelegation -eq $true} `
                                             -Properties Name, ObjectClass, DistinguishedName, "msDS-AllowedToDelegateTo", whenCreated, whenChanged, Description -ErrorAction SilentlyContinue
        
        foreach ($obj in $ConstrainedDelegation) {
            $pathParts = $obj.DistinguishedName -split ','
            $containerPath = ""
            for ($i = 1; $i -lt $pathParts.Count; $i++) {
                $containerPath += ($pathParts[$i] -replace '^(CN=|OU=|DC=)', '') + " → "
            }
            $containerPath = $containerPath.TrimEnd(" → ")
            
            $allowedServices = ""
            if ($obj.'msDS-AllowedToDelegateTo' -and $obj.'msDS-AllowedToDelegateTo'.Count -gt 0) {
                $allowedServices = ($obj.'msDS-AllowedToDelegateTo' | ForEach-Object { "  ├─ $_" }) -join "`n"
            } else {
                $allowedServices = "  └─ No services specified"
            }
            
            $serviceCount = if ($obj.'msDS-AllowedToDelegateTo') { $obj.'msDS-AllowedToDelegateTo'.Count } else { 0 }
            $riskLevel = if ($serviceCount -gt 5) { "High" } elseif ($serviceCount -gt 2) { "Medium" } else { "Low" }
            
            $ageInDays = 0
            if ($obj.whenCreated) {
                $ageInDays = [math]::Round((New-TimeSpan -Start $obj.whenCreated -End (Get-Date)).TotalDays)
            }
            
            $DelegatedObjects += [PSCustomObject]@{
                TreeStructure = "🟡 CONSTRAINED DELEGATION`n$allowedServices"
                ObjectName = [string]$obj.Name
                ObjectType = [string]$obj.ObjectClass
                DelegationType = [string]"Constrained Delegation"
                ContainerPath = [string]$containerPath
                AllowedServices = [string]"$serviceCount service(s) specified"
                RiskLevel = [string]$riskLevel
                SecurityImpact = [string]"Limited impersonation to specified services only"
                Created = if ($obj.whenCreated) { $obj.whenCreated.ToString("dd.MM.yyyy HH:mm") } else { "Unknown" }
                LastModified = if ($obj.whenChanged) { $obj.whenChanged.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                AccountAge = [int]$ageInDays
                Description = if ([string]::IsNullOrWhiteSpace($obj.Description)) { "No description" } else { [string]$obj.Description }
                Recommendation = [string]"Review allowed services and minimize delegation scope"
                ComplianceNote = [string]"Monitor for service enumeration and lateral movement"
            }
        }
        
        # Resource-Based Constrained Delegation Analysis (High Risk)
        Write-ADReportLog -Message "Scanning for resource-based constrained delegation..." -Type Info
        $ResourceBasedDelegation = Get-ADObject -Filter {msDS-AllowedToActOnBehalfOfOtherIdentity -like "*"} `
                                               -Properties Name, ObjectClass, DistinguishedName, "msDS-AllowedToActOnBehalfOfOtherIdentity", whenCreated, whenChanged, Description -ErrorAction SilentlyContinue
        
        foreach ($obj in $ResourceBasedDelegation) {
            $pathParts = $obj.DistinguishedName -split ','
            $containerPath = ""
            for ($i = 1; $i -lt $pathParts.Count; $i++) {
                $containerPath += ($pathParts[$i] -replace '^(CN=|OU=|DC=)', '') + " → "
            }
            $containerPath = $containerPath.TrimEnd(" → ")
            
            # Parse the security descriptor to get allowed principals
            $allowedPrincipals = "  └─ Security descriptor configured"
            try {
                if ($obj.'msDS-AllowedToActOnBehalfOfOtherIdentity') {
                    $allowedPrincipals = "  ├─ Resource-based delegation active`n  └─ Check security descriptor for details"
                }
            } catch {
                $allowedPrincipals = "  └─ Could not parse security descriptor"
            }
            
            $ageInDays = 0
            if ($obj.whenCreated) {
                $ageInDays = [math]::Round((New-TimeSpan -Start $obj.whenCreated -End (Get-Date)).TotalDays)
            }
            
            $DelegatedObjects += [PSCustomObject]@{
                TreeStructure = "🟠 RESOURCE-BASED DELEGATION`n$allowedPrincipals"
                ObjectName = [string]$obj.Name
                ObjectType = [string]$obj.ObjectClass
                DelegationType = [string]"Resource-Based Constrained Delegation"
                ContainerPath = [string]$containerPath
                AllowedServices = [string]"Resource-controlled delegation"
                RiskLevel = [string]"High"
                SecurityImpact = [string]"Target resource controls which principals can delegate"
                Created = if ($obj.whenCreated) { $obj.whenCreated.ToString("dd.MM.yyyy HH:mm") } else { "Unknown" }
                LastModified = if ($obj.whenChanged) { $obj.whenChanged.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                AccountAge = [int]$ageInDays
                Description = if ([string]::IsNullOrWhiteSpace($obj.Description)) { "No description" } else { [string]$obj.Description }
                Recommendation = [string]"Audit resource-based delegation permissions and principals"
                ComplianceNote = [string]"Modern delegation method - review configuration regularly"
            }
        }
        
        # Sortierung nach Risiko und dann nach Namen
        $SortedResults = $DelegatedObjects | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, ObjectName
        
        # Status-Update für GUI
        try {
            if ($Global:TextBlockStatus) {
                $criticalCount = ($DelegatedObjects | Where-Object { $_.RiskLevel -eq "Critical" }).Count
                $highCount = ($DelegatedObjects | Where-Object { $_.RiskLevel -eq "High" }).Count
                $mediumCount = ($DelegatedObjects | Where-Object { $_.RiskLevel -eq "Medium" }).Count
                
                $Global:TextBlockStatus.Text = "Delegation analysis completed. $($DelegatedObjects.Count) delegated objects found. Critical: $criticalCount, High: $highCount, Medium: $mediumCount"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI status: $($_.Exception.Message)" -Type Warning
        }
        
        Write-ADReportLog -Message "Delegation analysis completed. $($DelegatedObjects.Count) delegated objects found." -Type Info -Terminal
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing delegation settings: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Fehler-Status für GUI
        try {
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI error status: $($_.Exception.Message)" -Type Warning
        }
        
        return @()
    }
}

# --- Erweiterte Privilegien-Eskalation Analyse ---
Function Get-DCSyncRights {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DCSync rights..." -Type Info -Terminal
        
        $Domain = Get-ADDomain
        $DomainDN = $Domain.DistinguishedName
        $Results = @()
        
        # Get ACL for domain object
        try {
            $DomainACL = Get-Acl "AD:\$DomainDN" -ErrorAction Stop
        } catch {
            Write-ADReportLog -Message "Could not access domain ACL. Requires appropriate permissions." -Type Warning -Terminal
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"🔐 DCSync Rights Analysis"
                Identity = [string]"Access Error"
                ObjectType = [string]"Permission Error"
                Permissions = [string]"Could not read domain ACL"
                IsPrivileged = [string]"Unknown"
                HasFullDCSync = [string]"False"
                RiskLevel = [string]"Critical"
                SecurityImpact = [string]"Cannot verify DCSync permissions - potential security blind spot"
                Remediation = [string]"Run analysis with Domain Admin privileges to access domain ACL"
                LastModified = [string]"Unknown"
                ComplianceNote = [string]"ACL access required for proper security assessment"
            }
            return $Results
        }
        
        Write-ADReportLog -Message "Domain ACL accessed successfully. Analyzing replication permissions..." -Type Info
        
        # DCSync requires specific replication permissions
        $ReplicationGUIDs = @{
            "DS-Replication-Get-Changes" = "1131f6aa-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-All" = "1131f6ad-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-In-Filtered-Set" = "89e95b76-444d-4c62-991a-0facbeda640c"
        }
        
        # Define expected/safe principals (English and German names)
        $ExpectedPrincipals = @(
            'Domain Controllers',
            'Domänencontroller',
            'Enterprise Domain Controllers',
            'Organisations-Domänencontroller',
            'Administrators',
            'Administratoren',
            'Domain Admins',
            'Domänen-Admins',
            'Enterprise Admins',
            'Organisations-Admins',
            'SYSTEM',
            'NT AUTHORITY\SYSTEM'
        )
        
        # Collect all replication permissions
        $ReplicationPermissions = @()
        
        foreach ($ace in $DomainACL.Access) {
            if ($ace.AccessControlType -eq "Allow" -and $ace.ObjectType) {
                $aceObjectType = $ace.ObjectType.ToString()
                
                foreach ($guid in $ReplicationGUIDs.GetEnumerator()) {
                    if ($aceObjectType -eq $guid.Value) {
                        $identity = $ace.IdentityReference.Value
                        
                        # Resolve identity details
                        try {
                            $identityName = if ($identity.Contains('\')) { $identity.Split('\')[1] } else { $identity }
                            $adObject = Get-ADObject -Filter "SamAccountName -eq '$identityName'" -Properties ObjectClass, AdminCount, whenChanged, Description -ErrorAction SilentlyContinue
                            
                            if (-not $adObject) {
                                # Try to find by name if SamAccountName failed
                                $adObject = Get-ADObject -Filter "Name -eq '$identityName'" -Properties ObjectClass, AdminCount, whenChanged, Description -ErrorAction SilentlyContinue
                            }
                            
                            $objectType = if ($adObject) { 
                                if ($adObject.ObjectClass -is [array]) { $adObject.ObjectClass[-1] } else { $adObject.ObjectClass }
                            } else { "Unknown" }
                            $isAdmin = if ($adObject -and $adObject.AdminCount -eq 1) { "Yes" } else { "No" }
                            $lastModified = if ($adObject -and $adObject.whenChanged) { $adObject.whenChanged.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
                            $description = if ($adObject -and $adObject.Description) { $adObject.Description } else { "No description" }
                        } catch {
                            $objectType = "Unknown"
                            $isAdmin = "Unknown"
                            $lastModified = "Unknown"
                            $description = "Could not resolve identity"
                        }
                        
                        # Check if identity is expected/safe
                        $isExpected = $false
                        foreach ($expectedPrincipal in $ExpectedPrincipals) {
                            if ($identity -like "*$expectedPrincipal*") {
                                $isExpected = $true
                                break
                            }
                        }
                        
                        $ReplicationPermissions += [PSCustomObject]@{
                            Identity = $identity
                            Permission = $guid.Key
                            ObjectType = $objectType
                            IsPrivileged = $isAdmin
                            IsExpected = $isExpected
                            LastModified = $lastModified
                            Description = $description
                        }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Found $($ReplicationPermissions.Count) replication permissions. Analyzing DCSync capabilities..." -Type Info
        
        # Group by identity and analyze DCSync capability
        $IdentityGroups = $ReplicationPermissions | Group-Object Identity
        
        foreach ($group in $IdentityGroups) {
            $identity = $group.Name
            $permissions = $group.Group
            $firstPerm = $permissions[0]
            
            # Check for full DCSync capability
            $hasGetChanges = $permissions | Where-Object { $_.Permission -eq "DS-Replication-Get-Changes" }
            $hasGetChangesAll = $permissions | Where-Object { $_.Permission -eq "DS-Replication-Get-Changes-All" }
            $hasGetChangesFiltered = $permissions | Where-Object { $_.Permission -eq "DS-Replication-Get-Changes-In-Filtered-Set" }
            
            $hasFullDCSync = ($hasGetChanges -and $hasGetChangesAll)
            $permissionsList = ($permissions.Permission | Sort-Object -Unique) -join ", "
            
            # Determine risk level
            $riskLevel = "Low"
            $securityImpact = "Standard replication permission"
            $recommendation = "Monitor for unauthorized changes"
            
            if ($hasFullDCSync) {
                if (-not $firstPerm.IsExpected) {
                    $riskLevel = "Critical"
                    $securityImpact = "Full DCSync capability - can extract all domain credentials including KRBTGT"
                    $recommendation = "IMMEDIATE ACTION: Remove DCSync rights from this identity unless explicitly required"
                } else {
                    $riskLevel = "Medium"
                    $securityImpact = "Expected DCSync capability for system account"
                    $recommendation = "Monitor for unauthorized usage and ensure account security"
                }
            } elseif ($hasGetChanges) {
                if (-not $firstPerm.IsExpected) {
                    $riskLevel = "High"
                    $securityImpact = "Partial replication rights - potential for credential extraction"
                    $recommendation = "Review necessity and remove if not required"
                } else {
                    $riskLevel = "Low"
                    $securityImpact = "Expected partial replication permission"
                    $recommendation = "Standard monitoring"
                }
            }
            
            # Create tree structure for GUI display
            $treeStructure = if ($hasFullDCSync) {
                "🔥 DCSync Rights > Full Capability > $identity"
            } elseif ($hasGetChanges) {
                "⚠️ DCSync Rights > Partial Rights > $identity"
            } else {
                "ℹ️ DCSync Rights > Other Replication > $identity"
            }
            
            $Results += [PSCustomObject]@{
                TreeStructure = [string]$treeStructure
                Identity = [string]$identity
                ObjectType = [string]$firstPerm.ObjectType
                Permissions = [string]$permissionsList
                IsPrivileged = [string]$firstPerm.IsPrivileged
                HasFullDCSync = [string]$(if ($hasFullDCSync) { "Yes" } else { "No" })
                RiskLevel = [string]$riskLevel
                SecurityImpact = [string]$securityImpact
                Remediation = [string]$recommendation
                LastModified = [string]$firstPerm.LastModified
                ComplianceNote = [string]$(if ($firstPerm.IsExpected) { "Expected system permission" } else { "Requires security review" })
                Description = [string]$firstPerm.Description
            }
        }
        
        # Add summary information
        if ($Results.Count -eq 0) {
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"ℹ️ DCSync Rights Analysis > Summary"
                Identity = [string]"No DCSync Rights Found"
                ObjectType = [string]"Analysis Result"
                Permissions = [string]"None detected"
                IsPrivileged = [string]"N/A"
                HasFullDCSync = [string]"No"
                RiskLevel = [string]"Info"
                SecurityImpact = [string]"No explicit DCSync permissions found - using default permissions only"
                Remediation = [string]"Continue monitoring for new permissions"
                LastModified = [string]"N/A"
                ComplianceNote = [string]"Normal state - DCSync typically restricted to system accounts"
                Description = [string]"Analysis completed successfully"
            }
        } else {
            # Add statistics
            $criticalCount = ($Results | Where-Object { $_.RiskLevel -eq "Critical" }).Count
            $highCount = ($Results | Where-Object { $_.RiskLevel -eq "High" }).Count
            $fullDCSyncCount = ($Results | Where-Object { $_.HasFullDCSync -eq "Yes" }).Count
            
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"📊 DCSync Rights Analysis > Statistics"
                Identity = [string]"Analysis Summary"
                ObjectType = [string]"Summary"
                Permissions = [string]"$($Results.Count) identities analyzed"
                IsPrivileged = [string]"$($Results | Where-Object { $_.IsPrivileged -eq 'Yes' }).Count privileged"
                HasFullDCSync = [string]"$fullDCSyncCount with full capability"
                RiskLevel = [string]$(if ($criticalCount -gt 0) { "Critical" } elseif ($highCount -gt 0) { "High" } else { "Medium" })
                SecurityImpact = [string]"Critical: $criticalCount, High: $highCount, Full DCSync: $fullDCSyncCount"
                Remediation = [string]"Review all non-expected identities with DCSync rights"
                LastModified = [string]$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                ComplianceNote = [string]"Regular monitoring required for DCSync permissions"
                Description = [string]"Complete analysis of domain replication permissions"
            }
        }
        
        # Sort results for tree display
        $SortedResults = $Results | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    "Info" { 5 }
                    default { 6 }
                }
            }
        }, TreeStructure
        
        Write-ADReportLog -Message "DCSync rights analysis completed. Found $($Results.Count-1) identities with replication permissions." -Type Info -Terminal
        
        # Log security findings
        if ($Results.Count -gt 1) {
            $securityFindings = $Results | Where-Object { $_.RiskLevel -in @("Critical", "High") }
            if ($securityFindings.Count -gt 0) {
                Write-ADReportLog -Message "SECURITY ALERT: $($securityFindings.Count) high-risk DCSync permissions found" -Type Warning -Terminal
            }
        }
        
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing DCSync rights: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Return error result
        return @([PSCustomObject]@{
            TreeStructure = [string]"❌ DCSync Rights Analysis > Error"
            Identity = [string]"Analysis Failed"
            ObjectType = [string]"Error"
            Permissions = [string]"Could not analyze"
            IsPrivileged = [string]"Unknown"
            HasFullDCSync = [string]"Unknown"
            RiskLevel = [string]"Critical"
            SecurityImpact = [string]"Analysis failed - security status unknown"
            Remediation = [string]"Retry analysis with appropriate permissions"
            LastModified = [string]$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            ComplianceNote = [string]"Analysis failure requires investigation"
            Description = [string]$ErrorMessage
        })
    }
}

Function Get-SchemaAdminPaths {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Schema Admin access paths..." -Type Info -Terminal
        
        $Results = @()
        
        # Get Schema Admins group members - try both English and German names
        $SchemaAdmins = $null
        $SchemaAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.SchemaAdmins
        
        if ($SchemaAdminsGroup) {
            Write-ADReportLog -Message "Found Schema Admins group as: $($SchemaAdminsGroup.Name)" -Type Info -Terminal
            try {
                $SchemaAdmins = Get-ADGroupMember -Identity $SchemaAdminsGroup -Recursive -ErrorAction SilentlyContinue
            } catch {
                Write-ADReportLog -Message "Could not enumerate Schema Admins members: $($_.Exception.Message)" -Type Warning
            }
        }
        
        if (-not $SchemaAdminsGroup) {
            Write-ADReportLog -Message "Schema Admins group not found in domain." -Type Warning -Terminal
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"🔍 Schema Admins Analysis"
                PathType = [string]"Group Status"
                ObjectName = [string]"Schema Admins Group"
                ObjectType = [string]"Security Group"
                Status = [string]"Not Found"
                RiskLevel = [string]"Medium"
                SecurityImpact = [string]"Cannot verify Schema Admin membership - group might not exist or be renamed"
                LastActivity = [string]"Unknown"
                AccountAge = [string]"Unknown"
                PathDetails = [string]"Schema Admins group not found in domain structure"
                Recommendation = [string]"Verify Schema Admins group existence and naming conventions"
                ComplianceNote = [string]"Schema Admins group should exist but remain empty except during schema modifications"
            }
            return $Results
        }
        
        # Analyze current Schema Admins members
        if ($SchemaAdmins -and $SchemaAdmins.Count -gt 0) {
            Write-ADReportLog -Message "Found $($SchemaAdmins.Count) Schema Admin(s). Analyzing members..." -Type Info
            
            foreach ($admin in $SchemaAdmins) {
                try {
                    $userDetails = Get-ADUser -Identity $admin.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet, AdminCount, whenCreated, Description, Department -ErrorAction SilentlyContinue
                    
                    if ($userDetails) {
                        # Calculate account age
                        $ageInDays = 0
                        if ($userDetails.whenCreated) {
                            $ageInDays = [math]::Round((New-TimeSpan -Start $userDetails.whenCreated -End (Get-Date)).TotalDays)
                        }
                        
                        # Determine risk level based on activity and status
                        $riskLevel = "Critical"
                        $securityImpact = ""
                        $recommendation = ""
                        
                        if ($userDetails.Enabled) {
                            if ($userDetails.LastLogonDate -and $userDetails.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                                $riskLevel = "Critical"
                                $securityImpact = "Active Schema Admin with recent logon activity - full schema modification privileges"
                                $recommendation = "URGENT: Remove from Schema Admins immediately unless actively performing schema modifications"
                            } elseif ($userDetails.LastLogonDate -and $userDetails.LastLogonDate -gt (Get-Date).AddDays(-90)) {
                                $riskLevel = "High"
                                $securityImpact = "Active Schema Admin with moderate activity - potential security risk"
                                $recommendation = "Remove from Schema Admins group - account should only be added during schema changes"
                            } else {
                                $riskLevel = "High"
                                $securityImpact = "Enabled Schema Admin with minimal activity - unnecessary privilege retention"
                                $recommendation = "Remove from Schema Admins group and monitor for legitimate schema modification needs"
                            }
                        } else {
                            $riskLevel = "Medium"
                            $securityImpact = "Disabled Schema Admin account - reduced immediate risk but still holds privileged group membership"
                            $recommendation = "Remove disabled account from Schema Admins group"
                        }
                        
                        $Results += [PSCustomObject]@{
                            TreeStructure = [string]"🔴 CURRENT SCHEMA ADMIN"
                            PathType = [string]"Direct Member"
                            ObjectName = [string]$admin.Name
                            ObjectType = [string]$admin.ObjectClass
                            Status = if ($userDetails.Enabled) { [string]"Enabled" } else { [string]"Disabled" }
                            RiskLevel = [string]$riskLevel
                            SecurityImpact = [string]$securityImpact
                            LastActivity = if ($userDetails.LastLogonDate) { [string]$userDetails.LastLogonDate.ToString("MM/dd/yyyy HH:mm") } else { [string]"Never" }
                            AccountAge = [string]"$ageInDays days"
                            PathDetails = [string]"Direct member of Schema Admins group with full schema modification privileges"
                            Recommendation = [string]$recommendation
                            ComplianceNote = [string]"Schema Admins should be empty except during active schema modifications"
                        }
                    }
                } catch {
                    Write-ADReportLog -Message "Error analyzing Schema Admin member $($admin.SamAccountName): $($_.Exception.Message)" -Type Warning
                    
                    $Results += [PSCustomObject]@{
                        TreeStructure = [string]"🔴 SCHEMA ADMIN (ERROR)"
                        PathType = [string]"Direct Member"
                        ObjectName = [string]$admin.Name
                        ObjectType = [string]$admin.ObjectClass
                        Status = [string]"Analysis Failed"
                        RiskLevel = [string]"Unknown"
                        SecurityImpact = [string]"Cannot analyze account details - potential orphaned or corrupted member"
                        LastActivity = [string]"Unknown"
                        AccountAge = [string]"Unknown"
                        PathDetails = [string]"Schema Admin member with analysis errors - investigate manually"
                        Recommendation = [string]"Investigate account status and remove if orphaned"
                        ComplianceNote = [string]"Failed analysis indicates potential security issue"
                    }
                }
            }
        } else {
            # Schema Admins group exists but is empty (best practice)
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"✅ SCHEMA ADMINS COMPLIANCE"
                PathType = [string]"Empty Group"
                ObjectName = [string]$SchemaAdminsGroup.Name
                ObjectType = [string]"Security Group"
                Status = [string]"Empty (Compliant)"
                RiskLevel = [string]"Low"
                SecurityImpact = [string]"Schema Admins group is empty - follows security best practices"
                LastActivity = [string]"N/A"
                AccountAge = [string]"N/A"
                PathDetails = [string]"Schema Admins group exists but contains no members"
                Recommendation = [string]"Maintain empty state - only add members during active schema modifications"
                ComplianceNote = [string]"COMPLIANT: Schema Admins should remain empty except during schema changes"
            }
        }
        
        # Analyze Enterprise Admins as potential elevation path
        Write-ADReportLog -Message "Analyzing Enterprise Admin elevation paths to Schema Admin..." -Type Info
        $EnterpriseAdmins = $null
        $EnterpriseAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.EnterpriseAdmins
        
        if ($EnterpriseAdminsGroup) {
            try {
                $EnterpriseAdmins = Get-ADGroupMember -Identity $EnterpriseAdminsGroup -Recursive -ErrorAction SilentlyContinue
            } catch {
                Write-ADReportLog -Message "Could not enumerate Enterprise Admins members: $($_.Exception.Message)" -Type Warning
            }
        }
        
        if ($EnterpriseAdmins) {
            Write-ADReportLog -Message "Found $($EnterpriseAdmins.Count) Enterprise Admin(s). Checking elevation potential..." -Type Info
            
            foreach ($ea in $EnterpriseAdmins) {
                # Check if not already in Schema Admins
                $isAlreadySchemaAdmin = $false
                if ($SchemaAdmins) {
                    $isAlreadySchemaAdmin = $SchemaAdmins.SamAccountName -contains $ea.SamAccountName
                }
                
                if (-not $isAlreadySchemaAdmin) {
                    try {
                        $userDetails = Get-ADUser -Identity $ea.SamAccountName -Properties Enabled, LastLogonDate, whenCreated, Description -ErrorAction SilentlyContinue
                        
                        if ($userDetails) {
                            # Calculate account age
                            $ageInDays = 0
                            if ($userDetails.whenCreated) {
                                $ageInDays = [math]::Round((New-TimeSpan -Start $userDetails.whenCreated -End (Get-Date)).TotalDays)
                            }
                            
                            # Determine risk level for elevation potential
                            $riskLevel = "High"
                            $securityImpact = ""
                            
                            if ($userDetails.Enabled) {
                                if ($userDetails.LastLogonDate -and $userDetails.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                                    $riskLevel = "High"
                                    $securityImpact = "Active Enterprise Admin can elevate to Schema Admin privileges at any time"
                                } else {
                                    $riskLevel = "Medium"
                                    $securityImpact = "Enterprise Admin with elevation potential but minimal recent activity"
                                }
                            } else {
                                $riskLevel = "Low"
                                $securityImpact = "Disabled Enterprise Admin - reduced elevation risk"
                            }
                            
                            $Results += [PSCustomObject]@{
                                TreeStructure = [string]"🟡 ELEVATION PATH"
                                PathType = [string]"Enterprise Admin"
                                ObjectName = [string]$ea.Name
                                ObjectType = [string]$ea.ObjectClass
                                Status = if ($userDetails.Enabled) { [string]"Enabled" } else { [string]"Disabled" }
                                RiskLevel = [string]$riskLevel
                                SecurityImpact = [string]$securityImpact
                                LastActivity = if ($userDetails.LastLogonDate) { [string]$userDetails.LastLogonDate.ToString("MM/dd/yyyy HH:mm") } else { [string]"Never" }
                                AccountAge = [string]"$ageInDays days"
                                PathDetails = [string]"Enterprise Admin can add themselves to Schema Admins group"
                                Recommendation = [string]"Monitor Enterprise Admin activities and restrict membership"
                                ComplianceNote = [string]"Enterprise Admins have inherent ability to elevate to Schema Admin"
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error analyzing Enterprise Admin member $($ea.SamAccountName): $($_.Exception.Message)" -Type Warning
                    }
                }
            }
        }
        
        # Check for users with explicit permissions to modify Schema Admins group
        Write-ADReportLog -Message "Checking for explicit permissions to modify Schema Admins group..." -Type Info
        try {
            $SchemaAdminsACL = Get-ACL -Path "AD:\$($SchemaAdminsGroup.DistinguishedName)" -ErrorAction SilentlyContinue
            
            if ($SchemaAdminsACL) {
                $ExplicitPermissions = $SchemaAdminsACL.Access | Where-Object {
                    $_.ActiveDirectoryRights -match "WriteProperty|GenericWrite|WriteDacl|WriteOwner" -and
                    $_.AccessControlType -eq "Allow" -and
                    $_.IdentityReference -notmatch "SYSTEM|Administrators|Enterprise Admins|Schema Admins"
                }
                
                foreach ($perm in $ExplicitPermissions) {
                    $Results += [PSCustomObject]@{
                        TreeStructure = [string]"🔶 EXPLICIT PERMISSION"
                        PathType = [string]"ACL Permission"
                        ObjectName = [string]$perm.IdentityReference
                        ObjectType = [string]"Security Principal"
                        Status = [string]"Has Permissions"
                        RiskLevel = [string]"High"
                        SecurityImpact = [string]"Has explicit permissions to modify Schema Admins group membership"
                        LastActivity = [string]"Unknown"
                        AccountAge = [string]"Unknown"
                        PathDetails = [string]"Permission: $($perm.ActiveDirectoryRights) on Schema Admins group"
                        Recommendation = [string]"Review and remove unnecessary explicit permissions"
                        ComplianceNote = [string]"Non-standard permissions detected on Schema Admins group"
                    }
                }
            }
        } catch {
            Write-ADReportLog -Message "Could not analyze Schema Admins ACL: $($_.Exception.Message)" -Type Warning
        }
        
        # Add summary information
        $totalDirectMembers = ($Results | Where-Object { $_.PathType -eq "Direct Member" }).Count
        $totalElevationPaths = ($Results | Where-Object { $_.PathType -eq "Enterprise Admin" }).Count
        $totalExplicitPerms = ($Results | Where-Object { $_.PathType -eq "ACL Permission" }).Count
        
        $Results += [PSCustomObject]@{
            TreeStructure = [string]"📊 ANALYSIS SUMMARY"
            PathType = [string]"Summary"
            ObjectName = [string]"Schema Admin Security Assessment"
            ObjectType = [string]"Report Summary"
            Status = [string]"Analysis Complete"
            RiskLevel = if ($totalDirectMembers -gt 0) { [string]"Critical" } elseif ($totalElevationPaths -gt 0) { [string]"High" } else { [string]"Low" }
            SecurityImpact = [string]"Direct Members: $totalDirectMembers | Elevation Paths: $totalElevationPaths | Explicit Permissions: $totalExplicitPerms"
            LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
            AccountAge = [string]"N/A"
            PathDetails = [string]"Complete analysis of Schema Admin access paths and security risks"
            Recommendation = if ($totalDirectMembers -gt 0) { 
                [string]"CRITICAL: Remove all direct Schema Admin members immediately" 
            } else { 
                [string]"Monitor Enterprise Admin activities and maintain Schema Admins as empty group" 
            }
            ComplianceNote = [string]"Schema Admins should only contain members during active schema modification periods"
        }
        
        Write-ADReportLog -Message "Schema Admin paths analysis completed. $($Results.Count) findings generated." -Type Info -Terminal
        return $Results | Sort-Object @{Expression={
            switch($_.RiskLevel) {
                "Critical" { 1 }
                "High" { 2 }
                "Medium" { 3 }
                "Low" { 4 }
                default { 5 }
            }
        }}, TreeStructure
        
    } catch {
        $ErrorMessage = "Error analyzing Schema Admin paths: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        return @([PSCustomObject]@{
            TreeStructure = [string]"❌ ANALYSIS ERROR"
            PathType = [string]"Error"
            ObjectName = [string]"Schema Admin Analysis"
            ObjectType = [string]"Error Report"
            Status = [string]"Failed"
            RiskLevel = [string]"Unknown"
            SecurityImpact = [string]"Analysis failed - cannot assess Schema Admin security"
            LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
            AccountAge = [string]"N/A"
            PathDetails = [string]"Error: $ErrorMessage"
            Recommendation = [string]"Retry analysis or investigate manually"
            ComplianceNote = [string]"Failed security analysis requires immediate attention"
        })
    }
}

Function Get-CertificateSecurityAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Certificate Services security..." -Type Info -Terminal
        
        $CertAnalysis = @()
        
        # Check if AD CS is installed
        try {
            $Domain = Get-ADDomain
            $ConfigDN = "CN=Configuration,$($Domain.DistinguishedName)"
            $CertConfig = Get-ADObject -Filter "objectClass -eq 'pKIEnrollmentService'" -SearchBase $ConfigDN -Properties * -ErrorAction SilentlyContinue
        } catch {
            $CertConfig = $null
        }
        
        if (-not $CertConfig) {
            Write-ADReportLog -Message "No Certificate Services found in this domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                Component = "AD CS Status"
                Finding = "Not Installed"
                Risk = "N/A"
                Details = "Active Directory Certificate Services not detected in this domain"
            })
        }
        
        # Analyze certificate templates
        $CertTemplates = Get-ADObject -Filter "objectClass -eq 'pKICertificateTemplate'" -Properties * -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,CN=Configuration,$((Get-ADDomain).DistinguishedName)" -ErrorAction SilentlyContinue
        
        foreach ($template in $CertTemplates) {
            # Check for dangerous configurations
            $risks = @()
            $riskLevel = "Low"
            
            # Check if template allows requestor to specify SAN
            if ($template.'msPKI-Certificate-Name-Flag' -band 1) {
                $risks += "Allows SAN specification"
                $riskLevel = "High"
            }
            
            # Check for client authentication
            if ($template.'msPKI-Certificate-Application-Policy' -contains "1.3.6.1.5.5.7.3.2") {
                $risks += "Allows client authentication"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
            }
            
            # Check enrollment permissions
            $templateACL = Get-Acl "AD:$($template.DistinguishedName)" -ErrorAction SilentlyContinue
            $enrollmentRights = $templateACL.Access | Where-Object { $_.ActiveDirectoryRights -match "ExtendedRight" }
            
            if ($enrollmentRights) {
                $enrollers = $enrollmentRights.IdentityReference.Value -join ", "
                if ($enrollers -match "Authenticated Users|Domain Users") {
                    $risks += "Wide enrollment permissions"
                    $riskLevel = "Critical"
                }
            }
            
            if ($risks.Count -gt 0) {
                $CertAnalysis += [PSCustomObject]@{
                    TemplateName = $template.Name
                    DisplayName = $template.DisplayName
                    Risks = $risks -join "; "
                    RiskLevel = $riskLevel
                    EnrollmentRights = $enrollers
                    Remediation = "Review template permissions and configuration"
                }
            }
        }
        
        # Check for web enrollment
        $WebEnrollment = Get-ADObject -Filter "objectClass -eq 'certificationAuthority'" -Properties * -ErrorAction SilentlyContinue
        
        if ($WebEnrollment) {
            $CertAnalysis += [PSCustomObject]@{
                Component = "Web Enrollment"
                Status = "Detected"
                Risk = "Medium"
                Details = "Web enrollment increases attack surface"
                Remediation = "Ensure web enrollment is properly secured with HTTPS and authentication"
            }
        }
        
        Write-ADReportLog -Message "Certificate security analysis completed. $($CertAnalysis.Count) findings." -Type Info -Terminal
        return $CertAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing Certificate Services: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Performance & Health Monitoring ---
Function Get-SYSVOLHealthCheck {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Performing SYSVOL health check..." -Type Info -Terminal
        
        $SYSVOLHealth = @()
        $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
        
        foreach ($DC in $DomainControllers) {
            try {
                # Test SYSVOL share accessibility
                $SYSVOLPath = "\\$($DC.HostName)\SYSVOL"
                $SYSVOLAccessible = Test-Path $SYSVOLPath -ErrorAction SilentlyContinue
                
                # Check DFSR or FRS
                $ReplicationMethod = "Unknown"
                try {
                    $DFSRCheck = Get-WmiObject -ComputerName $DC.HostName -Namespace "root\microsoftdfs" -Class "dfsrreplicatedfolderinfo" -ErrorAction SilentlyContinue
                    if ($DFSRCheck) {
                        $ReplicationMethod = "DFSR"
                    } else {
                        $FRSCheck = Get-WmiObject -ComputerName $DC.HostName -Class "Win32_Service" -Filter "Name='NtFrs'" -ErrorAction SilentlyContinue
                        if ($FRSCheck -and $FRSCheck.State -eq "Running") {
                            $ReplicationMethod = "FRS (Legacy)"
                        }
                    }
                } catch {
                    $ReplicationMethod = "Unable to determine"
                }
                
                # Check for orphaned GPO folders
                $orphanedGPOs = 0
                if ($SYSVOLAccessible) {
                    try {
                        $GPOFolders = Get-ChildItem "$SYSVOLPath\$($DC.Domain)\Policies" -Directory -ErrorAction SilentlyContinue
                        $ADGPOs = Get-GPO -All -Domain $DC.Domain -ErrorAction SilentlyContinue
                        $ADGPOIds = $ADGPOs | ForEach-Object { "{$($_.Id)}" }
                        
                        foreach ($folder in $GPOFolders) {
                            if ($folder.Name -ne "PolicyDefinitions" -and $ADGPOIds -notcontains $folder.Name) {
                                $orphanedGPOs++
                            }
                        }
                    } catch {
                        $orphanedGPOs = "Unable to check"
                    }
                }
                
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $SYSVOLAccessible
                    ReplicationMethod = $ReplicationMethod
                    OrphanedGPOFolders = $orphanedGPOs
                    NetlogonShare = Test-Path "\\$($DC.HostName)\NETLOGON" -ErrorAction SilentlyContinue
                    Status = if ($SYSVOLAccessible -and $ReplicationMethod -ne "Unknown") { "Healthy" } else { "Issues Detected" }
                    Remediation = if ($ReplicationMethod -eq "FRS (Legacy)") { "Migrate from FRS to DFSR" } 
                                 elseif ($orphanedGPOs -gt 0) { "Clean up orphaned GPO folders" }
                                 else { "None required" }
                }
                
            } catch {
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $false
                    Status = "Error"
                    Error = $_.Exception.Message
                }
            }
        }
        
        Write-ADReportLog -Message "SYSVOL health check completed for $($SYSVOLHealth.Count) domain controllers." -Type Info -Terminal
        return $SYSVOLHealth
        
    } catch {
        $ErrorMessage = "Error performing SYSVOL health check: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DNSHealthAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DNS health..." -Type Info -Terminal
        
        $DNSHealth = @()
        $Domain = Get-ADDomain
        
        # Check DNS zones - only if DNS cmdlets are available
        $DNSZones = $null
        try {
            if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
                $DNSZones = Get-DnsServerZone -ComputerName $Domain.PDCEmulator -ErrorAction SilentlyContinue
            } else {
                Write-ADReportLog -Message "DNS Server cmdlets not available. Skipping DNS zone analysis." -Type Info -Terminal
            }
        } catch {
            Write-ADReportLog -Message "Could not query DNS zones: $($_.Exception.Message)" -Type Warning -Terminal
        }
        
        if ($DNSZones) {
            # Group zones by type for tree structure
            $forwardZones = $DNSZones | Where-Object { -not $_.IsReverseLookupZone }
            $reverseZones = $DNSZones | Where-Object { $_.IsReverseLookupZone }
            
            # Process Forward Zones
            foreach ($zone in $forwardZones) {
                $issuesList = [System.Collections.ArrayList]::new()
                
                # Security check
                if ($zone.DynamicUpdate -eq "NonsecureAndSecure") {
                    $issuesList.Add("Allows non-secure dynamic updates") | Out-Null
                    $status = "Security Risk"
                } else {
                    $status = "Healthy"
                }
                
                # Aging/Scavenging check
                if ($zone.AgingEnabled) {
                    $scavengingEnabled = $true
                    $noRefreshInterval = $zone.NoRefreshInterval
                    $refreshInterval = $zone.RefreshInterval
                } else {
                    $issuesList.Add("Scavenging not enabled") | Out-Null
                    if ($status -eq "Healthy") { $status = "Warning" }
                    $scavengingEnabled = $false
                }
                
                $DNSHealth += [PSCustomObject]@{
                    Category = "Forward Zones"
                    ZoneName = $zone.ZoneName
                    Type = "Forward Lookup"
                    ZoneType = $zone.ZoneType
                    IsDsIntegrated = $zone.IsDsIntegrated
                    DynamicUpdate = $zone.DynamicUpdate
                    Status = $status
                    Issues = if ($issuesList.Count -gt 0) { $issuesList -join "; " } else { "None" }
                    ScavengingEnabled = $scavengingEnabled
                    NoRefreshInterval = if ($scavengingEnabled) { $noRefreshInterval } else { "N/A" }
                    RefreshInterval = if ($scavengingEnabled) { $refreshInterval } else { "N/A" }
                }
            }
            
            # Process Reverse Zones
            foreach ($zone in $reverseZones) {
                $issuesList = [System.Collections.ArrayList]::new()
                
                if ($zone.DynamicUpdate -eq "NonsecureAndSecure") {
                    $issuesList.Add("Allows non-secure dynamic updates") | Out-Null
                    $status = "Security Risk"
                } else {
                    $status = "Healthy"
                }
                
                if ($zone.AgingEnabled) {
                    $scavengingEnabled = $true
                    $noRefreshInterval = $zone.NoRefreshInterval
                    $refreshInterval = $zone.RefreshInterval
                } else {
                    $issuesList.Add("Scavenging not enabled") | Out-Null
                    if ($status -eq "Healthy") { $status = "Warning" }
                    $scavengingEnabled = $false
                }
                
                $DNSHealth += [PSCustomObject]@{
                    Category = "Reverse Zones"
                    ZoneName = $zone.ZoneName
                    Type = "Reverse Lookup"
                    ZoneType = $zone.ZoneType
                    IsDsIntegrated = $zone.IsDsIntegrated
                    DynamicUpdate = $zone.DynamicUpdate
                    Status = $status
                    Issues = if ($issuesList.Count -gt 0) { $issuesList -join "; " } else { "None" }
                    ScavengingEnabled = $scavengingEnabled
                    NoRefreshInterval = if ($scavengingEnabled) { $noRefreshInterval } else { "N/A" }
                    RefreshInterval = if ($scavengingEnabled) { $refreshInterval } else { "N/A" }
                }
            }
        }
        
        # Add DNS Health Summary
        $DNSHealth += [PSCustomObject]@{
            Category = "DNS Health Summary"
            ZoneName = "Overall Status"
            Type = "System Check"
            ZoneType = "N/A"
            IsDsIntegrated = "N/A" 
            DynamicUpdate = "N/A"
            Status = if ($DNSHealth.Status -contains "Security Risk") { "Critical" } 
                    elseif ($DNSHealth.Status -contains "Warning") { "Warning" }
                    else { "Healthy" }
            Issues = "Stale records check required"
            ScavengingEnabled = "N/A"
            NoRefreshInterval = "N/A"
            RefreshInterval = "N/A"
        }
        
        Write-ADReportLog -Message "DNS health analysis completed. $($DNSHealth.Count) items analyzed." -Type Info -Terminal
        return $DNSHealth
        
    } catch {
        $ErrorMessage = "Error analyzing DNS health: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Backup & Recovery Readiness ---
Function Get-BackupReadinessStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD backup readiness..." -Type Info -Terminal
        
        $BackupStatus = @()
        $Domain = Get-ADDomain
        
        # Check System State backup on PDC
        $PDC = $Domain.PDCEmulator
        
        # Tombstone lifetime
        $ConfigNC = "CN=Configuration,$($Domain.DistinguishedName)"
        $TombstoneObject = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigNC" -Properties tombstoneLifetime -ErrorAction SilentlyContinue
        $TombstoneLifetime = if ($TombstoneObject.tombstoneLifetime) { $TombstoneObject.tombstoneLifetime } else { 180 }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Tombstone Lifetime"
            Value = "$TombstoneLifetime days"
            Status = if ($TombstoneLifetime -ge 180) { "Good" } else { "Warning" }
            Details = "Backups older than tombstone lifetime cannot be restored"
            Remediation = if ($TombstoneLifetime -lt 180) { "Consider increasing tombstone lifetime" } else { "None required" }
        }
        
        # Check Deleted Objects container size
        try {
            $DeletedObjects = Get-ADObject -Filter * -IncludeDeletedObjects -SearchBase "CN=Deleted Objects,$($Domain.DistinguishedName)" -ErrorAction SilentlyContinue
            $DeletedCount = if ($DeletedObjects) { @($DeletedObjects).Count } else { 0 }
            
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "$DeletedCount objects"
                Status = if ($DeletedCount -gt 10000) { "Warning" } else { "Good" }
                Details = "Large number of deleted objects can impact restore performance"
                Remediation = if ($DeletedCount -gt 10000) { "Consider garbage collection" } else { "None required" }
            }
        } catch {
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "Unable to access"
                Status = "Unknown"
                Details = "Requires elevated permissions"
            }
        }
        
        # DSRM password age check
        $DomainControllers = Get-ADDomainController -Filter *
        
        foreach ($DC in $DomainControllers) {
            $BackupStatus += [PSCustomObject]@{
                Component = "DSRM Password"
                DomainController = $DC.Name
                Status = "Manual Check Required"
                Details = "DSRM password age cannot be checked remotely"
                Remediation = "Ensure DSRM passwords are documented and regularly updated"
            }
        }
        
        # Backup GPO existence
        $BackupGPOs = Get-GPO -All | Where-Object { $_.DisplayName -like "*Backup*" -or $_.DisplayName -like "*Restore*" }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Backup Procedures"
            Value = if ($BackupGPOs) { "$($BackupGPOs.Count) backup-related GPOs found" } else { "No backup GPOs found" }
            Status = if ($BackupGPOs) { "Good" } else { "Warning" }
            Details = "GPOs can help standardize backup procedures"
            Remediation = if (-not $BackupGPOs) { "Consider creating backup procedure GPOs" } else { "Review existing backup GPOs" }
        }
        
        Write-ADReportLog -Message "Backup readiness check completed." -Type Info -Terminal
        return $BackupStatus
        
    } catch {
        $ErrorMessage = "Error checking backup readiness: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Schema & Trusts Analysis ---
Function Get-SchemaAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD schema..." -Type Info -Terminal
        
        $SchemaAnalysis = @()
        $RootDSE = Get-ADRootDSE
        $SchemaNC = $RootDSE.schemaNamingContext
        
        # Get schema version and details
        $Schema = Get-ADObject -Identity $SchemaNC -Properties objectVersion, msDS-Behavior-Version
        
        # Define all possible schema versions
        $AllSchemaVersions = @{
            30 = @{
                Name = "Windows Server 2003"
                Details = "Base Windows Server 2003 Schema"
            }
            31 = @{
                Name = "Windows Server 2003 R2" 
                Details = "Supports Read-Only Domain Controllers"
            }
            44 = @{
                Name = "Windows Server 2008"
                Details = "Supports Fine-Grained Password Policies"
            }
            47 = @{
                Name = "Windows Server 2008 R2"
                Details = "Supports Managed Service Accounts, Authentication Mechanism Assurance"
            }
            56 = @{
                Name = "Windows Server 2012"
                Details = "Supports Group Managed Service Accounts, Kerberos Armoring"
            }
            69 = @{
                Name = "Windows Server 2012 R2"
                Details = "Supports Dynamic Access Control, Kerberos KDC Support for Claims"
            }
            87 = @{
                Name = "Windows Server 2016"
                Details = "Supports Privileged Access Management, Microsoft Passport for Work"
            }
            88 = @{
                Name = "Windows Server 2019/2022"
                Details = "Supports FIDO2 Authentication, Group MSA Enhancements"
            }
            91 = @{
                Name = "Windows Server 2025"
                Details = "Supports Enhanced Security Features, Latest AD Improvements"
            }
        }

        # Only display the active schema version
        if ($AllSchemaVersions.ContainsKey($Schema.objectVersion)) {
            $SchemaAnalysis += [PSCustomObject]@{
                Name = $AllSchemaVersions[$Schema.objectVersion].Name
                Details = $AllSchemaVersions[$Schema.objectVersion].Details
                Status = "Active"
                IsCurrent = $true
            }
        }
        
        # Count custom schema extensions
        $AllSchemaObjects = Get-ADObject -Filter * -SearchBase $SchemaNC -Properties whenCreated, adminDescription
        $CustomSchema = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -and 
            $_.adminDescription -notlike "Microsoft*" -and 
            $_.Name -notlike "ms-*" 
        }
        
        $SchemaAnalysis += [PSCustomObject]@{
            Name = "Custom Extensions"
            Details = "$($CustomSchema.Count) custom schema objects found"
            Status = "Info"
            IsCurrent = $true
        }
        
        # Recent schema changes
        $RecentChanges = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -gt (Get-Date).AddDays(-90) 
        }
        
        if ($RecentChanges) {
            foreach ($change in $RecentChanges) {
                $SchemaAnalysis += [PSCustomObject]@{
                    Name = "Recent Change: $($change.Name)"
                    Details = "Schema change on $($change.whenCreated.ToString('MM/dd/yyyy'))"
                    Status = "Info"
                    IsCurrent = $true
                }
            }
        }
        
        Write-ADReportLog -Message "Schema analysis completed." -Type Info -Terminal
        return $SchemaAnalysis
        
    } catch {
        $ErrorMessage = "Error in schema analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-TrustRelationshipAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing trust relationships..." -Type Info -Terminal
        
        $TrustAnalysis = @()
        
        # Get all trusts
        $Trusts = Get-ADTrust -Filter * -Properties * -ErrorAction Stop
        
        foreach ($trust in $Trusts) {
            # Analyze trust properties
            $trustHealth = "Unknown"
            $issues = @()
            
            # Check trust type and transitivity
            if ($trust.TrustType -eq "External" -and $trust.Transitivity -eq "Transitive") {
                $issues += "External trust is transitive"
                $trustHealth = "Security Risk"
            }
            
            # Check SID filtering
            if ($trust.SIDFilteringForestAware -eq $false -or $trust.SIDFilteringQuarantined -eq $false) {
                $issues += "SID filtering may be disabled"
                if ($trustHealth -ne "Security Risk") { $trustHealth = "Warning" }
            }
            
            # Check selective authentication
            if ($trust.SelectiveAuthentication -eq $false -and $trust.TrustType -eq "External") {
                $issues += "Selective authentication not enabled"
                if ($trustHealth -eq "Unknown") { $trustHealth = "Review Needed" }
            }
            
            if ($trustHealth -eq "Unknown" -and $issues.Count -eq 0) { $trustHealth = "Healthy" }
            
            $TrustAnalysis += [PSCustomObject]@{
                TrustPartner = $trust.Target
                Direction = $trust.Direction
                TrustType = $trust.TrustType
                Transitivity = if ($trust.Transitivity) { "Transitive" } else { "Non-Transitive" }
                SIDFiltering = if ($trust.SIDFilteringQuarantined) { "Enabled" } else { "Check Required" }
                SelectiveAuth = if ($trust.SelectiveAuthentication) { "Enabled" } else { "Disabled" }
                Created = $trust.Created
                Status = $trustHealth
                Issues = $issues -join "; "
                Remediation = if ($issues.Count -gt 0) { "Review trust configuration for security best practices" } else { "None required" }
            }
        }
        
        # Forest trust insights
        $Forest = Get-ADForest
        if ($Forest.Domains.Count -gt 1) {
            $TrustAnalysis += [PSCustomObject]@{
                Component = "Forest Structure"
                Domains = $Forest.Domains.Count
                GlobalCatalogs = $Forest.GlobalCatalogs.Count
                Sites = $Forest.Sites.Count
                Status = "Multi-Domain Forest"
                Details = "Complex forest structure requires careful trust management"
            }
        }
        
        Write-ADReportLog -Message "Trust relationship analysis completed. $($Trusts.Count) trusts analyzed." -Type Info -Terminal
        
        # Ensure we always return an array
        if ($TrustAnalysis.Count -eq 0) {
            return @()
        }
        return @($TrustAnalysis)
        
    } catch {
        $ErrorMessage = "Error analyzing trust relationships: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @()
    }
}

Function Get-QuotasAndLimits {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD quotas and limits..." -Type Info -Terminal
        
        $QuotaAnalysis = @()
        $Domain = Get-ADDomain
        $Forest = Get-ADForest
        
        # RID Pool Analysis - Critical for Domain Controllers
        Write-ADReportLog -Message "Analyzing RID Pool..." -Type Info -Terminal
        $DomainControllers = Get-ADDomainController -Filter *
        
        foreach ($DC in $DomainControllers) {
            try {
                $RIDInfo = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($Domain.DistinguishedName)" -Properties rIDAvailablePool -Server $DC.HostName -ErrorAction SilentlyContinue
                
                if ($RIDInfo -and $RIDInfo.rIDAvailablePool) {
                    try {
                        if ($RIDInfo.rIDAvailablePool -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool[0]
                        } else {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool
                        }
                        
                        [int64]$pow32 = [math]::Pow(2, 32)
                        [int64]$pow30 = [math]::Pow(2, 30)
                        
                        [int64]$totalSIDS = [math]::Floor($ridPoolValue / $pow32)
                        [int64]$totalRIDS = $totalSIDS * $pow30
                        [int64]$currentRIDPoolCount = $ridPoolValue % $pow32
                        
                        $percentUsed = if ($totalRIDS -gt 0) {
                            [math]::Round((($totalRIDS - $currentRIDPoolCount) / $totalRIDS * 100), 2)
                        } else { 0 }
                        
                        $QuotaAnalysis += [PSCustomObject]@{
                            Category = "1. RID Pool Management"
                            Subcategory = "1.1 Domain Controllers"
                            Component = "DC $($DC.Name) → RID Pool Status"
                            CurrentValue = "$totalRIDS Total, $currentRIDPoolCount Remaining ($percentUsed% Used)"
                            RecommendedValue = "< 80% Usage"
                            Status = if ($percentUsed -gt 80) { "Critical" } elseif ($percentUsed -gt 60) { "Warning" } else { "Healthy" }
                            Description = "RID pool exhaustion can prevent new object creation"
                            Remediation = if ($percentUsed -gt 80) { "RID pool nearly exhausted - plan RID recovery" } else { "Monitor RID usage" }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error calculating RID pool for $($DC.Name): $($_.Exception.Message)" -Type Warning
                        continue
                    }
                }
            } catch {
                Write-ADReportLog -Message "Could not retrieve RID pool information from $($DC.Name)" -Type Warning
            }
        }

        # LDAP Query Limits and Performance Restrictions
        Write-ADReportLog -Message "Analyzing LDAP Query Policies..." -Type Info -Terminal
        
        $DefaultPolicy = $null
        
        # Try to find Default Query Policy with error handling
        try {
            $ConfigPath = "CN=Configuration,$($Forest.RootDomain)"
            $QueryPolicyPath = "CN=Default Query Policy,CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services,$ConfigPath"
            $DefaultPolicy = Get-ADObject -Identity $QueryPolicyPath -Properties * -ErrorAction SilentlyContinue
        } catch {
            Write-ADReportLog -Message "Default Query Policy not found in standard location. Trying alternative search..." -Type Warning
            
            # Alternative search method
            try {
                $QueryPolicies = Get-ADObject -Filter "objectClass -eq 'queryPolicy'" -SearchBase "CN=Configuration,$($Forest.RootDomain)" -Properties * -ErrorAction SilentlyContinue
                $DefaultPolicy = $QueryPolicies | Where-Object { $_.Name -eq "Default Query Policy" } | Select-Object -First 1
            } catch {
                Write-ADReportLog -Message "Could not locate Default Query Policy through alternative search" -Type Warning
            }
        }
        
        if ($DefaultPolicy) {
            Write-ADReportLog -Message "Found Default Query Policy at: $($DefaultPolicy.DistinguishedName)" -Type Info
            
            $MaxPageSize = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxPageSize=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxQueryDuration = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxQueryDuration=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxResults = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxResultSetSize=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxTempTableSize = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxTempTableSize=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxPoolThreads = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxPoolThreads=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxDatagramRecv = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxDatagramRecv=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            
            # LDAP MaxPageSize
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.1 MaxPageSize → Objects per Query Page"
                CurrentValue = if ($MaxPageSize) { $MaxPageSize } else { "1000 (Default)" }
                RecommendedValue = "1000-5000"
                Status = if ($MaxPageSize -and $MaxPageSize -lt 100) { "Low" } elseif ($MaxPageSize -and $MaxPageSize -gt 10000) { "High" } else { "OK" }
                Description = "Maximum objects returned per LDAP query page"
                Remediation = if ($MaxPageSize -and ($MaxPageSize -lt 100 -or $MaxPageSize -gt 10000)) { "Adjust MaxPageSize to optimize query performance" } else { "Current setting is appropriate" }
            }
            
            # LDAP MaxQueryDuration
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.2 MaxQueryDuration → Query Timeout"
                CurrentValue = if ($MaxQueryDuration) { "$MaxQueryDuration seconds" } else { "120 seconds (Default)" }
                RecommendedValue = "120-600 seconds"
                Status = if ($MaxQueryDuration -and $MaxQueryDuration -lt 60) { "Low" } elseif ($MaxQueryDuration -and $MaxQueryDuration -gt 3600) { "High" } else { "OK" }
                Description = "Maximum execution time for LDAP queries"
                Remediation = if ($MaxQueryDuration -and ($MaxQueryDuration -lt 60 -or $MaxQueryDuration -gt 3600)) { "Adjust query duration to balance performance and timeout prevention" } else { "Current setting is appropriate" }
            }
            
            # LDAP MaxResultSetSize
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.3 MaxResultSetSize → Result Limit"
                CurrentValue = if ($MaxResults) { $MaxResults } else { "262144 (Default)" }
                RecommendedValue = "10000-1000000"
                Status = if ($MaxResults -and $MaxResults -lt 1000) { "Low" } elseif ($MaxResults -and $MaxResults -gt 2000000) { "High" } else { "OK" }
                Description = "Maximum objects returned in single LDAP query"
                Remediation = if ($MaxResults -and ($MaxResults -lt 1000 -or $MaxResults -gt 2000000)) { "Review result set size limits for query optimization" } else { "Current setting is appropriate" }
            }
            
            # LDAP MaxTempTableSize
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.4 MaxTempTableSize → Temporary Table"
                CurrentValue = if ($MaxTempTableSize) { "$MaxTempTableSize rows" } else { "10000 rows (Default)" }
                RecommendedValue = "10000-100000 rows"
                Status = if ($MaxTempTableSize -and $MaxTempTableSize -lt 5000) { "Low" } elseif ($MaxTempTableSize -and $MaxTempTableSize -gt 500000) { "High" } else { "OK" }
                Description = "Maximum temporary table size for complex queries"
                Remediation = if ($MaxTempTableSize -and ($MaxTempTableSize -lt 5000 -or $MaxTempTableSize -gt 500000)) { "Adjust temp table size for complex query performance" } else { "Current setting is appropriate" }
            }
            
            # LDAP MaxPoolThreads
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.5 MaxPoolThreads → Thread Pool"
                CurrentValue = if ($MaxPoolThreads) { $MaxPoolThreads } else { "4 (Default)" }
                RecommendedValue = "4-20"
                Status = if ($MaxPoolThreads -and $MaxPoolThreads -lt 2) { "Low" } elseif ($MaxPoolThreads -and $MaxPoolThreads -gt 50) { "High" } else { "OK" }
                Description = "Maximum threads for LDAP operations"
                Remediation = if ($MaxPoolThreads -and ($MaxPoolThreads -lt 2 -or $MaxPoolThreads -gt 50)) { "Adjust thread pool size based on server capacity" } else { "Current setting is appropriate" }
            }
            
            # LDAP MaxDatagramRecv
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.6 MaxDatagramRecv → UDP Datagram Size"
                CurrentValue = if ($MaxDatagramRecv) { "$MaxDatagramRecv bytes" } else { "4096 bytes (Default)" }
                RecommendedValue = "4096-65536 bytes"
                Status = if ($MaxDatagramRecv -and $MaxDatagramRecv -lt 1024) { "Low" } elseif ($MaxDatagramRecv -and $MaxDatagramRecv -gt 131072) { "High" } else { "OK" }
                Description = "Maximum UDP datagram size for LDAP"
                Remediation = if ($MaxDatagramRecv -and ($MaxDatagramRecv -lt 1024 -or $MaxDatagramRecv -gt 131072)) { "Adjust datagram size for network optimization" } else { "Current setting is appropriate" }
            }
        } else {
            Write-ADReportLog -Message "Default Query Policy not found or accessible. Using standard defaults for analysis." -Type Warning
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "2. LDAP Query Policies"
                Subcategory = "2.1 Query Performance Settings"
                Component = "2.1.0 Policy Access → Access Status"
                CurrentValue = "Not Accessible"
                RecommendedValue = "Accessible"
                Status = "Warning"
                Description = "Default Query Policy could not be retrieved"
                Remediation = "Check access permissions to configuration partition"
            }
        }

        # Default Domain Password Policy
        Write-ADReportLog -Message "Analyzing Password Policies..." -Type Info -Terminal
        
        $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue
        
        if ($PasswordPolicy) {
            # Minimum Password Length
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.1 MinPasswordLength → Minimum Length"
                CurrentValue = "$($PasswordPolicy.MinPasswordLength) characters"
                RecommendedValue = "12+ characters"
                Status = if ($PasswordPolicy.MinPasswordLength -lt 8) { "Weak" } elseif ($PasswordPolicy.MinPasswordLength -lt 12) { "Fair" } else { "Strong" }
                Description = "Minimum required password length"
                Remediation = if ($PasswordPolicy.MinPasswordLength -lt 12) { "Increase minimum password length to 12+ characters" } else { "Current setting meets security recommendations" }
            }
            
            # Maximum Password Age
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.2 MaxPasswordAge → Password Expiration"
                CurrentValue = "$($PasswordPolicy.MaxPasswordAge.Days) days"
                RecommendedValue = "60-365 days"
                Status = if ($PasswordPolicy.MaxPasswordAge.Days -lt 30 -or $PasswordPolicy.MaxPasswordAge.Days -eq 0) { "Review" } elseif ($PasswordPolicy.MaxPasswordAge.Days -gt 365) { "Long" } else { "OK" }
                Description = "Maximum time before password must be changed"
                Remediation = if ($PasswordPolicy.MaxPasswordAge.Days -lt 30) { "Consider longer password age for user convenience" } elseif ($PasswordPolicy.MaxPasswordAge.Days -gt 365) { "Consider shorter password age for security" } else { "Current setting is balanced" }
            }
            
            # Minimum Password Age
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.3 MinPasswordAge → Minimum Age"
                CurrentValue = "$($PasswordPolicy.MinPasswordAge.Days) days"
                RecommendedValue = "1-7 days"
                Status = if ($PasswordPolicy.MinPasswordAge.Days -eq 0) { "Disabled" } elseif ($PasswordPolicy.MinPasswordAge.Days -gt 7) { "High" } else { "OK" }
                Description = "Minimum time before password can be changed again"
                Remediation = if ($PasswordPolicy.MinPasswordAge.Days -eq 0) { "Enable minimum password age to prevent rapid password cycling" } elseif ($PasswordPolicy.MinPasswordAge.Days -gt 7) { "Consider reducing minimum age for user flexibility" } else { "Current setting is appropriate" }
            }
            
            # Password History Count
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.4 PasswordHistoryCount → Password History"
                CurrentValue = "$($PasswordPolicy.PasswordHistoryCount) passwords"
                RecommendedValue = "12-24 passwords"
                Status = if ($PasswordPolicy.PasswordHistoryCount -lt 12) { "Low" } elseif ($PasswordPolicy.PasswordHistoryCount -gt 50) { "High" } else { "OK" }
                Description = "Number of previous passwords that cannot be reused"
                Remediation = if ($PasswordPolicy.PasswordHistoryCount -lt 12) { "Increase password history to prevent password reuse" } elseif ($PasswordPolicy.PasswordHistoryCount -gt 50) { "Consider reducing history count for practicality" } else { "Current setting is appropriate" }
            }
            
            # Account Lockout Duration
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.5 LockoutDuration → Lockout Duration"
                CurrentValue = if ($PasswordPolicy.LockoutDuration.TotalMinutes -gt 0) { "$($PasswordPolicy.LockoutDuration.TotalMinutes) minutes" } else { "Until admin unlocks" }
                RecommendedValue = "15-60 minutes"
                Status = if ($PasswordPolicy.LockoutDuration.TotalMinutes -eq 0) { "Manual" } elseif ($PasswordPolicy.LockoutDuration.TotalMinutes -lt 5) { "Short" } elseif ($PasswordPolicy.LockoutDuration.TotalMinutes -gt 180) { "Long" } else { "OK" }
                Description = "Time account remains locked after failed attempts"
                Remediation = if ($PasswordPolicy.LockoutDuration.TotalMinutes -eq 0) { "Consider automatic unlock to reduce admin overhead" } elseif ($PasswordPolicy.LockoutDuration.TotalMinutes -lt 5) { "Increase lockout duration for better security" } elseif ($PasswordPolicy.LockoutDuration.TotalMinutes -gt 180) { "Consider shorter duration for user convenience" } else { "Current setting is balanced" }
            }
            
            # Account Lockout Threshold
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.6 LockoutThreshold → Lockout Threshold"
                CurrentValue = if ($PasswordPolicy.LockoutThreshold -gt 0) { "$($PasswordPolicy.LockoutThreshold) attempts" } else { "Disabled" }
                RecommendedValue = "3-10 attempts"
                Status = if ($PasswordPolicy.LockoutThreshold -eq 0) { "Disabled" } elseif ($PasswordPolicy.LockoutThreshold -lt 3) { "Strict" } elseif ($PasswordPolicy.LockoutThreshold -gt 20) { "Lenient" } else { "OK" }
                Description = "Failed logon attempts before lockout"
                Remediation = if ($PasswordPolicy.LockoutThreshold -eq 0) { "Enable account lockout to prevent brute force attacks" } elseif ($PasswordPolicy.LockoutThreshold -lt 3) { "Consider higher threshold to reduce false lockouts" } elseif ($PasswordPolicy.LockoutThreshold -gt 20) { "Consider lower threshold for better security" } else { "Current setting is appropriate" }
            }
            
            # Password Complexity
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.7 ComplexityEnabled → Password Complexity"
                CurrentValue = if ($PasswordPolicy.ComplexityEnabled) { "Enabled" } else { "Disabled" }
                RecommendedValue = "Enabled"
                Status = if ($PasswordPolicy.ComplexityEnabled) { "OK" } else { "Weak" }
                Description = "Requires complex password requirements"
                Remediation = if (-not $PasswordPolicy.ComplexityEnabled) { "Enable password complexity for stronger passwords" } else { "Current setting follows security best practices" }
            }
            
            # Reversible Encryption
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.8 ReversibleEncryption → Reversible Encryption"
                CurrentValue = if ($PasswordPolicy.ReversibleEncryptionEnabled) { "Enabled" } else { "Disabled" }
                RecommendedValue = "Disabled"
                Status = if ($PasswordPolicy.ReversibleEncryptionEnabled) { "Risk" } else { "Secure" }
                Description = "Stores passwords using reversible encryption"
                Remediation = if ($PasswordPolicy.ReversibleEncryptionEnabled) { "Disable reversible encryption unless required by legacy applications" } else { "Current setting follows security best practices" }
            }
        } else {
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.1 Default Domain Policy"
                Component = "3.1.0 Policy Access → Access Status"
                CurrentValue = "Not Available"
                RecommendedValue = "Available"
                Status = "Error"
                Description = "Could not retrieve default domain password policy"
                Remediation = "Check domain controller connectivity and permissions"
            }
        }

        # Fine-Grained Password Policies (FGPP)
        Write-ADReportLog -Message "Analyzing Fine-Grained Password Policies..." -Type Info -Terminal
        $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue
        
        if ($FGPPs) {
            foreach ($fgpp in $FGPPs) {
                $QuotaAnalysis += [PSCustomObject]@{
                    Category = "3. Password Policies"
                    Subcategory = "3.2 Fine-Grained Password Policies"
                    Component = "3.2.1 FGPP → $($fgpp.Name)"
                    CurrentValue = "Precedence: $($fgpp.Precedence), MinLength: $($fgpp.MinPasswordLength), MaxAge: $($fgpp.MaxPasswordAge.Days) days"
                    RecommendedValue = "MinLength: 12+, Review precedence"
                    Status = if ($fgpp.MinPasswordLength -lt 12) { "Review" } else { "OK" }
                    Description = "Custom password policy with higher precedence"
                    Remediation = if ($fgpp.MinPasswordLength -lt 12) { "Review FGPP settings for compliance with security standards" } else { "Policy configuration meets recommendations" }
                }
            }
        } else {
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "3. Password Policies"
                Subcategory = "3.2 Fine-Grained Password Policies"
                Component = "3.2.0 FGPP Count → Policy Count"
                CurrentValue = "0 policies"
                RecommendedValue = "Consider for privileged accounts"
                Status = "Not Configured"
                Description = "No FGPP found - only default domain policy active"
                Remediation = "Consider implementing FGPP for privileged accounts"
            }
        }

        # Kerberos Settings
        Write-ADReportLog -Message "Analyzing Kerberos Configuration..." -Type Info -Terminal
        
        $LargeTokenUsers = Get-ADUser -Filter * -Properties MemberOf -ErrorAction SilentlyContinue | Where-Object { $_.MemberOf.Count -gt 100 }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "4. Kerberos Configuration"
            Subcategory = "4.1 Ticket Lifetimes"
            Component = "4.1.1 MaxTicketAge → Max Ticket Age"
            CurrentValue = "10 hours (Standard)"
            RecommendedValue = "8-24 hours"
            Status = "Standard"
            Description = "Maximum lifetime of Kerberos service tickets"
            Remediation = "Monitor for authentication issues if modified"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "4. Kerberos Configuration"
            Subcategory = "4.1 Ticket Lifetimes"
            Component = "4.1.2 MaxRenewalAge → Max Renewal Age"
            CurrentValue = "7 days (Standard)"
            RecommendedValue = "7-30 days"
            Status = "Standard"
            Description = "Maximum time a ticket can be renewed"
            Remediation = "Standard configuration is appropriate"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "4. Kerberos Configuration"
            Subcategory = "4.2 Token Management"
            Component = "4.2.1 LargeTokenUsers → Users with Large Token"
            CurrentValue = if ($LargeTokenUsers) { "$($LargeTokenUsers.Count) users" } else { "0 users" }
            RecommendedValue = "< 10 users"
            Status = if ($LargeTokenUsers -and $LargeTokenUsers.Count -gt 10) { "Warning" } else { "OK" }
            Description = "Users with 100+ group memberships"
            Remediation = if ($LargeTokenUsers -and $LargeTokenUsers.Count -gt 10) { "Review group memberships for users with excessive assignments" } else { "Token sizes are within normal ranges" }
        }

        # Tombstone Lifetime
        Write-ADReportLog -Message "Analyzing Tombstone Lifetime..." -Type Info -Terminal
        
        $TombstoneLifetime = $null
        try {
            $TombstoneLifetime = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($Forest.RootDomain)" -Properties tombstoneLifetime -ErrorAction SilentlyContinue
        } catch {
            Write-ADReportLog -Message "Could not retrieve tombstone lifetime from forest configuration" -Type Warning
        }
        
        $TombstoneValue = if ($TombstoneLifetime -and $TombstoneLifetime.tombstoneLifetime) { 
            $TombstoneLifetime.tombstoneLifetime 
        } else { 
            60 
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "5. Backup and Recovery"
            Subcategory = "5.1 Object Recovery"
            Component = "5.1.1 TombstoneLifetime → Tombstone Lifetime"
            CurrentValue = "$TombstoneValue days"
            RecommendedValue = "180+ days"
            Status = if ($TombstoneValue -lt 60) { "Critical" } elseif ($TombstoneValue -lt 180) { "Warning" } else { "Good" }
            Description = "Time deleted objects remain recoverable"
            Remediation = if ($TombstoneValue -lt 60) { "Increase tombstone lifetime to at least 60 days" } elseif ($TombstoneValue -lt 180) { "Consider increasing to 180+ days for better backup safety" } else { "Configuration provides good backup recovery window" }
        }

        # Replication Configuration
        Write-ADReportLog -Message "Analyzing Replication Configuration..." -Type Info -Terminal
        
        $SiteLinks = Get-ADReplicationSiteLink -Filter * -ErrorAction SilentlyContinue
        if ($SiteLinks) {
            # Safe average calculation for ReplicationInterval
            $NumericIntervals = @()
            foreach ($link in $SiteLinks) {
                if ($link.ReplicationInterval -is [System.Int32] -or $link.ReplicationInterval -is [System.Int64]) {
                    $NumericIntervals += $link.ReplicationInterval
                } elseif ($link.ReplicationInterval -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                    foreach ($interval in $link.ReplicationInterval) {
                        if ($interval -is [System.Int32] -or $interval -is [System.Int64]) {
                            $NumericIntervals += $interval
                        }
                    }
                }
            }
            
            if ($NumericIntervals.Count -gt 0) {
                $AverageInterval = ($NumericIntervals | Measure-Object -Average).Average
                $MinInterval = ($NumericIntervals | Measure-Object -Minimum).Minimum
                $MaxInterval = ($NumericIntervals | Measure-Object -Maximum).Maximum
            } else {
                $AverageInterval = 180
                $MinInterval = 15
                $MaxInterval = 180
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "6. Replication Topology"
                Subcategory = "6.1 Site Links"
                Component = "6.1.1 SiteLinkCount → Site Link Count"
                CurrentValue = "$($SiteLinks.Count) links"
                RecommendedValue = "Minimize complexity"
                Status = if ($SiteLinks.Count -gt 20) { "Complex" } else { "Standard" }
                Description = "Total replication site links configured"
                Remediation = if ($SiteLinks.Count -gt 20) { "Review site topology for optimization" } else { "Site link configuration is manageable" }
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "6. Replication Topology"
                Subcategory = "6.1 Site Links"
                Component = "6.1.2 AvgReplicationInterval → Average Replication Interval"
                CurrentValue = "$([math]::Round($AverageInterval, 2)) minutes"
                RecommendedValue = "15-180 minutes"
                Status = if ($AverageInterval -gt 360) { "Slow" } elseif ($AverageInterval -lt 15) { "Fast" } else { "OK" }
                Description = "Average replication frequency across links"
                Remediation = if ($AverageInterval -gt 360) { "Long intervals may cause replication delays" } elseif ($AverageInterval -lt 15) { "Very short intervals may increase network load" } else { "Replication timing is appropriate" }
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "6. Replication Topology"
                Subcategory = "6.1 Site Links"
                Component = "6.1.3 IntervalRange → Fastest/Slowest Interval"
                CurrentValue = "$MinInterval min / $MaxInterval min"
                RecommendedValue = "15+ min / < 1440 min"
                Status = if ($MinInterval -lt 15 -or $MaxInterval -gt 1440) { "Review" } else { "OK" }
                Description = "Range of replication intervals"
                Remediation = if ($MinInterval -lt 15) { "Very frequent replication may impact performance" } elseif ($MaxInterval -gt 1440) { "Very long intervals may cause delays" } else { "Interval range is acceptable" }
            }
        }

        # Forest and Domain Structure
        Write-ADReportLog -Message "Analyzing Forest Structure..." -Type Info -Terminal
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "7. Forest Structure"
            Subcategory = "7.1 Functional Levels"
            Component = "7.1.1 ForestMode → Forest Functional Level"
            CurrentValue = "$($Forest.ForestMode)"
            RecommendedValue = "Latest supported level"
            Status = "Informational"
            Description = "Forest-wide feature availability"
            Remediation = "Ensure level supports required features"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "7. Forest Structure"
            Subcategory = "7.1 Functional Levels"
            Component = "7.1.2 DomainMode → Domain Functional Level"
            CurrentValue = "$($Domain.DomainMode)"
            RecommendedValue = "Latest supported level"
            Status = "Informational"
            Description = "Domain-specific feature availability"
            Remediation = "Ensure level supports required features"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "7. Forest Structure"
            Subcategory = "7.2 Topology Metrics"
            Component = "7.2.1 DomainCount → Domain Count"
            CurrentValue = "$($Forest.Domains.Count) domains"
            RecommendedValue = "Minimize complexity"
            Status = if ($Forest.Domains.Count -gt 5) { "Complex" } elseif ($Forest.Domains.Count -gt 1) { "Multi-Domain" } else { "Single-Domain" }
            Description = "Total domains in forest"
            Remediation = if ($Forest.Domains.Count -gt 5) { "Complex structure requires careful management" } else { "Domain structure is manageable" }
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "7. Forest Structure"
            Subcategory = "7.2 Topology Metrics"
            Component = "7.2.2 GlobalCatalogCount → Global Catalog Count"
            CurrentValue = "$($Forest.GlobalCatalogs.Count) servers"
            RecommendedValue = "At least 1 per site"
            Status = if ($Forest.GlobalCatalogs.Count -lt $Forest.Sites.Count) { "Insufficient" } else { "Adequate" }
            Description = "Global Catalog servers in forest"
            Remediation = if ($Forest.GlobalCatalogs.Count -lt $Forest.Sites.Count) { "Consider adding Global Catalogs for better performance" } else { "Global Catalog distribution is adequate" }
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "7. Forest Structure"
            Subcategory = "7.2 Topology Metrics"
            Component = "7.2.3 SiteCount → Site Count"
            CurrentValue = "$($Forest.Sites.Count) sites"
            RecommendedValue = "Match physical topology"
            Status = if ($Forest.Sites.Count -gt 50) { "Many Sites" } else { "Standard" }
            Description = "Active Directory sites configured"
            Remediation = if ($Forest.Sites.Count -gt 50) { "Large number of sites requires careful replication management" } else { "Site count is manageable" }
        }

        # Database and Object Limits
        Write-ADReportLog -Message "Analyzing Database Limits..." -Type Info -Terminal
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "8. Database Limits"
            Subcategory = "8.1 Theoretical Maximums"
            Component = "8.1.1 MaxDatabaseSize → Maximum Database Size"
            CurrentValue = "16 TB (theoretical)"
            RecommendedValue = "< 100 GB for optimal performance"
            Status = "Informational"
            Description = "Maximum NTDS.dit database size"
            Remediation = "Monitor database growth and defragmentation"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "8. Database Limits"
            Subcategory = "8.1 Theoretical Maximums"
            Component = "8.1.2 MaxObjectsPerDomain → Maximum Objects Per Domain"
            CurrentValue = "~2.1 billion objects (theoretical)"
            RecommendedValue = "Monitor growth patterns"
            Status = "Informational"
            Description = "Theoretical object limit per domain"
            Remediation = "Monitor object growth and plan scalability"
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "8. Database Limits"
            Subcategory = "8.1 Theoretical Maximums"
            Component = "8.1.3 MaxAttributesPerObject → Maximum Attributes Per Object"
            CurrentValue = "1024 attributes"
            RecommendedValue = "Use standard attributes"
            Status = "Informational"
            Description = "Attribute limit per AD object"
            Remediation = "Avoid excessive custom attributes"
        }

        # Group Policy Analysis
        Write-ADReportLog -Message "Analyzing Group Policy Limits..." -Type Info -Terminal
        
        $AllGPOs = Get-GPO -All -ErrorAction SilentlyContinue
        if ($AllGPOs) {
            $GPOCount = $AllGPOs.Count
            $LargeGPOs = $AllGPOs | Where-Object { $_.Description -and $_.Description.Length -gt 500 }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "9. Group Policy Objects"
                Subcategory = "9.1 GPO Statistics"
                Component = "9.1.1 TotalGPOCount → Total GPO Count"
                CurrentValue = "$GPOCount GPOs"
                RecommendedValue = "< 999 GPOs per domain"
                Status = if ($GPOCount -gt 800) { "High" } elseif ($GPOCount -gt 500) { "Review" } else { "OK" }
                Description = "Total Group Policy Objects in domain"
                Remediation = if ($GPOCount -gt 500) { "Review GPO structure for consolidation" } else { "GPO count is within reasonable limits" }
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "9. Group Policy Objects"
                Subcategory = "9.1 GPO Statistics"
                Component = "9.1.2 LargeDescriptionGPOs → GPOs with Long Descriptions"
                CurrentValue = if ($LargeGPOs) { "$($LargeGPOs.Count) GPOs" } else { "0 GPOs" }
                RecommendedValue = "Keep descriptions concise"
                Status = if ($LargeGPOs -and $LargeGPOs.Count -gt 10) { "Review" } else { "OK" }
                Description = "GPOs with descriptions > 500 characters"
                Remediation = if ($LargeGPOs -and $LargeGPOs.Count -gt 10) { "Review GPO descriptions for clarity" } else { "GPO descriptions are reasonably sized" }
            }
        }

        # Schema Analysis
        Write-ADReportLog -Message "Analyzing Schema Extensions..." -Type Info -Terminal
        
        $Schema = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) -Filter "objectClass -eq 'attributeSchema'" -ErrorAction SilentlyContinue
        if ($Schema) {
            $CustomAttributes = $Schema | Where-Object { $_.Name -notmatch '^ms-' -and $_.Name -notmatch '^ou' -and $_.Name -notmatch '^system' }
            $SchemaVersion = (Get-ADObject -Identity ((Get-ADRootDSE).schemaNamingContext) -Properties objectVersion -ErrorAction SilentlyContinue).objectVersion
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "10. Schema Extensions"
                Subcategory = "10.1 Schema Metrics"
                Component = "10.1.1 TotalAttributes → Total Schema Attributes"
                CurrentValue = "$($Schema.Count) attributes"
                RecommendedValue = "Monitor extensions"
                Status = "Informational"
                Description = "Total attributes in schema"
                Remediation = "Track schema changes for impact assessment"
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "10. Schema Extensions"
                Subcategory = "10.1 Schema Metrics"
                Component = "10.1.2 CustomAttributes → Custom Schema Attributes"
                CurrentValue = if ($CustomAttributes) { "$($CustomAttributes.Count) attributes" } else { "0 attributes" }
                RecommendedValue = "Minimize custom extensions"
                Status = if ($CustomAttributes -and $CustomAttributes.Count -gt 100) { "Review" } elseif ($CustomAttributes -and $CustomAttributes.Count -gt 50) { "Monitor" } else { "OK" }
                Description = "Non-Microsoft schema attributes"
                Remediation = if ($CustomAttributes -and $CustomAttributes.Count -gt 100) { "Review custom schema extensions for necessity" } else { "Custom extensions are within reasonable limits" }
            }
            
            $QuotaAnalysis += [PSCustomObject]@{
                Category = "10. Schema Extensions"
                Subcategory = "10.1 Schema Metrics"
                Component = "10.1.3 SchemaVersion → Schema Version"
                CurrentValue = if ($SchemaVersion) { "Version $SchemaVersion" } else { "Unknown" }
                RecommendedValue = "Track version changes"
                Status = "Informational"
                Description = "Current AD schema version"
                Remediation = "Document schema version for change tracking"
            }
        }

        # External Integrations
        Write-ADReportLog -Message "Checking External Integrations..." -Type Info -Terminal
        
        # Exchange Integration
        $ExchangeObjects = Get-ADObject -Filter "objectClass -eq 'msExchOrganizationContainer'" -ErrorAction SilentlyContinue
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "11. External Integrations"
            Subcategory = "11.1 Exchange Integration"
            Component = "11.1.1 ExchangeOrganization → Exchange Organization"
            CurrentValue = if ($ExchangeObjects) { "Detected" } else { "Not Present" }
            RecommendedValue = if ($ExchangeObjects) { "Monitor Exchange objects" } else { "No action required" }
            Status = if ($ExchangeObjects) { "Present" } else { "Not Detected" }
            Description = "Microsoft Exchange organization in AD"
            Remediation = if ($ExchangeObjects) { "Monitor Exchange-specific AD objects" } else { "No Exchange-specific monitoring required" }
        }
        
        # DNS Integration
        $DNSZones = Get-ADObject -Filter "objectClass -eq 'dnsZone'" -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$($Domain.DistinguishedName)" -ErrorAction SilentlyContinue
        $QuotaAnalysis += [PSCustomObject]@{
            Category = "11. External Integrations"
            Subcategory = "11.2 DNS Integration"
            Component = "11.2.1 ADIntegratedZones → AD-Integrated DNS Zones"
            CurrentValue = if ($DNSZones) { "$($DNSZones.Count) zones" } else { "0 zones" }
            RecommendedValue = "Minimize complexity"
            Status = if ($DNSZones -and $DNSZones.Count -gt 100) { "Many Zones" } else { "Standard" }
            Description = "DNS zones integrated with AD"
            Remediation = if ($DNSZones -and $DNSZones.Count -gt 100) { "Review DNS zone structure" } else { "DNS zone count is manageable" }
        }

        Write-ADReportLog -Message "Comprehensive quota and limits analysis completed. $($QuotaAnalysis.Count) components analyzed." -Type Info -Terminal
        return $QuotaAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing quotas and limits: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "0. Analysis Error"
            Subcategory = "0.1 System Error"
            Component = "0.1.1 AnalysisFailure → Analysis Failure"
            CurrentValue = "Failed"
            RecommendedValue = "Successful completion"
            Status = "Failed"
            Description = "Could not complete quota and limits analysis"
            Remediation = "Check domain connectivity and permissions. Verify LDAP access to configuration containers."
        })
    }
}

# --- AD-Health Funktionen ---
Function Get-FSMORoles {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Retrieving FSMO role holders..." -Type Info -Terminal

        # Load localized messages
        $msgTable = data {
            #culture="en-US" 
            ConvertFrom-StringData @'
            ForestWide = Forest-wide
            DomainSpecific = Domain-specific
            Online = Online
            Offline = Offline
            SchemaDesc = Manages the Active Directory schema
            DomainNamingDesc = Manages adding and removing domains
            PDCDesc = Time synchronization and password changes
            RIDDesc = Distributes RID pools to domain controllers
            InfraDesc = Manages cross-domain references
'@
        }

        $Forest = Get-ADForest
        $Domain = Get-ADDomain
        
        $FSMORoles = @()
        
        # Forest-wide FSMO roles
        $FSMORoles += [PSCustomObject]@{
            Role = "Schema Master"
            Type = $msgTable.ForestWide
            Server = $Forest.SchemaMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.SchemaDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Domain Naming Master"
            Type = $msgTable.ForestWide
            Server = $Forest.DomainNamingMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.DomainNamingDesc
        }
        
        # Domain-specific FSMO roles
        $FSMORoles += [PSCustomObject]@{
            Role = "PDC Emulator"
            Type = $msgTable.DomainSpecific
            Server = $Domain.PDCEmulator
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.PDCDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "RID Master"
            Type = $msgTable.DomainSpecific
            Server = $Domain.RIDMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.RIDDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Infrastructure Master"
            Type = $msgTable.DomainSpecific
            Server = $Domain.InfrastructureMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.InfraDesc
        }
        
        Write-ADReportLog -Message "$($FSMORoles.Count) FSMO roles found." -Type Info -Terminal
        return $FSMORoles
        
    } catch {
        $ErrorMessage = "Error retrieving FSMO roles: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DomainControllerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Collecting AD Health Data..." -Type Info -Terminal
        
        # Initialize AD Health Report array
        $ADHealthReport = @()
        
        # 1. Forest Information
        try {
            $Forest = Get-ADForest
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information" 
                Parameter = "Forest Name"
                Value = $Forest.Name
                Status = "OK"
                Details = "Forest Functional Level: $($Forest.ForestMode)"
            }

            # Get Schema Version
            $schemaVersion = "Unknown"
            if ($Forest.SchemaVersion) {
                $schemaVersion = if ($Forest.SchemaVersion -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                    $Forest.SchemaVersion[0]
                } else {
                    $Forest.SchemaVersion.ToString()
                }
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information"
                Parameter = "Schema Version"
                Value = $schemaVersion
                Status = "OK" 
                Details = "Current schema version"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information"
                Parameter = "Forest Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 2. Domain Information
        try {
            $Domain = Get-ADDomain
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "Domain Name"
                Value = $Domain.NetBIOSName
                Status = "OK"
                Details = "FQDN: $($Domain.DNSRoot), Level: $($Domain.DomainMode)"
            }

            # PDC Emulator Check
            $pdcStatus = Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if ($pdcStatus) { "OK" } else { "Warning" }
                Details = if ($pdcStatus) { "PDC Emulator is reachable" } else { "PDC Emulator not reachable" }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "Domain Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 3. Domain Controller Health Check
        try {
            $DCs = Get-ADDomainController -Filter * -ErrorAction Stop
            
            $DCCount = @($DCs).Count
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controllers"
                Parameter = "Total DCs"
                Value = $DCCount
                Status = if ($DCCount -ge 2) { "OK" } else { "Warning" }
                Details = "Minimum recommended: 2 DCs for redundancy"
            }

            foreach ($DC in $DCs) {
                # Connectivity Test
                $pingTest = Test-Connection -ComputerName $DC.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controllers"
                    Parameter = "DC Connectivity"
                    Value = $DC.Name
                    Status = if ($pingTest) { "OK" } else { "Critical" }
                    Details = if ($pingTest) { "DC is responding" } else { "DC not responding" }
                }

                # LDAP Service Check
                $ldapTest = Test-NetConnection -ComputerName $DC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controllers"
                    Parameter = "LDAP Service"
                    Value = "$($DC.Name):389"
                    Status = if ($ldapTest) { "OK" } else { "Critical" }
                    Details = if ($ldapTest) { "LDAP service available" } else { "LDAP service unavailable" }
                }

                # Global Catalog Check
                if ($DC.IsGlobalCatalog) {
                    $gcTest = Test-NetConnection -ComputerName $DC.HostName -Port 3268 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    $ADHealthReport += [PSCustomObject]@{
                        Category = "Domain Controllers"
                        Parameter = "Global Catalog"
                        Value = "$($DC.Name):3268"
                        Status = if ($gcTest) { "OK" } else { "Warning" }
                        Details = if ($gcTest) { "Global Catalog service available" } else { "Global Catalog service unavailable" }
                    }
                }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controllers"
                Parameter = "DC Health Check"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 4. FSMO Roles Health Check
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Schema Master"
                Value = $Forest.SchemaMaster
                Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide FSMO role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Domain Naming Master"
                Value = $Forest.DomainNamingMaster
                Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide FSMO role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Domain-wide role for time sync and password changes"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "RID Master"
                Value = $Domain.RIDMaster
                Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Manages RID pool assignments"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Infrastructure Master"
                Value = $Domain.InfrastructureMaster
                Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Manages cross-domain object references"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "FSMO Check"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 5. DNS Health Check
        try {
            $dnsCheck = Resolve-DnsName -Name $Domain.DNSRoot -Type A -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "Domain DNS Resolution"
                Value = $Domain.DNSRoot
                Status = if ($dnsCheck) { "OK" } else { "Warning" }
                Details = if ($dnsCheck) { "DNS resolution successful" } else { "DNS resolution failed" }
            }

            # Check SRV Records
            $srvCheck = Resolve-DnsName -Name "_ldap._tcp.$($Domain.DNSRoot)" -Type SRV -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "LDAP SRV Records"
                Value = "_ldap._tcp.$($Domain.DNSRoot)"
                Status = if ($srvCheck) { "OK" } else { "Warning" }
                Details = if ($srvCheck) { "$($srvCheck.Count) SRV records found" } else { "No SRV records found" }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "DNS Check"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }

        # 6. Replication Health
        try {
            $replData = Get-ADReplicationPartnerMetadata -Target (Get-ADDomainController).HostName -Partition (Get-ADDomain).DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 3
            
            if ($replData) {
                $recentReplCount = ($replData | Where-Object { $_.LastReplicationSuccess -gt (Get-Date).AddHours(-24) }).Count
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Status"
                    Value = "$recentReplCount/$($replData.Count) Partners"
                    Status = if ($recentReplCount -eq $replData.Count) { "OK" } elseif ($recentReplCount -gt 0) { "Warning" } else { "Critical" }
                    Details = "Successful replication in last 24 hours"
                }
            }
            else {
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Status"
                    Value = "No Partners"
                    Status = "Warning"
                    Details = "No replication partners found"
                }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Replication"
                Parameter = "Replication Check"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }

        # 7. System Information
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "Server Name"
                Value = $env:COMPUTERNAME
                Status = "Info"
                Details = "Context: $($env:USERDOMAIN)\$($env:USERNAME)"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "PowerShell Version"
                Value = $PSVersionTable.PSVersion.ToString()
                Status = "Info"
                Details = "AD Module: $(if (Get-Module -ListAvailable -Name ActiveDirectory) { 'Available' } else { 'Not Available' })"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "System Time"
                Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Info"
                Details = "Time Zone: $((Get-TimeZone).DisplayName)"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "System Check"
                Value = "Partial Error"
                Status = "Info"
                Details = $_.Exception.Message
            }
        }

        # Generate Summary
        $criticalCount = ($ADHealthReport | Where-Object { $_.Status -eq "Critical" }).Count
        $warningCount = ($ADHealthReport | Where-Object { $_.Status -eq "Warning" }).Count
        $okCount = ($ADHealthReport | Where-Object { $_.Status -eq "OK" }).Count

        $summary = [PSCustomObject]@{
            Category = "=== SUMMARY ==="
            Parameter = "AD Health Status"
            Value = if ($criticalCount -eq 0 -and $warningCount -eq 0) { "Healthy" } 
                   elseif ($criticalCount -eq 0) { "Minor Issues" } 
                   else { "Critical Issues" }
            Status = if ($criticalCount -eq 0 -and $warningCount -eq 0) { "OK" }
                    elseif ($criticalCount -eq 0) { "Warning" }
                    else { "Critical" }
            Details = "OK: $okCount, Warnings: $warningCount, Critical: $criticalCount"
        }

        # Add summary to beginning of report
        $FinalReport = @($summary) + $ADHealthReport

        Write-ADReportLog -Message "AD Health Check completed: $($FinalReport.Count) checks performed." -Type Info -Terminal
        return $FinalReport

    }
    catch {
        $ErrorMessage = "Critical error during AD Health Check: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "ERROR"
            Parameter = "AD Health Check"
            Value = "Failed"
            Status = "Critical"
            Details = $ErrorMessage
        })
    }
}

Function Get-ReplicationStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD replication status..." -Type Info -Terminal
        
        $ReplicationData = @()
        $DomainControllers = Get-ADDomainController -Filter *
        
        # 1. Collect replication failures for all DCs
        Write-ADReportLog -Message "Checking replication failures..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $ReplicationFailures = Get-ADReplicationFailure -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($ReplicationFailures) {
                    foreach ($failure in $ReplicationFailures) {
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Replication Failures"
                            SourceDC = $DC.Name
                            TargetDC = $failure.Partner
                            Partition = $failure.PartitionDN
                            Status = "ERROR"
                            ErrorType = $failure.FailureType
                            ErrorCount = $failure.FailureCount
                            FirstFailure = $failure.FirstFailureTime
                            LastFailure = $failure.LastFailureTime
                            Details = $failure.LastError
                        }
                    }
                } else {
                    # No errors for this DC
                    $ReplicationData += [PSCustomObject]@{
                        Category = "Replication Status"
                        SourceDC = $DC.Name
                        TargetDC = "All Partners"
                        Partition = "All Partitions"
                        Status = "OK"
                        ErrorType = "None"
                        ErrorCount = 0
                        FirstFailure = $null
                        LastFailure = $null
                        Details = "No replication failures found"
                    }
                }
            } catch {
                $ReplicationData += [PSCustomObject]@{
                    Category = "System Error"
                    SourceDC = $DC.Name
                    TargetDC = "N/A"
                    Partition = "N/A"
                    Status = "CRITICAL"
                    ErrorType = "Query Error"
                    ErrorCount = 1
                    FirstFailure = Get-Date
                    LastFailure = Get-Date
                    Details = "Error retrieving replication data: $($_.Exception.Message)"
                }
            }
        }
        
        # 2. Collect replication partner metadata
        Write-ADReportLog -Message "Collecting replication partner metadata..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $PartnerMetadata = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($PartnerMetadata) {
                    foreach ($partner in $PartnerMetadata) {
                        $timeSinceLastSync = "Unknown"
                        $syncStatus = "Unknown"
                        
                        if ($partner.LastReplicationSuccess) {
                            $timeSince = (Get-Date) - $partner.LastReplicationSuccess
                            $timeSinceLastSync = "$([math]::Round($timeSince.TotalHours, 1)) hours"
                            
                            # Status based on time since last sync
                            if ($timeSince.TotalHours -lt 1) {
                                $syncStatus = "EXCELLENT"
                            } elseif ($timeSince.TotalHours -lt 6) {
                                $syncStatus = "GOOD"
                            } elseif ($timeSince.TotalHours -lt 24) {
                                $syncStatus = "WARNING"
                            } else {
                                $syncStatus = "CRITICAL"
                            }
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Partner Metadata"
                            SourceDC = $DC.Name
                            TargetDC = $partner.Partner -replace ".*CN=NTDS Settings,CN=([^,]+),.*", '$1'
                            Partition = $partner.Partition
                            Status = $syncStatus
                            ErrorType = if ($partner.ConsecutiveReplicationFailures -gt 0) { "Consecutive Failures" } else { "None" }
                            ErrorCount = $partner.ConsecutiveReplicationFailures
                            FirstFailure = $partner.LastReplicationAttempt
                            LastFailure = $partner.LastReplicationSuccess
                            Details = "Last sync: $timeSinceLastSync ago | USN: $($partner.LastReplicationSuccess)"
                        }
                    }
                }
            } catch {
                Write-Warning "Error retrieving partner metadata for $($DC.Name): $($_.Exception.Message)"
            }
        }
        
        # 3. Additional replication health checks
        Write-ADReportLog -Message "Performing extended replication diagnostics..." -Type Info -Terminal
        
        # DC-to-DC connectivity tests
        foreach ($SourceDC in $DomainControllers) {
            foreach ($TargetDC in $DomainControllers) {
                if ($SourceDC.Name -ne $TargetDC.Name) {
                    try {
                        # Test LDAP connectivity
                        $ldapTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        # Test RPC connectivity (for replication)
                        $rpcTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 135 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        $connectionStatus = "OK"
                        $connectionDetails = "LDAP: OK, RPC: OK"
                        
                        if (-not $ldapTest -and -not $rpcTest) {
                            $connectionStatus = "CRITICAL"
                            $connectionDetails = "LDAP: FAILED, RPC: FAILED"
                        } elseif (-not $ldapTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: FAILED, RPC: OK"
                        } elseif (-not $rpcTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: OK, RPC: FAILED"
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "DC Connectivity"
                            SourceDC = $SourceDC.Name
                            TargetDC = $TargetDC.Name
                            Partition = "Network Connectivity"
                            Status = $connectionStatus
                            ErrorType = if ($connectionStatus -ne "OK") { "Network Issue" } else { "None" }
                            ErrorCount = if ($connectionStatus -eq "CRITICAL") { 2 } elseif ($connectionStatus -eq "WARNING") { 1 } else { 0 }
                            FirstFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            LastFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            Details = $connectionDetails
                        }
                    } catch {
                        Write-Warning "Error testing connectivity between $($SourceDC.Name) and $($TargetDC.Name)"
                    }
                }
            }
        }
        
        # 4. Create replication summary
        $totalErrors = ($ReplicationData | Where-Object { $_.Status -eq "CRITICAL" -or $_.Status -eq "ERROR" }).Count
        $warnings = ($ReplicationData | Where-Object { $_.Status -eq "WARNING" }).Count
        $okCount = ($ReplicationData | Where-Object { $_.Status -eq "OK" -or $_.Status -eq "GOOD" -or $_.Status -eq "EXCELLENT" }).Count
        
        # Add summary to the beginning
        $summary = [PSCustomObject]@{
            Category = "SUMMARY"
            SourceDC = "All DCs"
            TargetDC = "All Partners"
            Partition = "N/A"
            Status = if ($totalErrors -eq 0 -and $warnings -eq 0) { "HEALTHY" } elseif ($totalErrors -eq 0) { "MINOR ISSUES" } else { "CRITICAL ISSUES" }
            ErrorType = "Overview"
            ErrorCount = $totalErrors
            FirstFailure = $null
            LastFailure = $null
            Details = "OK/Good: $okCount | Warnings: $warnings | Critical/Error: $totalErrors | DCs: $($DomainControllers.Count)"
        }
        
        # Final report with summary
        $FinalReplicationData = @($summary) + $ReplicationData
        
        Write-ADReportLog -Message "Replication analysis completed: $($FinalReplicationData.Count) entries found." -Type Info -Terminal
        return $FinalReplicationData
        
    } catch {
        $ErrorMessage = "Critical error during replication analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "SYSTEM ERROR"
            SourceDC = "Unknown"
            TargetDC = "Unknown"
            Partition = "N/A"
            Status = "CRITICAL"
            ErrorType = "System Error"
            ErrorCount = 1
            FirstFailure = Get-Date
            LastFailure = Get-Date
            Details = $ErrorMessage
        })
    }
}

# --- Funktion zum Abrufen von OU-Hierarchie-Berichten als Baumstruktur ---
Function Get-ADOUHierarchyReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Generating OU hierarchy report as tree structure..." -Type Info -Terminal
    try {
        # Get OUs only, no containers
        $AllOUs = Get-ADOrganizationalUnit -Filter * `
            -Properties DistinguishedName, Name, Description, whenCreated, whenChanged, 
                       ProtectedFromAccidentalDeletion, LinkedGroupPolicyObjects, ManagedBy -ErrorAction Stop
        
        if (-not $AllOUs) {
            Write-ADReportLog -Message "No organizational units found in AD." -Type Warning -Terminal
            return $null
        }

        # Get domain information for root level
        $Domain = Get-ADDomain -ErrorAction Stop
        $DomainDN = $Domain.DistinguishedName
        
        # Create hierarchy dictionary and root OUs list
        $HierarchyDict = @{}
        $RootOUs = [System.Collections.ArrayList]@()
        
        foreach ($OU in $AllOUs) {
            $Level = 0
            $ParentDN = $OU.DistinguishedName.Substring($OU.DistinguishedName.IndexOf(',') + 1)
            
            # Calculate indentation depth based on OU path
            $Level = ($OU.DistinguishedName.Split(',').Count - $DomainDN.Split(',').Count)
            $Indent = "-- " * $Level

            # Get linked GPOs count
            $GPOCount = if ($OU.LinkedGroupPolicyObjects) { 
                @($OU.LinkedGroupPolicyObjects).Count 
            } else { 
                0 
            }

            # Get manager info if exists
            $Manager = if ($OU.ManagedBy) {
                try {
                    $mgr = Get-ADObject $OU.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    $mgr.DisplayName
                } catch {
                    "Unknown"
                }
            } else {
                "Not Set"
            }
            
            $Entry = [PSCustomObject]@{
                Level = $Level
                Name = $OU.Name
                FullPath = $OU.DistinguishedName
                DisplayName = "$Indent$($OU.Name)"
                Description = $OU.Description
                Created = $OU.whenCreated
                Modified = $OU.whenChanged
                ParentDN = if ($ParentDN -eq $DomainDN) { "Root" } else { $ParentDN }
                Protected = $OU.ProtectedFromAccidentalDeletion
                LinkedGPOs = $GPOCount
                ManagedBy = $Manager
                Status = if ($OU.ProtectedFromAccidentalDeletion) { "Protected" } else { "Not Protected" }
            }
            
            if ($ParentDN -eq $DomainDN) {
                [void]$RootOUs.Add($Entry)
            } else {
                if (-not $HierarchyDict.ContainsKey($ParentDN)) {
                    $HierarchyDict[$ParentDN] = [System.Collections.ArrayList]@()
                }
                [void]$HierarchyDict[$ParentDN].Add($Entry)
            }
        }
        
        # Create sorted output list
        $Results = [System.Collections.ArrayList]@()
        
        # Add domain root as first entry
        [void]$Results.Add([PSCustomObject]@{
            Level = 0
            DisplayName = $Domain.NetBIOSName
            Description = "Domain Root"
            Created = $null
            Modified = $null
            FullPath = $DomainDN
            ParentDN = $null
            Protected = $true
            LinkedGPOs = 0
            ManagedBy = "System"
            Status = "Domain Root"
        })
        
        # Recursive function to build tree structure
        function Add-ChildOUs {
            param($ParentDN)
            
            $Children = $HierarchyDict[$ParentDN] | Sort-Object Name
            foreach ($Child in $Children) {
                [void]$Results.Add($Child)
                if ($HierarchyDict.ContainsKey($Child.FullPath)) {
                    Add-ChildOUs -ParentDN $Child.FullPath
                }
            }
        }
        
        # Process root level OUs
        foreach ($RootOU in ($RootOUs | Sort-Object Name)) {
            [void]$Results.Add($RootOU)
            if ($HierarchyDict.ContainsKey($RootOU.FullPath)) {
                Add-ChildOUs -ParentDN $RootOU.FullPath
            }
        }
        
        Write-ADReportLog -Message "OU hierarchy report successfully created. Found $($Results.Count) entries." -Type Info -Terminal
        return $Results
    }
    catch {
        Write-ADReportLog -Message "Error creating OU hierarchy report: $($_.Exception.Message)" -Type Error -Terminal
        return $null
    }
}

# --- Function to get AD Sites and Subnets Report ---
Function Get-ADSitesAndSubnetsReport {
    [CmdletBinding()]
    param()

    # Initialize message table for localization
    $msgTable = data {
        #culture="en-US" 
        ConvertFrom-StringData @'
        GatheringInfo = Gathering AD Sites and Subnets information...
        NoServersFound = Could not get servers for site {0}: {1}
        NoSiteLinksFound = Could not get site links for site {0}: {1}
        SitesRetrieved = Successfully retrieved {0} AD Replication Sites.
        NoSitesFound = No AD Replication Sites found.
        SubnetsRetrieved = Successfully retrieved {0} AD Replication Subnets.
        NoSubnetsFound = No AD Replication Subnets found.
        ReportGenerated = Successfully generated Sites and Subnets Report for {0} entries.
        NoDataFound = No data found for Sites and Subnets Report.
        ErrorGenerating = Error generating Sites and Subnets Report: {0}
'@
    }

    Write-ADReportLog -Message $msgTable.GatheringInfo -Type Info -Terminal
    Initialize-ResultCounters

    try {
        $Report = [System.Collections.ArrayList]@()

        # Get Sites
        $Sites = Get-ADReplicationSite -Filter * -Properties Description, Options, InterSiteTopologyGenerator -ErrorAction Stop
        if ($Sites -and $Sites.Count -gt 0) {
            foreach ($Site in $Sites) {
                # Get servers for this site separately
                $ServersCount = 0
                try {
                    $Servers = Get-ADDomainController -Filter { Site -eq $Site.Name } -ErrorAction Stop
                    $ServersCount = if ($Servers) { @($Servers).Count } else { 0 }
                } catch {
                    Write-ADReportLog -Message ($msgTable.NoServersFound -f $Site.Name, $_.Exception.Message) -Type Warning
                }

                # Get Site Links separately
                $SiteLinksText = ""
                try {
                    $SiteLinks = Get-ADReplicationSiteLink -Filter "Sites -eq '$($Site.DistinguishedName)'" -ErrorAction Stop
                    $SiteLinksText = if ($SiteLinks) { ($SiteLinks | ForEach-Object { $_.Name }) -join ", " }
                } catch {
                    Write-ADReportLog -Message ($msgTable.NoSiteLinksFound -f $Site.Name, $_.Exception.Message) -Type Warning
                }

                [void]$Report.Add([PSCustomObject]@{
                    Type = "Site"
                    Name = $Site.Name
                    DistinguishedName = $Site.DistinguishedName
                    Description = $Site.Description
                    ServersInSiteCount = $ServersCount
                    InterSiteTopologyGenerator = if ($Site.InterSiteTopologyGenerator) { $Site.InterSiteTopologyGenerator } else { "N/A" }
                    Options = if ($Site.Options) { $Site.Options } else { "None" }
                    SiteLinks = if ($SiteLinksText) { $SiteLinksText } else { "None" }
                    Location = "N/A"
                    AssociatedSite = "N/A"
                })
            }
            Write-ADReportLog -Message ($msgTable.SitesRetrieved -f $Sites.Count) -Type Info -Terminal
        } else {
            Write-ADReportLog -Message $msgTable.NoSitesFound -Type Warning -Terminal
        }

        # Get Subnets
        $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location, Site -ErrorAction Stop
        if ($Subnets -and $Subnets.Count -gt 0) {
            foreach ($Subnet in $Subnets) {
                [void]$Report.Add([PSCustomObject]@{
                    Type = "Subnet"
                    Name = $Subnet.Name
                    DistinguishedName = $Subnet.DistinguishedName
                    Description = if ($Subnet.Description) { $Subnet.Description } else { "N/A" }
                    ServersInSiteCount = "N/A"
                    InterSiteTopologyGenerator = "N/A"
                    Options = "N/A"
                    SiteLinks = "N/A"
                    Location = if ($Subnet.Location) { $Subnet.Location } else { "N/A" }
                    AssociatedSite = if ($Subnet.Site) {
                        try { 
                            (Get-ADReplicationSite -Identity $Subnet.Site -ErrorAction Stop).Name 
                        } catch { 
                            $Subnet.Site 
                        }
                    } else { "N/A" }
                })
            }
            Write-ADReportLog -Message ($msgTable.SubnetsRetrieved -f $Subnets.Count) -Type Info -Terminal
        } else {
            Write-ADReportLog -Message $msgTable.NoSubnetsFound -Type Warning -Terminal
        }

        if ($Report.Count -gt 0) {
            Write-ADReportLog -Message ($msgTable.ReportGenerated -f $Report.Count) -Type Info
            return $Report | Sort-Object Type, Name
        } else {
            Write-ADReportLog -Message $msgTable.NoDataFound -Type Info
            return @(
                [PSCustomObject]@{
                    Type = "Information"
                    Name = "No Data"
                    DistinguishedName = "N/A"
                    Description = "No Sites or Subnets found in Active Directory"
                    ServersInSiteCount = 0
                    InterSiteTopologyGenerator = "N/A"
                    Options = "N/A"
                    SiteLinks = "N/A"
                    Location = "N/A"
                    AssociatedSite = "N/A"
                }
            )
        }
    }
    catch {
        $ErrorMessage = $msgTable.ErrorGenerating -f $_.Exception.Message
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return @(
            [PSCustomObject]@{
                Type = "Error"
                Name = "System Error"
                DistinguishedName = "N/A" 
                Description = $ErrorMessage
                ServersInSiteCount = 0
                InterSiteTopologyGenerator = "N/A"
                Options = "N/A"
                SiteLinks = "N/A"
                Location = "N/A"
                AssociatedSite = "N/A"
            }
        )
    }
}


# --- Neue Roadmap Features - Benutzer-Reports ---
Function Get-StalePasswords {
    [CmdletBinding()]
    param([int]$Days = 90)
    
    try {
        Write-ADReportLog -Message "Analyzing users with stale passwords (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, DisplayName, SamAccountName, Department, Title -ErrorAction Stop
        $StalePasswordUsers = $Users | Where-Object { $_.PasswordLastSet -lt $CutoffDate }
        
        $Results = $StalePasswordUsers | Select-Object DisplayName, SamAccountName, Department, Title, 
            @{Name="PasswordLastSet";Expression={$_.PasswordLastSet}},
            @{Name="DaysSinceLastChange";Expression={(New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing stale passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NeverChangingPasswords {
    [CmdletBinding()]
    param([int]$Days = 365)
    
    try {
        Write-ADReportLog -Message "Analyzing users with passwords that never change or are very old (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, WhenCreated, DisplayName, SamAccountName, Department, Title, PasswordNeverExpires -ErrorAction Stop
        
        $NeverChangedUsers = $Users | Where-Object { 
            # Passwort nie geÃ¤ndert seit Erstellung
            ($_.PasswordLastSet -eq $_.WhenCreated) -or 
            # Passwort ist null
            ($_.PasswordLastSet -eq $null) -or
            # Passwort ist sehr alt
            ($_.PasswordLastSet -lt $CutoffDate) -or
            # Passwort lÃ¤uft nie ab und ist alt
            ($_.PasswordNeverExpires -eq $true -and $_.PasswordLastSet -lt $CutoffDate)
        }
        
        $Results = foreach ($user in $NeverChangedUsers) {
            $daysSinceChange = if ($user.PasswordLastSet) { 
                (New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).Days 
            } else { 
                if ($user.WhenCreated) { (New-TimeSpan -Start $user.WhenCreated -End (Get-Date)).Days } else { 9999 }
            }
            
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Title = $user.Title
                WhenCreated = $user.WhenCreated
                PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet } else { "Never" }
                PasswordNeverExpires = $user.PasswordNeverExpires
                DaysSinceLastChange = $daysSinceChange
                RiskLevel = if ($daysSinceChange -gt 730) { "Critical" } elseif ($daysSinceChange -gt 365) { "High" } else { "Medium" }
            }
        }
        
        Write-ADReportLog -Message "Never changing passwords analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastChange -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing never changing passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ExpiringAccounts {
    [CmdletBinding()]
    param([int]$Days = 30)
    
    try {
        Write-ADReportLog -Message "Analyzing accounts expiring within $Days days..." -Type Info -Terminal
        $FutureDate = (Get-Date).AddDays($Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true -and AccountExpirationDate -like '*'" -Properties AccountExpirationDate, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ExpiringUsers = $Users | Where-Object { $_.AccountExpirationDate -le $FutureDate -and $_.AccountExpirationDate -gt (Get-Date) }
        
        $Results = $ExpiringUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AccountExpirationDate";Expression={$_.AccountExpirationDate}},
            @{Name="DaysUntilExpiration";Expression={(New-TimeSpan -Start (Get-Date) -End $_.AccountExpirationDate).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing expiring accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ReversibleEncryptionUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with reversible encryption..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties AllowReversiblePasswordEncryption, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ReversibleUsers = $Users | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true }
        
        $Results = $ReversibleUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AllowReversiblePasswordEncryption";Expression={$_.AllowReversiblePasswordEncryption}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing reversible encryption users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-KerberosDESUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing all users with Kerberos DES encryption (enabled, disabled, and system accounts)..." -Type Info -Terminal
        
        # Get ALL users (enabled, disabled, system accounts)
        $AllUsers = Get-ADUser -Filter * -Properties UserAccountControl, DisplayName, SamAccountName, Department, Enabled, Description, ObjectClass, LastLogonDate, PasswordLastSet -ErrorAction Stop
        
        # Filter users with DES encryption flag (0x200000)
        $DESUsers = $AllUsers | Where-Object { ($_.UserAccountControl -band 0x200000) -ne 0 }
        
        $Results = foreach ($user in $DESUsers) {
            try {
                # Determine account type
                $accountType = "Regular User"
                $riskLevel = "High"
                
                # Check if it's a system/service account
                if ($user.SamAccountName -match '^\$' -or 
                    $user.SamAccountName -like "*svc*" -or 
                    $user.SamAccountName -like "*service*" -or
                    $user.Description -match "(service|system|computer)" -or
                    $user.SamAccountName -like "*krbtgt*") {
                    $accountType = "System/Service Account"
                    $riskLevel = "Medium"  # System accounts with DES are less critical but still risky
                }
                
                # Check if it's a computer account
                if ($user.SamAccountName -match '\$$' -and $user.ObjectClass -eq "user") {
                    $accountType = "Computer Account"
                    $riskLevel = "Low"  # Computer accounts with DES are legacy but lower risk
                }
                
                # Special handling for krbtgt account
                if ($user.SamAccountName -eq "krbtgt") {
                    $accountType = "Domain Controller Service Account"
                    $riskLevel = "Critical"  # krbtgt with DES is very dangerous
                }
                
                # Adjust risk based on enabled status
                $statusNote = ""
                if (-not $user.Enabled) {
                    $statusNote = " (Disabled)"
                    # Reduce risk for disabled accounts but keep some level
                    if ($riskLevel -eq "Critical") { $riskLevel = "High" }
                    elseif ($riskLevel -eq "High") { $riskLevel = "Medium" }
                    elseif ($riskLevel -eq "Medium") { $riskLevel = "Low" }
                }
                
                # Create recommendation
                $recommendation = switch ($riskLevel) {
                    "Critical" { "URGENT: Disable DES encryption immediately for this critical account" }
                    "High" { "High Priority: Remove DES encryption support and use AES encryption" }
                    "Medium" { "Medium Priority: Upgrade encryption to AES when possible" }
                    "Low" { "Low Priority: Consider upgrading to modern encryption methods" }
                }
                
                [PSCustomObject]@{
                    DisplayName = if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName }
                    SamAccountName = $user.SamAccountName
                    Department = if ($user.Department) { $user.Department } else { "Not Specified" }
                    AccountType = $accountType + $statusNote
                    Enabled = $user.Enabled
                    UsesDESEncryption = "Yes"
                    RiskLevel = $riskLevel
                    Recommendation = $recommendation
                    LastLogonDate = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                    PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                }
                
            } catch {
                Write-ADReportLog -Message "Error processing DES user $($user.SamAccountName): $($_.Exception.Message)" -Type Warning
                
                # Return error entry
                [PSCustomObject]@{
                    DisplayName = if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName }
                    SamAccountName = $user.SamAccountName
                    Department = "Error"
                    AccountType = "Error Processing"
                    Enabled = "Unknown"
                    UsesDESEncryption = "Error"
                    RiskLevel = "Critical"
                    Recommendation = "Error analyzing account - manual review required"
                    LastLogonDate = "Error"
                    PasswordLastSet = "Error"
                    Description = "Error during analysis"
                }
            }
        }
        
        # Sort by risk level and then by account type
        $SortedResults = $Results | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, SamAccountName
        
        Write-ADReportLog -Message "Kerberos DES encryption analysis completed. Found $($SortedResults.Count) accounts with DES encryption." -Type Info -Terminal
        
        # Log statistics
        if ($SortedResults.Count -gt 0) {
            $criticalCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "Critical" }).Count
            $highCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "High" }).Count
            $enabledCount = ($SortedResults | Where-Object { $_.Enabled -eq $true }).Count
            $systemCount = ($SortedResults | Where-Object { $_.AccountType -match "System|Service|Computer" }).Count
            
            Write-ADReportLog -Message "DES Statistics - Critical: $criticalCount, High: $highCount, Enabled: $enabledCount, System/Service Accounts: $systemCount" -Type Info -Terminal
        }
        
        return $SortedResults
        
    } catch {
        Write-ADReportLog -Message "Error analyzing Kerberos DES users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-KerberoastableAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Kerberoastable accounts..." -Type Info -Terminal
        
        # Get domain for language detection
        $domain = Get-ADDomain
        $isGermanDomain = $domain.DNSRoot -match "\.de$" -or $domain.NetBIOSName -match "DE$"

        # Define privileged groups in both languages
        $privilegedGroups = @(
            if($isGermanDomain) {
                "Domänen-Admins", "Schema-Admins", "Organisations-Admins", "Administratoren"
            }
            "Domain Admins", "Schema Admins", "Enterprise Admins", "Administrators"
        )

        # Get accounts with SPNs
        $kerberoastableAccounts = Get-ADUser -Filter {ServicePrincipalName -like "*" -and Enabled -eq $true} -Properties ServicePrincipalName, LastLogonDate, PasswordLastSet, MemberOf, Department, Description -ErrorAction Stop

        $Results = foreach($account in $kerberoastableAccounts) {
            # Calculate password age
            $passwordAge = if($account.PasswordLastSet) {
                [math]::Round((New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days, 0)
            } else { 999 }

            # Check privileged group membership
            $isPrivileged = $false
            foreach($group in $account.MemberOf) {
                $groupName = (Get-ADGroup $group).Name
                if($privilegedGroups -contains $groupName) {
                    $isPrivileged = $true
                    break
                }
            }

            # Determine risk level
            $riskLevel = switch($true) {
                { $isPrivileged -and $passwordAge -gt 30 } { "Critical" }
                { $isPrivileged } { "High" }
                { $passwordAge -gt 180 } { "High" }
                { $passwordAge -gt 90 } { "Medium" }
                default { "Low" }
            }

            # Create recommendation based on findings
            $recommendation = switch($true) {
                { $isPrivileged -and $passwordAge -gt 30 } { 
                    "Critical: Privileged account with SPN - Change password immediately and review necessity"
                }
                { $isPrivileged } {
                    "High Risk: Privileged account with SPN - Review if SPN is required"
                }
                { $passwordAge -gt 180 } {
                    "Change password (not changed for >180 days)"
                }
                { $passwordAge -gt 90 } {
                    "Consider password change (not changed for >90 days)"
                }
                default {
                    "Monitor account regularly"
                }
            }

            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Department = if($account.Department) { $account.Department } else { "Not specified" }
                SPNCount = ($account.ServicePrincipalName | Measure-Object).Count
                PasswordAge = $passwordAge
                LastLogon = if($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                IsPrivileged = $isPrivileged
                RiskLevel = $riskLevel
                Description = $account.Description
                Recommendation = $recommendation
            }
        }

        Write-ADReportLog -Message "Found $($Results.Count) Kerberoastable accounts" -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, PasswordAge -Descending

    } catch {
        Write-ADReportLog -Message "Error analyzing Kerberoastable accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}


Function Get-UsersWithSPN {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Starting comprehensive SPN security analysis..." -Type Info -Terminal

        # Get domain for language detection
        $domain = Get-ADDomain
        $isGermanDomain = $domain.DNSRoot -match "\.de$" -or $domain.NetBIOSName -match "DE$"

        # Define privileged groups in both languages
        $privilegedGroups = @(
            if($isGermanDomain) {
                "Domänen-Admins", "Schema-Admins", "Organisations-Admins", "Administratoren"
            }
            "Domain Admins", "Schema Admins", "Enterprise Admins", "Administrators" 
        )

        Write-ADReportLog -Message "Searching for service accounts and accounts with SPNs..." -Type Info -Terminal

        # Build filter string for service accounts
        $filterString = @"
            (Enabled -eq `$true) -and (
                (ServicePrincipalName -like '*') -or
                (SamAccountName -like '*svc*') -or
                (SamAccountName -like '*service*') -or
                (SamAccountName -like '*srv*') -or
                (SamAccountName -like '*dienst*') -or
                (SamAccountName -like '*konto*') -or
                (SamAccountName -like '*automated*') -or
                (SamAccountName -like '*automation*') -or
                (SamAccountName -like '*auto*') -or
                (SamAccountName -like '*batch*') -or
                (SamAccountName -like '*task*') -or
                (SamAccountName -like '*job*') -or
                (SamAccountName -like '*admin*') -or
                (SamAccountName -like '*system*') -or
                (SamAccountName -like '*backup*') -or
                (SamAccountName -like '*report*') -or
                (SamAccountName -like '*account*') -or
                (Description -like '*service*') -or
                (Description -like '*dienst*') -or
                (Description -like '*konto*') -or
                (Description -like '*automated*') -or
                (Description -like '*automation*') -or
                (Description -like '*auto*') -or
                (Description -like '*batch*') -or
                (Description -like '*task*') -or
                (Description -like '*job*') -or
                (Description -like '*admin*') -or
                (Description -like '*system*') -or
                (Description -like '*backup*') -or
                (Description -like '*report*') -or
                (Description -like '*account*') -or
                (SamAccountName -like '*skript*') -or
                (Description -like '*skript*') -or
                (SamAccountName -like '*prozess*') -or
                (Description -like '*prozess*')
            )
"@

        # Get all user accounts that could be service accounts
        $Users = Get-ADUser -Filter $filterString -Properties @(
            'ServicePrincipalName',
            'DisplayName',
            'SamAccountName',
            'Department',
            'Description',
            'Enabled',
            'LastLogonDate',
            'PasswordLastSet',
            'memberOf',
            'userAccountControl',
            'msDS-SupportedEncryptionTypes'
        ) -ErrorAction Stop

        Write-ADReportLog -Message "Found $($Users.Count) potential service accounts for analysis" -Type Info -Terminal

        $Results = foreach ($user in $Users) {
            # Initialize risk tracking
            $riskFactors = [System.Collections.ArrayList]@()
            $riskLevel = "Low"
            $detectionReason = [System.Collections.ArrayList]@()

            # Check why this account was detected as service account
            if ($user.ServicePrincipalName) {
                $detectionReason.Add("Has Service Principal Name (SPN) configuration") | Out-Null
            }
            if ($user.SamAccountName -match "(svc|service|srv|dienst|konto|automated|automation|auto|batch|task|job|admin|system|backup|report|account|skript|prozess)") {
                $detectionReason.Add("Username indicates service account") | Out-Null
            }
            if ($user.Description -match "(service|dienst|konto|automated|automation|auto|batch|task|job|admin|system|backup|report|account|skript|prozess)") {
                $detectionReason.Add("Description indicates service account") | Out-Null
            }

            # Check SPN configuration
            if ($user.ServicePrincipalName) {
                $riskFactors.Add("Has active SPN configuration") | Out-Null
                $riskLevel = "Medium"

                # Check for weak encryption
                if (-not $user.'msDS-SupportedEncryptionTypes' -or 
                    ($user.'msDS-SupportedEncryptionTypes' -band 0x3) -eq 0) {
                    $riskFactors.Add("Weak Kerberos encryption") | Out-Null
                    $riskLevel = "High"
                }
            }

            # Check account status and privileges
            if (-not $user.Enabled -and $user.ServicePrincipalName) {
                $riskFactors.Add("Disabled account with active SPN") | Out-Null
                $riskLevel = "High"
            }

            # Check privileged group memberships
            foreach ($group in $privilegedGroups) {
                if ($user.memberOf -match $group) {
                    $riskFactors.Add("Member of privileged group: $group") | Out-Null
                    $riskLevel = "Critical"
                    break
                }
            }

            # Check password age
            if ($user.PasswordLastSet) {
                $passwordAge = (Get-Date) - $user.PasswordLastSet
                if ($passwordAge.Days -gt 365) {
                    $riskFactors.Add("Password older than 1 year") | Out-Null
                    if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                }
            }

            # Create result object
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = if ($user.Department) { $user.Department } else { "Not specified" }
                SPNCount = if ($user.ServicePrincipalName) { @($user.ServicePrincipalName).Count } else { 0 }
                ServicePrincipalNames = if ($user.ServicePrincipalName) { @($user.ServicePrincipalName) -join "; " } else { "None" }
                AccountEnabled = $user.Enabled
                LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet } else { "Never" }
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors.Count -gt 0) { $riskFactors -join " | " } else { "None" }
                DetectionReason = if ($detectionReason.Count -gt 0) { $detectionReason -join " | " } else { "Unknown" }
                SecurityRecommendation = switch ($riskLevel) {
                    "Critical" { "URGENT: Review privileged access and SPN configuration immediately" }
                    "High" { "High Priority: Verify account status and update security settings" }
                    "Medium" { "Review SPN configuration and account permissions" }
                    default { "Monitor account for security changes" }
                }
            }
        }

        Write-ADReportLog -Message "SPN security analysis completed. Found $($Results.Count) accounts requiring attention." -Type Info -Terminal
        return $Results | Sort-Object { 
            switch ($_.RiskLevel) {
                "Critical" { 0 }
                "High" { 1 }
                "Medium" { 2 }
                "Low" { 3 }
                default { 4 }
            }
        }

    } catch {
        Write-ADReportLog -Message "Error in SPN security analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere User Report Funktionen ---
Function Get-GuestAccountStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info -Terminal
        
        # Suche nach dem Guest Account (kann verschiedene Namen haben)
        $GuestAccounts = @()
        
        # Standard Guest Account Namen in verschiedenen Sprachen
        $GuestNames = @("Guest", "Gast", "InvitÃ©", "Invitado", "Ospite")
        
        foreach ($name in $GuestNames) {
            $account = Get-ADUser -Filter "SamAccountName -eq '$name'" -Properties * -ErrorAction SilentlyContinue
            if ($account) {
                $GuestAccounts += $account
            }
        }
        
        # Auch nach SID suchen (Well-known Guest SID endet mit -501)
        $DomainSID = (Get-ADDomain).DomainSID.Value
        $GuestSID = "$DomainSID-501"
        
        try {
            $GuestBySID = Get-ADUser -Identity $GuestSID -Properties * -ErrorAction SilentlyContinue
            if ($GuestBySID -and -not ($GuestAccounts | Where-Object { $_.SID -eq $GuestBySID.SID })) {
                $GuestAccounts += $GuestBySID
            }
        } catch {
            # SID nicht gefunden ist OK
        }
        
        if ($GuestAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No Guest account found in domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                AccountName = "No Guest Account"
                Status = "Not Found"
                SecurityRisk = "Low"
                Recommendation = "Guest account not present in domain (good security practice)"
            })
        }
        
        $Results = foreach ($guest in $GuestAccounts) {
            # Analysiere Sicherheitsrisiken
            $risks = @()
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($guest.Enabled) {
                $risks += "Account is enabled"
                $riskLevel = "High"
                $recommendations += "Disable Guest account immediately"
            }
            
            if ($guest.PasswordNeverExpires) {
                $risks += "Password never expires"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
                $recommendations += "Set password expiration"
            }
            
            if ($guest.PasswordNotRequired) {
                $risks += "Password not required"
                $riskLevel = "Critical"
                $recommendations += "Enforce password requirement"
            }
            
            if ($guest.LastLogonDate -and $guest.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $risks += "Recently used (within 30 days)"
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $recommendations += "Investigate recent usage"
            }
            
            # PrÃ¼fe Gruppenmitgliedschaften
            $groups = Get-ADPrincipalGroupMembership -Identity $guest -ErrorAction SilentlyContinue
            $privilegedGroups = @()
            
            foreach ($group in $groups) {
                if ($group.Name -ne "Domain Guests" -and $group.Name -ne "Guests") {
                    $privilegedGroups += $group.Name
                }
            }
            
            if ($privilegedGroups.Count -gt 0) {
                $risks += "Member of additional groups"
                $riskLevel = "Critical"
                $recommendations += "Remove from all groups except Domain Guests"
            }
            
            [PSCustomObject]@{
                AccountName = $guest.SamAccountName
                DisplayName = $guest.DisplayName
                Enabled = $guest.Enabled
                PasswordLastSet = $guest.PasswordLastSet
                LastLogonDate = $guest.LastLogonDate
                PasswordNeverExpires = $guest.PasswordNeverExpires
                PasswordNotRequired = $guest.PasswordNotRequired
                AccountLockedOut = $guest.LockedOut
                SID = $guest.SID
                Groups = if ($privilegedGroups) { $privilegedGroups -join ", " } else { "Domain Guests only" }
                RiskLevel = $riskLevel
                SecurityIssues = if ($risks) { $risks -join "; " } else { "None" }
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Account properly secured" }
                WhenCreated = $guest.WhenCreated
                WhenChanged = $guest.WhenChanged
            }
        }
        
        Write-ADReportLog -Message "Guest account analysis completed. $($Results.Count) account(s) found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing Guest account status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByDepartment {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Department-Attribut
        $Users = Get-ADUser -Filter * -Properties Department, DisplayName, SamAccountName, Title, Enabled, LastLogonDate, Manager -ErrorAction Stop
        
        # Gruppiere nach Department und sortiere Departments alphabetisch
        $DepartmentGroups = $Users | Group-Object Department | Sort-Object Name
        
        $Results = foreach ($deptGroup in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($deptGroup.Name)) { "(No Department)" } else { $deptGroup.Name }
            
            # Department-Header erstellen
            [PSCustomObject]@{
                Type = "Department"
                Department = $deptName
                DisplayName = ""
                SamAccountName = ""
                Title = ""
                Manager = ""
                Enabled = $null
                LastLogonDate = $null
                DeptUserCount = $deptGroup.Count
                DeptEnabledCount = if ($deptName -eq "(No Department)") { $null } else { ($deptGroup.Group | Where-Object { $_.Enabled }).Count }
                DeptActiveCount = if ($deptName -eq "(No Department)") { $null } else { ($deptGroup.Group | Where-Object { $_.LastLogonDate -gt (Get-Date).AddDays(-30) }).Count }
                ActivityStatus = ""
            }
            
            # Manager Dictionary erstellen
            $managersDict = @{}
            foreach ($user in $deptGroup.Group) {
                if ($user.Manager) {
                    if (-not $managersDict.ContainsKey($user.Manager)) {
                        try {
                            $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName -ErrorAction SilentlyContinue
                            if ($managerObj) {
                                $managersDict[$user.Manager] = $managerObj.DisplayName
                            }
                        } catch {
                            # Manager nicht gefunden
                        }
                    }
                }
            }
            
            # Benutzer innerhalb der Abteilung alphabetisch sortieren und ausgeben
            $sortedUsers = $deptGroup.Group | Sort-Object DisplayName
            foreach ($user in $sortedUsers) {
                $managerName = "None"
                if ($user.Manager -and $managersDict.ContainsKey($user.Manager)) {
                    $managerName = $managersDict[$user.Manager]
                }
                
                [PSCustomObject]@{
                    Type = "User"
                    Department = $deptName
                    DisplayName = "    ├─ $($user.DisplayName)" # Einrückung für Baumstruktur
                    SamAccountName = $user.SamAccountName
                    Title = $user.Title
                    Manager = $managerName
                    Enabled = $user.Enabled
                    LastLogonDate = $user.LastLogonDate
                    DeptUserCount = $null # Nur im Department-Header anzeigen
                    DeptEnabledCount = $null
                    DeptActiveCount = $null
                    ActivityStatus = if ($user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                   elseif ($user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" }
                                   else { "Very Inactive" }
                }
            }
        }
        
        Write-ADReportLog -Message "Users by department analysis completed. Found users in $($DepartmentGroups.Count) departments." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by department: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByManager {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by manager..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Manager-Attribut
        $Users = Get-ADUser -Filter * -Properties Manager, DisplayName, SamAccountName, Department, Title, Enabled, LastLogonDate -ErrorAction Stop
        
        # Erstelle Dictionary für Manager-Namen
        $managersDict = @{}
        $usersWithManager = $Users | Where-Object { $_.Manager }
        
        foreach ($user in $usersWithManager) {
            if (-not $managersDict.ContainsKey($user.Manager)) {
                try {
                    $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName, Department, Title -ErrorAction SilentlyContinue
                    if ($managerObj) {
                        $managersDict[$user.Manager] = $managerObj
                    }
                } catch {
                    # Manager nicht gefunden
                }
            }
        }
        
        # Gruppiere nach Manager
        $ManagerGroups = $usersWithManager | Group-Object Manager
        
        $Results = @()
        
        # Verarbeite Benutzer mit Manager
        foreach ($mgrGroup in $ManagerGroups) {
            $managerInfo = $managersDict[$mgrGroup.Name]
            $managerName = if ($managerInfo) { $managerInfo.DisplayName } else { "Unknown Manager" }
            $managerDept = if ($managerInfo) { $managerInfo.Department } else { "Unknown" }
            $managerTitle = if ($managerInfo) { $managerInfo.Title } else { "Unknown" }
            
            # Manager-Header
            $Results += [PSCustomObject]@{
                ManagerName = $managerName
                ManagerDepartment = $managerDept
                ManagerTitle = $managerTitle
                DirectReports = $mgrGroup.Count
                UserDisplayName = "└─ Manager: $managerName"
                UserSamAccountName = ""
                UserDepartment = $managerDept
                UserTitle = $managerTitle
                UserEnabled = $true
                UserLastLogon = $null
            }
            
            # Sortierte Mitarbeiter unter dem Manager
            $sortedUsers = $mgrGroup.Group | Sort-Object DisplayName
            foreach ($user in $sortedUsers) {
                $Results += [PSCustomObject]@{
                    ManagerName = $managerName
                    ManagerDepartment = $managerDept
                    ManagerTitle = $managerTitle
                    DirectReports = $null
                    UserDisplayName = "    ├─ $($user.DisplayName)"
                    UserSamAccountName = $user.SamAccountName
                    UserDepartment = $user.Department
                    UserTitle = $user.Title
                    UserEnabled = $user.Enabled
                    UserLastLogon = $user.LastLogonDate
                }
            }
        }
        
        # Benutzer ohne Manager als separate Gruppe
        $usersWithoutManager = $Users | Where-Object { -not $_.Manager } | Sort-Object DisplayName
        if ($usersWithoutManager) {
            # Header für Benutzer ohne Manager
            $Results += [PSCustomObject]@{
                ManagerName = "(No Manager)"
                ManagerDepartment = "N/A" 
                ManagerTitle = "N/A"
                DirectReports = $usersWithoutManager.Count
                UserDisplayName = "└─ Users without Manager"
                UserSamAccountName = ""
                UserDepartment = "N/A"
                UserTitle = "N/A"
                UserEnabled = $true
                UserLastLogon = $null
            }
            
            foreach ($user in $usersWithoutManager) {
                $Results += [PSCustomObject]@{
                    ManagerName = "(No Manager)"
                    ManagerDepartment = "N/A"
                    ManagerTitle = "N/A"
                    DirectReports = $null
                    ActiveReports = $null
                    UserDisplayName = "    ├─ $($user.DisplayName)"
                    UserSamAccountName = $user.SamAccountName
                    UserDepartment = $user.Department
                    UserTitle = $user.Title
                    UserEnabled = $user.Enabled
                    UserLastLogon = $user.LastLogonDate
                }
            }
        }
        
        Write-ADReportLog -Message "Users by manager analysis completed. $($Results.Count) users analyzed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by manager: $($_.Exception.Message)" -Type Error
        return @()
    }
}
Function Get-EmptyGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing empty groups..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties Members, Name, GroupCategory, GroupScope, whenCreated, whenChanged -ErrorAction Stop
        
        $EmptyGroups = $Groups | Where-Object { 
            # PrÃ¼fe sowohl direkte Members als auch Ã¼ber Get-ADGroupMember
            $_.Members.Count -eq 0 -and 
            @(Get-ADGroupMember -Identity $_.DistinguishedName -ErrorAction SilentlyContinue).Count -eq 0
        }
        
        $Results = foreach ($group in $EmptyGroups) {
            $ageInDays = if ($group.whenCreated) { (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days } else { 0 }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                Description = $group.Description
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                AgeInDays = $ageInDays
                MemberCount = 0
                CleanupRecommended = if ($ageInDays -gt 90 -and [string]::IsNullOrWhiteSpace($group.Description)) { "Yes" } else { "Review" }
            }
        }
        
        Write-ADReportLog -Message "Empty groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object AgeInDays -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing empty groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-CircularNestedGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing nested group memberships..." -Type Info -Terminal
        
        # Verschachtelte Gruppenmitgliedschaften analysieren
        # Nested Groups können Sicherheitsrisiken darstellen:
        # - Komplexität der Berechtigung wird schwer überschaubar
        # - Potentielle zirkuläre Abhängigkeiten (Gruppe A in B, B in A)
        # - Schwierige Nachverfolgung von Berechtigungen für Compliance
        # - Performance-Probleme bei der Gruppenerweiterung
        # - Erschwerte Fehlerbehebung bei Zugriffsproblemen
        
        $Groups = Get-ADGroup -Filter * -Properties MemberOf, Name, GroupCategory, Description -ErrorAction Stop
        $NestedGroups = $Groups | Where-Object { $_.MemberOf.Count -gt 0 }
        
        $Results = foreach ($group in $NestedGroups) {
            # Analyse der Verschachtelungstiefe
            $nestingLevel = 0
            $currentGroups = @($group.MemberOf)
            $allParentGroups = @()
            
            while ($currentGroups.Count -gt 0 -and $nestingLevel -lt 10) {
                $nestingLevel++
                $nextLevelGroups = @()
                
                foreach ($parentDN in $currentGroups) {
                    try {
                        $parentGroup = Get-ADGroup -Identity $parentDN -Properties MemberOf -ErrorAction SilentlyContinue
                        if ($parentGroup) {
                            $allParentGroups += $parentGroup.Name
                            if ($parentGroup.MemberOf.Count -gt 0) {
                                $nextLevelGroups += $parentGroup.MemberOf
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Warning: Could not analyze parent group $parentDN" -Type Warning
                    }
                }
                
                $currentGroups = $nextLevelGroups | Select-Object -Unique
            }
            
            # Risikobewertung
            $riskLevel = switch ($true) {
                ($nestingLevel -ge 5) { "High - Deep nesting detected" }
                ($nestingLevel -ge 3) { "Medium - Multiple nesting levels" }
                ($nestingLevel -eq 2) { "Low - Standard nesting" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                Description = $group.Description
                MemberOfGroups = $group.MemberOf.Count
                NestingLevel = $nestingLevel
                RiskLevel = $riskLevel
                ParentGroups = ($group.MemberOf | ForEach-Object { 
                    try { (Get-ADGroup $_ -ErrorAction SilentlyContinue).Name } catch { "Unknown" }
                }) -join "; "
                AllParentHierarchy = ($allParentGroups | Select-Object -Unique) -join " → "
                SecurityConcerns = if ($nestingLevel -ge 3) { 
                    "Complex permission inheritance, difficult troubleshooting, potential circular dependencies" 
                } else { 
                    "Standard nesting structure" 
                }
                Recommendations = switch ($true) {
                    ($nestingLevel -ge 5) { "Review and flatten group structure immediately" }
                    ($nestingLevel -ge 3) { "Consider simplifying group hierarchy" }
                    default { "Monitor for changes" }
                }
            }
        }
        
        # Statistiken für Logging
        $totalNested = $Results.Count
        $highRisk = ($Results | Where-Object { $_.RiskLevel -like "High*" }).Count
        $mediumRisk = ($Results | Where-Object { $_.RiskLevel -like "Medium*" }).Count
        $deepestNesting = if ($Results) { ($Results | Measure-Object NestingLevel -Maximum).Maximum } else { 0 }
        
        Write-ADReportLog -Message "Nested group analysis completed:" -Type Info -Terminal
        Write-ADReportLog -Message "  Total nested groups: $totalNested" -Type Info -Terminal
        Write-ADReportLog -Message "  High risk (deep nesting): $highRisk" -Type Info -Terminal
        Write-ADReportLog -Message "  Medium risk: $mediumRisk" -Type Info -Terminal
        Write-ADReportLog -Message "  Deepest nesting level found: $deepestNesting" -Type Info -Terminal
        
        # Zusätzlich Zusammenfassung für GUI anzeigen
        try {
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "Nested group analysis completed. $totalNested groups analyzed - High risk: $highRisk, Medium risk: $mediumRisk, Deepest nesting: $deepestNesting levels."
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI status: $($_.Exception.Message)" -Type Warning
        }
        
        return $Results | Sort-Object NestingLevel -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing nested groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DynamicDistributionGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing dynamic distribution groups..." -Type Info -Terminal
        
        # Suche nach dynamischen Verteilerlisten anhand typischer Attribute
        # Exchange-Attribute nur verwenden, wenn Exchange installiert ist
        # Prüfe, ob Exchange-Attribute verfügbar sind
        $exchangeInstalled = $false
        try {
            $testGroup = Get-ADGroup -Filter * -Properties msExchRecipientDisplayType -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $testGroup.PSObject.Properties['msExchRecipientDisplayType']) {
                $exchangeInstalled = $true
            }
        } catch {
            $exchangeInstalled = $false
        }
        
        if ($exchangeInstalled) {
            $DynamicGroups = Get-ADGroup -Filter {
                (GroupCategory -eq 'Distribution') -and 
                ((msExchRecipientDisplayType -eq '7') -or 
                 (msExchRecipientTypeDetails -eq '8589934592'))
            } -Properties Name, whenCreated, whenChanged, msExchRecipientDisplayType, msExchRecipientTypeDetails -ErrorAction Stop
        } else {
            # Ohne Exchange können wir keine echten Dynamic Distribution Groups identifizieren
            Write-ADReportLog -Message "Exchange attributes not available. Cannot identify dynamic distribution groups." -Type Warning
            $DynamicGroups = @()
        }
        
        $Results = foreach ($group in $DynamicGroups) {
            [PSCustomObject]@{
                Name = $group.Name
                Description = $group.Description
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                AgeInDays = if ($group.whenCreated) {
                    [math]::Round((New-TimeSpan -Start $group.whenCreated -End (Get-Date)).TotalDays, 0)
                } else { 0 }
                RecipientType = "Dynamic Distribution Group"
                Status = "Active"
            }
        }
        
        Write-ADReportLog -Message "Dynamic distribution groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object Name
    } catch {
        Write-ADReportLog -Message "Error analyzing dynamic distribution groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NestedGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing nested group structure and hierarchy..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties MemberOf, Members, Name, GroupCategory, Description, whenCreated -ErrorAction Stop
        
        # Build group hierarchy tree structure
        $GroupHierarchy = @{}
        $NestedGroupResults = [System.Collections.ArrayList]@()
        
        # Initialize hierarchy tracking
        foreach ($group in $Groups) {
            $GroupHierarchy[$group.Name] = @{
                Group = $group
                Parents = @()
                Children = @()
                Level = 0
                Path = ""
                Interconnections = 0
            }
        }
        
        # Map parent-child relationships
        foreach ($group in $Groups) {
            if ($group.MemberOf) {
                foreach ($parentDN in $group.MemberOf) {
                    try {
                        $parentGroup = Get-ADGroup $parentDN -ErrorAction SilentlyContinue
                        if ($parentGroup) {
                            $GroupHierarchy[$group.Name].Parents += $parentGroup.Name
                            if ($GroupHierarchy.ContainsKey($parentGroup.Name)) {
                                $GroupHierarchy[$parentGroup.Name].Children += $group.Name
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Warning: Could not resolve parent group for $($group.Name)" -Type Warning
                    }
                }
            }
        }
        
        # Calculate nesting levels and tree structure
        function Get-GroupLevel($groupName, $visited = @()) {
            if ($visited -contains $groupName) {
                return 999  # Circular reference detected
            }
            
            $visited += $groupName
            $parents = $GroupHierarchy[$groupName].Parents
            
            if ($parents.Count -eq 0) {
                return 0  # Root level
            }
            
            $maxParentLevel = 0
            foreach ($parent in $parents) {
                if ($GroupHierarchy.ContainsKey($parent)) {
                    $parentLevel = Get-GroupLevel $parent $visited
                    if ($parentLevel -gt $maxParentLevel) {
                        $maxParentLevel = $parentLevel
                    }
                }
            }
            
            return $maxParentLevel + 1
        }
        
        # Generate tree path visualization
        function Get-TreePath($groupName, $level = 0, $prefix = "", $isLast = $true) {
            $connector = ""
            if ($level -eq 0) { 
                $connector = "├─ " 
            } elseif ($isLast) { 
                $connector = "└─ " 
            } else { 
                $connector = "├─ " 
            }
            $treePath = $prefix + $connector + $groupName
            
            $children = $GroupHierarchy[$groupName].Children
            if ($children.Count -gt 0) {
                $childPrefix = ""
                if ($isLast) {
                    $childPrefix = $prefix + "   "
                } else {
                    $childPrefix = $prefix + "│  "
                }
                
                for ($i = 0; $i -lt $children.Count; $i++) {
                    $childIsLast = ($i -eq $children.Count - 1)
                    $childPath = Get-TreePath $children[$i] ($level + 1) $childPrefix $childIsLast
                    $treePath += "`n" + $childPath
                }
            }
            
            return $treePath
        }
        
        # Analyze each group's nesting characteristics
        foreach ($group in $Groups) {
            if ($group.MemberOf.Count -gt 0 -or $GroupHierarchy[$group.Name].Children.Count -gt 0) {
                
                # Calculate nesting metrics
                $nestingLevel = Get-GroupLevel $group.Name
                $totalConnections = $group.MemberOf.Count + $GroupHierarchy[$group.Name].Children.Count
                $GroupHierarchy[$group.Name].Level = $nestingLevel
                $GroupHierarchy[$group.Name].Interconnections = $totalConnections
                
                # Risk assessment based on complexity
                $riskLevel = ""
                if ($nestingLevel -eq 999) { 
                    $riskLevel = "Critical - Circular Reference" 
                } elseif ($nestingLevel -gt 5 -or $totalConnections -gt 15) { 
                    $riskLevel = "High" 
                } elseif ($nestingLevel -gt 3 -or $totalConnections -gt 8) { 
                    $riskLevel = "Medium" 
                } elseif ($nestingLevel -gt 1 -or $totalConnections -gt 3) { 
                    $riskLevel = "Low" 
                } else { 
                    $riskLevel = "Minimal" 
                }
                
                # Generate tree visualization for this group's hierarchy
                $treeStructure = ""
                if ($GroupHierarchy[$group.Name].Parents.Count -eq 0) {
                    $treeStructure = Get-TreePath $group.Name
                } else {
                    $treeStructure = "Child of: " + ($GroupHierarchy[$group.Name].Parents -join ", ")
                }
                
                # Security and management recommendations
                $recommendation = ""
                if ($nestingLevel -eq 999) { 
                    $recommendation = "IMMEDIATE ACTION: Resolve circular group membership to prevent authentication issues" 
                } elseif ($nestingLevel -gt 5) { 
                    $recommendation = "Review deep nesting structure - exceeds recommended depth of 5 levels" 
                } elseif ($totalConnections -gt 15) { 
                    $recommendation = "Simplify group interconnections - high complexity affects performance and auditing" 
                } elseif ($nestingLevel -gt 3) { 
                    $recommendation = "Consider flattening group structure for better management" 
                } else { 
                    $recommendation = "Monitor for permission inheritance complexity" 
                }
                
                # Determine management complexity
                $managementComplexity = ""
                if ($totalConnections -gt 10) { 
                    $managementComplexity = "High - Difficult to manage" 
                } elseif ($totalConnections -gt 5) { 
                    $managementComplexity = "Medium - Requires attention" 
                } else { 
                    $managementComplexity = "Low - Manageable" 
                }
                
                # Calculate age in days
                $ageInDays = 0
                if ($group.whenCreated) { 
                    $ageInDays = [int][math]::Round((New-TimeSpan -Start $group.whenCreated -End (Get-Date)).TotalDays)
                }
                
                # Format when created
                $whenCreatedFormatted = ""
                if ($group.whenCreated) { 
                    $whenCreatedFormatted = $group.whenCreated.ToString("dd.MM.yyyy HH:mm") 
                } else { 
                    $whenCreatedFormatted = "Unknown" 
                }
                
                # Build comprehensive result object
                $groupResult = [PSCustomObject]@{
                    TreeStructure = [string]$treeStructure
                    GroupName = [string]$group.Name
                    GroupCategory = [string]$group.GroupCategory
                    NestingLevel = [int]$nestingLevel
                    ParentGroups = [int]$group.MemberOf.Count
                    ChildGroups = [int]$GroupHierarchy[$group.Name].Children.Count
                    TotalInterconnections = [int]$totalConnections
                    ParentGroupNames = [string](($GroupHierarchy[$group.Name].Parents | Select-Object -First 5) -join "; ")
                    ChildGroupNames = [string](($GroupHierarchy[$group.Name].Children | Select-Object -First 5) -join "; ")
                    RiskLevel = [string]$riskLevel
                    SecurityImpact = [string]"Complex nesting creates inheritance chains affecting access control and audit trails"
                    ManagementComplexity = [string]$managementComplexity
                    Recommendation = [string]$recommendation
                    WhenCreated = $whenCreatedFormatted
                    AgeInDays = $ageInDays
                }
                
                $null = $NestedGroupResults.Add($groupResult)
            }
        }
        
        # Sort results by complexity (nesting level, then total connections)
        $SortedResults = $NestedGroupResults | Sort-Object @{
            Expression = { 
                if ($_.NestingLevel -eq 999) { 
                    1000 
                } else { 
                    $_.NestingLevel 
                } 
            }
        }, TotalInterconnections -Descending
        
        # Generate summary statistics
        $totalNested = $SortedResults.Count
        $criticalCount = ($SortedResults | Where-Object { $_.RiskLevel -like "Critical*" }).Count
        $highRiskCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "High" }).Count
        $maxNestingLevel = 0
        $maxNestingLevelResult = $SortedResults | Where-Object { $_.NestingLevel -ne 999 } | Measure-Object NestingLevel -Maximum
        if ($maxNestingLevelResult.Maximum) {
            $maxNestingLevel = $maxNestingLevelResult.Maximum
        }
        $circularRefs = ($SortedResults | Where-Object { $_.NestingLevel -eq 999 }).Count
        
        Write-ADReportLog -Message "Nested group analysis completed. Found $totalNested groups with nesting relationships." -Type Info -Terminal
        Write-ADReportLog -Message "Complexity Analysis - Maximum nesting depth: $maxNestingLevel levels, Critical issues: $criticalCount, High risk: $highRiskCount" -Type Info -Terminal
        if ($circularRefs -gt 0) {
            Write-ADReportLog -Message "WARNING: $circularRefs circular group references detected - immediate attention required!" -Type Warning -Terminal
        }
        Write-ADReportLog -Message "Tree structure visualization included for hierarchy mapping and security audit trails." -Type Info -Terminal
        
        return $SortedResults
    } catch {
        Write-ADReportLog -Message "Error analyzing nested group structure: $($_.Exception.Message)" -Type Error
        return @()
    }
}
# --- Neue Roadmap Features - Computer-Reports ---
Function Get-OSSummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating Operating System summary..." -Type Info -Terminal
        $Computers = Get-ADComputer -Filter "Enabled -eq `$true" -Properties OperatingSystem, OperatingSystemVersion -ErrorAction Stop
        
        $OSGroups = $Computers | Group-Object OperatingSystem | Sort-Object Count -Descending
        $Results = $OSGroups | Select-Object @{Name="OperatingSystem";Expression={$_.Name}},
            @{Name="Count";Expression={$_.Count}},
            @{Name="Percentage";Expression={[math]::Round(($_.Count / $Computers.Count) * 100, 2)}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error generating OS summary: $($_.Exception.Message)" -Type Error
        return @()
    }
}


# --- Service Account Reports ---
Function Get-ServiceAccountsOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive Service Accounts overview..." -Type Info -Terminal
        
        # Identify service accounts through various criteria
        $PotentialServiceAccounts = @()
        
        # 1. Name-based identification (common service account naming patterns)
        $NamePatterns = @(
            '*svc*', '*service*', '*sql*', '*iis*', '*app*', '*web*', '*batch*', '*task*', 
            '*auto*', '*sched*', '*daemon*', '*system*', '*admin*', '*db*', '*database*',
            'Administrator', 'krbtgt', '*admin*' # Hinzugefügt: Standard Admin und Kerberos Ticket Granting Account
        )
        
        # Get all users and filter afterwards to avoid multiple AD queries
        $AllUsers = Get-ADUser -Filter * -Properties Description, LastLogonDate, 
            PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, 
            Department, AccountExpirationDate, LockedOut, BadLogonCount -ErrorAction Stop
            
        # Spezielle AD-Konten direkt hinzufügen
        $SpecialAccounts = $AllUsers | Where-Object { 
            $_.SamAccountName -eq "Administrator" -or 
            $_.SamAccountName -eq "krbtgt" -or
            $_.SamAccountName -like "*admin*"
        }
        foreach($specialAccount in $SpecialAccounts) {
            if($specialAccount -notin $PotentialServiceAccounts) {
                $PotentialServiceAccounts += $specialAccount
                Write-ADReportLog -Message "Found special account: $($specialAccount.Name)" -Type Info
            }
        }
            
        # Name pattern matching
        foreach($pattern in $NamePatterns) {
            $matchedUsers = $AllUsers | Where-Object { $_.Name -like $pattern }
            foreach($user in $matchedUsers) {
                if($user -notin $PotentialServiceAccounts) {
                    $PotentialServiceAccounts += $user
                    Write-ADReportLog -Message "Account '$($user.Name)' matched name pattern '$pattern'" -Type Info
                }
            }
        }
        
        # SPN check
        $SPNAccounts = $AllUsers | Where-Object { $_.ServicePrincipalName.Count -gt 0 }
        foreach($user in $SPNAccounts) {
            if($user -notin $PotentialServiceAccounts) {
                $PotentialServiceAccounts += $user
                Write-ADReportLog -Message "Account '$($user.Name)' has $($user.ServicePrincipalName.Count) SPNs" -Type Info
            }
        }
        
        # Description pattern matching
        $DescriptionPatterns = @(
            '*service*', '*application*', '*sql*', '*database*', '*automated*', '*system*',
            '*batch*', '*scheduled*', '*task*', '*daemon*', '*process*', '*admin*'
        )
        
        foreach($pattern in $DescriptionPatterns) {
            $matchedUsers = $AllUsers | Where-Object { $_.Description -like $pattern }
            foreach($user in $matchedUsers) {
                if($user -notin $PotentialServiceAccounts) {
                    $PotentialServiceAccounts += $user
                    Write-ADReportLog -Message "Account '$($user.Name)' matched description pattern '$pattern'" -Type Info
                }
            }
        }

        # Return template for no results
        if ($PotentialServiceAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No service accounts found in the domain" -Type Info -Terminal
            return @(
                [PSCustomObject]@{
                    Name = "No Service Accounts"
                    SamAccountName = "N/A"
                    Description = "No service accounts were found in the Active Directory domain"
                    Department = "N/A"
                    Enabled = $false
                    LastLogonDate = "N/A"
                    DaysSinceLastLogon = 0
                    PasswordLastSet = "N/A"
                    DaysSincePasswordChange = 0
                    PasswordNeverExpires = $false
                    HasSPN = $false
                    SPNCount = 0
                    RiskLevel = "N/A"
                    RiskFactors = "No service accounts to analyze"
                    WhenCreated = "N/A"
                    DetectionReason = "No service accounts detected"
                    AccountExpiration = "N/A"
                    IsLocked = $false
                    BadLogonCount = 0
                }
            )
        }
        
        # Process each account
        $Results = foreach ($account in $PotentialServiceAccounts) {
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
            } else { 9999 }
            
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            # Enhanced risk assessment
            $riskScore = 0
            $riskFactors = [System.Collections.ArrayList]@()
            
            # Special account risks
            if ($account.SamAccountName -eq "Administrator" -or $account.SamAccountName -eq "krbtgt") {
                $riskScore += 4
                [void]$riskFactors.Add("Built-in privileged account")
            }
            
            # Password-related risks
            if ($daysSincePasswordChange -gt 365) { 
                $riskScore += 3
                [void]$riskFactors.Add("Password not changed for $daysSincePasswordChange days")
            }
            if ($account.PasswordNeverExpires) { 
                $riskScore += 2
                [void]$riskFactors.Add("Password never expires")
            }
            
            # Activity-related risks
            if ($daysSinceLastLogon -gt 90) { 
                $riskScore += 2
                [void]$riskFactors.Add("No logon for $daysSinceLastLogon days")
            }
            if ($account.BadLogonCount -gt 5) {
                $riskScore += 2
                [void]$riskFactors.Add("High failed login count: $($account.BadLogonCount)")
            }
            if ($account.LockedOut) {
                $riskScore += 2
                [void]$riskFactors.Add("Account is locked out")
            }
            
            # SPN-related risks
            if ($account.ServicePrincipalName.Count -gt 0) { 
                $riskScore += 1
                [void]$riskFactors.Add("Has $($account.ServicePrincipalName.Count) SPNs")
            }
            
            # Account state risks
            if (-not $account.Enabled) { 
                $riskScore -= 2
                [void]$riskFactors.Add("Account disabled")
            }
            
            # Risk level classification
            $riskLevel = switch ($riskScore) {
                { $_ -ge 6 } { "Critical"; break }
                { $_ -ge 4 } { "High"; break }
                { $_ -ge 2 } { "Medium"; break }
                default { "Low" }
            }

            # Detection reasoning with detailed patterns
            $detectionReason = [System.Collections.ArrayList]@()
            
            # Special account detection
            if ($account.SamAccountName -eq "Administrator" -or $account.SamAccountName -eq "krbtgt") {
                [void]$detectionReason.Add("Built-in privileged account")
            }
            
            # Name pattern matches
            foreach($pattern in $NamePatterns) {
                if($account.Name -like $pattern) {
                    [void]$detectionReason.Add("Name matches pattern '$pattern'")
                }
            }
            
            # SPN check
            if ($account.ServicePrincipalName.Count -gt 0) {
                [void]$detectionReason.Add("Has $($account.ServicePrincipalName.Count) Service Principal Names")
            }
            
            # Description pattern matches
            foreach($pattern in $DescriptionPatterns) {
                if($account.Description -like $pattern) {
                    [void]$detectionReason.Add("Description matches pattern '$pattern'")
                }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordChange
                PasswordNeverExpires = $account.PasswordNeverExpires
                HasSPN = $account.ServicePrincipalName.Count -gt 0
                SPNCount = $account.ServicePrincipalName.Count
                RiskLevel = $riskLevel
                RiskFactors = ($riskFactors -join " | ")
                WhenCreated = $account.whenCreated
                DetectionReason = ($detectionReason -join " | ")
                AccountExpiration = if ($account.AccountExpirationDate) { $account.AccountExpirationDate } else { "Never" }
                IsLocked = $account.LockedOut
                BadLogonCount = $account.BadLogonCount
            }
        }
        
        Write-ADReportLog -Message "Service Accounts overview completed. Found $($Results.Count) potential service accounts." -Type Info -Terminal
        return $Results | Sort-Object { switch ($_.RiskLevel) {
            'Critical' { 4 }
            'High' { 3 }
            'Medium' { 2 }
            'Low' { 1 }
            default { 0 }
        }} -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing service accounts: $($_.Exception.Message)" -Type Error
        return @(
            [PSCustomObject]@{
                Name = "Error"
                SamAccountName = "N/A"
                Description = "Error: $($_.Exception.Message)"
                Department = "N/A"
                Enabled = $false
                LastLogonDate = "N/A"
                DaysSinceLastLogon = 0
                PasswordLastSet = "N/A"
                DaysSincePasswordChange = 0
                PasswordNeverExpires = $false
                HasSPN = $false
                SPNCount = 0
                RiskLevel = "Error"
                RiskFactors = "Analysis failed"
                WhenCreated = "N/A"
                DetectionReason = "Error during detection"
                AccountExpiration = "N/A"
                IsLocked = $false
                BadLogonCount = 0
            }
        )
    }
}

Function Get-ManagedServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Managed Service Accounts..." -Type Info -Terminal
        
        # gMSA und sMSA haben spezielle ObjectClasses
        $MSAs = Get-ADObject -Filter "ObjectClass -eq 'msDS-ManagedServiceAccount' -or ObjectClass -eq 'msDS-GroupManagedServiceAccount'" -Properties Name, ObjectClass, Created, Modified, DistinguishedName -ErrorAction Stop
        
        $Results = foreach ($msa in $MSAs) {
            [PSCustomObject]@{
                Name = $msa.Name
                AccountType = if($msa.ObjectClass -eq "msDS-GroupManagedServiceAccount") {"Group MSA"} else {"Standalone MSA"}
                Created = $msa.Created
                Modified = $msa.Modified
                DistinguishedName = $msa.DistinguishedName
                Status = "Active" # MSAs sind immer aktiv
            }
        }
        
        Write-ADReportLog -Message "Found $($Results.Count) managed service accounts" -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing managed service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-CompromiseIndicators {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing compromise indicators..." -Type Info -Terminal
        
        $Results = @()
        $ProcessedAccounts = @() # Track processed accounts to avoid duplicates
        
        # Header entry for tree structure
        $Results += [PSCustomObject]@{
            TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS"
            AccountName = "Security Assessment Overview"
            DisplayName = "Active Directory Compromise Detection"
            IndicatorType = "Analysis Report"
            Details = "Comprehensive security indicator analysis initiated"
            Severity = "Info"
            RiskLevel = "Info"
            SecurityImpact = "Scanning for potential security compromises"
            Recommendation = "Review all findings and investigate high-risk indicators immediately"
            ComplianceNote = "Security monitoring and incident response procedures"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        # 1. After-hours login analysis
        try {
            Write-ADReportLog -Message "Analyzing unusual login times..." -Type Info
            $AfterHoursLogons = Get-ADUser -Filter * -Properties LastLogonDate, BadLogonCount, LockedOut, PasswordLastSet, whenCreated, DisplayName, Department -ErrorAction SilentlyContinue | 
                Where-Object { 
                    $null -ne $_.LastLogonDate -and
                    $_.LastLogonDate -is [datetime] -and
                    ($_.LastLogonDate.Hour -lt 6 -or $_.LastLogonDate.Hour -gt 20) -and
                    $_.LastLogonDate -gt (Get-Date).AddDays(-30) -and  # Only recent after-hours logins
                    $_.SamAccountName -notin $ProcessedAccounts
                }
            
            if ($AfterHoursLogons) {
                foreach ($user in $AfterHoursLogons) {
                    $ProcessedAccounts += $user.SamAccountName
                    $lastLogonStr = if ($null -ne $user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
                    $riskLevel = if ($user.LastLogonDate.Hour -lt 3 -or $user.LastLogonDate.Hour -gt 23) { "High" } else { "Medium" }
                    
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🕒 Unusual Login Times ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "After-Hours Login Activity"
                        Details = "Login at: $lastLogonStr (outside business hours 06:00-20:00)"
                        Severity = $riskLevel
                        RiskLevel = $riskLevel
                        SecurityImpact = "Potential unauthorized access or compromised credentials"
                        Recommendation = "Verify login legitimacy with user and review access logs for additional suspicious activity"
                        ComplianceNote = "Monitor for patterns of after-hours access"
                        Timestamp = $lastLogonStr
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🕒 Unusual Login Times"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No After-Hours Logins"
                    IndicatorType = "Login Time Analysis"
                    Details = "No after-hours login activity detected in the last 30 days"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "No security concerns identified"
                    Recommendation = "Continue monitoring login patterns"
                    ComplianceNote = "Good - No unusual login times detected"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing after-hours logins: $($_.Exception.Message)" -Type Warning
        }
        
        # 2. Failed login attempts analysis
        try {
            Write-ADReportLog -Message "Analyzing failed login attempts..." -Type Info
            $HighBadLogonCount = Get-ADUser -Filter "BadLogonCount -gt 5" -Properties BadLogonCount, LastLogonDate, LockedOut, DisplayName, Department, whenCreated -ErrorAction SilentlyContinue |
                Where-Object { $_.SamAccountName -notin $ProcessedAccounts }
            
            if ($HighBadLogonCount) {
                foreach ($user in $HighBadLogonCount) {
                    $ProcessedAccounts += $user.SamAccountName
                    $riskLevel = switch ($user.BadLogonCount) {
                        {$_ -gt 20} { "Critical" }
                        {$_ -gt 10} { "High" }
                        default { "Medium" }
                    }
                    
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🚫 Failed Login Attempts ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "Excessive Failed Login Attempts"
                        Details = "Failed login attempts: $($user.BadLogonCount) | Account locked: $(if ($user.LockedOut) {'Yes'} else {'No'})"
                        Severity = $riskLevel
                        RiskLevel = $riskLevel
                        SecurityImpact = "Possible brute force attack or password guessing attempt"
                        Recommendation = "Investigate source of failed attempts, consider account lockout policies, and verify user identity"
                        ComplianceNote = "Monitor for coordinated attacks across multiple accounts"
                        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🚫 Failed Login Attempts"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Excessive Failed Logins"
                    IndicatorType = "Failed Login Analysis"
                    Details = "No accounts with excessive failed login attempts (>5) detected"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "No brute force indicators identified"
                    Recommendation = "Continue monitoring failed login patterns"
                    ComplianceNote = "Good - No excessive failed login attempts"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing failed login attempts: $($_.Exception.Message)" -Type Warning
        }
        
        # 3. Suspicious password changes
        try {
            Write-ADReportLog -Message "Analyzing suspicious password changes..." -Type Info
            $SuspiciousPasswordChanges = Get-ADUser -Filter * -Properties PasswordLastSet, LastLogonDate, whenCreated, DisplayName, Department -ErrorAction SilentlyContinue |
                Where-Object {
                    $null -ne $_.PasswordLastSet -and 
                    $null -ne $_.LastLogonDate -and
                    $_.PasswordLastSet -gt (Get-Date).AddDays(-7) -and
                    $_.LastLogonDate -lt (Get-Date).AddDays(-90) -and
                    $_.SamAccountName -notin $ProcessedAccounts
                }
            
            if ($SuspiciousPasswordChanges) {
                foreach ($user in $SuspiciousPasswordChanges) {
                    $ProcessedAccounts += $user.SamAccountName
                    $daysSinceLastLogon = [math]::Round(((Get-Date) - $user.LastLogonDate).TotalDays)
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🔑 Suspicious Password Changes ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "Password Change After Long Inactivity"
                        Details = "Password changed recently but last login was $daysSinceLastLogon days ago"
                        Severity = "High"
                        RiskLevel = "High"
                        SecurityImpact = "Possible account takeover or unauthorized password reset"
                        Recommendation = "Verify password change legitimacy with user and check for unauthorized access"
                        ComplianceNote = "Investigate password change authorization and method"
                        Timestamp = $user.PasswordLastSet.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🔑 Suspicious Password Changes"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Suspicious Password Changes"
                    IndicatorType = "Password Change Analysis"
                    Details = "No suspicious password change patterns detected"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "No unauthorized password changes identified"
                    Recommendation = "Continue monitoring password change patterns"
                    ComplianceNote = "Good - No suspicious password activity"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing password changes: $($_.Exception.Message)" -Type Warning
        }
        
        # 4. Service account interactive logins
        try {
            Write-ADReportLog -Message "Analyzing service account activities..." -Type Info
            $SuspiciousServiceAccounts = Get-ADUser -Filter "Name -like '*svc*' -or Name -like '*service*' -or Name -like '*srv*' -or Name -like '*dienst*'" -Properties LastLogonDate, PasswordLastSet, DisplayName, ServicePrincipalNames, UserAccountControl -ErrorAction SilentlyContinue |
                Where-Object { 
                    $null -ne $_.LastLogonDate -and 
                    $_.LastLogonDate -gt (Get-Date).AddDays(-7) -and
                    -not ($_.UserAccountControl -band 0x0200) -and  # Not a computer account
                    $_.SamAccountName -notin $ProcessedAccounts
                }
            
            if ($SuspiciousServiceAccounts) {
                foreach ($user in $SuspiciousServiceAccounts) {
                    $ProcessedAccounts += $user.SamAccountName
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🛠️ Service Account Activity ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "Service Account Interactive Login"
                        Details = "Service account logged in interactively on $($user.LastLogonDate.ToString('yyyy-MM-dd HH:mm:ss'))"
                        Severity = "High"
                        RiskLevel = "High"
                        SecurityImpact = "Service accounts should not log in interactively - possible compromise"
                        Recommendation = "Verify legitimacy of interactive login and review service account security"
                        ComplianceNote = "Service accounts should only be used for automated processes"
                        Timestamp = $user.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🛠️ Service Account Activity"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Service Account Interactive Logins"
                    IndicatorType = "Service Account Analysis"
                    Details = "No interactive logins detected from service accounts"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "Service accounts are being used appropriately"
                    Recommendation = "Continue monitoring service account usage"
                    ComplianceNote = "Good - Service accounts not used interactively"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing service accounts: $($_.Exception.Message)" -Type Warning
        }
        
        # 5. Recently created privileged accounts
        try {
            Write-ADReportLog -Message "Analyzing recently created privileged accounts..." -Type Info
            $RecentPrivilegedAccounts = @()
            
            # Check for recently created accounts in privileged groups (German and English names)
            $PrivilegedGroupNames = @(
                "Domain Admins", "Domänen-Admins",
                "Enterprise Admins", "Organisations-Admins", 
                "Schema Admins", "Schema-Admins",
                "Administrators", "Administratoren",
                "Account Operators", "Konten-Operatoren",
                "Backup Operators", "Sicherungs-Operatoren",
                "Server Operators", "Server-Operatoren",
                "Print Operators", "Druck-Operatoren"
            )
            
            foreach ($groupName in $PrivilegedGroupNames) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $members = Get-ADGroupMember -Identity $group -ErrorAction SilentlyContinue | 
                            Where-Object { $_.objectClass -eq "user" }
                        
                        foreach ($member in $members) {
                            try {
                                if ($member.SamAccountName -notin $ProcessedAccounts) {
                                    $user = Get-ADUser -Identity $member.SamAccountName -Properties whenCreated, DisplayName -ErrorAction SilentlyContinue
                                    if ($user -and $user.whenCreated -gt (Get-Date).AddDays(-30)) {
                                        $ProcessedAccounts += $user.SamAccountName
                                        $RecentPrivilegedAccounts += [PSCustomObject]@{
                                            User = $user
                                            Group = $group.Name
                                            Created = $user.whenCreated
                                        }
                                    }
                                }
                            } catch {
                                # Continue with next member
                            }
                        }
                    }
                } catch {
                    # Continue with next group
                }
            }
            
            if ($RecentPrivilegedAccounts) {
                foreach ($account in $RecentPrivilegedAccounts) {
                    $daysOld = [math]::Round(((Get-Date) - $account.Created).TotalDays)
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 👑 Recent Privileged Accounts ► $($account.User.SamAccountName)"
                        AccountName = $account.User.SamAccountName
                        DisplayName = $(if ($account.User.DisplayName) { $account.User.DisplayName } else { $account.User.SamAccountName })
                        IndicatorType = "Recently Created Privileged Account"
                        Details = "Created $daysOld days ago and added to '$($account.Group)' group"
                        Severity = "Medium"
                        RiskLevel = "Medium"
                        SecurityImpact = "New privileged accounts require verification and monitoring"
                        Recommendation = "Verify business justification for privileged access and review approval process"
                        ComplianceNote = "All privileged account creation should be documented and approved"
                        Timestamp = $account.Created.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 👑 Recent Privileged Accounts"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Recent Privileged Accounts"
                    IndicatorType = "Privileged Account Analysis"
                    Details = "No recently created privileged accounts detected in the last 30 days"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "No unexpected privileged account creation"
                    Recommendation = "Continue monitoring privileged account creation"
                    ComplianceNote = "Good - No unauthorized privileged account creation"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing recent privileged accounts: $($_.Exception.Message)" -Type Warning
        }
        
        # 6. Dormant accounts with privileges (Trimarc recommendation)
        try {
            Write-ADReportLog -Message "Analyzing dormant privileged accounts..." -Type Info
            $DormantPrivilegedAccounts = @()
            
            foreach ($groupName in $PrivilegedGroupNames) {
                try {
                    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($group) {
                        $members = Get-ADGroupMember -Identity $group -ErrorAction SilentlyContinue | 
                            Where-Object { $_.objectClass -eq "user" }
                        
                        foreach ($member in $members) {
                            try {
                                if ($member.SamAccountName -notin $ProcessedAccounts) {
                                    $user = Get-ADUser -Identity $member.SamAccountName -Properties LastLogonDate, DisplayName, PasswordLastSet -ErrorAction SilentlyContinue
                                    if ($user -and (($null -eq $user.LastLogonDate) -or ($user.LastLogonDate -lt (Get-Date).AddDays(-90)))) {
                                        $ProcessedAccounts += $user.SamAccountName
                                        $DormantPrivilegedAccounts += [PSCustomObject]@{
                                            User = $user
                                            Group = $group.Name
                                            LastLogon = $user.LastLogonDate
                                        }
                                    }
                                }
                            } catch {
                                # Continue with next member
                            }
                        }
                    }
                } catch {
                    # Continue with next group
                }
            }
            
            if ($DormantPrivilegedAccounts) {
                foreach ($account in $DormantPrivilegedAccounts) {
                    $daysSinceLogon = if ($account.LastLogon) { [math]::Round(((Get-Date) - $account.LastLogon).TotalDays) } else { "Never" }
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 💤 Dormant Privileged Accounts ► $($account.User.SamAccountName)"
                        AccountName = $account.User.SamAccountName
                        DisplayName = $(if ($account.User.DisplayName) { $account.User.DisplayName } else { $account.User.SamAccountName })
                        IndicatorType = "Dormant Privileged Account"
                        Details = "Privileged account in '$($account.Group)' - Last logon: $daysSinceLogon days ago"
                        Severity = "High"
                        RiskLevel = "High"
                        SecurityImpact = "Dormant privileged accounts pose security risks if compromised"
                        Recommendation = "Disable or remove unused privileged accounts and review access regularly"
                        ComplianceNote = "Privileged accounts should be actively monitored and unused accounts removed"
                        Timestamp = $(if ($account.LastLogon) { $account.LastLogon.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" })
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 💤 Dormant Privileged Accounts"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Dormant Privileged Accounts"
                    IndicatorType = "Dormant Account Analysis"
                    Details = "All privileged accounts show recent activity"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "Privileged accounts are being actively used"
                    Recommendation = "Continue monitoring privileged account usage"
                    ComplianceNote = "Good - No dormant privileged accounts found"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing dormant privileged accounts: $($_.Exception.Message)" -Type Warning
        }
        
        # 7. Users with old passwords (Trimarc recommendation)
        try {
            Write-ADReportLog -Message "Analyzing accounts with old passwords..." -Type Info
            $OldPasswordAccounts = Get-ADUser -Filter * -Properties PasswordLastSet, DisplayName, LastLogonDate -ErrorAction SilentlyContinue |
                Where-Object {
                    $null -ne $_.PasswordLastSet -and
                    $_.PasswordLastSet -lt (Get-Date).AddDays(-365) -and
                    $_.Enabled -eq $true -and
                    $_.SamAccountName -notin $ProcessedAccounts
                }
            
            if ($OldPasswordAccounts) {
                foreach ($user in $OldPasswordAccounts) {
                    $ProcessedAccounts += $user.SamAccountName
                    $daysOld = [math]::Round(((Get-Date) - $user.PasswordLastSet).TotalDays)
                    $riskLevel = if ($daysOld -gt 730) { "Critical" } elseif ($daysOld -gt 365) { "High" } else { "Medium" }
                    
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🗓️ Old Passwords ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "Extremely Old Password"
                        Details = "Password is $daysOld days old (last changed: $($user.PasswordLastSet.ToString('yyyy-MM-dd')))"
                        Severity = $riskLevel
                        RiskLevel = $riskLevel
                        SecurityImpact = "Old passwords increase risk of compromise through various attack vectors"
                        Recommendation = "Force password change and implement stronger password policies"
                        ComplianceNote = "Passwords should be changed regularly according to security policies"
                        Timestamp = $user.PasswordLastSet.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🗓️ Old Passwords"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Extremely Old Passwords"
                    IndicatorType = "Password Age Analysis"
                    Details = "No passwords older than 1 year detected"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "Password policies appear to be working effectively"
                    Recommendation = "Continue monitoring password ages and enforcing policies"
                    ComplianceNote = "Good - No extremely old passwords found"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing old passwords: $($_.Exception.Message)" -Type Warning
        }
        
        # 8. Accounts with reversible encryption (Trimarc recommendation)
        try {
            Write-ADReportLog -Message "Analyzing accounts with reversible encryption..." -Type Info
            $ReversibleEncryptionAccounts = Get-ADUser -Filter {UserAccountControl -band 0x0080} -Properties DisplayName, PasswordLastSet -ErrorAction SilentlyContinue |
                Where-Object { $_.SamAccountName -notin $ProcessedAccounts }
            
            if ($ReversibleEncryptionAccounts) {
                foreach ($user in $ReversibleEncryptionAccounts) {
                    $ProcessedAccounts += $user.SamAccountName
                    $Results += [PSCustomObject]@{
                        TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🔓 Reversible Encryption ► $($user.SamAccountName)"
                        AccountName = $user.SamAccountName
                        DisplayName = $(if ($user.DisplayName) { $user.DisplayName } else { $user.SamAccountName })
                        IndicatorType = "Reversible Password Encryption"
                        Details = "Account has reversible password encryption enabled - passwords stored in cleartext equivalent"
                        Severity = "Critical"
                        RiskLevel = "Critical"
                        SecurityImpact = "Password can be easily retrieved in cleartext format"
                        Recommendation = "Disable reversible encryption immediately and force password change"
                        ComplianceNote = "Reversible encryption violates security best practices"
                        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    }
                }
            } else {
                $Results += [PSCustomObject]@{
                    TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 🔓 Reversible Encryption"
                    AccountName = "No Issues Found"
                    DisplayName = "Clean - No Reversible Encryption"
                    IndicatorType = "Encryption Analysis"
                    Details = "No accounts with reversible password encryption detected"
                    Severity = "Low"
                    RiskLevel = "Low"
                    SecurityImpact = "Password storage security is appropriate"
                    Recommendation = "Continue monitoring for reversible encryption settings"
                    ComplianceNote = "Good - No reversible encryption found"
                    Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
        } catch {
            Write-ADReportLog -Message "Error analyzing reversible encryption: $($_.Exception.Message)" -Type Warning
        }
        
        # Summary statistics
        $criticalCount = ($Results | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highCount = ($Results | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumCount = ($Results | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        $lowCount = ($Results | Where-Object { $_.RiskLevel -eq "Low" }).Count
        
        Write-ADReportLog -Message "Compromise indicators analysis completed. Found $criticalCount critical, $highCount high, $mediumCount medium, and $lowCount low risk indicators." -Type Info -Terminal
        
        # Add summary entry
        $Results += [PSCustomObject]@{
            TreeStructure = "🔍 COMPROMISE INDICATORS ANALYSIS ► 📊 Summary"
            AccountName = "Analysis Summary"
            DisplayName = "Security Assessment Results"
            IndicatorType = "Summary Report"
            Details = "Critical: $criticalCount | High: $highCount | Medium: $mediumCount | Low: $lowCount"
            Severity = $(if ($criticalCount -gt 0) { "Critical" } elseif ($highCount -gt 0) { "High" } elseif ($mediumCount -gt 0) { "Medium" } else { "Low" })
            RiskLevel = $(if ($criticalCount -gt 0) { "Critical" } elseif ($highCount -gt 0) { "High" } elseif ($mediumCount -gt 0) { "Medium" } else { "Low" })
            SecurityImpact = "Overall security posture assessment based on compromise indicators"
            Recommendation = $(if ($criticalCount -gt 0 -or $highCount -gt 0) { "Immediate attention required for high-risk findings" } else { "Continue monitoring and maintaining current security practices" })
            ComplianceNote = "Regular security assessments help maintain AD security posture"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        
        return $Results | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    "Info" { 5 }
                    default { 6 }
                }
            }
        }, TreeStructure
        
    } catch {
        $ErrorMessage = "Error during compromise indicators analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        return @([PSCustomObject]@{
            TreeStructure = "❌ COMPROMISE INDICATORS ANALYSIS ► Error"
            AccountName = "Analysis Failed"
            DisplayName = "Error during security analysis"
            IndicatorType = "System Error"
            Details = $ErrorMessage
            Severity = "Critical"
            RiskLevel = "Critical"
            SecurityImpact = "Unable to assess compromise indicators - security status unknown"
            Recommendation = "Retry analysis with appropriate permissions and check system connectivity"
            ComplianceNote = "Security analysis failure requires immediate attention"
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        })
    }
}

Function Get-SecurityDashboard {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Creating security assessment dashboard..." -Type Info -Terminal
        
        $SecurityMetrics = [System.Collections.ArrayList]::new()
        
        # 1. Password Security
        $TotalUsers = (Get-ADUser -Filter * -ErrorAction SilentlyContinue | Measure-Object).Count
        $PasswordNeverExpires = (Get-ADUser -Filter "PasswordNeverExpires -eq 'True'" -ErrorAction SilentlyContinue | Measure-Object).Count
        $ExpiredPasswords = (Get-ADUser -Filter "PasswordExpired -eq 'True'" -ErrorAction SilentlyContinue | Measure-Object).Count
        $ReversibleEncryption = (Get-ADUser -Filter {UserAccountControl -band 0x0080} -ErrorAction SilentlyContinue | Measure-Object).Count
        $PasswordNotRequired = (Get-ADUser -Filter {UserAccountControl -band 0x0020} -ErrorAction SilentlyContinue | Measure-Object).Count
        
        # Fix for PasswordLastSet filter syntax
        $cutoffDate = (Get-Date).AddDays(-90)
        $WeakPasswords = (Get-ADUser -Filter * -Properties PasswordLastSet -ErrorAction SilentlyContinue | 
            Where-Object { $_.PasswordLastSet -lt $cutoffDate } | 
            Measure-Object).Count
        
        $PasswordScore = 100
        if ($TotalUsers -gt 0) {
            $PasswordScore -= [math]::Round(($PasswordNeverExpires / $TotalUsers) * 25)
            $PasswordScore -= [math]::Round(($ExpiredPasswords / $TotalUsers) * 15)
            $PasswordScore -= [math]::Round(($ReversibleEncryption / $TotalUsers) * 30)
            $PasswordScore -= [math]::Round(($PasswordNotRequired / $TotalUsers) * 30)
            $PasswordScore -= [math]::Round(($WeakPasswords / $TotalUsers) * 20)
        }
        
        $null = $SecurityMetrics.Add([PSCustomObject]@{
            Category = "Password Security"
            Score = [math]::Max(0, $PasswordScore)
            Status = if ($PasswordScore -ge 80) { "Good" } elseif ($PasswordScore -ge 60) { "Fair" } else { "Critical" }
            Details = "Issues found: $PasswordNeverExpires non-expiring, $ExpiredPasswords expired, $ReversibleEncryption reversible encryption, $PasswordNotRequired not required, $WeakPasswords weak passwords"
            Recommendation = if ($PasswordScore -lt 80) { 
                "Review password policies and implement stronger requirements. Consider enabling password complexity and expiration." 
            } else { 
                "Password policies are adequate" 
            }
            ChartData = @{
                Labels = @("Non-expiring", "Expired", "Reversible", "Not Required", "Weak")
                Values = @($PasswordNeverExpires, $ExpiredPasswords, $ReversibleEncryption, $PasswordNotRequired, $WeakPasswords)
                Colors = @("#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0", "#FF9F40")
            }
        })
        
        # 2. Account Security
        $DisabledUsers = (Get-ADUser -Filter "Enabled -eq 'False'" -ErrorAction SilentlyContinue | Measure-Object).Count
        $LockedUsers = (Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue | Measure-Object).Count
        $InactiveUsers = (Get-ADUser -Filter * -Properties LastLogonDate -ErrorAction SilentlyContinue | 
            Where-Object { $_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90) } | Measure-Object).Count
        $KerberosDESUsers = (Get-ADUser -Filter {UserAccountControl -band 0x200000} -ErrorAction SilentlyContinue | Measure-Object).Count
        $SmartcardNotRequired = (Get-ADUser -Filter {-not(UserAccountControl -band 0x40000)} -ErrorAction SilentlyContinue | Measure-Object).Count
        
        $AccountScore = 100
        if ($TotalUsers -gt 0) {
            $AccountScore -= [math]::Round(($InactiveUsers / $TotalUsers) * 30)
            $AccountScore -= [math]::Round(($KerberosDESUsers / $TotalUsers) * 20)
            $AccountScore -= [math]::Round(($SmartcardNotRequired / $TotalUsers) * 15)
            if ($LockedUsers -gt 5) { $AccountScore -= 15 }
        }
        
        $null = $SecurityMetrics.Add([PSCustomObject]@{
            Category = "Account Management"
            Score = [math]::Max(0, $AccountScore)
            Status = if ($AccountScore -ge 80) { "Good" } elseif ($AccountScore -ge 60) { "Fair" } else { "Critical" }
            Details = "$InactiveUsers inactive (>90 days), $LockedUsers locked, $KerberosDESUsers using DES encryption, $SmartcardNotRequired without smartcard"
            Recommendation = if ($AccountScore -lt 80) { 
                "Clean up inactive accounts, review security settings, and consider enforcing smartcard authentication" 
            } else { 
                "Account management is effective" 
            }
            ChartData = @{
                Labels = @("Inactive", "Locked", "DES Encryption", "No Smartcard")
                Values = @($InactiveUsers, $LockedUsers, $KerberosDESUsers, $SmartcardNotRequired)
                Colors = @("#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0")
            }
        })
        
        # 3. Privileged Accounts with MFA Check
        $AdminGroups = @(
            $Global:ADGroupNames.DomainAdmins,
            $Global:ADGroupNames.EnterpriseAdmins,
            $Global:ADGroupNames.SchemaAdmins,
            $Global:ADGroupNames.Administrators
        )
        
        $PrivilegedUsers = 0
        $ProtectedAdmins = 0
        $AdminsWithoutMFA = 0
        $AdminsWithWeakAuth = 0
        
        foreach ($groupName in $AdminGroups) {
            try {
                $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                if ($group) {
                    $members = Get-ADGroupMember -Identity $group -ErrorAction SilentlyContinue
                    if ($members) { 
                        $PrivilegedUsers += $members.Count
                        
                        foreach ($member in $members) {
                            $user = Get-ADUser $member.SamAccountName -Properties UserAccountControl, msDS-AuthNPolicySiloMembersBL -ErrorAction SilentlyContinue
                            
                            # Check for Protected Admin status
                            if ($user.UserAccountControl -band 0x1000000) {
                                $ProtectedAdmins++
                            }
                            
                            # Check for MFA enforcement
                            if (-not ($user.UserAccountControl -band 0x40000) -and -not $user.'msDS-AuthNPolicySiloMembersBL') {
                                $AdminsWithoutMFA++
                            }
                            
                            # Check for weak authentication methods
                            if ($user.UserAccountControl -band 0x80000) { # DONT_REQ_PREAUTH
                                $AdminsWithWeakAuth++
                            }
                        }
                    }
                }
            } catch { 
                Write-ADReportLog -Message "Error checking admin group $groupName : $($_.Exception.Message)" -Type Warning
            }
        }
        
        $PrivilegeScore = 100
        if ($PrivilegedUsers -gt 10) { $PrivilegeScore -= 20 }
        if ($PrivilegedUsers -gt 20) { $PrivilegeScore -= 30 }
        if ($ProtectedAdmins -lt $PrivilegedUsers) { $PrivilegeScore -= 25 }
        if ($AdminsWithoutMFA -gt 0) { $PrivilegeScore -= 25 }
        if ($AdminsWithWeakAuth -gt 0) { $PrivilegeScore -= 20 }
        
        $null = $SecurityMetrics.Add([PSCustomObject]@{
            Category = "Privileged Accounts"
            Score = [math]::Max(0, $PrivilegeScore)
            Status = if ($PrivilegeScore -ge 80) { "Good" } elseif ($PrivilegeScore -ge 60) { "Fair" } else { "Critical" }
            Details = "$PrivilegedUsers admins total, $ProtectedAdmins protected, $AdminsWithoutMFA without MFA, $AdminsWithWeakAuth with weak auth"
            Recommendation = if ($PrivilegeScore -lt 80) { 
                "Configure AD FS additional authentication methods for admin accounts, enforce MFA, and protect all privileged accounts from delegation" 
            } else { 
                "Admin account security is well configured" 
            }
            ChartData = @{
                Labels = @("Total Admins", "Protected", "Without MFA", "Weak Auth")
                Values = @($PrivilegedUsers, $ProtectedAdmins, $AdminsWithoutMFA, $AdminsWithWeakAuth)
                Colors = @("#36A2EB", "#4BC0C0", "#FF6384", "#FFCE56")
            }
        })
        
        # 4. Group Security
        $TotalGroups = (Get-ADGroup -Filter * -ErrorAction SilentlyContinue | Measure-Object).Count
        $EmptyGroups = 0
        $NestedGroups = 0
        $UnprotectedGroups = 0
        
        Get-ADGroup -Filter * -Properties adminCount -ErrorAction SilentlyContinue | ForEach-Object {
            $members = Get-ADGroupMember -Identity $_ -ErrorAction SilentlyContinue
            if (-not $members) { 
                $EmptyGroups++ 
            } else {
                $NestedGroups += (@($members | Where-Object { $_.objectClass -eq 'group' })).Count
            }
            if ($_.adminCount -ne 1) { $UnprotectedGroups++ }
        }
        
        $GroupScore = 100
        if ($TotalGroups -gt 0) {
            $GroupScore -= [math]::Round(($EmptyGroups / $TotalGroups) * 20)
            if ($NestedGroups -gt 50) { $GroupScore -= 20 }
            $GroupScore -= [math]::Round(($UnprotectedGroups / $TotalGroups) * 15)
        }
        
        $null = $SecurityMetrics.Add([PSCustomObject]@{
            Category = "Group Management"
            Score = [math]::Max(0, $GroupScore)
            Status = if ($GroupScore -ge 80) { "Good" } elseif ($GroupScore -ge 60) { "Fair" } else { "Critical" }
            Details = "$EmptyGroups empty groups, $NestedGroups nested group memberships, $UnprotectedGroups unprotected groups"
            Recommendation = if ($GroupScore -lt 80) { 
                "Review empty groups, simplify group nesting, and ensure proper group protection" 
            } else { 
                "Group structure is clean" 
            }
            ChartData = @{
                Labels = @("Empty Groups", "Nested Groups", "Unprotected")
                Values = @($EmptyGroups, $NestedGroups, $UnprotectedGroups)
                Colors = @("#FF6384", "#36A2EB", "#FFCE56")
            }
        })
        
        # Overall Score
        $OverallScore = [math]::Round(($SecurityMetrics | Measure-Object -Property Score -Average).Average)
        
        $null = $SecurityMetrics.Add([PSCustomObject]@{
            Category = "OVERALL SCORE"
            Score = $OverallScore
            Status = if ($OverallScore -ge 80) { "Good" } elseif ($OverallScore -ge 60) { "Fair" } else { "Critical" }
            Details = "Comprehensive security assessment score"
            Recommendation = if ($OverallScore -lt 80) { 
                "Multiple security issues require immediate attention. Review detailed recommendations above." 
            } else { 
                "Security posture is strong. Continue monitoring and maintaining security controls." 
            }
            ChartData = @{
                Labels = @("Overall Security Score")
                Values = @($OverallScore)
                Colors = @(if ($OverallScore -ge 80) { "#4BC0C0" } elseif ($OverallScore -ge 60) { "#FFCE56" } else { "#FF6384" })
            }
        })
        
        Write-ADReportLog -Message "Security assessment completed. Overall score: $OverallScore" -Type Info -Terminal
        return $SecurityMetrics | Sort-Object Score
        
    } catch {
        Write-ADReportLog -Message "Error during security assessment: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-AuthProtocolAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing authentication protocols..." -Type Info -Terminal
        
        # Note: This analysis requires advanced logging configuration
        # Simulating data for demonstration
        
        $AuthProtocols = @()
        
        # Analyze UserAccountControl for Kerberos settings
        $KerberosUsers = Get-ADUser -Filter * -Properties UserAccountControl -ErrorAction SilentlyContinue | 
            Where-Object { -not ($_.UserAccountControl -band 0x80000) } # DONT_REQ_PREAUTH not set
        
        $NTLMOnlyUsers = Get-ADUser -Filter * -Properties UserAccountControl -ErrorAction SilentlyContinue |
            Where-Object { $_.UserAccountControl -band 0x20000 } # DONT_USE_KERBEROS set
        
        $TotalUsers = (Get-ADUser -Filter * -ErrorAction SilentlyContinue | Measure-Object).Count
        
        $AuthProtocols += [PSCustomObject]@{
            Protocol = "Kerberos"
            UserCount = $KerberosUsers.Count
            Percentage = if ($TotalUsers -gt 0) { [math]::Round(($KerberosUsers.Count / $TotalUsers) * 100, 2) } else { 0 }
            SecurityLevel = "High"
            Status = "Recommended"
            Details = "Modern, secure authentication"
            Recommendation = "Continue using as standard"
        }
        
        $AuthProtocols += [PSCustomObject]@{
            Protocol = "NTLM (forced)"
            UserCount = $NTLMOnlyUsers.Count
            Percentage = if ($TotalUsers -gt 0) { [math]::Round(($NTLMOnlyUsers.Count / $TotalUsers) * 100, 2) } else { 0 }
            SecurityLevel = "Low"
            Status = if ($NTLMOnlyUsers.Count -gt 0) { "Risk" } else { "OK" }
            Details = "Legacy protocol, vulnerable to attacks"
            Recommendation = if ($NTLMOnlyUsers.Count -gt 0) { "Migration to Kerberos recommended" } else { "No NTLM-only users found" }
        }
        
        # Analyze computer authentication
        $ComputerAuth = Get-ADComputer -Filter * -Properties OperatingSystem -ErrorAction SilentlyContinue |
            Group-Object -Property OperatingSystem
        
        foreach ($osGroup in $ComputerAuth) {
            $AuthProtocols += [PSCustomObject]@{
                Protocol = "Computer-Auth"
                UserCount = $osGroup.Count
                Percentage = "N/A"
                SecurityLevel = "Variable"
                Status = "Info"
                Details = "$($osGroup.Name): $($osGroup.Count) computers"
                Recommendation = "Ensure all systems support current authentication protocols"
            }
        }
        
        Write-ADReportLog -Message "Authentication protocol analysis completed." -Type Info -Terminal
        return $AuthProtocols
        
    } catch {
        Write-ADReportLog -Message "Error during authentication protocol analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-FailedAuthPatterns {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing failed authentication patterns..." -Type Info -Terminal
        
        [System.Collections.Generic.List[PSObject]]$FailedAuthPatterns = @()
        
        # Analyze BadLogonCount
        $UsersWithFailedLogons = Get-ADUser -Filter "BadLogonCount -gt 0" -Properties BadLogonCount, LastLogonDate, LockedOut, DisplayName, Department -ErrorAction SilentlyContinue
        
        # Group by BadLogonCount ranges
        $FailureRanges = @{
            "1-3 Failed Attempts" = $UsersWithFailedLogons | Where-Object { $_.BadLogonCount -ge 1 -and $_.BadLogonCount -le 3 }
            "4-5 Failed Attempts" = $UsersWithFailedLogons | Where-Object { $_.BadLogonCount -ge 4 -and $_.BadLogonCount -le 5 }
            "6-10 Failed Attempts" = $UsersWithFailedLogons | Where-Object { $_.BadLogonCount -ge 6 -and $_.BadLogonCount -le 10 }
            "Over 10 Failed Attempts" = $UsersWithFailedLogons | Where-Object { $_.BadLogonCount -gt 10 }
        }
        
        # Analyze patterns
        foreach ($range in $FailureRanges.GetEnumerator()) {
            if ($range.Value.Count -gt 0) {
                $topFailures = $range.Value | Sort-Object BadLogonCount -Descending | Select-Object -First 5
                
                foreach ($user in $topFailures) {
                    $pattern = "Normal"
                    $risk = "Low"
                    
                    if ($user.BadLogonCount -gt 10) {
                        $pattern = "Possible Brute Force Attack"
                        $risk = "High"
                    } elseif ($user.BadLogonCount -gt 5) {
                        $pattern = "Suspicious Activity"
                        $risk = "Medium"
                    } elseif ($user.LockedOut) {
                        $pattern = "Account Locked"
                        $risk = "Medium"
                    }
                    
                    $FailedAuthPatterns.Add([PSCustomObject]@{
                        AccountName = $user.SamAccountName
                        DisplayName = $user.DisplayName
                        Department = if ($user.Department) { $user.Department } else { "N/A" }
                        FailedAttempts = $user.BadLogonCount
                        Pattern = $pattern
                        Risk = $risk
                        LockedOut = $user.LockedOut
                        LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                        Recommendation = switch ($risk) {
                            "High" { "Immediate investigation required - possible attack" }
                            "Medium" { "Check account and consider password reset" }
                            default { "Inform user about secure passwords" }
                        }
                    })
                }
            }
        }
        
        # Add summary
        $summary = [PSCustomObject]@{
            AccountName = "SUMMARY"
            DisplayName = "Overview"
            Department = "All"
            FailedAttempts = ($UsersWithFailedLogons | Measure-Object -Property BadLogonCount -Sum).Sum
            Pattern = "Statistics"
            Risk = if ($FailureRanges["Over 10 Failed Attempts"].Count -gt 0) { "High" } elseif ($FailureRanges["6-10 Failed Attempts"].Count -gt 0) { "Medium" } else { "Low" }
            LockedOut = ($UsersWithFailedLogons | Where-Object { $_.LockedOut }).Count
            LastLogon = "N/A"
            Recommendation = "$($UsersWithFailedLogons.Count) users with failed login attempts"
        }
        
        $FailedAuthPatterns.Insert(0, $summary)
        
        if ($FailedAuthPatterns.Count -eq 1) {
            return @([PSCustomObject]@{
                AccountName = "None found"
                DisplayName = "N/A"
                Department = "N/A"
                FailedAttempts = 0
                Pattern = "No suspicious patterns"
                Risk = "None"
                LockedOut = $false
                LastLogon = "N/A"
                Recommendation = "No failed authentication patterns detected"
            })
        }
        
        Write-ADReportLog -Message "Failed authentication analysis completed. $($FailedAuthPatterns.Count - 1) patterns found." -Type Info -Terminal
        return $FailedAuthPatterns | Sort-Object -Property Risk, FailedAttempts -Descending
        
    } catch {
        Write-ADReportLog -Message "Error during failed authentication analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ConditionalAccessPolicies {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing conditional access policies..." -Type Info -Terminal
        
        # Note: Conditional Access Policies are primarily an Azure AD/Entra ID feature
        # For local AD we can analyze similar concepts
        
        [System.Collections.Generic.List[PSObject]]$ConditionalPolicies = @()
        
        # 1. Analyze account restrictions
        # First check which properties are available
        $availableProps = @('LogonWorkstations', 'AccountExpirationDate')
        
        # Check if AllowedToDelegateTo is available
        $testUser = Get-ADUser -Filter * -ResultSetSize 1 -ErrorAction SilentlyContinue
        if ($testUser) {
            try {
                $null = Get-ADUser -Identity $testUser.SamAccountName -Properties AllowedToDelegateTo -ErrorAction Stop
                $availableProps += 'AllowedToDelegateTo'
            } catch {
                # Property not available
            }
        }
        
        $RestrictedUsers = Get-ADUser -Filter * -Properties $availableProps -ErrorAction SilentlyContinue |
            Where-Object { 
                $_.LogonWorkstations -or 
                ($availableProps -contains 'AllowedToDelegateTo' -and $_.AllowedToDelegateTo) -or 
                $_.AccountExpirationDate 
            }
        
        foreach ($user in $RestrictedUsers) {
            $restrictions = @()
            
            if ($user.LogonWorkstations) {
                $restrictions += "Workstation Restriction"
            }
            if ($availableProps -contains 'AllowedToDelegateTo' -and $user.PSObject.Properties['AllowedToDelegateTo'] -and $user.AllowedToDelegateTo) {
                $restrictions += "Delegation Allowed"
            }
            if ($user.AccountExpirationDate) {
                $restrictions += "Account Expiration Set"
            }
            
            $ConditionalPolicies.Add([PSCustomObject]@{
                ObjectName = $user.SamAccountName
                ObjectType = "User"
                PolicyType = "Account Restrictions"
                Conditions = $restrictions -join ", "
                Status = "Active"
                Details = if ($user.LogonWorkstations) { "Allowed Workstations: $($user.LogonWorkstations)" } else { "Various Restrictions" }
                Recommendation = "Review if these restrictions are still required"
            })
        }
        
        # 2. Analyze group-based policies
        $AdminGroups = Get-ADGroup -Filter "Name -like '*admin*'" -Properties ManagedBy, Description -ErrorAction SilentlyContinue
        
        foreach ($group in $AdminGroups) {
            $ConditionalPolicies.Add([PSCustomObject]@{
                ObjectName = $group.Name
                ObjectType = "Group"
                PolicyType = "Administrative Group"
                Conditions = "Membership requires elevated rights"
                Status = if ($group.ManagedBy) { "Managed" } else { "Unmanaged" }
                Details = if ($group.Description) { $group.Description } else { "No Description" }
                Recommendation = if (-not $group.ManagedBy) { "Assign a manager" } else { "Regular member review" }
            })
        }
        
        # 3. Fine-Grained Password Policies as conditional access control
        try {
            $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue
            foreach ($fgpp in $FGPPs) {
                $ConditionalPolicies.Add([PSCustomObject]@{
                    ObjectName = $fgpp.Name
                    ObjectType = "Password Policy"
                    PolicyType = "Fine-Grained Password Policy"
                    Conditions = "Precedence: $($fgpp.Precedence)"
                    Status = "Active"
                    Details = "Min. Password Length: $($fgpp.MinPasswordLength), Max. Age: $($fgpp.MaxPasswordAge.Days) days"
                    Recommendation = "Ensure policy meets current security standards and best practices"
                })
            }
        } catch {
            Write-ADReportLog -Message "Fine-Grained Password Policies not available" -Type Warning -Terminal
        }
        
        if ($ConditionalPolicies.Count -eq 0) {
            return @([PSCustomObject]@{
                ObjectName = "None found"
                ObjectType = "N/A"
                PolicyType = "None"
                Conditions = "No conditional policies configured"
                Status = "N/A"
                Details = "No local conditional access policies found"
                Recommendation = "Consider implementing access controls"
            })
        }
        
        Write-ADReportLog -Message "Conditional access policy analysis completed. $($ConditionalPolicies.Count) policy(ies) found." -Type Info -Terminal
        return $ConditionalPolicies | Sort-Object PolicyType, ObjectName
        
    } catch {
        Write-ADReportLog -Message "Error during conditional access policy analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-HoneyTokens {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential honey token accounts and suspicious activities..." -Type Info -Terminal
        
        # Initialize empty results list
        [System.Collections.Generic.List[PSObject]]$PotentialHoneyTokens = [System.Collections.Generic.List[PSObject]]::new()
        
        # Define suspicious patterns
        $SuspiciousPatterns = @(
            "*honey*", "*canary*", "*trap*", "*decoy*", "*bait*", 
            "*test*", "*dummy*", "*fake*", "*monitor*", "*audit*",
            "*detect*", "*alert*", "*sensor*", "*probe*", "*track*",
            "*watch*", "*guard*", "*check*", "*scan*", "*suspicious*"
        )
        
        # 1. Check accounts with suspicious names
        foreach ($pattern in $SuspiciousPatterns) {
            $accounts = Get-ADUser -Filter "Name -like '$pattern' -or SamAccountName -like '$pattern'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department, LogonCount, BadLogonCount, PasswordLastSet, AccountExpirationDate, userAccountControl -ErrorAction SilentlyContinue
            
            if ($null -ne $accounts) {
                foreach ($account in $accounts) {
                    if ($null -ne $account) {
                        $PotentialHoneyTokens.Add($account)
                    }
                }
            }
        }
        
        # 2. Check accounts with suspicious descriptions
        $DescriptionBasedAccounts = Get-ADUser -Filter {
            Description -like "*honey*" -or 
            Description -like "*canary*" -or 
            Description -like "*monitoring*" -or 
            Description -like "*security*" -or 
            Description -like "*test*" -or 
            Description -like "*trap*" -or 
            Description -like "*decoy*" -or 
            Description -like "*detect*" -or 
            Description -like "*alert*"
        } -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department, LogonCount, BadLogonCount, PasswordLastSet, AccountExpirationDate, userAccountControl -ErrorAction SilentlyContinue
        
        if ($null -ne $DescriptionBasedAccounts) {
            foreach ($account in $DescriptionBasedAccounts) {
                if ($null -ne $account) {
                    $PotentialHoneyTokens.Add($account)
                }
            }
        }
        
        # Remove duplicates and null entries
        $UniqueAccounts = $PotentialHoneyTokens | Where-Object {$null -ne $_} | Sort-Object DistinguishedName -Unique
        
        # Return default object if no results found
        if ($null -eq $UniqueAccounts -or $UniqueAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No potential honey token accounts found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                SamAccountName = "N/A" 
                Description = "No potential honey token accounts found"
                PotentialHoneyToken = "None"
                SuspicionLevel = "N/A"
                Indicators = "No suspicious patterns detected"
            })
        }
        
        # Process results
        $Results = foreach ($account in $UniqueAccounts) {
            if ($null -eq $account) { continue }
            
            # Calculate dates
            $ageInDays = if ($null -ne $account.whenCreated) {
                [math]::Round((New-TimeSpan -Start $account.whenCreated -End (Get-Date)).TotalDays, 0)
            } else {
                0
            }
            
            $daysSinceLastLogon = if ($null -ne $account.LastLogonDate) {
                [math]::Round((New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).TotalDays, 0)
            } else {
                9999
            }
            
            # Calculate suspicion score
            [int]$suspicionScore = 0
            [System.Collections.Generic.List[string]]$indicators = [System.Collections.Generic.List[string]]::new()
            
            # Add indicators based on patterns
            if ($account.Name -match "(honey|canary|trap|decoy|bait)") { 
                $suspicionScore += 5
                $indicators.Add("Suspicious Name Pattern")
            }
            
            if ($account.PasswordNeverExpires) {
                $suspicionScore += 2
                $indicators.Add("Password Never Expires")
            }
            
            if ($account.LogonCount -eq 0) {
                $suspicionScore += 3
                $indicators.Add("Never Logged On")
            }
            
            # Create result object
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = if ($null -ne $account.Description) { $account.Description } else { "No Description" }
                Department = if ($null -ne $account.Department) { $account.Department } else { "No Department" }
                Enabled = $account.Enabled
                LastLogon = if ($null -ne $account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                AgeInDays = $ageInDays
                SuspicionLevel = switch ($suspicionScore) {
                    {$_ -ge 8} { "High" }
                    {$_ -ge 5} { "Medium" }
                    default { "Low" }
                }
                Indicators = if ($indicators.Count -gt 0) { $indicators -join ", " } else { "None" }
                PotentialHoneyToken = if ($suspicionScore -ge 5) { "Likely" } else { "Unlikely" }
            }
        }
        
        Write-ADReportLog -Message "Honey Token Analysis completed. Found $($Results.Count) potential accounts." -Type Info -Terminal
        return $Results | Sort-Object SuspicionLevel -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing honey tokens: $($_.Exception.Message)" -Type Error -Terminal
        return @()
    }
}
Function Get-GPOOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive GPO overview..." -Type Info -Terminal
        
        # PrÃ¼fe ob GroupPolicy Module verfÃ¼gbar ist
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "GroupPolicy PowerShell module not available. Attempting alternative approach..." -Type Warning
            
            # Alternative: Ãœber AD direkt abfragen
            $GPOs = Get-ADObject -SearchBase "CN=Policies,CN=System,$((Get-ADDomain).DistinguishedName)" -Filter "ObjectClass -eq 'groupPolicyContainer'" -Properties DisplayName, whenCreated, whenChanged, gPCFileSysPath -ErrorAction Stop
            
            $Results = foreach ($gpo in $GPOs) {
                [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Id = $gpo.Name
                    CreationTime = $gpo.whenCreated
                    ModificationTime = $gpo.whenChanged
                    GPOStatus = "Unknown (Manual Check Required)"
                    FileSysPath = $gpo.gPCFileSysPath
                    AgeInDays = if ($gpo.whenCreated) { (New-TimeSpan -Start $gpo.whenCreated -End (Get-Date)).Days } else { 0 }
                    DaysSinceModification = if ($gpo.whenChanged) { (New-TimeSpan -Start $gpo.whenChanged -End (Get-Date)).Days } else { 0 }
                }
            }
            
            return $Results | Sort-Object DaysSinceModification
        }
        
        Import-Module GroupPolicy -ErrorAction Stop
        $GPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $GPOs) {
            # GPO Links analysieren
            $gpoLinks = @()
            try {
                $gpoReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
                if ($gpoReport) {
                    # Vereinfachte Link-Analyse
                    $linkCount = ([xml]$gpoReport).GPO.LinksTo.SOMPath.Count
                } else {
                    $linkCount = 0
                }
            } catch {
                $linkCount = 0
            }
            
            $ageInDays = if ($gpo.CreationTime) { (New-TimeSpan -Start $gpo.CreationTime -End (Get-Date)).Days } else { 0 }
            $daysSinceModification = if ($gpo.ModificationTime) { (New-TimeSpan -Start $gpo.ModificationTime -End (Get-Date)).Days } else { 0 }
            
            # Status-Bewertung
            $healthStatus = "Good"
            $issues = @()
            
            if ($daysSinceModification -gt 365) { $issues += "Not modified for over 1 year"; $healthStatus = "Review" }
            if ($linkCount -eq 0) { $issues += "No links found"; $healthStatus = "Orphaned" }
            if ($gpo.GpoStatus -eq "AllSettingsDisabled") { $issues += "All settings disabled"; $healthStatus = "Disabled" }
            
            [PSCustomObject]@{
                DisplayName = $gpo.DisplayName
                Id = $gpo.Id
                CreationTime = $gpo.CreationTime
                ModificationTime = $gpo.ModificationTime
                GPOStatus = $gpo.GpoStatus
                AgeInDays = $ageInDays
                DaysSinceModification = $daysSinceModification
                LinkCount = $linkCount
                HealthStatus = $healthStatus
                Issues = $issues -join "; "
                Domain = $gpo.DomainName
                Owner = $gpo.Owner
            }
        }
        
        Write-ADReportLog -Message "GPO overview completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object HealthStatus, DaysSinceModification -Descending
    } catch {
        Write-ADReportLog -Message "Error generating GPO overview: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Gruppen-Report Funktionen ---
Function Get-CircularGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing circular group references..." -Type Info -Terminal
        
        # Dictionary zur Speicherung der Gruppenmitgliedschaften
        $groupMemberships = @{}
        $circularReferences = @()
        
        # Lade alle Gruppen mit ihren Mitgliedschaften
        $allGroups = Get-ADGroup -Filter * -Properties MemberOf, Members, GroupCategory, GroupScope -ErrorAction Stop
        
        # Baue Dictionary auf
        foreach ($group in $allGroups) {
            $groupMemberships[$group.DistinguishedName] = @{
                Group = $group
                MemberOf = $group.MemberOf
                Visited = $false
                InStack = $false
            }
        }
        
        # Funktion zur rekursiven Suche nach zirkulären Referenzen
        function Find-CircularReference {
            param(
                [string]$GroupDN,
                [System.Collections.ArrayList]$Path
            )
            
            if (-not $groupMemberships.ContainsKey($GroupDN)) { return }
            
            $groupData = $groupMemberships[$GroupDN]
            
            # Wenn bereits im Stack, haben wir eine zirkuläre Referenz gefunden
            if ($groupData.InStack) {
                $circularPath = $Path.Clone()
                $circularPath.Add($GroupDN) | Out-Null
                
                # Finde den Start des Zyklus
                $startIndex = $circularPath.IndexOf($GroupDN)
                $cycle = $circularPath[$startIndex..($circularPath.Count - 1)]
                
                # Erstelle eine eindeutige ID für den Zyklus
                $cycleId = ($cycle | Sort-Object) -join '|'
                
                if (-not ($circularReferences | Where-Object { $_.CycleId -eq $cycleId })) {
                    $groupNames = $cycle | ForEach-Object {
                        if ($groupMemberships.ContainsKey($_)) {
                            $groupMemberships[$_].Group.Name
                        } else {
                            "Unknown"
                        }
                    }
                    
                    $circularReferences += [PSCustomObject]@{
                        CycleId = $cycleId
                        CircularPath = $groupNames -join " → "
                        NumberOfGroups = $cycle.Count
                        FirstGroup = $groupNames[0]
                        GroupsInvolved = $groupNames -join ", "
                        Risk = if ($cycle.Count -gt 3) { "High" } elseif ($cycle.Count -eq 3) { "Medium" } else { "Low" }
                        Recommendation = "Break circular reference by removing one membership link"
                    }
                }
                return
            }
            
            # Markiere als besucht und im Stack
            $groupData.Visited = $true
            $groupData.InStack = $true
            $Path.Add($GroupDN) | Out-Null
            
            # Durchlaufe alle Parent-Gruppen
            foreach ($parentDN in $groupData.MemberOf) {
                Find-CircularReference -GroupDN $parentDN -Path $Path
            }
            
            # Entferne aus Stack
            $groupData.InStack = $false
            $Path.RemoveAt($Path.Count - 1)
        }
        
        # Suche nach zirkulären Referenzen für alle Gruppen
        foreach ($groupDN in $groupMemberships.Keys) {
            if (-not $groupMemberships[$groupDN].Visited) {
                $path = New-Object System.Collections.ArrayList
                Find-CircularReference -GroupDN $groupDN -Path $path
            }
        }
        
        if ($circularReferences.Count -eq 0) {
            Write-ADReportLog -Message "No circular group references found." -Type Info -Terminal
            return @([PSCustomObject]@{
                CircularPath = "None Found"
                NumberOfGroups = 0
                FirstGroup = "N/A"
                GroupsInvolved = "No circular references detected"
                Risk = "None"
                Recommendation = "Environment is healthy - no circular references"
            })
        }
        
        Write-ADReportLog -Message "Circular group reference analysis completed. $($circularReferences.Count) circular reference(s) found." -Type Info -Terminal
        return $circularReferences | Sort-Object NumberOfGroups, FirstGroup
        
    } catch {
        Write-ADReportLog -Message "Error analyzing circular group references: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsByTypeScope {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups by type and scope..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties GroupCategory, GroupScope, ManagedBy, whenCreated, whenChanged, Members, MemberOf -ErrorAction Stop
        
        # Gruppiere nach Type und Scope
        $GroupStats = $Groups | Group-Object -Property GroupCategory, GroupScope
        
        $Results = foreach ($statGroup in $GroupStats) {
            $category = $statGroup.Group[0].GroupCategory
            $scope = $statGroup.Group[0].GroupScope
            
            # Konvertiere Enum-Werte zu String für ToUpper()
            $categoryString = $category.ToString()
            $scopeString = $scope.ToString()
            
            # Berechne durchschnittliche Mitgliederzahl
            $memberCounts = @()
            foreach ($group in $statGroup.Group) {
                try {
                    $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
                    $memberCounts += $memberCount
                } catch {
                    $memberCounts += 0
                }
            }
            
            $avgMembers = if ($memberCounts.Count -gt 0) { 
                [math]::Round(($memberCounts | Measure-Object -Average).Average, 1) 
            } else { 0 }
            
            $maxMembers = if ($memberCounts.Count -gt 0) { 
                ($memberCounts | Measure-Object -Maximum).Maximum 
            } else { 0 }
            
            # Best Practice Empfehlungen
            $recommendation = switch ("$categoryString-$scopeString") {
                "Security-Global" { "Standard configuration for security groups - use for user assignments" }
                "Security-DomainLocal" { "Good for resource permissions - assign to local resources" }
                "Security-Universal" { "Appropriate for cross-forest scenarios - monitor replication load" }
                "Distribution-Global" { "Standard for email distribution - maintain for organizational communication" }
                "Distribution-DomainLocal" { "Rarely used configuration - consider converting to global scope" }
                "Distribution-Universal" { "Suitable for cross-forest distribution - monitor mail flow" }
                default { "Review group configuration and usage patterns" }
            }
            
            # Erstelle strukturierte Gruppenliste im gewünschten Format
            $groupListHeader = "$($categoryString.ToUpper())-$($scopeString.ToUpper())"
            $separator = "-" * $groupListHeader.Length
            $sortedGroupNames = ($statGroup.Group | Sort-Object Name | ForEach-Object { "  $($_.Name)" }) -join "`r`n"
            $formattedGroupList = "$groupListHeader`r`n$separator`r`n$sortedGroupNames`r`n"
            
            [PSCustomObject]@{
                Type = "$categoryString-$scopeString"
                TotalCount = $statGroup.Count
                Percentage = [math]::Round(($statGroup.Count / $Groups.Count) * 100, 2)
                AverageMemberCount = $avgMembers
                MaxMemberCount = $maxMembers
                GroupList = $formattedGroupList
                Recommendation = $recommendation
            }
        }
        
        Write-ADReportLog -Message "Groups by type/scope analysis completed. $($Results.Count) category combinations found." -Type Info -Terminal
        
        return $Results | Sort-Object Type
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups by type and scope: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DynamicDistGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing dynamic distribution groups..." -Type Info -Terminal
        
        # Hinweis: Dynamic Distribution Groups sind ein Exchange-Feature
        # Diese Funktion würde Exchange-Cmdlets benötigen
        
        Write-ADReportLog -Message "Dynamic Distribution Groups require Exchange PowerShell module." -Type Warning
        
        return @([PSCustomObject]@{
            Name = "Not Available"
            Status = "Exchange Required"
            Description = "Dynamic Distribution Groups are an Exchange feature"
            Recommendation = "Use Exchange Management Shell for this report"
        })
        
    } catch {
        Write-ADReportLog -Message "Error analyzing dynamic distribution groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-MailEnabledGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing mail-enabled groups..." -Type Info -Terminal
        
        # Prüfe, ob Exchange-Attribute verfügbar sind
        $exchangeInstalled = $false
        try {
            $testGroup = Get-ADGroup -Filter * -Properties msExchRecipientTypeDetails -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $testGroup.PSObject.Properties['msExchRecipientTypeDetails']) {
                $exchangeInstalled = $true
            }
        } catch {
            $exchangeInstalled = $false
        }
        
        # Basis-Properties, die immer verfügbar sind
        $baseProperties = @('mail', 'proxyAddresses', 'GroupCategory', 'GroupScope', 'ManagedBy')
        
        # Exchange-spezifische Properties nur wenn Exchange installiert ist
        if ($exchangeInstalled) {
            $properties = $baseProperties + @('legacyExchangeDN', 'msExchRecipientTypeDetails')
        } else {
            $properties = $baseProperties
        }
        
        # Erweiterte Filterung für Mail-aktivierte Gruppen
        $filterString = "(mail -like '*') -or (proxyAddresses -like '*SMTP:*')"
        $MailGroups = Get-ADGroup -Filter $filterString -Properties $properties -ErrorAction Stop
        
        if ($MailGroups.Count -eq 0) {
            Write-ADReportLog -Message "No mail-enabled groups found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Mail = "N/A"
                GroupCategory = "N/A"
                Status = "No mail-enabled groups found"
                Recommendation = "Mail-enabled groups may be managed through Exchange"
            })
        }
        
        $Results = foreach ($group in $MailGroups) {
            # Prüfe Gruppenmitgliederzahl
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Analysiere Mail-Eigenschaften
            $primarySMTP = $null
            if ($group.proxyAddresses) {
                $primarySMTP = $group.proxyAddresses | Where-Object { $_ -clike "SMTP:*" } | Select-Object -First 1
            }
            $effectiveMailAddress = if ($primarySMTP) { 
                $primarySMTP -replace "SMTP:", "" 
            } elseif ($group.mail) { 
                $group.mail 
            } else {
                "No valid email address"
            }
            
            # Manager auflösen
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Status und Empfehlungen
            $status = "Active"
            $recommendations = @()
            
            if (-not $effectiveMailAddress -or $effectiveMailAddress -eq "No valid email address") {
                $status = "Incomplete"
                $recommendations += "No valid email configuration found"
            }
            
            if ($memberCount -eq 0) {
                $recommendations += "Group has no members"
            }
            
            if ($memberCount -gt 1000) {
                $recommendations += "Large distribution list - consider breaking up"
            }
            
            if ([string]::IsNullOrWhiteSpace($group.ManagedBy)) {
                $recommendations += "No manager assigned"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                Mail = $effectiveMailAddress
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                ManagedBy = $managerName
                ProxyAddressCount = if ($group.proxyAddresses) { $group.proxyAddresses.Count } else { 0 }
                Status = $status
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Properly configured" }
            }
        }
        
        Write-ADReportLog -Message "Mail-enabled groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing mail-enabled groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsWithoutOwners {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups without owners/managers..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties ManagedBy, GroupCategory, GroupScope, whenCreated, whenChanged, info -ErrorAction Stop
        
        # Filtere Gruppen ohne ManagedBy
        $UnmanagedGroups = $Groups | Where-Object { [string]::IsNullOrWhiteSpace($_.ManagedBy) }
        
        $Results = foreach ($group in $UnmanagedGroups) {
            # Prüfe Gruppengröße
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Alter der Gruppe
            $ageInDays = if ($group.whenCreated) { 
                (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days 
            } else { 0 }
            
            # Risikobewertung
            $riskLevel = "Low"
            $riskFactors = @()
            
            if ($group.GroupCategory -eq "Security") {
                $riskLevel = "Medium"
                $riskFactors += "Security group"
            }
            
            if ($memberCount -gt 50) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $riskFactors += "Large membership ($memberCount members)"
            }
            
            if ($memberCount -gt 100 -and $group.GroupCategory -eq "Security") {
                $riskLevel = "High"
            }
            
            if ($ageInDays -gt 365) {
                $riskFactors += "Old group (>1 year)"
            }
            
            # Empfehlungen
            $recommendations = @("Assign a group manager/owner")
            
            if ($memberCount -eq 0 -and $ageInDays -gt 90) {
                $recommendations += "Consider deleting empty old group"
            }
            
            if ($group.GroupCategory -eq "Security" -and $memberCount -gt 20) {
                $recommendations += "Security groups should have designated owners"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                Description = if ($group.Description) { $group.Description } else { "(No description)" }
                WhenCreated = $group.whenCreated
                AgeInDays = $ageInDays
                LastModified = $group.whenChanged
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors) { $riskFactors -join "; " } else { "None" }
                Recommendations = $recommendations -join "; "
            }
        }
        
        Write-ADReportLog -Message "Groups without owners analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups without owners: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-LargeGroups {
    [CmdletBinding()]
    param([int]$Threshold = 100)
    
    try {
        Write-ADReportLog -Message "Analyzing large groups (threshold: $Threshold members)..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties GroupCategory, GroupScope, ManagedBy, whenCreated -ErrorAction Stop
        
        $LargeGroups = @()
        
        foreach ($group in $Groups) {
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction SilentlyContinue).Count
            
            if ($memberCount -ge $Threshold) {
                # Nested group analysis
                $directMembers = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue)
                $nestedGroups = @($directMembers | Where-Object { $_.objectClass -eq "group" })
                
                # Manager info
                $managerName = "None"
                if ($group.ManagedBy) {
                    try {
                        $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                        if ($manager) {
                            $managerName = $manager.DisplayName
                        }
                    } catch {
                        $managerName = "Unknown"
                    }
                }
                
                # Performance impact assessment
                $performanceImpact = "Low"
                if ($memberCount -gt 1000) {
                    $performanceImpact = "High"
                } elseif ($memberCount -gt 500) {
                    $performanceImpact = "Medium"
                }
                
                # Recommendations
                $recommendations = @()
                if ($memberCount -gt 5000) {
                    $recommendations += "Consider breaking into smaller groups"
                }
                if ($nestedGroups.Count -gt 10) {
                    $recommendations += "High number of nested groups may impact performance"
                }
                if ($group.GroupScope -eq "Global" -and $memberCount -gt 5000) {
                    $recommendations += "Consider using Universal scope for large groups"
                }
                
                $LargeGroups += [PSCustomObject]@{
                    Name = $group.Name
                    GroupCategory = $group.GroupCategory
                    GroupScope = $group.GroupScope
                    TotalMembers = $memberCount
                    DirectMembers = $directMembers.Count
                    NestedGroups = $nestedGroups.Count
                    ManagedBy = $managerName
                    WhenCreated = $group.whenCreated
                    PerformanceImpact = $performanceImpact
                    Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor group size" }
                }
            }
        }
        
        if ($LargeGroups.Count -eq 0) {
            Write-ADReportLog -Message "No large groups found (threshold: $Threshold members)." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                TotalMembers = 0
                Status = "No groups with $Threshold or more members found"
                Recommendations = "All groups are within size limits"
            })
        }
        
        Write-ADReportLog -Message "Large groups analysis completed. $($LargeGroups.Count) groups found." -Type Info -Terminal
        return $LargeGroups | Sort-Object TotalMembers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing large groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-RecentlyModifiedGroups {
    [CmdletBinding()]
    param([int]$Days = 7)
    
    try {
        Write-ADReportLog -Message "Analyzing recently modified groups (last $Days days)..." -Type Info -Terminal
        
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Groups = Get-ADGroup -Filter "whenChanged -gt '$CutoffDate'" -Properties whenChanged, whenCreated, ManagedBy, GroupCategory, GroupScope, modifyTimeStamp -ErrorAction Stop
        
        if ($Groups.Count -eq 0) {
            Write-ADReportLog -Message "No groups modified in the last $Days days." -Type Info -Terminal
            $Results = @()
            $Results += [PSCustomObject]@{
                Name = "No Results"
                Status = "No groups modified in the last $Days days"
                LastModified = "N/A"
                Recommendations = "Normal - no recent group modifications"
            }
            return $Results
        }
        
        $Results = foreach ($group in $Groups) {
            # Versuche die Art der Änderung zu ermitteln
            $changeType = "Modified"
            if ($group.whenCreated -and $group.whenCreated -gt $CutoffDate) {
                $changeType = "Created"
            }
            
            # Manager info
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Mitgliederzahl für Kontext
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Tage seit Änderung
            $daysSinceChange = [math]::Round((New-TimeSpan -Start $group.whenChanged -End (Get-Date)).TotalDays, 1)
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                ChangeType = $changeType
                LastModified = $group.whenChanged
                DaysSinceChange = $daysSinceChange
                CreatedDate = $group.whenCreated
                MemberCount = $memberCount
                ManagedBy = $managerName
                ModificationTime = $group.whenChanged.ToString("HH:mm:ss")
            }
        }
        
        Write-ADReportLog -Message "Recently modified groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object LastModified -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing recently modified groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Computer-Report Funktionen ---
Function Get-ComputersByOSVersion {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by OS version..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, Name -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            # Support-Status ermitteln
            $supportStatus = "Unknown"
            $eolDate = "Unknown"
            
            if ($computer.OperatingSystem -like "*Windows 7*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($computer.OperatingSystem -like "*Windows 8*" -and $computer.OperatingSystem -notlike "*8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2016-01-12"
            } elseif ($computer.OperatingSystem -like "*Windows 8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-01-10"
            } elseif ($computer.OperatingSystem -like "*Windows 10*") {
                $supportStatus = "Check Version"
                $eolDate = "Version-dependent"
            } elseif ($computer.OperatingSystem -like "*Windows 11*") {
                $supportStatus = "Supported"
                $eolDate = "Active"
            } elseif ($computer.OperatingSystem -like "*Server 2008*" -and $computer.OperatingSystem -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2015-07-14"
            } elseif ($computer.OperatingSystem -like "*Server 2008 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($computer.OperatingSystem -like "*Server 2012*" -and $computer.OperatingSystem -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($computer.OperatingSystem -like "*Server 2012 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($computer.OperatingSystem -like "*Server 2016*") {
                $supportStatus = "Supported"
                $eolDate = "2027-01-12"
            } elseif ($computer.OperatingSystem -like "*Server 2019*") {
                $supportStatus = "Supported"
                $eolDate = "2029-01-09"
            } elseif ($computer.OperatingSystem -like "*Server 2022*") {
                $supportStatus = "Supported"
                $eolDate = "2031-10-14"
            }
            
            [PSCustomObject]@{
                ComputerName = $computer.Name
                OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                Version = if ($computer.OperatingSystemVersion) { $computer.OperatingSystemVersion } else { "Unknown" }
                ServicePack = if ($computer.OperatingSystemServicePack) { $computer.OperatingSystemServicePack } else { "None" }
                SupportStatus = $supportStatus
                EndOfLifeDate = $eolDate
            }
        }
        
        Write-ADReportLog -Message "Computers by OS version analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object OperatingSystem, Version
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by OS version: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-OSSummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing OS summary..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, LastLogonDate, Enabled -ErrorAction Stop
        
        # Gruppiere nach OS
        $OSGroups = $Computers | Group-Object OperatingSystem
        
        $Results = foreach ($osGroup in $OSGroups) {
            $os = if ([string]::IsNullOrWhiteSpace($osGroup.Name)) { "Unknown" } else { $osGroup.Name }
            
            # Statistiken für dieses OS
            $enabledCount = @($osGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($osGroup.Group | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # Support-Status ermitteln
            $supportStatus = "Unknown"
            $eolDate = "Unknown"
            
            if ($os -like "*Windows 7*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Windows 8*" -and $os -notlike "*8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2016-01-12"
            } elseif ($os -like "*Windows 8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-01-10"
            } elseif ($os -like "*Windows 10*") {
                $supportStatus = "Check Version"
                $eolDate = "Version-dependent"
            } elseif ($os -like "*Windows 11*") {
                $supportStatus = "Supported"
                $eolDate = "Active"
            } elseif ($os -like "*Server 2008*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2015-07-14"
            } elseif ($os -like "*Server 2008 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Server 2012*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2012 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2016*") {
                $supportStatus = "Supported"
                $eolDate = "2027-01-12"
            } elseif ($os -like "*Server 2019*") {
                $supportStatus = "Supported"
                $eolDate = "2029-01-09"
            } elseif ($os -like "*Server 2022*") {
                $supportStatus = "Supported"
                $eolDate = "2031-10-14"
            }
            
            # Risikobewertung
            $riskLevel = "Low"
            if ($supportStatus -eq "End of Life") {
                $riskLevel = "Critical"
            } elseif ($supportStatus -eq "Check Version") {
                $riskLevel = "Medium"
            }
            
            [PSCustomObject]@{
                OperatingSystem = $os
                Count = $osGroup.Count
                EnabledCount = $enabledCount
                ActiveCount = $activeCount
                InactiveCount = $osGroup.Count - $activeCount
                Percentage = [math]::Round(($osGroup.Count / $Computers.Count) * 100, 2)
                SupportStatus = $supportStatus
                EndOfLifeDate = $eolDate
                RiskLevel = $riskLevel
                Recommendation = if ($supportStatus -eq "End of Life") { 
                    "Urgent: Upgrade or replace systems" 
                } elseif ($supportStatus -eq "Check Version") { 
                    "Verify specific version support status" 
                } else { 
                    "Keep systems updated" 
                }
            }
        }
        
        Write-ADReportLog -Message "OS summary analysis completed. $($Results.Count) unique OS versions found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, Count -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing OS summary: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-BitLockerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing BitLocker status..." -Type Info -Terminal
        
        # Nur die relevanten BitLocker-Attribute abfragen
        $Computers = Get-ADComputer -Filter * -Properties Name, DistinguishedName -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            $bitlockerStatus = "Unknown"
            $recoveryKey = "Not Found"
            
            # BitLocker-Recovery-Informationen suchen
            try {
                $recoveryObjects = Get-ADObject -Filter {objectClass -eq 'msFVE-RecoveryInformation'} -SearchBase $computer.DistinguishedName -Properties "msFVE-RecoveryPassword" -ErrorAction Stop
                
                if ($recoveryObjects) {
                    $recoveryKey = $recoveryObjects."msFVE-RecoveryPassword"
                    $bitlockerStatus = "Enabled"
                }
            }
            catch {
                Write-ADReportLog -Message "Warning: Could not retrieve BitLocker recovery info for $($computer.Name): $($_.Exception.Message)" -Type Warning
            }
            
            [PSCustomObject]@{
                ComputerName = $computer.Name
                BitLockerStatus = $bitlockerStatus
                RecoveryKey = $recoveryKey
                Recommendation = if ($bitlockerStatus -eq "Unknown") {
                    "Enable BitLocker encryption"
                } elseif ($recoveryKey -eq "Not Found") {
                    "BitLocker enabled but recovery key not stored in AD"
                } else {
                    "BitLocker properly configured with recovery key"
                }
            }
        }
        
        Write-ADReportLog -Message "BitLocker status analysis completed. $($Results.Count) computers analyzed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing BitLocker status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-StaleComputerPasswords {
    [CmdletBinding()]
    param([int]$Days = 90)
    
    try {
        Write-ADReportLog -Message "Analyzing stale computer passwords..." -Type Info -Terminal
        
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Computers = Get-ADComputer -Filter * -Properties PasswordLastSet, LastLogonDate, OperatingSystem -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            if ($computer.PasswordLastSet -lt $CutoffDate) {
                [PSCustomObject]@{
                    ComputerName = $computer.Name
                    PasswordLastSet = $computer.PasswordLastSet
                    DaysSinceLastChange = [math]::Round((New-TimeSpan -Start $computer.PasswordLastSet -End (Get-Date)).TotalDays, 1)
                    LastLogon = $computer.LastLogonDate
                    OperatingSystem = $computer.OperatingSystem
                    RiskLevel = if ((New-TimeSpan -Start $computer.PasswordLastSet -End (Get-Date)).TotalDays -gt 180) { "High" } else { "Medium" }
                    Recommendation = "Reset computer password and verify network connectivity"
                }
            }
        }
        
        Write-ADReportLog -Message "Stale computer passwords analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastChange -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing stale computer passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersNeverLoggedOn {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers that never logged on..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter "LastLogonDate -notlike '*'" -Properties LastLogonDate, whenCreated, OperatingSystem, Enabled, Description, DistinguishedName -ErrorAction Stop
        
        if ($Computers.Count -eq 0) {
            Write-ADReportLog -Message "No computers found that never logged on." -Type Info -Terminal
            $Results = @()
            $Results += [PSCustomObject]@{
                Name = "No Results"
                Status = "All computers have logged on at least once"
                Recommendation = "Environment is healthy"
            }
            return $Results
        }
        
        $Results = foreach ($computer in $Computers) {
            # Alter des Computer-Objekts
            $ageInDays = if ($computer.whenCreated) {
                (New-TimeSpan -Start $computer.whenCreated -End (Get-Date)).Days
            } else { "Unknown" }
            
            # OU-Pfad extrahieren
            $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
            
            # Risikobewertung
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($computer.Enabled) {
                $riskLevel = "Medium"
                $recommendations += "Consider disabling unused computer"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 90) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                if ($computer.Enabled) { $riskLevel = "High" }
                $recommendations += "Old unused computer account - consider deletion"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 180) {
                $riskLevel = "High"
                $recommendations = @("Delete unused computer account")
            }
            
            [PSCustomObject]@{
                Name = $computer.Name
                OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                Enabled = $computer.Enabled
                WhenCreated = $computer.whenCreated
                AgeInDays = $ageInDays
                Description = if ($computer.Description) { $computer.Description } else { "(No description)" }
                OrganizationalUnit = $ouPath
                RiskLevel = $riskLevel
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor account" }
                Status = if ($computer.Enabled) { "Enabled but never used" } else { "Disabled and never used" }
            }
        }
        
        Write-ADReportLog -Message "Computers never logged on analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, AgeInDays -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers never logged on: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DuplicateComputerNames {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing duplicate computer names..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties DNSHostName, Enabled, OperatingSystem, LastLogonDate, whenCreated, DistinguishedName, IPv4Address -ErrorAction Stop
        
        # Suche nach Duplikaten basierend auf Name (ohne $)
        $ComputerNames = $Computers | ForEach-Object { 
            [PSCustomObject]@{
                Computer = $_
                CleanName = $_.Name.TrimEnd('$').ToUpper()
            }
        }
        
        # Gruppiere nach bereinigtem Namen
        $DuplicateGroups = $ComputerNames | Group-Object CleanName | Where-Object { $_.Count -gt 1 }
        
        if ($DuplicateGroups.Count -eq 0) {
            Write-ADReportLog -Message "No duplicate computer names found." -Type Info -Terminal
            $Results = @()
            $Results += [PSCustomObject]@{
                Name = "No Duplicates"
                Status = "No duplicate computer names detected"
                Count = 0
                Recommendation = "Environment is healthy"
            }
            return $Results
        }
        
        $Results = foreach ($dupGroup in $DuplicateGroups) {
            foreach ($item in $dupGroup.Group) {
                $computer = $item.Computer
                
                # OU-Pfad extrahieren
                $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
                
                # Bestimme welcher Computer der "aktive" ist
                $isActive = $computer.Enabled -and $computer.LastLogonDate -and 
                           $computer.LastLogonDate -gt (Get-Date).AddDays(-30)
                
                [PSCustomObject]@{
                    Name = $computer.Name
                    CleanName = $item.CleanName
                    DNSHostName = $computer.DNSHostName
                    IPv4Address = if ($computer.IPv4Address) { $computer.IPv4Address } else { "No IP" }
                    Enabled = $computer.Enabled
                    OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                    LastLogonDate = $computer.LastLogonDate
                    WhenCreated = $computer.whenCreated
                    OrganizationalUnit = $ouPath
                    DuplicateCount = $dupGroup.Count
                    Status = if ($isActive) { "Active Duplicate" } 
                            elseif ($computer.Enabled) { "Enabled but Inactive" } 
                            else { "Disabled" }
                    Recommendation = if (-not $computer.Enabled -and (-not $computer.LastLogonDate -or 
                                      $computer.LastLogonDate -lt (Get-Date).AddDays(-90))) {
                                        "Delete disabled/inactive duplicate"
                                    } elseif ($dupGroup.Count -gt 2) {
                                        "Multiple duplicates - immediate investigation required"
                                    } else {
                                        "Investigate and resolve duplicate"
                                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Duplicate computer names analysis completed. $($DuplicateGroups.Count) duplicate groups found." -Type Info -Terminal
        return $Results | Sort-Object CleanName, Status
        
    } catch {
        Write-ADReportLog -Message "Error analyzing duplicate computer names: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersByLocation {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by location..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties Location, Enabled, OperatingSystem, LastLogonDate, Description, DistinguishedName -ErrorAction Stop
        
        # Gruppiere nach Location
        $LocationGroups = $Computers | Group-Object Location
        
        $Results = foreach ($locGroup in $LocationGroups) {
            $location = if ([string]::IsNullOrWhiteSpace($locGroup.Name)) { "(No Location)" } else { $locGroup.Name }
            
            # Statistiken für diese Location
            $enabledCount = @($locGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($locGroup.Group | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # OS-Verteilung in dieser Location
            $osDistribution = $locGroup.Group | Group-Object OperatingSystem | 
                             Sort-Object Count -Descending | 
                             Select-Object -First 3 | 
                             ForEach-Object { "$($_.Name): $($_.Count)" }
            
            # Durchschnittliche Inaktivität
            $inactiveDays = @()
            foreach ($comp in $locGroup.Group) {
                if ($comp.LastLogonDate) {
                    $inactiveDays += (New-TimeSpan -Start $comp.LastLogonDate -End (Get-Date)).Days
                }
            }
            
            $avgInactiveDays = if ($inactiveDays.Count -gt 0) {
                [math]::Round(($inactiveDays | Measure-Object -Average).Average, 1)
            } else { "N/A" }
            
            [PSCustomObject]@{
                Location = $location
                TotalComputers = $locGroup.Count
                EnabledComputers = $enabledCount
                ActiveComputers = $activeCount
                InactiveComputers = $locGroup.Count - $activeCount
                Percentage = [math]::Round(($locGroup.Count / $Computers.Count) * 100, 2)
                TopOperatingSystems = if ($osDistribution) { $osDistribution -join "; " } else { "Unknown" }
                AverageInactiveDays = $avgInactiveDays
                HealthStatus = if ($activeCount -lt ($enabledCount * 0.5)) { "Poor" } 
                              elseif ($activeCount -lt ($enabledCount * 0.8)) { "Fair" } 
                              else { "Good" }
                Recommendation = if ($enabledCount -gt 0 -and $activeCount -lt ($enabledCount * 0.5)) {
                                   "Many inactive computers - review and cleanup"
                               } elseif ([string]::IsNullOrWhiteSpace($locGroup.Name)) {
                                   "Update location information for better tracking"
                               } else {
                                   "Location properly maintained"
                               }
            }
        }
        
        Write-ADReportLog -Message "Computers by location analysis completed. $($Results.Count) locations found." -Type Info -Terminal
        return $Results | Sort-Object TotalComputers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by location: $($_.Exception.Message)" -Type Error
        return @()
    }
}
Function Get-ServiceAccountsSPN {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Starting analysis of accounts with SPNs..." -Type Info -Terminal
        
        # Get ALL accounts with SPNs including system accounts and computer accounts
        $allAccounts = Get-ADObject -Filter {ServicePrincipalName -like '*'} -Properties ServicePrincipalName, Description, ObjectClass, whenCreated, Name, DistinguishedName -ErrorAction Stop
        
        if (-not $allAccounts -or $allAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No accounts with SPNs found - this is highly unusual and indicates a potential AD issue" -Type Warning -Terminal
            return @(
                [PSCustomObject]@{
                    Name = "No SPN accounts found"
                    ObjectType = "N/A" 
                    Description = "No accounts with SPNs found. This is highly unusual as even standard AD accounts like 'krbtgt' should have SPNs."
                    SPNCount = 0
                    SPNs = "None"
                    RiskLevel = "Critical"
                    RiskFactors = "No SPNs found - possible AD configuration issue or query permission problem"
                    Created = "N/A"
                    Location = "N/A"
                }
            )
        }

        # Process each account type separately
        $Results = foreach ($account in $allAccounts) {
            
            # Get additional properties based on object type
            $additionalInfo = switch ($account.ObjectClass) {
                "user" {
                    $userObj = Get-ADUser -Identity $account.DistinguishedName -Properties LastLogonDate, PasswordLastSet, Enabled, Department -ErrorAction SilentlyContinue
                    @{
                        LastLogon = $userObj.LastLogonDate
                        PasswordLastSet = $userObj.PasswordLastSet
                        Enabled = $userObj.Enabled
                        Department = $userObj.Department
                    }
                }
                "computer" {
                    $compObj = Get-ADComputer -Identity $account.DistinguishedName -Properties LastLogonDate, OperatingSystem -ErrorAction SilentlyContinue
                    @{
                        LastLogon = $compObj.LastLogonDate
                        OperatingSystem = $compObj.OperatingSystem
                        Enabled = $compObj.Enabled
                    }
                }
                default { @{} }
            }

            # Risk assessment
            $riskFactors = [System.Collections.Generic.List[string]]@()
            $riskScore = 0

            # Check SPN count
            if ($account.ServicePrincipalName.Count -gt 10) {
                $riskScore += 3
                $riskFactors.Add("Unusually high number of SPNs ($($account.ServicePrincipalName.Count))")
            }

            # Check for known risky SPN patterns
            $riskySpnPatterns = @('*MSSQL*', '*TERMSRV*', '*HTTP*', '*WSMAN*', '*ADMIN*')
            foreach ($pattern in $riskySpnPatterns) {
                if ($account.ServicePrincipalName -like $pattern) {
                    $riskScore += 1
                    $riskFactors.Add("Contains sensitive service type: $pattern")
                }
            }

            # Additional user-specific checks
            if ($account.ObjectClass -eq "user") {
                if ($additionalInfo.LastLogon -and (New-TimeSpan -Start $additionalInfo.LastLogon -End (Get-Date)).Days -gt 90) {
                    $riskScore += 2
                    $riskFactors.Add("No logon in 90+ days")
                }
                if ($additionalInfo.PasswordLastSet -and (New-TimeSpan -Start $additionalInfo.PasswordLastSet -End (Get-Date)).Days -gt 180) {
                    $riskScore += 2
                    $riskFactors.Add("Password not changed in 180+ days")
                }
            }

            # Calculate risk level as string
            $riskLevel = switch ($riskScore) {
                { $_ -ge 5 } { "Critical"; break }
                { $_ -ge 3 } { "High"; break }
                { $_ -ge 1 } { "Medium"; break }
                default { "Low" }
            }

            [PSCustomObject]@{
                Name = $account.Name
                ObjectType = $account.ObjectClass
                Description = if ($account.Description) { $account.Description } else { "No description" }
                SPNCount = $account.ServicePrincipalName.Count
                LastLogon = if ($additionalInfo.LastLogon) { $additionalInfo.LastLogon } else { "N/A" }
                PasswordLastSet = if ($additionalInfo.PasswordLastSet) { $additionalInfo.PasswordLastSet } else { "N/A" }
                Enabled = if ($null -ne $additionalInfo.Enabled) { $additionalInfo.Enabled } else { "N/A" }
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors) { $riskFactors -join ", " } else { "None" }
                Created = $account.whenCreated
                Location = ($account.DistinguishedName -split ',')[1..99] -join ','
                SPNs = $account.ServicePrincipalName -join "; "
            }
        }

        Write-ADReportLog -Message "SPN analysis completed. Found $($Results.Count) accounts with SPNs." -Type Info -Terminal
        return $Results | Sort-Object { 
            switch ($_.RiskLevel) {
                'Critical' { 4 }
                'High' { 3 }
                'Medium' { 2 }
                'Low' { 1 }
                default { 0 }
            }
        } -Descending

    } catch {
        Write-ADReportLog -Message "Error analyzing SPN accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-HighPrivServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing high privileged accounts and service accounts..." -Type Info -Terminal
        
        # Define privileged groups and accounts to check
        $PrivGroups = @(
            # Critical Built-in Groups
            "Domain Admins", "Domänen-Admins",
            "Enterprise Admins", "Organisations-Admins",
            "Schema Admins", "Schema-Admins",
            "Administrators", "Administratoren",
            # High-Risk Built-in Groups  
            "Account Operators", "Konten-Operatoren",
            "Backup Operators", "Sicherungs-Operatoren",
            "Server Operators", "Server-Operatoren",
            "Print Operators", "Druck-Operatoren",
            "Domain Controllers", "Domänencontroller"
        )

        # Get all privileged accounts
        $Results = [System.Collections.Generic.List[PSObject]]@()
        
        # First check built-in admin accounts
        $builtInAccounts = Get-ADUser -Filter {
            (SamAccountName -eq "Administrator") -or 
            (SamAccountName -eq "krbtgt") -or
            (Name -like "*admin*")
        } -Properties Description, LastLogonDate, PasswordLastSet, Enabled, 
            whenCreated, Department, ServicePrincipalName, MemberOf, 
            PasswordNeverExpires, BadLogonCount -ErrorAction Stop

        foreach($account in $builtInAccounts) {
            $riskFactors = [System.Collections.Generic.List[string]]@()
            $riskScore = 0

            # Check risk factors
            if($account.Enabled -eq $false) { 
                $riskFactors.Add("Account disabled")
            }
            if($account.PasswordNeverExpires) {
                $riskScore += 2
                $riskFactors.Add("Password never expires") 
            }
            if($account.BadLogonCount -gt 5) {
                $riskScore += 2
                $riskFactors.Add("High number of bad logon attempts")
            }
            if($account.LastLogonDate -and (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days -gt 90) {
                $riskScore += 2
                $riskFactors.Add("No logon in 90+ days")
            }
            if($account.PasswordLastSet -and (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days -gt 180) {
                $riskScore += 3
                $riskFactors.Add("Password not changed in 180+ days")
            }

            $Results.Add([PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName 
                AccountType = "Built-in Admin Account"
                Department = $account.Department
                Enabled = $account.Enabled
                LastLogonDate = if($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                PasswordLastSet = if($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                PasswordNeverExpires = $account.PasswordNeverExpires
                HasSPN = $account.ServicePrincipalName.Count -gt 0
                RiskLevel = switch($riskScore) {
                    {$_ -ge 5} { "Critical" }
                    {$_ -ge 3} { "High" }
                    {$_ -ge 1} { "Medium" }
                    default { "Low" }
                }
                RiskFactors = if($riskFactors) { $riskFactors -join " | " } else { "None" }
                WhenCreated = $account.whenCreated
                MemberOfGroups = ($account.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace 'CN=' }) -join ' | '
            })
        }

        # Then check group memberships
        foreach($group in $PrivGroups) {
            try {
                $groupMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop |
                    Where-Object { $_.objectClass -eq "user" } |
                    Get-ADUser -Properties Description, LastLogonDate, PasswordLastSet, Enabled,
                        whenCreated, Department, ServicePrincipalName, MemberOf,
                        PasswordNeverExpires, BadLogonCount

                foreach($account in $groupMembers) {
                    # Skip if already processed
                    if($Results.SamAccountName -contains $account.SamAccountName) {
                        continue
                    }

                    $riskFactors = [System.Collections.Generic.List[string]]@()
                    $riskScore = 0

                    # Service account indicators
                    $isServiceAccount = $account.SamAccountName -match "(svc|service|dienst)" -or
                                      $account.Description -match "(service|dienst)" -or
                                      $account.ServicePrincipalName.Count -gt 0

                    if($isServiceAccount) { 
                        $riskScore += 2
                        $riskFactors.Add("Service Account with high privileges")
                    }

                    # Additional risk checks
                    if($account.Enabled -eq $false) {
                        $riskFactors.Add("Account disabled")
                    }
                    if($account.PasswordNeverExpires) {
                        $riskScore += 2
                        $riskFactors.Add("Password never expires")
                    }
                    if($account.BadLogonCount -gt 5) {
                        $riskScore += 2
                        $riskFactors.Add("High number of bad logon attempts")
                    }
                    if($account.LastLogonDate -and (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days -gt 90) {
                        $riskScore += 2
                        $riskFactors.Add("No logon in 90+ days")
                    }
                    if($account.PasswordLastSet -and (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days -gt 180) {
                        $riskScore += 3
                        $riskFactors.Add("Password not changed in 180+ days")
                    }

                    $Results.Add([PSCustomObject]@{
                        Name = $account.Name
                        SamAccountName = $account.SamAccountName
                        AccountType = if($isServiceAccount) { "Service Account" } else { "User Account" }
                        Description = if($account.Description) { $account.Description } else { "No description" }
                        Department = $account.Department
                        Enabled = $account.Enabled
                        LastLogonDate = if($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                        PasswordLastSet = if($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                        PasswordNeverExpires = $account.PasswordNeverExpires
                        HasSPN = $account.ServicePrincipalName.Count -gt 0
                        RiskLevel = switch($riskScore) {
                            {$_ -ge 5} { "Critical" }
                            {$_ -ge 3} { "High" }
                            {$_ -ge 1} { "Medium" }
                            default { "Low" }
                        }
                        RiskFactors = if($riskFactors) { $riskFactors -join " | " } else { "None" }
                        WhenCreated = $account.whenCreated
                        MemberOfGroups = ($account.MemberOf | ForEach-Object { ($_ -split ',')[0] -replace 'CN=' }) -join ' | '
                    })
                }
            }
            catch {
                Write-ADReportLog -Message "Error processing group $group : $($_.Exception.Message)" -Type Warning
                continue
            }
        }

        Write-ADReportLog -Message "High privileged accounts analysis completed. Found $($Results.Count) accounts." -Type Info -Terminal
        return $Results | Sort-Object { 
            switch($_.RiskLevel) {
                'Critical' { 4 }
                'High' { 3 }
                'Medium' { 2 }
                'Low' { 1 }
                default { 0 }
            }
        } -Descending

    }
    catch {
        Write-ADReportLog -Message "Error analyzing high privileged accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ServiceAccountPasswordAge {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Starting comprehensive service account password age analysis..." -Type Info -Terminal
        
        # Separate filters for better error handling and multilingual support
        $serviceFilters = @(
            "SamAccountName -like '*svc*'",
            "SamAccountName -like '*service*'", 
            "SamAccountName -like '*srv*'",
            "SamAccountName -like '*admin*'",
            "SamAccountName -like '*dienst*'", # German
            "SamAccountName -like '*system*'",
            "SamAccountName -like '*backup*'",
            "Description -like '*service*'",
            "Description -like '*dienst*'" # German
        )
        
        $filterString = $serviceFilters -join " -or "
        
        # Get service accounts with extended properties
        $ServiceAccounts = Get-ADUser -Filter $filterString -Properties @(
            'Description',
            'LastLogonDate',
            'PasswordLastSet',
            'Enabled',
            'whenCreated',
            'Department',
            'ServicePrincipalName',
            'PasswordNeverExpires',
            'userAccountControl',
            'memberOf'
        ) -ErrorAction Stop

        if (-not $ServiceAccounts) {
            Write-ADReportLog -Message "No service accounts found with standard filters. Expanding search..." -Type Warning -Terminal
            
            # Fallback search for accounts with service indicators
            $ServiceAccounts = Get-ADUser -Filter {
                (Description -like '*automated*') -or
                (Description -like '*automation*') -or
                (Description -like '*batch*') -or
                (Description -like '*task*') -or
                (ServicePrincipalName -like '*') -or
                (PasswordNeverExpires -eq $true)
            } -Properties @(
                'Description',
                'LastLogonDate',
                'PasswordLastSet',
                'Enabled',
                'whenCreated',
                'Department',
                'ServicePrincipalName',
                'PasswordNeverExpires',
                'userAccountControl',
                'memberOf'
            ) -ErrorAction Stop
        }

        $Results = foreach ($account in $ServiceAccounts) {
            # Calculate password age
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                [math]::Round((New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days, 0)
            } else { 9999 }

            # Security risk assessment
            $securityRisks = [System.Collections.Generic.List[string]]@()
            
            if ($account.PasswordNeverExpires) {
                $securityRisks.Add("Password never expires")
            }
            if ($daysSincePasswordChange -gt 365) {
                $securityRisks.Add("Password not changed in >1 year")
            }
            if ($account.ServicePrincipalName.Count -gt 0) {
                $securityRisks.Add("Kerberoasting risk (SPN)")
            }
            if (-not $account.Enabled) {
                $securityRisks.Add("Account disabled but configured")
            }
            if ($account.memberOf -match "Domain Admins|Enterprise Admins|Schema Admins") {
                $securityRisks.Add("High privileged group member")
            }

            # Risk assessment
            $passwordStatus = if ($daysSincePasswordChange -gt 365) {
                "Critical"
            } elseif ($daysSincePasswordChange -gt 180) {
                "Warning"
            } elseif ($daysSincePasswordChange -gt 90) {
                "Review"
            } else {
                "Good"
            }

            $securityRisk = if ($securityRisks.Count -ge 3) {
                "Critical"
            } elseif ($securityRisks.Count -ge 2) {
                "High"
            } elseif ($securityRisks.Count -ge 1) {
                "Medium"
            } else {
                "Low"
            }

            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Department = if ($account.Department) { $account.Department } else { "Not specified" }
                Enabled = $account.Enabled
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordChange
                PasswordStatus = $passwordStatus.ToString()
                SecurityRisk = $securityRisk.ToString()
                HasSPN = $account.ServicePrincipalName.Count -gt 0
                PasswordNeverExpires = $account.PasswordNeverExpires
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                Created = $account.whenCreated
                SecurityIssues = if ($securityRisks) { $securityRisks -join " | " } else { "None" }
                Recommendation = if ($passwordStatus -eq "Critical") {
                    "Immediate password rotation required"
                } elseif ($passwordStatus -eq "Warning") {
                    "Schedule password change within 30 days"
                } elseif ($passwordStatus -eq "Review") {
                    "Review account usage and policy"
                } else {
                    "Monitor usage patterns"
                }
            }
        }
        
        Write-ADReportLog -Message "Service account analysis completed. Found $($Results.Count) accounts with potential security implications." -Type Info -Terminal
        return $Results | Sort-Object {
            switch($_.SecurityRisk) {
                'Critical' { 4 }
                'High' { 3 }
                'Medium' { 2 }
                'Low' { 1 }
                default { 0 }
            }
        } -Descending
        
    } catch {
        Write-ADReportLog -Message "Error in service account analysis: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UnusedServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Starting comprehensive analysis of unused service accounts..." -Type Info -Terminal

        # Get domain for language detection
        $domain = Get-ADDomain
        $isGermanDomain = $domain.DNSRoot -match "\.de$" -or $domain.NetBIOSName -match "DE$"

        # Build comprehensive filter for service accounts in German and English
        $filterString = @"
            (Enabled -eq `$true) -and (
                (SamAccountName -like '*svc*') -or
                (SamAccountName -like '*service*') -or
                (SamAccountName -like '*srv*') -or
                (SamAccountName -like '*dienst*') -or
                (SamAccountName -like '*konto*') -or
                (SamAccountName -like '*automated*') -or
                (SamAccountName -like '*automation*') -or
                (SamAccountName -like '*auto*') -or
                (SamAccountName -like '*batch*') -or
                (SamAccountName -like '*task*') -or
                (SamAccountName -like '*job*') -or
                (SamAccountName -like '*admin*') -or
                (SamAccountName -like '*adm*') -or
                (SamAccountName -like '*system*') -or
                (SamAccountName -like '*backup*') -or
                (SamAccountName -like '*report*') -or
                (Description -like '*service account*') -or
                (Description -like '*dienstkonto*') -or
                (Description -like '*system account*') -or
                (Description -like '*automated*') -or
                (Description -like '*automation*') -or
                (Description -like '*admin*') -or
                (ServicePrincipalName -like '*')
            )
"@

        # Get all potential service accounts with extended properties
        $ServiceAccounts = Get-ADUser -Filter $filterString -Properties @(
            'Description',
            'LastLogonDate',
            'PasswordLastSet',
            'Enabled',
            'whenCreated',
            'Department',
            'ServicePrincipalName',
            'LastLogonTimestamp',
            'PasswordNeverExpires',
            'userAccountControl'
        ) -ErrorAction Stop

        Write-ADReportLog -Message "Found $($ServiceAccounts.Count) potential service accounts for analysis" -Type Info -Terminal

        $Results = foreach ($account in $ServiceAccounts) {
            # Calculate days since last logon using both LastLogonDate and LastLogonTimestamp
            $lastLogon = if ($account.LastLogonDate) { $account.LastLogonDate } `
                        elseif ($account.LastLogonTimestamp) { [DateTime]::FromFileTime($account.LastLogonTimestamp) } `
                        else { $null }

            $daysSinceLastLogon = if ($lastLogon) { 
                [Math]::Abs((New-TimeSpan -Start $lastLogon -End (Get-Date)).Days)
            } else { 
                9999  # High number for accounts that never logged on
            }

            # Calculate password age
            $daysSincePasswordSet = if ($account.PasswordLastSet) {
                [Math]::Abs((New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days)
            } else {
                9999
            }

            # Determine risk level based on multiple factors
            $riskFactors = [System.Collections.ArrayList]@()
            
            if ($daysSinceLastLogon -gt 365) {
                $riskFactors.Add("No logon for over 1 year") | Out-Null
            } elseif ($daysSinceLastLogon -gt 180) {
                $riskFactors.Add("No logon for over 6 months") | Out-Null
            } elseif ($daysSinceLastLogon -gt 90) {
                $riskFactors.Add("No logon for over 3 months") | Out-Null
            }

            if ($daysSincePasswordSet -gt 365) {
                $riskFactors.Add("Password not changed for over 1 year") | Out-Null
            }

            if ($account.PasswordNeverExpires) {
                $riskFactors.Add("Password never expires") | Out-Null
            }

            if ($account.ServicePrincipalName) {
                $riskFactors.Add("Has SPN configuration") | Out-Null
            }

            # Determine overall risk level
            $riskLevel = switch ($true) {
                { $riskFactors.Count -ge 3 } { "Critical" }
                { $riskFactors.Count -eq 2 } { "High" }
                { $riskFactors.Count -eq 1 } { "Medium" }
                default { "Low" }
            }

            # Create result object with clear English output
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = if ($account.Description) { $account.Description } else { "Not specified" }
                Department = if ($account.Department) { $account.Department } else { "Not specified" }
                Enabled = $account.Enabled
                LastLogon = if ($lastLogon) { $lastLogon } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordSet
                Created = $account.whenCreated
                RiskLevel = $riskLevel  # Explicitly set as string
                RiskFactors = if ($riskFactors.Count -gt 0) { $riskFactors -join " | " } else { "None" }
                Recommendation = switch ($riskLevel) {
                    "Critical" { "Immediate security review required - Consider deactivation" }
                    "High" { "High priority review needed - Verify account necessity" }
                    "Medium" { "Schedule security review - Check account usage" }
                    "Low" { "Monitor account activity" }
                }
            }
        }

        Write-ADReportLog -Message "Unused service accounts analysis completed. Found $($Results.Count) accounts requiring attention." -Type Info -Terminal
        
        # Sort by risk level and days since last logon
        return $Results | Sort-Object {
            switch ($_.RiskLevel) {
                "Critical" { 4 }
                "High" { 3 }
                "Medium" { 2 }
                "Low" { 1 }
                default { 0 }
            }
        }, DaysSinceLastLogon -Descending

    } catch {
        Write-ADReportLog -Message "Error analyzing unused service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}
function Get-UnlinkedGPOs {
    try {
        Write-ADReportLog -Message "Analyzing unlinked GPOs..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            # XML-Report generieren um Links zu prüfen
            $report = Get-GPOReport -Guid $gpo.Id -ReportType XML
            if ($report -notmatch "<LinksTo>") {
                [PSCustomObject]@{
                    Name = $gpo.DisplayName
                    ID = $gpo.Id
                    CreationTime = $gpo.CreationTime
                    ModificationTime = $gpo.ModificationTime
                    Status = $gpo.GpoStatus
                    WMIFilter = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { "None" }
                    Owner = $gpo.Owner
                    Description = $gpo.Description
                    Recommendation = "Review and consider deletion if not needed"
                }
            }
        }
        
        Write-ADReportLog -Message "Unlinked GPOs analysis completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object ModificationTime -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing unlinked GPOs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-EmptyGPOs {
    try {
        Write-ADReportLog -Message "Analyzing empty GPOs..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            $report = [xml](Get-GPOReport -Guid $gpo.Id -ReportType Xml)
            
            # Prüfen ob Computer- und User-Einstellungen leer sind
            if (-not $report.GPO.Computer.ExtensionData -and -not $report.GPO.User.ExtensionData) {
                [PSCustomObject]@{
                    Name = $gpo.DisplayName
                    ID = $gpo.Id
                    CreationTime = $gpo.CreationTime
                    ModificationTime = $gpo.ModificationTime
                    Status = $gpo.GpoStatus
                    WMIFilter = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { "None" }
                    Owner = $gpo.Owner
                    Description = $gpo.Description
                    Recommendation = "Review and consider deletion"
                }
            }
        }
        
        Write-ADReportLog -Message "Empty GPOs analysis completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object ModificationTime -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing empty GPOs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-GPOPermissions {
    try {
        Write-ADReportLog -Message "Analyzing GPO permissions..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            $permissions = Get-GPPermission -Guid $gpo.Id -All
            
            foreach ($perm in $permissions) {
                [PSCustomObject]@{
                    GPOName = $gpo.DisplayName
                    GPOID = $gpo.Id
                    Trustee = $perm.Trustee.Name
                    TrusteeType = $perm.Trustee.SidType
                    Permission = $perm.Permission
                    Inherited = $perm.Inherited
                    RiskLevel = switch ($perm.Permission) {
                        "GpoEditDeleteModifySecurity" { "High" }
                        "GpoEdit" { "Medium" }
                        default { "Low" }
                    }
                    Recommendation = switch ($perm.Permission) {
                        "GpoEditDeleteModifySecurity" { "Review full control permissions" }
                        "GpoEdit" { "Verify edit permissions are required" }
                        default { "Standard permission level" }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "GPO permissions analysis completed. $($Results.Count) permissions found." -Type Info -Terminal
        return $Results | Sort-Object GPOName,RiskLevel
        
    } catch {
        Write-ADReportLog -Message "Error analyzing GPO permissions: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-PasswordPolicySummary {
    try {
        Write-ADReportLog -Message "Loading password policy summary..." -Type Info -Terminal
        
        $Results = @()
        
        # Get Default Domain Policy
        try {
            $defaultPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
            Write-ADReportLog -Message "Default domain policy successfully retrieved." -Type Info -Terminal
        } catch {
            Write-ADReportLog -Message "Error retrieving default domain policy: $($_.Exception.Message)" -Type Warning -Terminal
            $defaultPolicy = $null
        }
        
        # Analyze Default Domain Policy
        if ($defaultPolicy) {
            Write-ADReportLog -Message "Default domain policy found, analyzing..." -Type Info -Terminal
            
            # Debug-Ausgabe für alle verfügbaren Properties
            Write-ADReportLog -Message "Default Policy Properties: MinPasswordLength=$($defaultPolicy.MinPasswordLength), MaxPasswordAge=$($defaultPolicy.MaxPasswordAge), ComplexityEnabled=$($defaultPolicy.ComplexityEnabled)" -Type Info -Terminal
            
            # Sichere Konvertierung der Werte zur Vermeidung von System.Object
            $minPwdLength = if ($defaultPolicy.MinPasswordLength -is [int]) { $defaultPolicy.MinPasswordLength } else { 0 }
            $maxPwdAge = if ($defaultPolicy.MaxPasswordAge -and $defaultPolicy.MaxPasswordAge.TotalDays -gt 0) { [int]$defaultPolicy.MaxPasswordAge.TotalDays } else { 0 }
            $minPwdAge = if ($defaultPolicy.MinPasswordAge -and $defaultPolicy.MinPasswordAge.TotalDays -gt 0) { [int]$defaultPolicy.MinPasswordAge.TotalDays } else { 0 }
            $pwdHistoryCount = if ($defaultPolicy.PasswordHistoryCount -is [int]) { $defaultPolicy.PasswordHistoryCount } else { 0 }
            $complexityEnabled = if ($defaultPolicy.ComplexityEnabled -is [bool]) { $defaultPolicy.ComplexityEnabled } else { $false }
            $reversibleEncryption = if ($defaultPolicy.ReversibleEncryptionEnabled -is [bool]) { $defaultPolicy.ReversibleEncryptionEnabled } else { $false }
            $lockoutThreshold = if ($defaultPolicy.LockoutThreshold -is [int]) { $defaultPolicy.LockoutThreshold } else { 0 }
            $lockoutDuration = if ($defaultPolicy.LockoutDuration -and $defaultPolicy.LockoutDuration.TotalMinutes -gt 0) { [int]$defaultPolicy.LockoutDuration.TotalMinutes } else { 0 }
            $resetCounterAfter = if ($defaultPolicy.LockoutObservationWindow -and $defaultPolicy.LockoutObservationWindow.TotalMinutes -gt 0) { [int]$defaultPolicy.LockoutObservationWindow.TotalMinutes } else { 0 }
            
            # Explizite RiskLevel-Bestimmung
            $riskLevel = "Low"
            if ($minPwdLength -lt 8) {
                $riskLevel = "High"
            } elseif ($minPwdLength -lt 12) {
                $riskLevel = "Medium"
            }
            
            # Explizite Recommendation-Bestimmung
            $recommendation = "Policy meets basic security requirements"
            if ($minPwdLength -lt 8) {
                $recommendation = "Increase minimum password length to at least 12 characters"
            } elseif ($minPwdLength -lt 12) {
                $recommendation = "Consider increasing minimum password length"
            }
            
            $Results += [PSCustomObject]@{
                PolicyName = "Default Domain Policy"
                PolicyType = "Domain Default"
                Scope = "Domain-wide"
                MinPasswordLength = if ($minPwdLength -gt 0) { $minPwdLength.ToString() } else { "Not Set" }
                MaxPasswordAge = if ($maxPwdAge -gt 0) { $maxPwdAge.ToString() } else { "Not Set" }
                MinPasswordAge = if ($minPwdAge -gt 0) { $minPwdAge.ToString() } else { "Not Set" }
                PasswordHistoryCount = if ($pwdHistoryCount -gt 0) { $pwdHistoryCount.ToString() } else { "Not Set" }
                ComplexityEnabled = $complexityEnabled.ToString()
                ReversibleEncryption = $reversibleEncryption.ToString()
                LockoutThreshold = if ($lockoutThreshold -eq 0) { "Not Set" } else { $lockoutThreshold.ToString() }
                LockoutDuration = if ($lockoutDuration -gt 0) { $lockoutDuration.ToString() } else { "Not Set" }
                ResetCounterAfter = if ($resetCounterAfter -gt 0) { $resetCounterAfter.ToString() } else { "Not Set" }
                RiskLevel = $riskLevel
                Recommendation = $recommendation
            }
        } else {
            Write-ADReportLog -Message "No default domain policy found! Creating manual entry..." -Type Warning -Terminal
            
            # Fallback: Versuche direkt über Domain Controller GPO zu gehen
            try {
                $domain = Get-ADDomain -ErrorAction Stop
                $domainDN = $domain.DistinguishedName
                Write-ADReportLog -Message "Domain DN: $domainDN" -Type Info -Terminal
                
                # Default Domain Policy GPO direkt abfragen
                $defaultGPO = Get-GPO -Name "Default Domain Policy" -ErrorAction SilentlyContinue
                if ($defaultGPO) {
                    Write-ADReportLog -Message "Found Default Domain Policy GPO: $($defaultGPO.DisplayName)" -Type Info -Terminal
                    
                    $Results += [PSCustomObject]@{
                        PolicyName = "Default Domain Policy (via GPO)"
                        PolicyType = "Domain Default"
                        Scope = "Domain-wide"
                        MinPasswordLength = "Check via GPO Report"
                        MaxPasswordAge = "Check via GPO Report"
                        MinPasswordAge = "Check via GPO Report"
                        PasswordHistoryCount = "Check via GPO Report"
                        ComplexityEnabled = "Check via GPO Report"
                        ReversibleEncryption = "Check via GPO Report"
                        LockoutThreshold = "Check via GPO Report"
                        LockoutDuration = "Check via GPO Report"
                        ResetCounterAfter = "Check via GPO Report"
                        RiskLevel = "Review Required"
                        Recommendation = "Manually review Default Domain Policy GPO settings"
                    }
                }
            } catch {
                Write-ADReportLog -Message "Error accessing domain information: $($_.Exception.Message)" -Type Error -Terminal
            }
        }

        # Get Fine-Grained Password Policies
        try {
            $fgpps = Get-ADFineGrainedPasswordPolicy -Filter * -Properties * -ErrorAction Stop
            if ($fgpps) {
                Write-ADReportLog -Message "Found $($fgpps.Count) fine-grained password policies" -Type Info -Terminal
                foreach ($fgpp in $fgpps) {
                    try {
                        $appliesTo = @()
                        if ($fgpp.AppliesTo) {
                            foreach ($dn in $fgpp.AppliesTo) {
                                try {
                                    $obj = Get-ADObject $dn -ErrorAction SilentlyContinue
                                    if ($obj) {
                                        $appliesTo += $obj.Name
                                    } else {
                                        $appliesTo += $dn
                                    }
                                } catch {
                                    $appliesTo += $dn
                                }
                            }
                        }
                        
                        # Sichere Konvertierung der FGPP-Werte
                        $fgppMinPwdLength = if ($fgpp.MinPasswordLength -is [int]) { $fgpp.MinPasswordLength } else { 0 }
                        $fgppMaxPwdAge = if ($fgpp.MaxPasswordAge -and $fgpp.MaxPasswordAge.TotalDays -gt 0) { [int]$fgpp.MaxPasswordAge.TotalDays } else { 0 }
                        $fgppMinPwdAge = if ($fgpp.MinPasswordAge -and $fgpp.MinPasswordAge.TotalDays -gt 0) { [int]$fgpp.MinPasswordAge.TotalDays } else { 0 }
                        $fgppPwdHistoryCount = if ($fgpp.PasswordHistoryCount -is [int]) { $fgpp.PasswordHistoryCount } else { 0 }
                        $fgppComplexityEnabled = if ($fgpp.ComplexityEnabled -is [bool]) { $fgpp.ComplexityEnabled } else { $false }
                        $fgppReversibleEncryption = if ($fgpp.ReversibleEncryptionEnabled -is [bool]) { $fgpp.ReversibleEncryptionEnabled } else { $false }
                        $fgppLockoutThreshold = if ($fgpp.LockoutThreshold -is [int]) { $fgpp.LockoutThreshold } else { 0 }
                        $fgppLockoutDuration = if ($fgpp.LockoutDuration -and $fgpp.LockoutDuration.TotalMinutes -gt 0) { [int]$fgpp.LockoutDuration.TotalMinutes } else { 0 }
                        $fgppResetCounterAfter = if ($fgpp.LockoutObservationWindow -and $fgpp.LockoutObservationWindow.TotalMinutes -gt 0) { [int]$fgpp.LockoutObservationWindow.TotalMinutes } else { 0 }
                        
                        # Explizite RiskLevel-Bestimmung für FGPP
                        $fgppRiskLevel = "Low"
                        if ($fgppMinPwdLength -lt 8) {
                            $fgppRiskLevel = "High"
                        } elseif ($fgppMinPwdLength -lt 12) {
                            $fgppRiskLevel = "Medium"
                        }
                        
                        # Explizite Recommendation-Bestimmung für FGPP
                        $fgppRecommendation = "Policy meets basic security requirements"
                        if ($fgppMinPwdLength -lt 8) {
                            $fgppRecommendation = "Increase minimum password length to at least 12 characters"
                        } elseif ($fgppMinPwdLength -lt 12) {
                            $fgppRecommendation = "Consider increasing minimum password length"
                        }
                        
                        $Results += [PSCustomObject]@{
                            PolicyName = $fgpp.Name
                            PolicyType = "Fine-Grained"
                            Scope = if ($appliesTo.Count -gt 0) { $appliesTo -join ", " } else { "Not Applied" }
                            MinPasswordLength = if ($fgppMinPwdLength -gt 0) { $fgppMinPwdLength.ToString() } else { "Not Set" }
                            MaxPasswordAge = if ($fgppMaxPwdAge -gt 0) { $fgppMaxPwdAge.ToString() } else { "Not Set" }
                            MinPasswordAge = if ($fgppMinPwdAge -gt 0) { $fgppMinPwdAge.ToString() } else { "Not Set" }
                            PasswordHistoryCount = if ($fgppPwdHistoryCount -gt 0) { $fgppPwdHistoryCount.ToString() } else { "Not Set" }
                            ComplexityEnabled = $fgppComplexityEnabled.ToString()
                            ReversibleEncryption = $fgppReversibleEncryption.ToString()
                            LockoutThreshold = if ($fgppLockoutThreshold -eq 0) { "Not Set" } else { $fgppLockoutThreshold.ToString() }
                            LockoutDuration = if ($fgppLockoutDuration -gt 0) { $fgppLockoutDuration.ToString() } else { "Not Set" }
                            ResetCounterAfter = if ($fgppResetCounterAfter -gt 0) { $fgppResetCounterAfter.ToString() } else { "Not Set" }
                            RiskLevel = $fgppRiskLevel
                            Recommendation = $fgppRecommendation
                        }
                    } catch {
                        Write-ADReportLog -Message "Error processing FGPP '$($fgpp.Name)': $($_.Exception.Message)" -Type Warning -Terminal
                    }
                }
            } else {
                Write-ADReportLog -Message "No fine-grained password policies found." -Type Info -Terminal
            }
        } catch {
            Write-ADReportLog -Message "Error retrieving fine-grained password policies: $($_.Exception.Message)" -Type Warning -Terminal
        }

        # Get Password Policies from GPOs - erweiterte Suche
        try {
            Write-ADReportLog -Message "Searching for GPOs with password settings..." -Type Info -Terminal
            $allGpos = Get-GPO -All -ErrorAction Stop
            $passwordGpos = @()
            
            foreach ($gpo in $allGpos) {
                try {
                    $report = Get-GPOReport -Guid $gpo.Id -ReportType XML -ErrorAction SilentlyContinue
                    if ($report -and ($report -match "Password" -or $report -match "Lockout")) {
                        $passwordGpos += $gpo
                    }
                } catch {
                    # Ignoriere Fehler bei einzelnen GPOs
                    continue
                }
            }
            
            if ($passwordGpos.Count -gt 0) {
                Write-ADReportLog -Message "Found $($passwordGpos.Count) GPOs with password/lockout settings" -Type Info -Terminal
                foreach ($gpo in $passwordGpos) {
                    try {
                        $report = Get-GPOReport -Guid $gpo.Id -ReportType XML -ErrorAction SilentlyContinue
                        if ($report) {
                            $Results += [PSCustomObject]@{
                                PolicyName = $gpo.DisplayName
                                PolicyType = "GPO"
                                Scope = "GPO-linked OUs/Sites"
                                MinPasswordLength = if ($report -match "MinimumPasswordLength.*?(\d+)") { $matches[1] } else { "Not Set" }
                                MaxPasswordAge = if ($report -match "MaximumPasswordAge.*?(\d+)") { $matches[1] } else { "Not Set" }
                                MinPasswordAge = if ($report -match "MinimumPasswordAge.*?(\d+)") { $matches[1] } else { "Not Set" }
                                PasswordHistoryCount = if ($report -match "PasswordHistorySize.*?(\d+)") { $matches[1] } else { "Not Set" }
                                ComplexityEnabled = if ($report -match "PasswordComplexity.*?(\d+)") { if ($matches[1] -eq "1") { "True" } else { "False" } } else { "Not Set" }
                                ReversibleEncryption = if ($report -match "ClearTextPassword.*?(\d+)") { if ($matches[1] -eq "1") { "True" } else { "False" } } else { "Not Set" }
                                LockoutThreshold = if ($report -match "LockoutBadCount.*?(\d+)") { $matches[1] } else { "Not Set" }
                                LockoutDuration = if ($report -match "LockoutDuration.*?(\d+)") { $matches[1] } else { "Not Set" }
                                ResetCounterAfter = if ($report -match "ResetLockoutCount.*?(\d+)") { $matches[1] } else { "Not Set" }
                                RiskLevel = "Review Required"
                                Recommendation = "Review GPO settings and scope"
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error processing GPO '$($gpo.DisplayName)': $($_.Exception.Message)" -Type Warning -Terminal
                    }
                }
            } else {
                Write-ADReportLog -Message "No GPOs with password settings found." -Type Info -Terminal
            }
        } catch {
            Write-ADReportLog -Message "Error searching GPOs for password settings: $($_.Exception.Message)" -Type Warning -Terminal
        }
        
        # Fallback: Wenn immer noch keine Ergebnisse, erstelle eine Standardmeldung
        if ($Results.Count -eq 0) {
            Write-ADReportLog -Message "No password policies found via standard methods. Creating informational entry." -Type Warning -Terminal
            $Results += [PSCustomObject]@{
                PolicyName = "No Policies Found"
                PolicyType = "Information"
                Scope = "N/A"
                MinPasswordLength = "Unknown"
                MaxPasswordAge = "Unknown"
                MinPasswordAge = "Unknown"
                PasswordHistoryCount = "Unknown"
                ComplexityEnabled = "Unknown"
                ReversibleEncryption = "Unknown"
                LockoutThreshold = "Unknown"
                LockoutDuration = "Unknown"
                ResetCounterAfter = "Unknown"
                RiskLevel = "High"
                Recommendation = "Manual review required - Standard AD queries failed to retrieve password policies"
            }
        } else {
            Write-ADReportLog -Message "Password policy analysis completed. Found $($Results.Count) policies." -Type Info -Terminal
        }
        
        return $Results | Sort-Object PolicyType, PolicyName
        
    } catch {
        Write-ADReportLog -Message "Error analyzing password policies: $($_.Exception.Message)" -Type Error
        # Erstelle Fallback-Ergebnis auch bei Gesamtfehlern
        return @([PSCustomObject]@{
            PolicyName = "Error during analysis"
            PolicyType = "Error"
            Scope = "N/A"
            MinPasswordLength = "Error"
            MaxPasswordAge = "Error"
            MinPasswordAge = "Error"
            PasswordHistoryCount = "Error"
            ComplexityEnabled = "Error"
            ReversibleEncryption = "Error"
            LockoutThreshold = "Error"
            LockoutDuration = "Error"
            ResetCounterAfter = "Error"
            RiskLevel = "Critical"
            Recommendation = "Function failed: $($_.Exception.Message)"
        })
    }
}

function Get-AccountLockoutPolicies {
    try {
        Write-ADReportLog -Message "Analyzing account lockout policies..." -Type Info -Terminal
        
        # Default Domain Policy abrufen
        $defaultPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
        
        $Results = @()
        $Results += [PSCustomObject]@{
            PolicyName = "Default Domain Policy"
            LockoutThreshold = $defaultPolicy.LockoutThreshold
            LockoutDuration = if ($defaultPolicy.LockoutDuration) { $defaultPolicy.LockoutDuration.Minutes } else { 0 }
            ResetCounterAfter = if ($defaultPolicy.LockoutObservationWindow) { $defaultPolicy.LockoutObservationWindow.Minutes } else { 0 }
            RiskLevel = switch ($defaultPolicy.LockoutThreshold) {
                0 { "High" }
                {$_ -gt 10} { "Medium" }
                default { "Low" }
            }
            Recommendation = switch ($defaultPolicy.LockoutThreshold) {
                0 { "Enable account lockout policy" }
                {$_ -gt 10} { "Consider reducing lockout threshold" }
                default { "Policy meets security recommendations" }
            }
        }
        
        Write-ADReportLog -Message "Account lockout policy analysis completed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing account lockout policies: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-FineGrainedPasswordPolicies {
    try {
        Write-ADReportLog -Message "Analyzing fine-grained password policies..." -Type Info -Terminal
        
        # Fine-Grained Password Policies abrufen
        $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -Properties * -ErrorAction Stop
        
        $Results = foreach ($fgpp in $FGPPs) {
            [PSCustomObject]@{
                Name = $fgpp.Name
                Precedence = $fgpp.Precedence
                MinPasswordLength = $fgpp.MinPasswordLength
                MaxPasswordAge = $fgpp.MaxPasswordAge.Days
                MinPasswordAge = $fgpp.MinPasswordAge.Days
                PasswordHistoryCount = $fgpp.PasswordHistoryCount
                ComplexityEnabled = $fgpp.ComplexityEnabled
                ReversibleEncryption = $fgpp.ReversibleEncryptionEnabled
                LockoutThreshold = $fgpp.LockoutThreshold
                LockoutDuration = if ($fgpp.LockoutDuration) { $fgpp.LockoutDuration.Minutes } else { 0 }
                ResetCounterAfter = if ($fgpp.LockoutObservationWindow) { $fgpp.LockoutObservationWindow.Minutes } else { 0 }
                AppliesTo = ($fgpp.AppliesTo | ForEach-Object { (Get-ADObject $_).Name }) -join "; "
                RiskLevel = switch ($fgpp.MinPasswordLength) {
                    {$_ -lt 8} { "High" }
                    {$_ -lt 12} { "Medium" }
                    default { "Low" }
                }
                Recommendation = switch ($fgpp.MinPasswordLength) {
                    {$_ -lt 8} { "Increase minimum password length to at least 12 characters" }
                    {$_ -lt 12} { "Consider increasing minimum password length" }
                    default { "Policy meets basic security requirements" }
                }
            }
        }
        
        Write-ADReportLog -Message "Fine-grained password policies analysis completed. $($Results.Count) policies found." -Type Info -Terminal
        return $Results | Sort-Object Precedence
        
    } catch {
        Write-ADReportLog -Message "Error analyzing fine-grained password policies: $($_.Exception.Message)" -Type Error
        return @()
    }
}
function Get-PrivilegeEscalationPaths {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential privilege escalation paths..." -Type Info -Terminal
        
        $Results = @()
        
        # Get all privileged groups (both German and English names)
        $PrivilegedGroups = @()
        foreach ($groupCategory in $Global:ADGroupNames.GetEnumerator()) {
            foreach ($groupName in $groupCategory.Value) {
                try {
                    $groupObj = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                    if ($groupObj) {
                        $PrivilegedGroups += $groupObj
                        Write-ADReportLog -Message "Found privileged group: $($groupObj.Name)" -Type Info
                    }
                } catch {
                    # Continue with next group name
                }
            }
        }
        
        if ($PrivilegedGroups.Count -eq 0) {
            Write-ADReportLog -Message "No privileged groups found for analysis." -Type Warning -Terminal
            return @([PSCustomObject]@{
                TreeStructure = [string]"⚠️ PRIVILEGED GROUPS ANALYSIS"
                PathType = [string]"Analysis Status"
                ObjectName = [string]"No Privileged Groups Found"
                ObjectType = [string]"Warning"
                Status = [string]"Incomplete"
                RiskLevel = [string]"Medium"
                SecurityImpact = [string]"Cannot analyze privilege escalation paths - no privileged groups detected"
                LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
                AccountAge = [string]"N/A"
                PathDetails = [string]"Analysis requires access to privileged groups"
                Recommendation = [string]"Verify domain connectivity and user permissions"
                ComplianceNote = [string]"Privilege escalation analysis incomplete"
            })
        }
        
        # Analyze each privileged group for dangerous permissions
        foreach ($group in $PrivilegedGroups) {
            Write-ADReportLog -Message "Analyzing access rights for group: $($group.Name)" -Type Info
            
            try {
                $acl = Get-Acl -Path "AD:\$($group.DistinguishedName)" -ErrorAction Stop
                
                # Find dangerous permissions
                $dangerousPermissions = $acl.Access | Where-Object {
                    ($_.ActiveDirectoryRights -match "WriteProperty|GenericWrite|WriteDacl|WriteOwner|GenericAll|ExtendedRight") -and
                    ($_.AccessControlType -eq "Allow") -and
                    ($_.IdentityReference.Value -notmatch "SYSTEM|NT AUTHORITY|BUILTIN|S-1-5-32") -and
                    ($_.IdentityReference.Value -notmatch "Domain Admins|Enterprise Admins|Schema Admins|Administrators") -and
                    ($_.IdentityReference.Value -notmatch "Domänen-Admins|Organisations-Admins|Schema-Admins|Administratoren")
                }
                
                foreach ($permission in $dangerousPermissions) {
                    try {
                        $identityName = $permission.IdentityReference.Value
                        if ($identityName -like "*\*") {
                            $identityName = $identityName.Split('\')[-1]
                        }
                        
                        # Try to resolve the identity
                        $identityObject = $null
                        try {
                            $identityObject = Get-ADUser -Filter "SamAccountName -eq '$identityName' -or Name -eq '$identityName'" -Properties Department, Description, LastLogonDate, whenCreated, memberOf, AdminCount -ErrorAction SilentlyContinue
                        } catch {
                            # Might be a group
                            try {
                                $identityObject = Get-ADGroup -Filter "SamAccountName -eq '$identityName' -or Name -eq '$identityName'" -Properties Description, whenCreated, memberOf -ErrorAction SilentlyContinue
                            } catch {
                                # Continue with unknown object
                            }
                        }
                        
                        if ($identityObject) {
                            $objectType = if ($identityObject.ObjectClass -eq "user") { "User Account" } else { "Security Group" }
                            $isPrivileged = if ($identityObject.AdminCount -eq 1) { "Yes" } else { "No" }
                            
                            # Calculate risk level
                            $riskLevel = "High"
                            if ($permission.ActiveDirectoryRights -match "GenericAll|WriteDacl|WriteOwner") {
                                $riskLevel = "Critical"
                            } elseif ($permission.ActiveDirectoryRights -match "WriteProperty|GenericWrite") {
                                $riskLevel = "High"
                            } else {
                                $riskLevel = "Medium"
                            }
                            
                            # Create security impact description
                            $securityImpact = "Can modify $($group.Name) group membership or properties"
                            if ($permission.ActiveDirectoryRights -match "GenericAll") {
                                $securityImpact = "Has full control over $($group.Name) group - can add/remove members and modify all properties"
                            } elseif ($permission.ActiveDirectoryRights -match "WriteDacl") {
                                $securityImpact = "Can modify security permissions on $($group.Name) group"
                            } elseif ($permission.ActiveDirectoryRights -match "WriteOwner") {
                                $securityImpact = "Can take ownership of $($group.Name) group"
                            }
                            
                            $Results += [PSCustomObject]@{
                                TreeStructure = [string]"🚨 PRIVILEGE ESCALATION PATHS > $($group.Name) > $($identityObject.Name)"
                                PathType = [string]"Dangerous Permission"
                                ObjectName = [string]$identityObject.Name
                                ObjectType = [string]$objectType
                                Status = if ($identityObject.Enabled -eq $false) { [string]"Disabled" } else { [string]"Active" }
                                RiskLevel = [string]$riskLevel
                                SecurityImpact = [string]$securityImpact
                                LastActivity = if ($identityObject.LastLogonDate) { [string]$identityObject.LastLogonDate.ToString("MM/dd/yyyy HH:mm") } else { [string]"Never" }
                                AccountAge = if ($identityObject.whenCreated) { [string]"$([math]::Round((Get-Date - $identityObject.whenCreated).TotalDays)) days" } else { [string]"Unknown" }
                                PathDetails = [string]"Permission: $($permission.ActiveDirectoryRights) | Target: $($group.Name) | Privileged Account: $isPrivileged"
                                Recommendation = [string]"Remove dangerous permissions or justify business need with documentation"
                                ComplianceNote = [string]"Non-standard permissions on privileged group require review and approval"
                            }
                        } else {
                            # Unknown identity - still report it
                            $Results += [PSCustomObject]@{
                                TreeStructure = [string]"🚨 PRIVILEGE ESCALATION PATHS > $($group.Name) > Unknown Identity"
                                PathType = [string]"Unknown Permission"
                                ObjectName = [string]$identityName
                                ObjectType = [string]"Unknown Object"
                                Status = [string]"Unknown"
                                RiskLevel = [string]"Medium"
                                SecurityImpact = [string]"Unknown identity has permissions on privileged group $($group.Name)"
                                LastActivity = [string]"Unknown"
                                AccountAge = [string]"Unknown"
                                PathDetails = [string]"Permission: $($permission.ActiveDirectoryRights) | Target: $($group.Name) | Identity could not be resolved"
                                Recommendation = [string]"Investigate and remove unknown identity permissions"
                                ComplianceNote = [string]"Unresolved security principals pose security risks"
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error processing permission for identity $($permission.IdentityReference.Value): $($_.Exception.Message)" -Type Warning
                    }
                }
                
            } catch {
                Write-ADReportLog -Message "Error analyzing ACL for group $($group.Name): $($_.Exception.Message)" -Type Warning
                
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"❌ PRIVILEGE ESCALATION PATHS > $($group.Name) > Analysis Error"
                    PathType = [string]"Access Error"
                    ObjectName = [string]$group.Name
                    ObjectType = [string]"Privileged Group"
                    Status = [string]"Analysis Failed"
                    RiskLevel = [string]"Medium"
                    SecurityImpact = [string]"Cannot analyze permissions on privileged group - security status unknown"
                    LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
                    AccountAge = [string]"N/A"
                    PathDetails = [string]"Error: $($_.Exception.Message)"
                    Recommendation = [string]"Verify permissions and retry analysis"
                    ComplianceNote = [string]"Failed security analysis requires investigation"
                }
            }
        }
        
        # Check for orphaned SIDs with dangerous permissions
        Write-ADReportLog -Message "Checking for orphaned SIDs with dangerous permissions..." -Type Info
        try {
            foreach ($group in $PrivilegedGroups) {
                $acl = Get-Acl -Path "AD:\$($group.DistinguishedName)" -ErrorAction SilentlyContinue
                if ($acl) {
                    $orphanedSids = $acl.Access | Where-Object {
                        $_.IdentityReference.Value -match "S-1-5-21-\d+-\d+-\d+-\d+" -and
                        ($_.ActiveDirectoryRights -match "WriteProperty|GenericWrite|WriteDacl|WriteOwner|GenericAll") -and
                        ($_.AccessControlType -eq "Allow")
                    }
                    
                    foreach ($orphan in $orphanedSids) {
                        $Results += [PSCustomObject]@{
                            TreeStructure = [string]"💀 PRIVILEGE ESCALATION PATHS > $($group.Name) > Orphaned SID"
                            PathType = [string]"Orphaned Permission"
                            ObjectName = [string]$orphan.IdentityReference.Value
                            ObjectType = [string]"Orphaned SID"
                            Status = [string]"Orphaned"
                            RiskLevel = [string]"High"
                            SecurityImpact = [string]"Orphaned SID with dangerous permissions on privileged group"
                            LastActivity = [string]"Unknown"
                            AccountAge = [string]"Unknown"
                            PathDetails = [string]"Permission: $($orphan.ActiveDirectoryRights) | Target: $($group.Name) | SID no longer resolvable"
                            Recommendation = [string]"Remove orphaned SID permissions immediately"
                            ComplianceNote = [string]"Orphaned SIDs create security vulnerabilities and should be cleaned up"
                        }
                    }
                }
            }
        } catch {
            Write-ADReportLog -Message "Error checking for orphaned SIDs: $($_.Exception.Message)" -Type Warning
        }
        
        # Add summary information
        $totalEscalationPaths = ($Results | Where-Object { $_.PathType -ne "Summary" }).Count
        $criticalPaths = ($Results | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highRiskPaths = ($Results | Where-Object { $_.RiskLevel -eq "High" }).Count
        
        $Results += [PSCustomObject]@{
            TreeStructure = [string]"📊 PRIVILEGE ESCALATION ANALYSIS SUMMARY"
            PathType = [string]"Summary"
            ObjectName = [string]"Privilege Escalation Security Assessment"
            ObjectType = [string]"Report Summary"
            Status = [string]"Analysis Complete"
            RiskLevel = if ($criticalPaths -gt 0) { [string]"Critical" } elseif ($highRiskPaths -gt 0) { [string]"High" } else { [string]"Low" }
            SecurityImpact = [string]"Total Escalation Paths: $totalEscalationPaths | Critical: $criticalPaths | High Risk: $highRiskPaths"
            LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
            AccountAge = [string]"N/A"
            PathDetails = [string]"Analysis of $($PrivilegedGroups.Count) privileged groups completed"
            Recommendation = if ($totalEscalationPaths -gt 0) { 
                [string]"Review and remediate all identified privilege escalation paths" 
            } else { 
                [string]"No privilege escalation paths detected - maintain current security posture" 
            }
            ComplianceNote = [string]"Regular privilege escalation analysis is essential for maintaining security"
        }
        
        Write-ADReportLog -Message "Privilege escalation paths analysis completed. $($Results.Count) findings generated." -Type Info -Terminal
        
        # Log security findings
        if ($criticalPaths -gt 0) {
            Write-ADReportLog -Message "SECURITY ALERT: $criticalPaths critical privilege escalation paths found" -Type Warning -Terminal
        }
        if ($highRiskPaths -gt 0) {
            Write-ADReportLog -Message "SECURITY WARNING: $highRiskPaths high-risk privilege escalation paths found" -Type Warning -Terminal
        }
        
        return $Results | Sort-Object @{Expression={
            switch($_.RiskLevel) {
                "Critical" { 1 }
                "High" { 2 }
                "Medium" { 3 }
                "Low" { 4 }
                default { 5 }
            }
        }}, TreeStructure
        
    } catch {
        $ErrorMessage = "Error analyzing privilege escalation paths: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        return @([PSCustomObject]@{
            TreeStructure = [string]"❌ PRIVILEGE ESCALATION ANALYSIS ERROR"
            PathType = [string]"Error"
            ObjectName = [string]"Privilege Escalation Analysis"
            ObjectType = [string]"Error Report"
            Status = [string]"Failed"
            RiskLevel = [string]"Unknown"
            SecurityImpact = [string]"Analysis failed - cannot assess privilege escalation risks"
            LastActivity = [string](Get-Date).ToString("MM/dd/yyyy HH:mm")
            AccountAge = [string]"N/A"
            PathDetails = [string]"Error: $ErrorMessage"
            Recommendation = [string]"Retry analysis with appropriate permissions"
            ComplianceNote = [string]"Failed security analysis requires immediate attention"
        })
    }
}
# End of Selection

function Get-ExposedCredentials {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potentially exposed credentials..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit reversible Verschlüsselung
        $reversibleUsers = Get-ADUser -Filter {AllowReversiblePasswordEncryption -eq $true} -Properties Department,Description,LastLogonDate,whenCreated
        foreach ($user in $reversibleUsers) {
            $Results += [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                ExposureType = "Reversible Encryption"
                RiskLevel = "High"
                Description = $user.Description
                LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                Created = $user.whenCreated
                Recommendation = "Disable reversible password encryption"
            }
        }
        
        # Benutzer mit Kerberos DES
        $desUsers = Get-ADUser -Filter {UseDESKeyOnly -eq $true} -Properties Department,Description,LastLogonDate,whenCreated
        foreach ($user in $desUsers) {
            $Results += [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                ExposureType = "DES Encryption"
                RiskLevel = "High"
                Description = $user.Description
                LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                Created = $user.whenCreated
                Recommendation = "Disable DES encryption usage"
            }
        }
        
        # Ensure we always return an array
        if ($Results.Count -eq 0) {
            Write-ADReportLog -Message "No exposed credentials found." -Type Info -Terminal
            return @()
        }
        
        Write-ADReportLog -Message "Exposed credentials analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return @($Results | Sort-Object RiskLevel,ExposureType)
        
    } catch {
        Write-ADReportLog -Message "Error analyzing exposed credentials: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-SuspiciousLogons {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing suspicious logon patterns..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit ungewöhnlichen Anmeldezeiten (außerhalb 6-20 Uhr)
        $users = Get-ADUser -Filter * -Properties Department,Description,LastLogonDate,logonHours,whenCreated
        foreach ($user in $users) {
            if ($user.logonHours) {
                $logonHours = [System.BitConverter]::ToString($user.logonHours)
                if ($logonHours -match "FF") {
                    $Results += [PSCustomObject]@{
                        Name = $user.Name
                        SamAccountName = $user.SamAccountName
                        Department = $user.Department
                        Pattern = "24/7 Logon Hours"
                        RiskLevel = "Medium"
                        Description = $user.Description
                        LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                        Created = $user.whenCreated
                        Recommendation = "Review logon hour restrictions"
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Suspicious logon analysis completed. $($Results.Count) patterns found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel,Pattern
        
    } catch {
        Write-ADReportLog -Message "Error analyzing suspicious logons: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-ForeignSecurityPrincipals {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing foreign security principals..." -Type Info -Terminal
        
        $FSPs = Get-ADObject -Filter {objectClass -eq "foreignSecurityPrincipal"} -Properties memberOf,whenCreated
        
        $Results = foreach ($fsp in $FSPs) {
            $sid = $fsp.Name
            $memberOfGroups = $fsp.memberOf | ForEach-Object { (Get-ADObject $_).Name }
            
            # Versuche den Kontonamen aus der SID zu ermitteln
            try {
                $account = New-Object System.Security.Principal.SecurityIdentifier($sid)
                $accountName = $account.Translate([System.Security.Principal.NTAccount]).Value
            }
            catch {
                $accountName = "Unknown Account"
            }
            
            [PSCustomObject]@{
                Name = $accountName 
                ObjectSID = $sid
                MemberOfGroups = $memberOfGroups -join "; "
                Created = $fsp.whenCreated
                RiskLevel = if ($memberOfGroups -match "Admin|Schema|Enterprise") { "High" } else { "Medium" }
                Recommendation = "Review foreign security principal memberships"
            }
        }
        
        Write-ADReportLog -Message "Foreign security principals analysis completed. $($Results.Count) principals found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel
        
    } catch {
        Write-ADReportLog -Message "Error analyzing foreign security principals: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-SIDHistoryAbuse {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential SID history abuse..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit SID-History
        $users = Get-ADUser -Filter * -Properties SIDHistory,Department,Description,LastLogonDate,whenCreated
        foreach ($user in $users) {
            if ($user.SIDHistory) {
                $Results += [PSCustomObject]@{
                    Name = $user.Name
                    SamAccountName = $user.SamAccountName
                    Department = $user.Department
                    SIDHistoryCount = $user.SIDHistory.Count
                    RiskLevel = "High"
                    Description = $user.Description
                    LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                    Created = $user.whenCreated
                    Recommendation = "Review and clean up SID history"
                }
            }
        }
        
        Write-ADReportLog -Message "SID history abuse analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object SIDHistoryCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing SID history abuse: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-ACLAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing ACL permissions..." -Type Info -Terminal
        
        $Results = @()
        $DomainDN = (Get-ADDomain).DistinguishedName
        
        # Get critical AD containers for tree structure
        $CriticalContainers = @(
            "CN=Configuration,$((Get-ADRootDSE).configurationNamingContext)",
            "CN=Schema,$((Get-ADRootDSE).schemaNamingContext)",
            "CN=System,$DomainDN",
            "CN=AdminSDHolder,CN=System,$DomainDN",
            "CN=Users,$DomainDN",
            "CN=Computers,$DomainDN",
            "OU=Domain Controllers,$DomainDN"
        )
        
        # Get all AD objects with ACLs from critical containers
        $objects = @()
        foreach ($container in $CriticalContainers) {
            try {
                $containerObjects = Get-ADObject -SearchBase $container -Filter * -Properties nTSecurityDescriptor,distinguishedName -ErrorAction SilentlyContinue
                $objects += $containerObjects
            } catch {
                # Continue if container doesn't exist
            }
        }
        
        # Group objects by their parent containers for tree structure
        $ACLTree = @{}
        
        foreach ($object in $objects) {
            try {
                $acl = $object.nTSecurityDescriptor
                
                # Skip if no ACL
                if (-not $acl) { continue }
                
                # Determine the object's position in tree
                $dn = $object.DistinguishedName
                $pathParts = $dn -split ','
                $containerPath = ""
                
                # Build hierarchical path
                if ($dn -like "*CN=Configuration*") {
                    $containerPath = "└── Configuration Container"
                } elseif ($dn -like "*CN=Schema*") {
                    $containerPath = "└── Schema Container"
                } elseif ($dn -like "*CN=System*") {
                    $containerPath = "└── System Container"
                } elseif ($dn -like "*CN=Users*") {
                    $containerPath = "└── Users Container"
                } elseif ($dn -like "*CN=Computers*") {
                    $containerPath = "└── Computers Container"
                } elseif ($dn -like "*OU=Domain Controllers*") {
                    $containerPath = "└── Domain Controllers OU"
                } else {
                    $containerPath = "└── Domain Root"
                }
                
                # Check for unusual ACL entries
                foreach ($ace in $acl.Access) {
                    if ($ace.AccessControlType -eq "Allow" -and 
                        ($ace.ActiveDirectoryRights -match "GenericAll|WriteDacl|WriteOwner|CreateChild|DeleteChild")) {
                        
                        # Determine risk level based on rights and object type
                        $riskLevel = "Medium"
                        if ($ace.ActiveDirectoryRights -match "GenericAll|WriteDacl|WriteOwner") {
                            $riskLevel = "High"
                        }
                        if ($object.ObjectClass -eq "organizationalUnit" -or $object.ObjectClass -eq "container") {
                            $riskLevel = "Critical"
                        }
                        
                        # Build tree-like display name
                        $treeDisplayName = ""
                        if ($object.Name -eq $object.DistinguishedName.Split(',')[0].Replace('CN=','').Replace('OU=','')) {
                            $treeDisplayName = "$containerPath"
                        } else {
                            $treeDisplayName = "$containerPath`n    ├── $($object.Name)"
                        }
                        
                        # Enhanced recommendation based on context
                        $recommendation = ""
                        switch ($ace.ActiveDirectoryRights) {
                            { $_ -match "GenericAll" } { $recommendation = "Full Control - Review necessity and scope" }
                            { $_ -match "WriteDacl" } { $recommendation = "Can modify permissions - High security risk" }
                            { $_ -match "WriteOwner" } { $recommendation = "Can take ownership - Critical security risk" }
                            { $_ -match "CreateChild" } { $recommendation = "Can create objects - Monitor for abuse" }
                            { $_ -match "DeleteChild" } { $recommendation = "Can delete objects - Review and restrict" }
                            default { $recommendation = "Review permission necessity" }
                        }
                        
                        $Results += [PSCustomObject]@{
                            TreeStructure = [string]$treeDisplayName
                            ObjectName = [string]$object.Name
                            ObjectClass = [string]$object.ObjectClass
                            ContainerPath = [string]$containerPath.Replace('└── ','')
                            TrusteeName = [string]$ace.IdentityReference.ToString()
                            PermissionType = [string]$ace.ActiveDirectoryRights.ToString()
                            AccessType = [string]$ace.AccessControlType.ToString()
                            InheritanceType = [string]$ace.InheritanceType.ToString()
                            RiskLevel = [string]$riskLevel
                            SecurityImpact = [string]"Elevated permissions may allow unauthorized modifications to AD structure"
                            Recommendation = [string]$recommendation
                            DistinguishedName = [string]$object.DistinguishedName
                        }
                    }
                }
            } catch {
                # Skip objects that can't be processed
                continue
            }
        }
        
        Write-ADReportLog -Message "ACL analysis completed. $($Results.Count) suspicious entries found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel,ContainerPath,ObjectName
        
    } catch {
        Write-ADReportLog -Message "Error analyzing ACLs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-InheritanceBreaks {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing inheritance breaks..." -Type Info -Terminal
        
        $Results = @()
        $ProcessedCount = 0
        
        # Get domain information for proper tree structure
        $Domain = Get-ADDomain
        $DomainDN = $Domain.DistinguishedName
        
        # Find objects with inheritance protection enabled
        Write-ADReportLog -Message "Scanning AD objects for inheritance breaks..." -Type Info -Terminal
        $Objects = Get-ADObject -Filter * -Properties nTSecurityDescriptor,distinguishedName,objectClass,whenCreated -ErrorAction SilentlyContinue
        
        foreach ($Object in $Objects) {
            $ProcessedCount++
            if ($ProcessedCount % 1000 -eq 0) {
                Write-ADReportLog -Message "Processed $ProcessedCount objects..." -Type Info
            }
            
            try {
                $ACL = $Object.nTSecurityDescriptor
                
                # Check if inheritance is broken (AreAccessRulesProtected = True means inheritance is disabled)
                if ($ACL.AreAccessRulesProtected -eq $true) {
                    
                    # Build hierarchical path for tree structure
                    $DNParts = $Object.DistinguishedName -split ','
                    $ContainerPath = ""
                    $TreeStructure = ""
                    
                    # Create tree-like visualization
                    $PathElements = @()
                    for ($i = $DNParts.Count - 1; $i -ge 0; $i--) {
                        $Part = $DNParts[$i].Trim()
                        if ($Part -match '^(CN|OU|DC)=(.+)') {
                            $PathElements += $matches[2]
                        }
                    }
                    
                    # Build tree structure
                    if ($PathElements.Count -gt 1) {
                        $TreeStructure = $PathElements[0]
                        for ($j = 1; $j -lt $PathElements.Count - 1; $j++) {
                            $TreeStructure += "`n" + ("│  " * $j) + "├── " + $PathElements[$j]
                        }
                        $TreeStructure += "`n" + ("│  " * ($PathElements.Count - 1)) + "└── " + $PathElements[-1]
                        $ContainerPath = ($PathElements[0..($PathElements.Count-2)] -join " → ")
                    } else {
                        $TreeStructure = $Object.Name
                        $ContainerPath = "Root"
                    }
                    
                    # Determine risk level based on object type and location
                    $RiskLevel = ""
                    $SecurityImpact = ""
                    $Recommendation = ""
                    
                    switch ($Object.ObjectClass) {
                        "organizationalUnit" {
                            $RiskLevel = "High"
                            $SecurityImpact = "OU inheritance break can affect all child objects and their security"
                            $Recommendation = "Review OU inheritance settings - may block security policy propagation"
                        }
                        "container" {
                            $RiskLevel = "Medium"
                            $SecurityImpact = "Container inheritance break affects child object security inheritance"
                            $Recommendation = "Verify container inheritance is intentionally disabled"
                        }
                        "user" {
                            $RiskLevel = "Low"
                            $SecurityImpact = "User object has custom security settings that override inheritance"
                            $Recommendation = "Review custom permissions - ensure they follow security standards"
                        }
                        "group" {
                            $RiskLevel = "Medium"
                            $SecurityImpact = "Group object has non-inherited permissions that may affect access control"
                            $Recommendation = "Validate group permissions are properly configured"
                        }
                        "computer" {
                            $RiskLevel = "Low"
                            $SecurityImpact = "Computer object has custom security configuration"
                            $Recommendation = "Verify computer account permissions are appropriate"
                        }
                        default {
                            $RiskLevel = "Medium"
                            $SecurityImpact = "Object has inheritance disabled - custom security configuration active"
                            $Recommendation = "Review inheritance break necessity and security implications"
                        }
                    }
                    
                    # Special high-risk cases
                    if ($Object.DistinguishedName -match "CN=Users,DC=" -or 
                        $Object.DistinguishedName -match "CN=Computers,DC=" -or
                        $Object.DistinguishedName -match "OU=Domain Controllers,DC=") {
                        $RiskLevel = "Critical"
                        $SecurityImpact = "Inheritance break in critical system container affects domain security"
                        $Recommendation = "URGENT: Review inheritance break in system container - may compromise domain security"
                    }
                    
                    # Calculate age for additional context
                    $ObjectAge = ""
                    if ($Object.whenCreated) {
                        $AgeInDays = [int][math]::Round((New-TimeSpan -Start $Object.whenCreated -End (Get-Date)).TotalDays)
                        $ObjectAge = "$AgeInDays days old"
                    } else {
                        $ObjectAge = "Unknown age"
                    }
                    
                    $Results += [PSCustomObject]@{
                        TreeStructure = [string]$TreeStructure
                        ObjectName = [string]$Object.Name
                        ObjectClass = [string]$Object.ObjectClass
                        ContainerPath = [string]$ContainerPath
                        InheritanceStatus = [string]"Inheritance Disabled"
                        RiskLevel = [string]$RiskLevel
                        SecurityImpact = [string]$SecurityImpact
                        Recommendation = [string]$Recommendation
                        ObjectAge = [string]$ObjectAge
                        DistinguishedName = [string]$Object.DistinguishedName
                        WhenCreated = if ($Object.whenCreated) { [string]$Object.whenCreated.ToString("MM/dd/yyyy HH:mm") } else { [string]"Unknown" }
                    }
                }
            } catch {
                # Skip objects that can't be processed (access denied, etc.)
                continue
            }
        }
        
        # Sort results by risk level and container path for better overview
        $SortedResults = $Results | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, ContainerPath, ObjectName
        
        Write-ADReportLog -Message "Inheritance break analysis completed. $($Results.Count) objects with inheritance breaks found." -Type Info -Terminal
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing inheritance breaks: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @()
    }
}

function Get-AdminSDHolderObjects {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AdminSDHolder protected objects..." -Type Info -Terminal
        
        $Results = @()
        $Domain = Get-ADDomain
        $DomainDN = $Domain.DistinguishedName
        
        # Get AdminSDHolder container for reference
        $AdminSDHolderDN = "CN=AdminSDHolder,CN=System,$DomainDN"
        
        # Find all objects with adminCount=1 (AdminSDHolder protected)
        try {
            Write-ADReportLog -Message "Searching for AdminSDHolder protected objects (adminCount=1)..." -Type Info
            $protectedObjects = Get-ADObject -LDAPFilter "(adminCount=1)" -Properties adminCount,whenCreated,whenChanged,ObjectClass,Description,CanonicalName -ErrorAction Stop
            
            Write-ADReportLog -Message "Found $($protectedObjects.Count) AdminSDHolder protected objects." -Type Info
            
            # Categorize objects for tree structure
            $UserObjects = @()
            $GroupObjects = @()
            $ComputerObjects = @()
            $OtherObjects = @()
            
            foreach ($object in $protectedObjects) {
                try {
                    # Determine object age
                    $ObjectAge = "Unknown"
                    if ($object.whenCreated) {
                        $AgeInDays = [int][math]::Round((New-TimeSpan -Start $object.whenCreated -End (Get-Date)).TotalDays)
                        $ObjectAge = "$AgeInDays days old"
                    }
                    
                    # Determine last modified
                    $LastModified = "Unknown"
                    if ($object.whenChanged) {
                        $LastModified = $object.whenChanged.ToString("MM/dd/yyyy HH:mm")
                    }
                    
                    # Determine container path for context
                    $ContainerPath = "Unknown"
                    if ($object.CanonicalName) {
                        $pathParts = $object.CanonicalName -split '/'
                        if ($pathParts.Count -gt 1) {
                            $ContainerPath = ($pathParts[0..($pathParts.Count-2)] -join '/')
                        }
                    }
                    
                    # Determine risk level based on object class and properties
                    $RiskLevel = "Medium"
                    $SecurityImpact = "Standard AdminSDHolder protection"
                    $Recommendation = "Verify if AdminSDHolder protection is still required"
                    
                    # Higher risk for certain object types
                    switch ($object.ObjectClass) {
                        "user" {
                            $RiskLevel = "High"
                            $SecurityImpact = "User account with administrative privileges protection"
                            $Recommendation = "Review user's administrative roles and necessity of protection"
                        }
                        "computer" {
                            $RiskLevel = "Medium"
                            $SecurityImpact = "Computer account with elevated privileges"
                            $Recommendation = "Verify computer's role requires AdminSDHolder protection"
                        }
                        "group" {
                            $RiskLevel = "Critical"
                            $SecurityImpact = "Administrative group with inherited permissions control"
                            $Recommendation = "Audit group membership and ensure proper administrative oversight"
                        }
                    }
                    
                    # Create object for categorization
                    $ObjectInfo = [PSCustomObject]@{
                        TreeStructure = [string]""  # Will be set below
                        ObjectName = [string]$object.Name
                        ObjectClass = [string]$object.ObjectClass
                        ContainerPath = [string]$ContainerPath
                        AdminCount = [string]"1 (Protected)"
                        RiskLevel = [string]$RiskLevel
                        SecurityImpact = [string]$SecurityImpact
                        Recommendation = [string]$Recommendation
                        ObjectAge = [string]$ObjectAge
                        LastModified = [string]$LastModified
                        DistinguishedName = [string]$object.DistinguishedName
                        WhenCreated = [string]$(if ($object.whenCreated) { $object.whenCreated.ToString("MM/dd/yyyy HH:mm") } else { "Unknown" })
                    }
                    
                    # Categorize for tree structure
                    switch ($object.ObjectClass) {
                        "user" { $UserObjects += $ObjectInfo }
                        "group" { $GroupObjects += $ObjectInfo }
                        "computer" { $ComputerObjects += $ObjectInfo }
                        default { $OtherObjects += $ObjectInfo }
                    }
                    
                } catch {
                    Write-ADReportLog -Message "Error processing object $($object.Name): $($_.Exception.Message)" -Type Warning
                    continue
                }
            }
            
            # Build tree structure results
            if ($UserObjects.Count -gt 0) {
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                    ObjectName = [string]"User Accounts ($($UserObjects.Count))"
                    ObjectClass = [string]"Category"
                    ContainerPath = [string]"Various Containers"
                    AdminCount = [string]"Protected"
                    RiskLevel = [string]"High"
                    SecurityImpact = [string]"Administrative user accounts with inheritance protection"
                    Recommendation = [string]"Review all protected user accounts for necessity"
                    ObjectAge = [string]"Various"
                    LastModified = [string]"Various"
                    Description = [string]"User accounts protected by AdminSDHolder"
                    DistinguishedName = [string]"Multiple Objects"
                    WhenCreated = [string]"Various"
                }
                
                foreach ($userObj in ($UserObjects | Sort-Object ObjectName)) {
                    $userObj.TreeStructure = [string]"├── User Account"
                    $Results += $userObj
                }
            }
            
            if ($GroupObjects.Count -gt 0) {
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                    ObjectName = [string]"Administrative Groups ($($GroupObjects.Count))"
                    ObjectClass = [string]"Category"
                    ContainerPath = [string]"Various Containers"
                    AdminCount = [string]"Protected"
                    RiskLevel = [string]"Critical"
                    SecurityImpact = [string]"Administrative groups with inheritance protection"
                    Recommendation = [string]"Audit group memberships and permissions"
                    ObjectAge = [string]"Various"
                    LastModified = [string]"Various"
                    Description = [string]"Administrative groups protected by AdminSDHolder"
                    DistinguishedName = [string]"Multiple Objects"
                    WhenCreated = [string]"Various"
                }
                
                foreach ($groupObj in ($GroupObjects | Sort-Object ObjectName)) {
                    $groupObj.TreeStructure = [string]"├── Administrative Group"
                    $Results += $groupObj
                }
            }
            
            if ($ComputerObjects.Count -gt 0) {
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                    ObjectName = [string]"Computer Accounts ($($ComputerObjects.Count))"
                    ObjectClass = [string]"Category"
                    ContainerPath = [string]"Various Containers"
                    AdminCount = [string]"Protected"
                    RiskLevel = [string]"Medium"
                    SecurityImpact = [string]"Computer accounts with elevated privileges"
                    Recommendation = [string]"Verify computer roles require protection"
                    ObjectAge = [string]"Various"
                    LastModified = [string]"Various"
                    Description = [string]"Computer accounts protected by AdminSDHolder"
                    DistinguishedName = [string]"Multiple Objects"
                    WhenCreated = [string]"Various"
                }
                
                foreach ($computerObj in ($ComputerObjects | Sort-Object ObjectName)) {
                    $computerObj.TreeStructure = [string]"├── Computer Account"
                    $Results += $computerObj
                }
            }
            
            if ($OtherObjects.Count -gt 0) {
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                    ObjectName = [string]"Other Objects ($($OtherObjects.Count))"
                    ObjectClass = [string]"Category"
                    ContainerPath = [string]"Various Containers"
                    AdminCount = [string]"Protected"
                    RiskLevel = [string]"Medium"
                    SecurityImpact = [string]"Miscellaneous objects with AdminSDHolder protection"
                    Recommendation = [string]"Review necessity of protection for these objects"
                    ObjectAge = [string]"Various"
                    LastModified = [string]"Various"
                    Description = [string]"Other object types protected by AdminSDHolder"
                    DistinguishedName = [string]"Multiple Objects"
                    WhenCreated = [string]"Various"
                }
                
                foreach ($otherObj in ($OtherObjects | Sort-Object ObjectName)) {
                    $otherObj.TreeStructure = [string]"├── Other Object"
                    $Results += $otherObj
                }
            }
            
            # Add summary if no objects found
            if ($Results.Count -eq 0) {
                $Results += [PSCustomObject]@{
                    TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                    ObjectName = [string]"No Protected Objects Found"
                    ObjectClass = [string]"Information"
                    ContainerPath = [string]"N/A"
                    AdminCount = [string]"0"
                    RiskLevel = [string]"Low"
                    SecurityImpact = [string]"No objects currently protected by AdminSDHolder"
                    Recommendation = [string]"This is normal for domains without administrative object protection"
                    ObjectAge = [string]"N/A"
                    LastModified = [string]"N/A"
                    Description = [string]"No objects with adminCount=1 found"
                    DistinguishedName = [string]"N/A"
                    WhenCreated = [string]"N/A"
                }
            }
            
        } catch {
            Write-ADReportLog -Message "Error querying AdminSDHolder protected objects: $($_.Exception.Message)" -Type Error
            $Results += [PSCustomObject]@{
                TreeStructure = [string]"🔐 AdminSDHolder Protected Objects"
                ObjectName = [string]"Query Error"
                ObjectClass = [string]"Error"
                ContainerPath = [string]"N/A"
                AdminCount = [string]"Unknown"
                RiskLevel = [string]"Critical"
                SecurityImpact = [string]"Cannot analyze AdminSDHolder protection status"
                Recommendation = [string]"Verify Active Directory connectivity and permissions"
                ObjectAge = [string]"N/A"
                LastModified = [string]"N/A"
                Description = [string]$_.Exception.Message
                DistinguishedName = [string]"N/A"
                WhenCreated = [string]"N/A"
            }
        }
        
        # Sort results to maintain tree structure
        $SortedResults = $Results | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, TreeStructure, ObjectName
        
        # Status update for GUI
        try {
            if ($Global:TextBlockStatus) {
                $criticalCount = ($Results | Where-Object { $_.RiskLevel -eq "Critical" }).Count
                $highCount = ($Results | Where-Object { $_.RiskLevel -eq "High" }).Count
                $mediumCount = ($Results | Where-Object { $_.RiskLevel -eq "Medium" }).Count
                
                $Global:TextBlockStatus.Text = "AdminSDHolder analysis completed. $($Results.Count) entries found. Critical: $criticalCount, High: $highCount, Medium: $mediumCount"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI status: $($_.Exception.Message)" -Type Warning
        }
        
        Write-ADReportLog -Message "AdminSDHolder analysis completed. $($Results.Count) entries found." -Type Info -Terminal
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing AdminSDHolder objects: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Error status for GUI
        try {
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI error status: $($_.Exception.Message)" -Type Warning
        }
        
        return @()
    }
}

function Get-AdvancedDelegations {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing advanced delegations..." -Type Info -Terminal
        
        $Results = @()
        $DomainDN = (Get-ADDomain).DistinguishedName
        
        # Critical containers to analyze
        $CriticalContainers = @(
            "CN=AdminSDHolder,CN=System,$DomainDN",
            "CN=Users,$DomainDN",
            "CN=Computers,$DomainDN",
            "OU=Domain Controllers,$DomainDN",
            $((Get-ADRootDSE).configurationNamingContext)
        )
        
        # Analyze delegations in critical containers
        foreach ($container in $CriticalContainers) {
            try {
                $containerName = ($container -split ',')[0] -replace '^(CN=|OU=)', ''
                Write-ADReportLog -Message "Analyzing delegations in container: $containerName" -Type Info
                
                $objects = Get-ADObject -SearchBase $container -Filter * -Properties nTSecurityDescriptor,distinguishedName -ErrorAction SilentlyContinue
                
                foreach ($object in $objects) {
                    try {
                        $acl = $object.nTSecurityDescriptor
                        if (-not $acl) { continue }
                        
                        foreach ($ace in $acl.Access) {
                            # Filter for explicit (non-inherited) permissions to non-system accounts
                            $isNotInherited = $ace.IsInherited -eq $false
                            $isNotSystemAccount = $ace.IdentityReference -notmatch "BUILTIN|NT AUTHORITY|SYSTEM|S-1-5-32|S-1-1-0|S-1-5-11"
                            $isAllowPermission = $ace.AccessControlType -eq "Allow"
                            
                            if ($isNotInherited -and $isNotSystemAccount -and $isAllowPermission) {
                                
                                # Determine risk level based on permissions and location
                                $riskLevel = "Low"
                                $securityImpact = "Standard delegation"
                                
                                # High-risk permissions
                                $activeDirectoryRights = $ace.ActiveDirectoryRights.ToString()
                                if ($activeDirectoryRights -match "GenericAll|WriteDacl|WriteOwner|FullControl") {
                                    $riskLevel = "Critical"
                                    $securityImpact = "Full control delegation - can modify any attribute and permissions"
                                }
                                elseif ($activeDirectoryRights -match "WriteProperty.*userAccountControl|WriteProperty.*pwdLastSet") {
                                    $riskLevel = "High" 
                                    $securityImpact = "Can modify critical user account properties"
                                }
                                elseif ($activeDirectoryRights -match "WriteProperty|CreateChild|DeleteChild") {
                                    $riskLevel = "Medium"
                                    $securityImpact = "Can modify objects or create/delete child objects"
                                }
                                
                                # Increase risk for critical containers
                                if ($container -match "AdminSDHolder|Domain Controllers" -and $riskLevel -eq "Medium") {
                                    $riskLevel = "High"
                                    $securityImpact += " in critical container"
                                }
                                
                                # Generate recommendation based on risk
                                $recommendation = ""
                                switch ($riskLevel) {
                                    "Critical" { $recommendation = "IMMEDIATE REVIEW: Remove or restrict full control delegations" }
                                    "High" { $recommendation = "Review and validate business justification for elevated permissions" }
                                    "Medium" { $recommendation = "Verify delegation necessity and consider principle of least privilege" }
                                    "Low" { $recommendation = "Monitor delegation usage and review periodically" }
                                }
                                
                                # Build tree structure path
                                $treePath = "└── $containerName"
                                if ($object.DistinguishedName -ne $container) {
                                    $objectPath = $object.DistinguishedName -replace ",$container$", ""
                                    $pathParts = $objectPath -split ',' | ForEach-Object { $_ -replace '^(CN=|OU=)', '' }
                                    [array]::Reverse($pathParts)
                                    $treePath += " → " + ($pathParts -join " → ")
                                }
                                
                                # Determine audit priority
                                $auditPriority = ""
                                switch ($riskLevel) {
                                    "Critical" { $auditPriority = "Immediate" }
                                    "High" { $auditPriority = "High" }
                                    "Medium" { $auditPriority = "Medium" }
                                    default { $auditPriority = "Standard" }
                                }
                                
                                $Results += [PSCustomObject]@{
                                    TreeStructure = [string]$treePath
                                    ContainerName = [string]$containerName
                                    ObjectName = [string]$object.Name
                                    ObjectType = [string]$object.ObjectClass
                                    DelegatedTo = [string]$ace.IdentityReference.ToString()
                                    PermissionType = [string]$activeDirectoryRights
                                    AccessType = [string]$ace.AccessControlType.ToString()
                                    RiskLevel = [string]$riskLevel
                                    SecurityImpact = [string]$securityImpact
                                    Recommendation = [string]$recommendation
                                    IsInherited = [bool]$ace.IsInherited
                                    ObjectPath = [string]$object.DistinguishedName
                                    AuditPriority = [string]$auditPriority
                                }
                            }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error processing object $($object.Name): $($_.Exception.Message)" -Type Warning
                        continue
                    }
                }
            } catch {
                Write-ADReportLog -Message "Error accessing container ${container}: $($_.Exception.Message)" -Type Warning
                continue
            }
        }
        
        # Sort results by risk level and container for tree display
        $SortedResults = $Results | Sort-Object @{
            Expression = { 
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, ContainerName, ObjectName
        
        Write-ADReportLog -Message "Advanced delegation analysis completed. $($SortedResults.Count) delegations found." -Type Info -Terminal
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing advanced delegations: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @()
    }
}

function Get-SchemaPermissions {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing schema permissions for security assessment..." -Type Info -Terminal
        
        $Results = @()
        
        # Get Schema Naming Context
        $SchemaNC = (Get-ADRootDSE).schemaNamingContext
        
        if (-not $SchemaNC) {
            Write-ADReportLog -Message "Could not access Schema Naming Context" -Type Error
            return @()
        }
        
        Write-ADReportLog -Message "Scanning schema objects in: $SchemaNC" -Type Info -Terminal
        
        # Get all schema objects with security descriptors
        try {
            $SchemaObjects = Get-ADObject -SearchBase $SchemaNC -SearchScope Subtree -Filter * -Properties nTSecurityDescriptor,objectClass,whenCreated,whenChanged,adminDescription -ErrorAction SilentlyContinue
        } catch {
            Write-ADReportLog -Message "Error accessing schema objects: $($_.Exception.Message)" -Type Error
            return @()
        }
        
        if (-not $SchemaObjects) {
            Write-ADReportLog -Message "No schema objects found or insufficient permissions" -Type Warning
            return @()
        }
        
        Write-ADReportLog -Message "Analyzing $($SchemaObjects.Count) schema objects for security permissions..." -Type Info
        
        # Define critical schema permissions to check
        $CriticalPermissions = @(
            'WriteProperty',
            'WriteDacl', 
            'WriteOwner',
            'DeleteChild',
            'DeleteTree',
            'GenericAll',
            'ControlAccess'
        )
        
        # Define expected/safe principals (English and German names)
        $SafePrincipals = @(
            'Domain Admins',
            'Domänen-Admins',
            'Enterprise Admins',
            'Organisations-Admins',
            'Schema Admins',
            'Schema-Admins',
            'SYSTEM',
            'Administrators',
            'Administratoren',
            'CREATOR OWNER'
        )
        
        $processedCount = 0
        foreach ($object in $SchemaObjects) {
            $processedCount++
            
            if ($processedCount % 100 -eq 0) {
                Write-ADReportLog -Message "Processed $processedCount of $($SchemaObjects.Count) schema objects..." -Type Info
            }
            
            try {
                if (-not $object.nTSecurityDescriptor) {
                    continue
                }
                
                $acl = $object.nTSecurityDescriptor
                
                # Determine object hierarchy for tree structure
                $objectPath = $object.DistinguishedName
                $pathParts = $objectPath -split ','
                
                # Build tree structure path
                $treePath = "🔐 Schema Security Analysis"
                if ($object.objectClass -contains "classSchema") {
                    $treePath += " → 📋 Class Schemas"
                } elseif ($object.objectClass -contains "attributeSchema") {
                    $treePath += " → 🏷️ Attribute Schemas"
                } elseif ($object.objectClass -contains "subSchema") {
                    $treePath += " → ⚙️ Schema Configuration"
                } else {
                    $treePath += " → 📁 Other Schema Objects"
                }
                
                # Calculate object age
                $objectAge = 0
                if ($object.whenCreated) {
                    $objectAge = [math]::Round((New-TimeSpan -Start $object.whenCreated -End (Get-Date)).TotalDays)
                }
                
                foreach ($ace in $acl.Access) {
                    try {
                        # Skip inherited permissions from parent containers
                        if ($ace.IsInherited) {
                            continue
                        }
                        
                        # Check for critical permissions
                        $hasRiskyPermission = $false
                        $riskyPermissions = @()
                        
                        foreach ($permission in $CriticalPermissions) {
                            if ($ace.ActiveDirectoryRights -match $permission) {
                                $hasRiskyPermission = $true
                                $riskyPermissions += $permission
                            }
                        }
                        
                        if (-not $hasRiskyPermission) {
                            continue
                        }
                        
                        # Determine if the principal is expected/safe
                        $identity = $ace.IdentityReference.Value
                        $isSafePrincipal = $false
                        
                        foreach ($safePrincipal in $SafePrincipals) {
                            if ($identity -like "*$safePrincipal*") {
                                $isSafePrincipal = $true
                                break
                            }
                        }
                        
                        # Calculate risk level
                        $riskLevel = "Low"
                        $securityImpact = "Standard schema permission"
                        $recommendation = "Monitor for changes"
                        
                        if (-not $isSafePrincipal) {
                            if ($riskyPermissions -contains "GenericAll" -or $riskyPermissions -contains "WriteOwner") {
                                $riskLevel = "Critical"
                                $securityImpact = "Full control over schema object - can modify AD structure"
                                $recommendation = "IMMEDIATE: Remove excessive permissions from non-administrative accounts"
                            } elseif ($riskyPermissions -contains "WriteDacl" -or $riskyPermissions -contains "WriteProperty") {
                                $riskLevel = "High"
                                $securityImpact = "Can modify schema definitions or permissions"
                                $recommendation = "Review and restrict schema modification rights"
                            } else {
                                $riskLevel = "Medium"
                                $securityImpact = "Potential unauthorized schema access"
                                $recommendation = "Verify necessity of schema permissions"
                            }
                        } else {
                            if ($riskyPermissions -contains "GenericAll") {
                                $riskLevel = "Medium"
                                $securityImpact = "Administrative account with full schema control"
                                $recommendation = "Monitor for unauthorized schema changes"
                            }
                        }
                        
                        # Create result object
                        $Results += [PSCustomObject]@{
                            TreeStructure = [string]$treePath
                            SchemaObjectName = [string]$object.Name
                            ObjectType = [string]($object.objectClass | Select-Object -Last 1)
                            IdentityReference = [string]$identity
                            PermissionsGranted = [string]($riskyPermissions -join ", ")
                            AccessType = [string]$ace.AccessControlType.ToString()
                            RiskLevel = [string]$riskLevel
                            IsSafePrincipal = [bool]$isSafePrincipal
                            SecurityImpact = [string]$securityImpact
                            Recommendation = [string]$recommendation
                            ObjectCreated = if ($object.whenCreated) { $object.whenCreated.ToString("dd.MM.yyyy HH:mm") } else { "Unknown" }
                            ObjectAge = [int]$objectAge
                            ObjectPath = [string]$object.DistinguishedName
                            AdminDescription = if ([string]::IsNullOrWhiteSpace($object.adminDescription)) { "No description" } else { [string]$object.adminDescription }
                            LastModified = if ($object.whenChanged) { $object.whenChanged.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                            ComplianceNote = [string]"Schema permissions should be restricted to administrative accounts only"
                        }
                        
                    } catch {
                        Write-ADReportLog -Message "Error processing ACE for object $($object.Name): $($_.Exception.Message)" -Type Warning
                        continue
                    }
                }
                
            } catch {
                Write-ADReportLog -Message "Error processing schema object $($object.Name): $($_.Exception.Message)" -Type Warning
                continue
            }
        }
        
        # Sort results by risk level and tree structure for optimal display
        $SortedResults = $Results | Sort-Object @{
            Expression = { 
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }
        }, TreeStructure, SchemaObjectName
        
        Write-ADReportLog -Message "Schema permissions analysis completed. Found $($SortedResults.Count) permission entries across $($SchemaObjects.Count) schema objects" -Type Info -Terminal
        
        # Update GUI with results
        try {
            if ($SortedResults.Count -gt 0) {
                # Update DataGrid with results
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $SortedResults
                }
                
                # Update status with risk summary
                $criticalCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "Critical" }).Count
                $highCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "High" }).Count
                $mediumCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "Medium" }).Count
                
                if ($Global:TextBlockStatus) {
                    $Global:TextBlockStatus.Text = "Schema permissions analysis completed. $($SortedResults.Count) permission(s) found. Critical: $criticalCount, High: $highCount, Medium: $mediumCount"
                }
                
                # Set status indicator based on highest risk
                if ($Global:StatusIndicator) {
                    if ($criticalCount -gt 0) {
                        $Global:StatusIndicator.Fill = "#FFFF0000"  # Red for critical
                    } elseif ($highCount -gt 0) {
                        $Global:StatusIndicator.Fill = "#FFFF8000"  # Orange for high
                    } elseif ($mediumCount -gt 0) {
                        $Global:StatusIndicator.Fill = "#FFFFFF00"  # Yellow for medium
                    } else {
                        $Global:StatusIndicator.Fill = "#FF00C800"  # Green for low/safe
                    }
                }
            } else {
                # No results found
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = @()
                }
                if ($Global:TextBlockStatus) {
                    $Global:TextBlockStatus.Text = "Schema permissions analysis completed. No concerning permissions found."
                }
                if ($Global:StatusIndicator) {
                    $Global:StatusIndicator.Fill = "#FF00C800"  # Green for safe
                }
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI: $($_.Exception.Message)" -Type Warning
        }
        
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error analyzing schema permissions: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Update GUI error state
        try {
            if ($Global:StatusIndicator) {
                $Global:StatusIndicator.Fill = "#FFFF0000"  # Red for error
            }
            if ($Global:DataGridResults) {
                $Global:DataGridResults.ItemsSource = @()
            }
            if ($Global:TextBlockStatus) {
                $Global:TextBlockStatus.Text = "Error analyzing schema permissions: $ErrorMessage"
            }
        } catch {
            Write-ADReportLog -Message "Could not update GUI error state: $($_.Exception.Message)" -Type Warning
        }
        
        return @()
    }
}

Function Get-ExpiredPasswordUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing all users for password expiration status..." -Type Info -Terminal
        
        # Properties for password expiration analysis
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", 
            "PasswordLastSet", "LastLogonDate", "PasswordNotRequired",
            "AccountExpirationDate", "LockedOut", "BadLogonCount",
            "UserPrincipalName", "DistinguishedName", "Description",
            "whenCreated", "MemberOf"
        )
        
        # Get Domain Password Policy
        try {
            $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
            $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
            Write-ADReportLog -Message "Domain Password Policy - Max Password Age: $MaxPasswordAge days" -Type Info -Terminal
        } catch {
            Write-ADReportLog -Message "Could not retrieve domain password policy, using default 42 days" -Type Warning -Terminal
            $MaxPasswordAge = 42  # Fallback value
        }
        
        # Load ALL users (enabled and disabled)
        $AllUsers = Get-ADUser -Filter * -Properties $Properties -ErrorAction SilentlyContinue
        
        if (-not $AllUsers) {
            Write-ADReportLog -Message "No users found in the domain" -Type Warning -Terminal
            return @()
        }
        
        Write-ADReportLog -Message "$($AllUsers.Count) users loaded for password expiration analysis..." -Type Info -Terminal
        
        # Analyze ALL users for password status
        $PasswordUsers = @()
        $CurrentDate = Get-Date
        
        foreach ($user in $AllUsers) {
            try {
                $passwordStatus = "Active"
                $daysUntilExpiration = 0
                $passwordAge = "Unknown"
                $expirationDate = "Never"
                $expirationDateTime = $null
                $riskLevel = "Low"
                $statusDetails = ""
                
                # Check if password is required
                if ($user.PasswordNotRequired -eq $true) {
                    $passwordStatus = "Not Required"
                    $passwordAge = "N/A"
                    $riskLevel = "Critical"
                    $statusDetails = "Password not required"
                    $daysUntilExpiration = -9999  # Sort to top as critical
                }
                # Check if password never expires
                elseif ($user.PasswordNeverExpires -eq $true) {
                    $passwordStatus = "Never Expires"
                    $riskLevel = "Medium"
                    $statusDetails = "Password never expires"
                    $daysUntilExpiration = 9999  # Sort to bottom
                    $expirationDateTime = [DateTime]::MaxValue
                    
                    if ($user.PasswordLastSet -eq $null) {
                        $passwordAge = "Never set"
                        $riskLevel = "Critical"
                        $statusDetails = "Password never set and never expires"
                    } else {
                        $passwordAgeTimeSpan = $CurrentDate - $user.PasswordLastSet
                        $passwordAge = "$([Math]::Round($passwordAgeTimeSpan.TotalDays)) days"
                    }
                }
                # Password has expiration policy
                else {
                    if ($user.PasswordLastSet -eq $null) {
                        $passwordStatus = "Never Set"
                        $passwordAge = "Never set"
                        $riskLevel = "Critical"
                        $statusDetails = "Password never set"
                        $daysUntilExpiration = -9998  # Sort near top as critical
                        $expirationDateTime = [DateTime]::MinValue
                    } else {
                        # Calculate password age and expiration
                        $passwordAgeTimeSpan = $CurrentDate - $user.PasswordLastSet
                        $passwordAge = "$([Math]::Round($passwordAgeTimeSpan.TotalDays)) days"
                        
                        # Calculate expiration date
                        $calculatedExpirationDate = $user.PasswordLastSet.AddDays($MaxPasswordAge)
                        $expirationDate = $calculatedExpirationDate.ToString("dd.MM.yyyy")
                        $expirationDateTime = $calculatedExpirationDate
                        
                        # Calculate days until expiration (negative if expired)
                        $daysUntilExpiration = [Math]::Round(($calculatedExpirationDate - $CurrentDate).TotalDays)
                        
                        # Determine status and risk level
                        if ($daysUntilExpiration -lt 0) {
                            $passwordStatus = "Expired"
                            $riskLevel = "High"
                            $statusDetails = "Expired $([Math]::Abs($daysUntilExpiration)) days ago"
                            
                            # Critical if expired more than 90 days
                            if ([Math]::Abs($daysUntilExpiration) -gt 90) {
                                $riskLevel = "Critical"
                            }
                        } elseif ($daysUntilExpiration -eq 0) {
                            $passwordStatus = "Expires Today"
                            $riskLevel = "High"
                            $statusDetails = "Expires today"
                        } elseif ($daysUntilExpiration -le 7) {
                            $passwordStatus = "Expires Soon"
                            $riskLevel = "Medium"
                            $statusDetails = "Expires in $daysUntilExpiration days"
                        } elseif ($daysUntilExpiration -le 30) {
                            $passwordStatus = "Expires This Month"
                            $riskLevel = "Low"
                            $statusDetails = "Expires in $daysUntilExpiration days"
                        } else {
                            $passwordStatus = "Active"
                            $riskLevel = "Info"
                            $statusDetails = "Expires in $daysUntilExpiration days"
                        }
                    }
                }
                
                # Check for privileged groups
                $isPrivileged = $false
                $privilegedGroups = @()
                if ($user.MemberOf) {
                    $privilegedGroupNames = @(
                        "Domain Admins", "Enterprise Admins", "Schema Admins", 
                        "Administrators", "Account Operators", "Backup Operators",
                        "Server Operators", "Print Operators", "Domänen-Admins",
                        "Organisations-Admins", "Schema-Admins", "Administratoren",
                        "Konto-Operatoren", "Sicherungs-Operatoren", "Server-Operatoren",
                        "Druck-Operatoren"
                    )
                    
                    foreach ($groupDN in $user.MemberOf) {
                        try {
                            $group = Get-ADGroup -Identity $groupDN -ErrorAction SilentlyContinue
                            if ($group -and $privilegedGroupNames -contains $group.Name) {
                                $isPrivileged = $true
                                $privilegedGroups += $group.Name
                            }
                        } catch {
                            # Group could not be resolved - ignore
                        }
                    }
                }
                
                # Increase risk if privileged account has issues
                if ($isPrivileged -and ($passwordStatus -in @("Expired", "Never Set", "Not Required", "Expires Today", "Expires Soon"))) {
                    if ($riskLevel -ne "Critical") {
                        $riskLevel = "High"
                    }
                }
                
                # Add user to results (ALL users, not just problematic ones)
                $PasswordUsers += [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    PasswordStatus = $passwordStatus
                    PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                    PasswordAge = $passwordAge
                    ExpirationDate = $expirationDate
                    ExpirationDateTime = $expirationDateTime
                    DaysUntilExpiration = $daysUntilExpiration
                    StatusDetails = $statusDetails
                    RiskLevel = $riskLevel
                    IsPrivileged = $isPrivileged
                    PrivilegedGroups = if ($privilegedGroups.Count -gt 0) { $privilegedGroups -join ", " } else { "None" }
                    LastLogonDate = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("dd.MM.yyyy HH:mm") } else { "Never" }
                    AccountExpirationDate = if ($user.AccountExpirationDate) { $user.AccountExpirationDate.ToString("dd.MM.yyyy") } else { "Never" }
                    LockedOut = $user.LockedOut
                    BadLogonCount = $user.BadLogonCount
                    DistinguishedName = $user.DistinguishedName
                    WhenCreated = if ($user.whenCreated) { $user.whenCreated.ToString("dd.MM.yyyy") } else { "Unknown" }
                }
                
            } catch {
                Write-ADReportLog -Message "Error processing user $($user.SamAccountName): $($_.Exception.Message)" -Type Warning
                # Add user as error case
                $PasswordUsers += [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    PasswordStatus = "Error"
                    PasswordLastSet = "Error"
                    PasswordAge = "Error"
                    ExpirationDate = "Error"
                    ExpirationDateTime = [DateTime]::MaxValue
                    DaysUntilExpiration = -9997
                    StatusDetails = "Error processing user"
                    RiskLevel = "Critical"
                    IsPrivileged = "Unknown"
                    PrivilegedGroups = "Error"
                    LastLogonDate = "Error"
                    AccountExpirationDate = "Error"
                    LockedOut = "Error"
                    BadLogonCount = "Error"
                    Description = "Error processing user"
                    DistinguishedName = $user.DistinguishedName
                    WhenCreated = "Error"
                }
            }
        }
        
        # Sort results: 
        # 1. Critical issues first (by RiskLevel)
        # 2. Then by expiration date (earliest expiration first)
        # 3. Then by DaysUntilExpiration as fallback
        $SortedResults = $PasswordUsers | Sort-Object @{
            Expression = {
                switch ($_.RiskLevel) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    "Info" { 5 }
                    default { 6 }
                }
            }
        }, @{
            Expression = { $_.ExpirationDateTime }
            Ascending = $true
        }, DaysUntilExpiration
        
        Write-ADReportLog -Message "Password expiration analysis completed. Analyzed $($SortedResults.Count) users." -Type Info -Terminal
        
        # Log statistics
        if ($SortedResults.Count -gt 0) {
            $criticalCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "Critical" }).Count
            $highCount = ($SortedResults | Where-Object { $_.RiskLevel -eq "High" }).Count
            $expiredCount = ($SortedResults | Where-Object { $_.PasswordStatus -eq "Expired" }).Count
            $privilegedCount = ($SortedResults | Where-Object { $_.IsPrivileged -eq $true }).Count
            
            Write-ADReportLog -Message "Password Statistics - Critical: $criticalCount, High: $highCount, Expired: $expiredCount, Privileged Accounts: $privilegedCount" -Type Info -Terminal
        }
        
        return $SortedResults
        
    } catch {
        $ErrorMessage = "Error in Get-ExpiredPasswordUsers: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        
        # Fallback result on errors
        return @([PSCustomObject]@{
            DisplayName = "Error during analysis"
            SamAccountName = "ERROR"
            UserPrincipalName = "error@domain.local"
            Enabled = "Error"
            PasswordStatus = "Error"
            PasswordLastSet = "Error"
            PasswordAge = "Error"
            ExpirationDate = "Error"
            ExpirationDateTime = [DateTime]::MaxValue
            DaysUntilExpiration = -9999
            StatusDetails = "Function failed"
            RiskLevel = "Critical"
            IsPrivileged = "Unknown"
            PrivilegedGroups = "Error"
            LastLogonDate = "Error"
            AccountExpirationDate = "Error"
            LockedOut = "Error"
            BadLogonCount = "Error"
            Description = "Function failed: $($_.Exception.Message)"
            DistinguishedName = "CN=ERROR,DC=domain,DC=local"
            WhenCreated = "Error"
        })
    }
}



# ===================================
# NETWORK TOPOLOGY FUNCTIONS
# ===================================

Function Get-ADNetworkTopology {
    [CmdletBinding()]
    param(
        [ValidateSet("DomainControllers", "Sites", "OUHierarchy", "Trusts")]
        [string]$ViewType = "DomainControllers"
    )
    
    $topology = @{
        nodes = @()
        links = @()
    }
    
    try {
        switch ($ViewType) {
            "DomainControllers" {
                # Hole Domain und DCs
                $domain = Get-ADDomain
                $dcs = Get-ADDomainController -Filter *
                
                # Zentraler Domain-Knoten
                $topology.nodes += @{
                    id = "domain"
                    label = $domain.Name
                    type = "domain"
                    x = 400
                    y = 300
                    color = "#0078D7"
                    size = 40
                }
                
                # DC Knoten im Kreis anordnen
                $angleStep = 360 / $dcs.Count
                $radius = 200
                $centerX = 400
                $centerY = 300
                
                for ($i = 0; $i -lt $dcs.Count; $i++) {
                    $dc = $dcs[$i]
                    $angle = $i * $angleStep * [Math]::PI / 180
                    $x = $centerX + $radius * [Math]::Cos($angle)
                    $y = $centerY + $radius * [Math]::Sin($angle)
                    
                    $topology.nodes += @{
                        id = $dc.Name
                        label = $dc.Name
                        type = "dc"
                        x = $x
                        y = $y
                        color = if ($dc.IsGlobalCatalog) { "#10B981" } else { "#3B82F6" }
                        size = 30
                        properties = @{
                            Site = $dc.Site
                            IP = $dc.IPv4Address
                            OS = $dc.OperatingSystem
                            IsGC = $dc.IsGlobalCatalog
                            IsRODC = $dc.IsReadOnly
                        }
                    }
                    
                    # Link zum Domain-Knoten
                    $topology.links += @{
                        source = "domain"
                        target = $dc.Name
                        type = "contains"
                        color = "#9CA3AF"
                    }
                }
            }
            
            "Sites" {
                # AD Sites und Replikationslinks
                $sites = Get-ADReplicationSite -Filter *
                $siteLinks = Get-ADReplicationSiteLink -Filter *
                
                # Erstelle Site-Knoten
                $gridSize = [Math]::Ceiling([Math]::Sqrt($sites.Count))
                $spacing = 150
                
                for ($i = 0; $i -lt $sites.Count; $i++) {
                    $site = $sites[$i]
                    $row = [Math]::Floor($i / $gridSize)
                    $col = $i % $gridSize
                    
                    $topology.nodes += @{
                        id = $site.Name
                        label = $site.Name
                        type = "site"
                        x = 100 + ($col * $spacing)
                        y = 100 + ($row * $spacing)
                        color = "#10B981"
                        size = 35
                        properties = @{
                            Description = $site.Description
                            Location = $site.Location
                        }
                    }
                }
                
                # Erstelle Site-Links
                foreach ($link in $siteLinks) {
                    $sitesInLink = $link.SitesIncluded | ForEach-Object {
                        ($_ -split ',')[0] -replace 'CN=', ''
                    }
                    
                    for ($i = 0; $i -lt $sitesInLink.Count - 1; $i++) {
                        for ($j = $i + 1; $j -lt $sitesInLink.Count; $j++) {
                            $topology.links += @{
                                source = $sitesInLink[$i]
                                target = $sitesInLink[$j]
                                type = "replication"
                                color = "#3B82F6"
                                cost = $link.Cost
                            }
                        }
                    }
                }
            }
        }
        
        return $topology
    }
    catch {
        Write-Error "Fehler beim Erstellen der Netzwerktopologie: $_"
        return $topology
    }
}

# Funktion zum Zeichnen auf Canvas
Function Draw-TopologyOnCanvas {
    param(
        [System.Windows.Controls.Canvas]$Canvas,
        [hashtable]$Topology
    )
    
    $Canvas.Children.Clear()
    
    # Zeichne Links
    foreach ($link in $Topology.links) {
        $sourceNode = $Topology.nodes | Where-Object { $_.id -eq $link.source }
        $targetNode = $Topology.nodes | Where-Object { $_.id -eq $link.target }
        
        if ($sourceNode -and $targetNode) {
            $line = New-Object System.Windows.Shapes.Line
            $line.X1 = $sourceNode.x
            $line.Y1 = $sourceNode.y
            $line.X2 = $targetNode.x
            $line.Y2 = $targetNode.y
            $line.Stroke = $link.color
            $line.StrokeThickness = 2
            $line.Opacity = 0.6
            $Canvas.Children.Add($line)
        }
    }
    
    # Zeichne Knoten
    foreach ($node in $Topology.nodes) {
        # Knoten-Container
        $nodeGroup = New-Object System.Windows.Controls.Canvas
        
        # Kreis für Knoten
        $ellipse = New-Object System.Windows.Shapes.Ellipse
        $ellipse.Width = $node.size
        $ellipse.Height = $node.size
        $ellipse.Fill = $node.color
        $ellipse.Stroke = "#E5E7EB"
        $ellipse.StrokeThickness = 2
        
        # Schatten
        $shadow = New-Object System.Windows.Media.Effects.DropShadowEffect
        $shadow.BlurRadius = 8
        $shadow.ShadowDepth = 2
        $shadow.Opacity = 0.3
        $ellipse.Effect = $shadow
        
        # Text-Label
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $node.label
        $label.Foreground = "White"
        $label.FontWeight = "SemiBold"
        $label.FontSize = 10
        $label.TextAlignment = "Center"
        
        # Positionierung
        [System.Windows.Controls.Canvas]::SetLeft($ellipse, $node.x - $node.size/2)
        [System.Windows.Controls.Canvas]::SetTop($ellipse, $node.y - $node.size/2)
        [System.Windows.Controls.Canvas]::SetLeft($label, $node.x - 30)
        [System.Windows.Controls.Canvas]::SetTop($label, $node.y - 5)
        
        # Tooltip
        if ($node.properties) {
            $tooltip = ""
            foreach ($prop in $node.properties.GetEnumerator()) {
                $tooltip += "$($prop.Key): $($prop.Value)`n"
            }
            $ellipse.ToolTip = $tooltip.TrimEnd()
        }
        
        $Canvas.Children.Add($ellipse)
        $Canvas.Children.Add($label)
    }
}

# ===================================
# SECURITY HEATMAP FUNCTIONS
# ===================================

Function Get-SecurityHeatMapData {
    [CmdletBinding()]
    param(
        [ValidateSet("PasswordSecurity", "AccountActivity", "PrivilegeLevel", "Compliance")]
        [string]$MetricType = "PasswordSecurity",
        
        [ValidateSet("Department", "OU", "Manager")]
        [string]$GroupBy = "Department"
    )
    
    $heatMapData = @()
    
    try {
        # Hole Benutzer mit relevanten Attributen
        $users = Get-ADUser -Filter * -Properties *
        
        # Gruppiere Benutzer
        $groups = switch ($GroupBy) {
            "Department" { $users | Group-Object Department | Where-Object { $_.Name } }
            "OU" { $users | Group-Object { ($_.DistinguishedName -split ',')[1] } }
            "Manager" { $users | Group-Object Manager | Where-Object { $_.Name } }
        }
        
        foreach ($group in $groups) {
            $score = 0
            $details = @{}
            
            switch ($MetricType) {
                "PasswordSecurity" {
                    # Berechne Passwort-Sicherheitsscore
                    $pwdNeverExpires = ($group.Group | Where-Object { $_.PasswordNeverExpires }).Count
                    $pwdNotRequired = ($group.Group | Where-Object { $_.PasswordNotRequired }).Count
                    $oldPasswords = ($group.Group | Where-Object { 
                        $_.PasswordLastSet -and $_.PasswordLastSet -lt (Get-Date).AddDays(-180) 
                    }).Count
                    
                    # Score: 0-100 (100 = sehr sicher)
                    $totalUsers = $group.Count
                    $score = 100
                    $score -= ($pwdNeverExpires / $totalUsers) * 30
                    $score -= ($pwdNotRequired / $totalUsers) * 40
                    $score -= ($oldPasswords / $totalUsers) * 30
                    
                    $details = @{
                        "Total Users" = $totalUsers
                        "Password Never Expires" = $pwdNeverExpires
                        "Password Not Required" = $pwdNotRequired
                        "Old Passwords (>180d)" = $oldPasswords
                    }
                }
                
                "AccountActivity" {
                    # Berechne Account-Aktivitäts-Risiko
                    $inactiveUsers = ($group.Group | Where-Object { 
                        $_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90) 
                    }).Count
                    $neverLoggedOn = ($group.Group | Where-Object { -not $_.LastLogonDate }).Count
                    $disabledAccounts = ($group.Group | Where-Object { -not $_.Enabled }).Count
                    
                    $totalUsers = $group.Count
                    $score = 100
                    $score -= ($inactiveUsers / $totalUsers) * 40
                    $score -= ($neverLoggedOn / $totalUsers) * 30
                    $score -= ($disabledAccounts / $totalUsers) * 30
                    
                    $details = @{
                        "Total Users" = $totalUsers
                        "Inactive (>90d)" = $inactiveUsers
                        "Never Logged On" = $neverLoggedOn
                        "Disabled Accounts" = $disabledAccounts
                    }
                }
            }
            
            $heatMapData += @{
                Group = $group.Name
                Score = [Math]::Round($score, 1)
                Count = $group.Count
                Details = $details
                Color = Get-HeatMapColor -Score $score
            }
        }
        
        return $heatMapData | Sort-Object Score
    }
    catch {
        Write-Error "Fehler beim Erstellen der HeatMap-Daten: $_"
        return @()
    }
}

Function Get-HeatMapColor {
    param([double]$Score)
    
    # Farbverlauf von Rot (0) über Gelb (50) zu Grün (100)
    if ($Score -ge 80) { return "#10B981" }  # Grün
    elseif ($Score -ge 60) { return "#34D399" }  # Hellgrün
    elseif ($Score -ge 40) { return "#FDE047" }  # Gelb
    elseif ($Score -ge 20) { return "#FB923C" }  # Orange
    else { return "#EF4444" }  # Rot
}

Function Draw-SecurityHeatMap {
    param(
        [System.Windows.Controls.Grid]$Container,
        [array]$HeatMapData
    )
    
    $Container.Children.Clear()
    $Container.RowDefinitions.Clear()
    $Container.ColumnDefinitions.Clear()
    
    # Berechne Grid-Layout
    $itemCount = $HeatMapData.Count
    $cols = [Math]::Ceiling([Math]::Sqrt($itemCount))
    $rows = [Math]::Ceiling($itemCount / $cols)
    
    # Erstelle Grid-Definitionen
    for ($i = 0; $i -lt $rows; $i++) {
        $rowDef = New-Object System.Windows.RowDefinition
        $rowDef.Height = "1*"
        $Container.RowDefinitions.Add($rowDef)
    }
    
    for ($i = 0; $i -lt $cols; $i++) {
        $colDef = New-Object System.Windows.ColumnDefinition
        $colDef.Width = "1*"
        $Container.ColumnDefinitions.Add($colDef)
    }
    
    # Füge HeatMap-Zellen hinzu
    for ($i = 0; $i -lt $HeatMapData.Count; $i++) {
        $data = $HeatMapData[$i]
        $row = [Math]::Floor($i / $cols)
        $col = $i % $cols
        
        # Zellen-Container
        $border = New-Object System.Windows.Controls.Border
        $border.Margin = "5"
        $border.CornerRadius = "8"
        $border.Background = $data.Color
        $border.MinHeight = 100
        
        # Inhalt
        $stack = New-Object System.Windows.Controls.StackPanel
        $stack.VerticalAlignment = "Center"
        $stack.Margin = "10"
        
        # Gruppenname
        $groupLabel = New-Object System.Windows.Controls.TextBlock
        $groupLabel.Text = $data.Group
        $groupLabel.FontWeight = "Bold"
        $groupLabel.FontSize = 14
        $groupLabel.Foreground = "White"
        $groupLabel.TextWrapping = "Wrap"
        $groupLabel.HorizontalAlignment = "Center"
        $stack.Children.Add($groupLabel)
        
        # Score
        $scoreLabel = New-Object System.Windows.Controls.TextBlock
        $scoreLabel.Text = "$($data.Score)%"
        $scoreLabel.FontSize = 24
        $scoreLabel.FontWeight = "Bold"
        $scoreLabel.Foreground = "White"
        $scoreLabel.HorizontalAlignment = "Center"
        $scoreLabel.Margin = "0,5,0,0"
        $stack.Children.Add($scoreLabel)
        
        # Benutzeranzahl
        $countLabel = New-Object System.Windows.Controls.TextBlock
        $countLabel.Text = "$($data.Count) users"
        $countLabel.FontSize = 11
        $countLabel.Foreground = "White"
        $countLabel.Opacity = 0.8
        $countLabel.HorizontalAlignment = "Center"
        $stack.Children.Add($countLabel)
        
        $border.Child = $stack
        
        # Tooltip mit Details
        $tooltip = New-Object System.Windows.Controls.ToolTip
        $tooltipStack = New-Object System.Windows.Controls.StackPanel
        
        foreach ($detail in $data.Details.GetEnumerator()) {
            $detailText = New-Object System.Windows.Controls.TextBlock
            $detailText.Text = "$($detail.Key): $($detail.Value)"
            $detailText.Margin = "0,2"
            $tooltipStack.Children.Add($detailText)
        }
        
        $tooltip.Content = $tooltipStack
        $border.ToolTip = $tooltip
        
        # Position im Grid
        [System.Windows.Controls.Grid]::SetRow($border, $row)
        [System.Windows.Controls.Grid]::SetColumn($border, $col)
        
        $Container.Children.Add($border)
    }
}

# ===================================
# CUSTOM REPORT BUILDER FUNCTIONS
# ===================================

# Report Template Klasse
class ReportTemplate {
    [string]$Name
    [string]$Description
    [array]$Fields
    [hashtable]$Filters
    [string]$Layout
    
    ReportTemplate([string]$name) {
        $this.Name = $name
        $this.Fields = @()
        $this.Filters = @{}
        $this.Layout = "Table"
    }
}

Function Initialize-ReportBuilder {
    param(
        [System.Windows.Controls.ListBox]$AvailableFieldsList,
        [System.Windows.Controls.StackPanel]$ReportCanvas
    )
    
    # Verfügbare Felder laden
    $availableFields = @(
        @{Name="DisplayName"; Category="User"; Type="String"}
        @{Name="SamAccountName"; Category="User"; Type="String"}
        @{Name="Department"; Category="User"; Type="String"}
        @{Name="Title"; Category="User"; Type="String"}
        @{Name="Manager"; Category="User"; Type="String"}
        @{Name="LastLogonDate"; Category="User"; Type="DateTime"}
        @{Name="PasswordLastSet"; Category="User"; Type="DateTime"}
        @{Name="Enabled"; Category="User"; Type="Boolean"}
        @{Name="Name"; Category="Computer"; Type="String"}
        @{Name="OperatingSystem"; Category="Computer"; Type="String"}
        @{Name="LastLogonDate"; Category="Computer"; Type="DateTime"}
        @{Name="GroupName"; Category="Group"; Type="String"}
        @{Name="GroupCategory"; Category="Group"; Type="String"}
        @{Name="GroupScope"; Category="Group"; Type="String"}
        @{Name="Members"; Category="Group"; Type="Array"}
    )
    
    foreach ($field in $availableFields) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = "$($field.Category).$($field.Name)"
        $item.Tag = $field
        
        # Drag & Drop aktivieren
        $item.AllowDrop = $true
        $item.Add_MouseMove({
            param($sender, $e)
            if ($e.LeftButton -eq 'Pressed') {
                [System.Windows.DragDrop]::DoDragDrop($sender, $sender.Tag, 'Copy')
            }
        })
        
        $AvailableFieldsList.Items.Add($item)
    }
    
    # Drop-Handler für Canvas
    $ReportCanvas.AllowDrop = $true
    $ReportCanvas.Add_Drop({
        param($sender, $e)
        $field = $e.Data.GetData([hashtable])
        if ($field) {
            Add-FieldToReport -Field $field -Canvas $sender
        }
    })
    
    $ReportCanvas.Add_DragOver({
        param($sender, $e)
        $e.Effects = 'Copy'
        $e.Handled = $true
    })
}

Function Add-FieldToReport {
    param(
        [hashtable]$Field,
        [System.Windows.Controls.StackPanel]$Canvas
    )
    
    # Erstelle Feld-Container
    $fieldContainer = New-Object System.Windows.Controls.Border
    $fieldContainer.Background = "White"
    $fieldContainer.BorderBrush = "#E5E7EB"
    $fieldContainer.BorderThickness = "1"
    $fieldContainer.CornerRadius = "6"
    $fieldContainer.Margin = "0,5"
    $fieldContainer.Padding = "10"
    
    $grid = New-Object System.Windows.Controls.Grid
    $col1 = New-Object System.Windows.ColumnDefinition
    $col1.Width = "*"
    $col2 = New-Object System.Windows.ColumnDefinition
    $col2.Width = "Auto"
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    
    # Feldname
    $nameBlock = New-Object System.Windows.Controls.TextBlock
    $nameBlock.Text = "$($Field.Category).$($Field.Name)"
    $nameBlock.FontWeight = "Medium"
    $nameBlock.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($nameBlock, 0)
    $grid.Children.Add($nameBlock)
    
    # Entfernen-Button
    $removeBtn = New-Object System.Windows.Controls.Button
    $removeBtn.Content = "✕"
    $removeBtn.Width = 20
    $removeBtn.Height = 20
    $removeBtn.Background = "Transparent"
    $removeBtn.BorderThickness = 0
    $removeBtn.Tag = $fieldContainer
    $removeBtn.Add_Click({
        param($s, $e)
        $Canvas.Children.Remove($s.Tag)
    })
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 1)
    $grid.Children.Add($removeBtn)
    
    $fieldContainer.Child = $grid
    $Canvas.Children.Remove($Canvas.Children[0]) # Entferne Platzhalter-Text
    $Canvas.Children.Add($fieldContainer)
}

Function Build-CustomReport {
    param(
        [System.Windows.Controls.StackPanel]$ReportCanvas
    )
    
    $selectedFields = @()
    foreach ($child in $ReportCanvas.Children) {
        if ($child -is [System.Windows.Controls.Border]) {
            $textBlock = $child.Child.Children[0]
            $selectedFields += $textBlock.Text
        }
    }
    
    if ($selectedFields.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Feld aus.", "Report Builder", "OK", "Warning")
        return
    }
    
    # Generiere LDAP-Properties aus Feldnamen
    $properties = $selectedFields | ForEach-Object {
        ($_ -split '\.')[-1]
    } | Select-Object -Unique
    
    # Bestimme Objekttyp
    $categories = $selectedFields | ForEach-Object {
        ($_ -split '\.')[0]
    } | Select-Object -Unique
    
    $results = @()
    
    try {
        if ($categories -contains "User") {
            $users = Get-ADUser -Filter * -Properties $properties
            foreach ($user in $users) {
                $obj = [PSCustomObject]@{}
                foreach ($field in $selectedFields) {
                    if ($field -like "User.*") {
                        $propName = ($field -split '\.')[-1]
                        $obj | Add-Member -NotePropertyName $field -NotePropertyValue $user.$propName
                    }
                }
                $results += $obj
            }
        }
        
        # Ähnliche Logik für Computer und Groups...
        
        return $results
    }
    catch {
        Write-Error "Fehler beim Erstellen des benutzerdefinierten Berichts: $_"
        return @()
    }
}

Function Save-ReportTemplate {
    param(
        [System.Windows.Controls.StackPanel]$ReportCanvas,
        [string]$TemplateName,
        [string]$Description
    )
    
    $template = [ReportTemplate]::new($TemplateName)
    $template.Description = $Description
    
    foreach ($child in $ReportCanvas.Children) {
        if ($child -is [System.Windows.Controls.Border]) {
            $textBlock = $child.Child.Children[0]
            $template.Fields += $textBlock.Text
        }
    }
    
    # Speichere Template als JSON
    $templatesPath = "$env:APPDATA\AD Reporting Tool\Templates"
    if (-not (Test-Path $templatesPath)) {
        New-Item -ItemType Directory -Path $templatesPath -Force
    }
    
    $templateFile = Join-Path $templatesPath "$($TemplateName).json"
    $template | ConvertTo-Json | Out-File $templateFile -Encoding UTF8
    
    return $templateFile
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Initialize-ResultCounters {
    # GesamtergebniszÃ¤hler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle ZÃ¤hler zurÃ¼ckgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurÃ¼cksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$Results
    )

    if ($null -eq $Results) {
        if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = "0" }
        if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = "0" }
        if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = "0" }
        if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = "0" }
        return
    }

    $totalCount = $Results.Count
    $userCount = 0
    $computerCount = 0
    $groupCount = 0

    if ($totalCount -gt 0) {
        foreach ($item in $Results) {
            if ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Person,CN=Schema,CN=Configuration,*") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Computer,CN=Schema,CN=Configuration,*") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Group,CN=Schema,CN=Configuration,*") {
                $groupCount++
            }
        }
    }

    if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = $totalCount.ToString() }
    if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = $userCount.ToString() }
    if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = $computerCount.ToString() }
    if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = $groupCount.ToString() }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )

    # Stelle sicher, dass Results immer ein Array ist, auch bei NULL
    if ($null -eq $Results) {
        $Results = @()
    } elseif ($Results -isnot [array]) {
        $Results = @($Results)
    }

    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $Results
    }

    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
    
    Update-ResultCounters -Results $Results
    Update-ResultVisualization -Results $Results
}

# Hilfsfunktion zum sicheren HinzufÃ¼gen von Event-Handlern
Function Add-SafeEventHandler {
    param(
        [Parameter(Mandatory=$true)]
        $Button,
        [Parameter(Mandatory=$true)]
        [scriptblock]$Handler
    )
    
    if ($null -ne $Button) {
        try {
            $Button.add_Click($Handler)
        } catch {
            Write-ADReportLog -Message "Warnung: Konnte Event-Handler fÃ¼r Button nicht hinzufÃ¼gen: $($_.Exception.Message)" -Type Warning -Terminal
        }
   }
}

# --- Globale Variablen fÃ¼r UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # ÃœberprÃ¼fen, ob das Window-Objekt bereits existiert und zurÃ¼cksetzen
    if ($Global:Window) {
        Remove-Variable -Name Window -Scope Global -ErrorAction SilentlyContinue
    }
    
    $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
    $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Objekttyp Radio Buttons
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")

    # Filter und Attribute
    $Global:ComboBoxFilterAttribute1 = $Window.FindName("ComboBoxFilterAttribute1")
    $Global:ComboBoxFilterOperator1 = $Window.FindName("ComboBoxFilterOperator1")
    $Global:TextBoxFilterValue1 = $Window.FindName("TextBoxFilterValue1")
    $Global:ComboBoxFilterAttribute2 = $Window.FindName("ComboBoxFilterAttribute2")  
    $Global:ComboBoxFilterOperator2 = $Window.FindName("ComboBoxFilterOperator2")
    $Global:TextBoxFilterValue2 = $Window.FindName("TextBoxFilterValue2")
    $Global:RadioButtonAnd = $Window.FindName("RadioButtonAnd")
    $Global:RadioButtonOr = $Window.FindName("RadioButtonOr")
    $Global:CheckBoxUseSecondFilter = $Window.FindName("CheckBoxUseSecondFilter")
    $Global:SecondFilterPanel = $Window.FindName("SecondFilterPanel")

    # Bestehende UI-Elemente
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")
    $Global:DataGridResults = $Window.FindName("DataGridResults")
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")
    $Global:ButtonRefresh = $Window.FindName("ButtonRefresh")
    $Global:ButtonCopy = $Window.FindName("ButtonCopy")

    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    $Global:ButtonQuickRecentlyCreatedUsers = $Window.FindName("ButtonQuickRecentlyCreatedUsers")
    $Global:ButtonQuickPasswordExpiringSoon = $Window.FindName("ButtonQuickPasswordExpiringSoon")
    $Global:ButtonQuickExpiredPasswords = $Window.FindName("ButtonQuickExpiredPasswords")
    $Global:ButtonQuickNeverLoggedOn = $Window.FindName("ButtonQuickNeverLoggedOn")
    $Global:ButtonQuickRecentlyDeletedUsers = $Window.FindName("ButtonQuickRecentlyDeletedUsers")
    $Global:ButtonQuickRecentlyModifiedUsers = $Window.FindName("ButtonQuickRecentlyModifiedUsers")
    $Global:ButtonQuickUsersWithoutManager = $Window.FindName("ButtonQuickUsersWithoutManager")
    $Global:ButtonQuickUsersMissingRequiredAttributes = $Window.FindName("ButtonQuickUsersMissingRequiredAttributes")
    $Global:ButtonQuickUsersDuplicateLogonNames = $Window.FindName("ButtonQuickUsersDuplicateLogonNames")
    $Global:ButtonQuickOrphanedSIDsUsers = $Window.FindName("ButtonQuickOrphanedSIDsUsers")
    
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")

    $Global:ButtonQuickWeakPasswordPolicy = $Window.FindName("ButtonQuickWeakPasswordPolicy")
    $Global:ButtonQuickRiskyGroupMemberships = $Window.FindName("ButtonQuickRiskyGroupMemberships")
    $Global:ButtonQuickPrivilegedAccounts = $Window.FindName("ButtonQuickPrivilegedAccounts")

    $Global:ButtonQuickFSMORoles = $Window.FindName("ButtonQuickFSMORoles")
    $Global:ButtonQuickDCStatus = $Window.FindName("ButtonQuickDCStatus")
    $Global:ButtonQuickReplicationStatus = $Window.FindName("ButtonQuickReplicationStatus")
    $Global:ButtonQuickOUHierarchy = $Window.FindName("ButtonQuickOUHierarchy")
    $Global:ButtonQuickSitesSubnets = $Window.FindName("ButtonQuickSitesSubnets")
    
    $Global:ButtonQuickDepartmentStats = $Window.FindName("ButtonQuickDepartmentStats")
    $Global:ButtonQuickDepartmentSecurity = $Window.FindName("ButtonQuickDepartmentSecurity")
    $Global:ButtonQuickKerberoastable = $Window.FindName("ButtonQuickKerberoastable")
    $Global:ButtonQuickASREPRoastable = $Window.FindName("ButtonQuickASREPRoastable")
    $Global:ButtonQuickDelegation = $Window.FindName("ButtonQuickDelegation")
    $Global:ButtonQuickDCSyncRights = $Window.FindName("ButtonQuickDCSyncRights")
    $Global:ButtonQuickSchemaAdmins = $Window.FindName("ButtonQuickSchemaAdmins")
    $Global:ButtonQuickCertificateAnalysis = $Window.FindName("ButtonQuickCertificateAnalysis")
    $Global:ButtonQuickSYSVOLHealth = $Window.FindName("ButtonQuickSYSVOLHealth")
    $Global:ButtonQuickDNSHealth = $Window.FindName("ButtonQuickDNSHealth")
    $Global:ButtonQuickBackupStatus = $Window.FindName("ButtonQuickBackupStatus")
    $Global:ButtonQuickSchemaAnalysis = $Window.FindName("ButtonQuickSchemaAnalysis")
    $Global:ButtonQuickTrustRelationships = $Window.FindName("ButtonQuickTrustRelationships")
    $Global:ButtonQuickQuotasLimits = $Window.FindName("ButtonQuickQuotasLimits")

    $Global:ButtonQuickStalePasswords = $Window.FindName("ButtonQuickStalePasswords")
    $Global:ButtonQuickNeverChangingPasswords = $Window.FindName("ButtonQuickNeverChangingPasswords")
    $Global:ButtonQuickExpiringAccounts = $Window.FindName("ButtonQuickExpiringAccounts")
    $Global:ButtonQuickReversibleEncryption = $Window.FindName("ButtonQuickReversibleEncryption")
    $Global:ButtonQuickKerberosDES = $Window.FindName("ButtonQuickKerberosDES")
    $Global:ButtonQuickUsersWithSPN = $Window.FindName("ButtonQuickUsersWithSPN")
    $Global:ButtonQuickGuestAccountStatus = $Window.FindName("ButtonQuickGuestAccountStatus")
    $Global:ButtonQuickUsersByDepartment = $Window.FindName("ButtonQuickUsersByDepartment")
    $Global:ButtonQuickUsersByManager = $Window.FindName("ButtonQuickUsersByManager")
    $Global:ButtonQuickMobileDeviceUsers = $Window.FindName("ButtonQuickMobileDeviceUsers")
    $Global:ButtonQuickEmptyGroups = $Window.FindName("ButtonQuickEmptyGroups")
    $Global:ButtonQuickNestedGroups = $Window.FindName("ButtonQuickNestedGroups")
    $Global:ButtonQuickCircularGroups = $Window.FindName("ButtonQuickCircularGroups")
    $Global:ButtonQuickGroupsByTypeScope = $Window.FindName("ButtonQuickGroupsByTypeScope")
    $Global:ButtonQuickDynamicDistGroups = $Window.FindName("ButtonQuickDynamicDistGroups")
    $Global:ButtonQuickMailEnabledGroups = $Window.FindName("ButtonQuickMailEnabledGroups")
    $Global:ButtonQuickGroupsWithoutOwners = $Window.FindName("ButtonQuickGroupsWithoutOwners")
    $Global:ButtonQuickLargeGroups = $Window.FindName("ButtonQuickLargeGroups")
    $Global:ButtonQuickRecentlyModifiedGroups = $Window.FindName("ButtonQuickRecentlyModifiedGroups")
    

    $Global:ButtonQuickOSSummary = $Window.FindName("ButtonQuickOSSummary")
    $Global:ButtonQuickComputersByOSVersion = $Window.FindName("ButtonQuickComputersByOSVersion")
    $Global:ButtonQuickBitLockerStatus = $Window.FindName("ButtonQuickBitLockerStatus")
    $Global:ButtonQuickStaleComputerPasswords = $Window.FindName("ButtonQuickStaleComputerPasswords")
    $Global:ButtonQuickComputersNeverLoggedOn = $Window.FindName("ButtonQuickComputersNeverLoggedOn")
    $Global:ButtonQuickDuplicateComputerNames = $Window.FindName("ButtonQuickDuplicateComputerNames")
    $Global:ButtonQuickComputersByLocation = $Window.FindName("ButtonQuickComputersByLocation")
    
    # Neue Security Audit Buttons aus Roadmap
    $Global:ButtonQuickHoneyTokens = $Window.FindName("ButtonQuickHoneyTokens")
    $Global:ButtonQuickPrivilegeEscalation = $Window.FindName("ButtonQuickPrivilegeEscalation")
    $Global:ButtonQuickExposedCredentials = $Window.FindName("ButtonQuickExposedCredentials")
    $Global:ButtonQuickSuspiciousLogons = $Window.FindName("ButtonQuickSuspiciousLogons")
    $Global:ButtonQuickForeignSecurityPrincipals = $Window.FindName("ButtonQuickForeignSecurityPrincipals")
    $Global:ButtonQuickSIDHistoryAbuse = $Window.FindName("ButtonQuickSIDHistoryAbuse")
    
    # Neue Security Analysis Buttons
    $Global:ButtonQuickCompromiseIndicators = $Window.FindName("ButtonQuickCompromiseIndicators")
    $Global:ButtonQuickSecurityDashboard = $Window.FindName("ButtonQuickSecurityDashboard")
    $Global:ButtonQuickAuthProtocolAnalysis = $Window.FindName("ButtonQuickAuthProtocolAnalysis")
    $Global:ButtonQuickFailedAuthPatterns = $Window.FindName("ButtonQuickFailedAuthPatterns")
    
    # Service Account Buttons
    $Global:ButtonQuickServiceAccountsOverview = $Window.FindName("ButtonQuickServiceAccountsOverview")
    $Global:ButtonQuickManagedServiceAccounts = $Window.FindName("ButtonQuickManagedServiceAccounts")
    $Global:ButtonQuickServiceAccountsSPN = $Window.FindName("ButtonQuickServiceAccountsSPN")
    $Global:ButtonQuickHighPrivServiceAccounts = $Window.FindName("ButtonQuickHighPrivServiceAccounts")
    $Global:ButtonQuickServiceAccountPasswordAge = $Window.FindName("ButtonQuickServiceAccountPasswordAge")
    $Global:ButtonQuickUnusedServiceAccounts = $Window.FindName("ButtonQuickUnusedServiceAccounts")
    
    # GPO & Policy Buttons
    $Global:ButtonQuickGPOOverview = $Window.FindName("ButtonQuickGPOOverview")
    $Global:ButtonQuickUnlinkedGPOs = $Window.FindName("ButtonQuickUnlinkedGPOs")
    $Global:ButtonQuickEmptyGPOs = $Window.FindName("ButtonQuickEmptyGPOs")
    $Global:ButtonQuickGPOPermissions = $Window.FindName("ButtonQuickGPOPermissions")
    $Global:ButtonQuickPasswordPolicySummary = $Window.FindName("ButtonQuickPasswordPolicySummary")
    $Global:ButtonQuickAccountLockoutPolicies = $Window.FindName("ButtonQuickAccountLockoutPolicies")
    $Global:ButtonQuickFineGrainedPasswordPolicies = $Window.FindName("ButtonQuickFineGrainedPasswordPolicies")
    $Global:ButtonQuickConditionalAccessPolicies = $Window.FindName("ButtonQuickConditionalAccessPolicies")
        
    # Advanced Permissions Buttons
    $Global:ButtonQuickACLAnalysis = $Window.FindName("ButtonQuickACLAnalysis")
    $Global:ButtonQuickInheritanceBreaks = $Window.FindName("ButtonQuickInheritanceBreaks")
    $Global:ButtonQuickAdminSDHolderObjects = $Window.FindName("ButtonQuickAdminSDHolderObjects")
    $Global:ButtonQuickAdvancedDelegation = $Window.FindName("ButtonQuickAdvancedDelegation")
    $Global:ButtonQuickSchemaPermissions = $Window.FindName("ButtonQuickSchemaPermissions")
    
    # Attribute selection buttons
    $Global:ButtonSelectAllAttributes = $Window.FindName("ButtonSelectAllAttributes")
    $Global:ButtonSelectNoneAttributes = $Window.FindName("ButtonSelectNoneAttributes")
    $Global:TabControlAttributes = $Window.FindName("TabControlAttributes")
    $Global:ListBoxBasicAttributes = $Window.FindName("ListBoxBasicAttributes")
    $Global:ListBoxSecurityAttributes = $Window.FindName("ListBoxSecurityAttributes")
    $Global:ListBoxExtendedAttributes = $Window.FindName("ListBoxExtendedAttributes")
    
    # Help and About buttons
    $Global:ButtonHelp = $Window.FindName("ButtonHelp")
    $Global:ButtonAbout = $Window.FindName("ButtonAbout")
    
    # Footer elements
    $Global:StatusIndicator = $Window.FindName("StatusIndicator")
    $Global:TextBlockSelectedRows = $Window.FindName("TextBlockSelectedRows")
    $Global:TextBlockLastUpdate = $Window.FindName("TextBlockLastUpdate")

    # Event Handler fÃ¼r erweiterte Filter
    $Global:CheckBoxUseSecondFilter.add_Checked({
        $Global:SecondFilterPanel.IsEnabled = $true
    })
    
    $Global:CheckBoxUseSecondFilter.add_Unchecked({
        $Global:SecondFilterPanel.IsEnabled = $false
    })

    # Helper function to get all selected attributes from all ListBoxes
    function Get-AllSelectedAttributes {
        $selectedAttributes = @()
        
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        
        return $selectedAttributes | Select-Object -Unique
    }

    # Helper function to select specific attributes in the tabbed ListBoxes
    function Select-AttributesInListBoxes {
        param (
            [string[]]$Attributes
        )
        
        # Clear all selections
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        
        # Select requested attributes
        foreach ($attr in $Attributes) {
            # Check in Basic Attributes
            if ($Global:ListBoxBasicAttributes) {
                foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Security Attributes
            if ($Global:ListBoxSecurityAttributes) {
                foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Extended Attributes
            if ($Global:ListBoxExtendedAttributes) {
                foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
        }
    }

    # Helper function to populate attributes in the tabbed ListBoxes
    function Set-AttributesForObjectType {
        param (
            [string]$ObjectType
        )
        
        # Clear all ListBoxes
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        
        switch ($ObjectType) {
            "User" {
                # Basic Attributes
                @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("DisplayName", "SamAccountName", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonTimestamp", "PasswordExpired", "PasswordLastSet", "AccountExpirationDate", "badPwdCount", "lockoutTime", "UserAccountControl", "memberOf") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "Manager", "Company", "physicalDeliveryOfficeName", "telephoneNumber", "homeDirectory", "scriptPath") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Group" {
                # Basic Attributes
                @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "mail") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "SamAccountName")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("ManagedBy", "info", "memberOf", "members") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "distinguishedName", "objectGUID") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Computer" {
                # Basic Attributes
                @("Name", "DNSHostName", "OperatingSystem", "OperatingSystemVersion", "Enabled", "IPv4Address") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "OperatingSystem", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonDate", "PasswordLastSet", "userAccountControl") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "Description", "Location", "ManagedBy", "servicePrincipalName") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
        }
    }

    # RadioButton Event Handler
    $RadioButtonUser.add_Checked({
        Write-ADReportLog -Message "Object type changed to User" -Type Info -Terminal
        
        # Filter-Attribute für Benutzer
        $UserFilterAttributes = @("SamAccountName", "DisplayName", "GivenName", "Surname", "mail", "Department", "Title", "EmployeeID", "UserPrincipalName")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $UserFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "DisplayName"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "User"
        
        $Global:TextBlockStatus.Text = "Ready for user query"
    })

    $RadioButtonGroup.add_Checked({
        Write-ADReportLog -Message "Object type changed to Group" -Type Info -Terminal
        
        # Filter-Attribute für Gruppen
        $GroupFilterAttributes = @("SamAccountName", "Name", "Description", "GroupCategory", "GroupScope")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $GroupFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "Name"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Group"
        
        $Global:TextBlockStatus.Text = "Ready for group query"
    })

    $RadioButtonComputer.add_Checked({
        Write-ADReportLog -Message "Object type changed to Computer" -Type Info -Terminal
        
        # Filter-Attribute für Computer
        $ComputerFilterAttributes = @("Name", "DNSHostName", "OperatingSystem", "Description")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $ComputerFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf Name setzen (da Computer meist über Namen gesucht werden)
        $Global:ComboBoxFilterAttribute1.SelectedItem = "Name"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "DNSHostName"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Computer"
        
        $Global:TextBlockStatus.Text = "Ready for computer query"
    })

    $RadioButtonGroupMemberships.add_Checked({
        Write-ADReportLog -Message "Object type changed to GroupMemberships" -Type Info -Terminal
        
        # Filter-Attribute für Gruppenmitgliedschaften
        $MembershipFilterAttributes = @("SamAccountName", "Name", "DisplayName")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $MembershipFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "Name"
        
        # Clear all ListBoxes for Group Memberships (no attribute selection needed)
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        Write-ADReportLog -Message "Attribute selection disabled for GroupMemberships query." -Type Info
    })

    # Event Handler fÃ¼r ButtonQueryAD anpassen, um Objekttyp zu berÃ¼cksichtigen
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            # Hole Filter-Werte
            $SelectedFilterAttribute = if ($Global:ComboBoxFilterAttribute1.SelectedItem) { $Global:ComboBoxFilterAttribute1.SelectedItem.ToString() } else { "" }
            $FilterValue = $Global:TextBoxFilterValue1.Text
            $FilterOperator = if ($Global:ComboBoxFilterOperator1.SelectedItem) { $Global:ComboBoxFilterOperator1.SelectedItem.Content.ToString() } else { "Contains" }
            
            # Zweiter Filter (wenn aktiviert)
            $UseSecondFilter = $Global:CheckBoxUseSecondFilter.IsChecked
            $SelectedFilterAttribute2 = ""
            $FilterValue2 = ""
            $FilterOperator2 = "Contains"
            $FilterLogic = if ($Global:RadioButtonAnd.IsChecked) { "AND" } else { "OR" }
            
            if ($UseSecondFilter) {
                $SelectedFilterAttribute2 = if ($Global:ComboBoxFilterAttribute2.SelectedItem) { $Global:ComboBoxFilterAttribute2.SelectedItem.ToString() } else { "" }
                $FilterValue2 = $Global:TextBoxFilterValue2.Text
                $FilterOperator2 = if ($Global:ComboBoxFilterOperator2.SelectedItem) { $Global:ComboBoxFilterOperator2.SelectedItem.Content.ToString() } else { "Contains" }
            }
            
            # Get selected attributes from all three ListBoxes
            $SelectedAttributes = Get-AllSelectedAttributes
            # Debug-Info ins Logfile schreiben
            Write-DebugLog "Selected attributes: $($SelectedAttributes -join '; ')" -Category "GUI"
            $isUserSearch = $Global:RadioButtonUser.IsChecked

            if ($SelectedAttributes.Count -eq 0 -and $Global:RadioButtonGroupMemberships.IsChecked -eq $false) {
                [System.Windows.MessageBox]::Show("Please select at least one attribute for export.", "Warnung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            # Bestimme den aktuell ausgewÃ¤hlten Objekttyp
            $ObjectType = if ($Global:RadioButtonUser.IsChecked) { "User" } 
                        elseif ($Global:RadioButtonGroup.IsChecked) { "Group" } 
                        elseif ($Global:RadioButtonGroupMemberships.IsChecked) { "GroupMemberships" } 
                        else { "Computer" }
            
            # AD-Abfrage basierend auf Objekttyp durchfÃ¼hren
            $ReportData = $null
            switch ($ObjectType) {
                "User" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "User"
                }
                "Group" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Group"
                }
                "Computer" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Computer"
                }
                "GroupMemberships" {
                    if (-not ([string]::IsNullOrWhiteSpace($FilterValue)) -and -not ([string]::IsNullOrWhiteSpace($SelectedFilterAttribute))) {
                        $ReportData = Get-ADGroupMembershipsReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue
                    } else {
                        Write-ADReportLog -Message "Filter attribute or value is empty for GroupMemberships query. Please specify a filter." -Type Warning
                        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Filter (Attribut und Wert) fÃ¼r die Mitgliedschaftsabfrage an.", "Hinweis", "OK", "Information")
                        $ReportData = @()
                    }
                }
            }
            
            if ($ReportData) {
                try { # Inner try for processing $ReportData
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null # Clear items first
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() } # Clear columns
                    }
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # ZÃ¤hle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found for $ObjectType." -Type Info
                    
                    # ErgebniszÃ¤hler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                    
                    if ($isUserSearch -and $ReportCollection.Count -eq 1 -and $ReportCollection[0].PSObject.Properties['SamAccountName']) {
                        $userSamAccountName = $ReportCollection[0].SamAccountName
                        
                        # Check if the "Mitgliedschaften" RadioButton is checked
                        if ($Global:RadioButtonGroupMemberships.IsChecked -eq $true) {
                            Write-ADReportLog -Message "Rufe Gruppenmitgliedschaften fÃ¼r Benutzer $($userSamAccountName) ab (RadioButton 'Mitgliedschaften' ist aktiv)..." -Type Info
                            $GroupMemberships = Get-UserGroupMemberships -SamAccountName $userSamAccountName
                            
                            if ($GroupMemberships -and $GroupMemberships.Count -gt 0) {
                                Write-ADReportLog -Message "$($GroupMemberships.Count) Gruppenmitgliedschaften gefunden." -Type Info
                                $DisplayData = $GroupMemberships | Select-Object @{
                                    Name = 'Benutzer';
                                    Expression = {$_.UserDisplayName}
                                }, @{
                                    Name = 'Benutzer (SAM)';
                                    Expression = {$_.UserSamAccountName}
                                }, @{
                                    Name = 'Gruppe';
                                    Expression = {$_.GroupName}
                                }, @{
                                    Name = 'Gruppe (SAM)';
                                    Expression = {$_.GroupSamAccountName}
                                }, @{
                                    Name = 'Gruppentyp';
                                    Expression = {"$($_.GroupCategory) / $($_.GroupScope)"}
                                }, @{
                                    Name = 'Beschreibung';
                                }
                            } else {
                                # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                                Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
                                $Global:DataGridResults.ItemsSource = $ReportCollection
                                $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                                Update-ResultCounters -Results $ReportCollection
                                Update-ResultVisualization -Results $ReportCollection
                            }
                        } else {
                            # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                            Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
                            $Global:DataGridResults.ItemsSource = $ReportCollection
                            $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                            Update-ResultCounters -Results $ReportCollection
                            Update-ResultVisualization -Results $ReportCollection
                        }
                    } else {
                        # This is the original 'else' for cases other than single user search.
                        # It should remain as is, displaying the $ReportCollection.
                        $Global:DataGridResults.ItemsSource = $ReportCollection
                        Update-ResultCounters -Results $ReportCollection
                        Update-ResultVisualization -Results $ReportCollection # Ensure visualization is updated here too.
                    }
                } catch {
                    $ErrorMessage = "Error in query process: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                    }
                    Update-ResultCounters -Results @() # Leeres Array fÃ¼r die ZÃ¤hler
                } 
            } else {
                Write-ADReportLog -Message "No data returned from query for $ObjectType." -Type Info
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "No results found for $ObjectType." }
            }
        } catch { # Outer catch for the whole ButtonQueryAD.add_Click
            $OuterCatchErrorMessage = "An unexpected error occurred during query execution: $($_.Exception.Message)"
            Write-ADReportLog -Message $OuterCatchErrorMessage -Type Error
            try {
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "Error: Query failed." } 
            } catch {
                Write-ADReportLog -Message "CRITICAL: Error within the main query button's outer catch block: $($_.Exception.Message)" -Type Error
            }
        }
    }) # End of .add_Click({
            $ButtonQuickAllUsers.add_Click({
                Write-ADReportLog -Message "Loading all users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""
                    
                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Enabled", "LastLogonDate", "whenCreated", "LockedOut")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "All users loaded."
                    $Global:TextBlockStatus.Text = "All users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDisabledUsers.add_Click({
                Write-ADReportLog -Message "Loading disabled users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes
                    
                    $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Disabled users loaded."
                    $Global:TextBlockStatus.Text = "Disabled users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Disabled users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickLockedUsers.add_Click({
                Write-ADReportLog -Message "Loading locked out users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LockedOut", "LastLogonDate", "BadLogonCount")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende Search-ADAccount statt Get-ADUser fÃ¼r gesperrte Konten (basierend auf Microsoft-Dokumentation)
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue
                    
                    if ($LockedOutAccounts) {
                        # Hole detaillierte Informationen fÃ¼r jeden gesperrten Benutzer
                        $DetailedLockedAccounts = @()
                        foreach ($account in $LockedOutAccounts) {
                            $userDetails = Get-ADUser -Identity $account.SamAccountName -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                            if ($userDetails) {
                                $DetailedLockedAccounts += $userDetails
                            }
                        }
                        
                        if ($DetailedLockedAccounts.Count -gt 0) {
                            Update-ADReportResults -Results $DetailedLockedAccounts -StatusMessage "Locked out users loaded."
                            $Global:TextBlockStatus.Text = "Locked out users loaded. $($DetailedLockedAccounts.Count) record(s) found."
                            Write-ADReportLog -Message "Locked out users loaded. $($DetailedLockedAccounts.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                            $Global:TextBlockStatus.Text = "No locked out users found."
                            Write-ADReportLog -Message "No locked out users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                        $Global:TextBlockStatus.Text = "No locked out users found."
                        Write-ADReportLog -Message "No locked out users found." -Type Info
                    }
                } catch {
                    $ErrorMessage = "Error loading locked out users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverExpire.add_Click({
                Write-ADReportLog -Message "Loading users with password never expires..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordNeverExpires", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users with password never expires loaded."
                    $Global:TextBlockStatus.Text = "Users with password never expires loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users with password never expires loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users with password never expires: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveUsers.add_Click({
                Write-ADReportLog -Message "Loading inactive users (90 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende FileTime-Format fÃ¼r AD-Datumsvergleiche
                    $inactiveThreshold = (Get-Date).AddDays(-90).ToFileTime()
                    # Alternative: Lade alle Benutzer und filtere mit Where-Object
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveUsers = $AllUsers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveUsers -and $InactiveUsers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveUsers -StatusMessage "Inactive users loaded."
                            $Global:TextBlockStatus.Text = "Inactive users loaded. $($InactiveUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive users loaded. $($InactiveUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive users found."
                            $Global:TextBlockStatus.Text = "No users inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickAdminUsers.add_Click({
                Write-ADReportLog -Message "Loading admin users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Verbesserte Methode zum Finden von Admin-Benutzern
                    # Zuerst alle Benutzer laden
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                    
                    # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
                    $AdminGroups = @()
                    $AdminGroups += $Global:ADGroupNames.DomainAdmins
                    $AdminGroups += $Global:ADGroupNames.EnterpriseAdmins
                    $AdminGroups += $Global:ADGroupNames.Administrators
                    $AdminGroups += $Global:ADGroupNames.SchemaAdmins
                    $AdminGroups += @("AAD DC Administrators", "Azure AD-DC-Administratoren")
                    $AdminGroups += $Global:ADGroupNames.ServerOperators
                    $AdminGroups += $Global:ADGroupNames.AccountOperators
                    $AdminGroups += $Global:ADGroupNames.BackupOperators
                    
                    $AdminUsers = @()
                    # PrÃ¼fe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
                    $AdminGroupSIDs = @()
                    foreach ($groupName in $AdminGroups) {
                        try {
                            $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                            if ($group) {
                                $AdminGroupSIDs += $group.SID.Value
                            }
                        } catch {
                            # Ignoriere Fehler bei nicht existierenden Gruppen
                        }
                    }
                    
                    # Nur fortfahren, wenn wir Admin-Gruppen gefunden haben
                    if ($AdminGroupSIDs.Count -gt 0) {
                        # FÃ¼r jeden Benutzer die Gruppenmitgliedschaften prÃ¼fen
                        foreach ($user in $AllUsers) {
                            $memberOf = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                            foreach ($group in $memberOf) {
                                if ($AdminGroupSIDs -contains $group.SID.Value) {
                                    # Benutzer ist Mitglied einer Admin-Gruppe
                                    $adminUser = $user | Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, 
                                        @{Name='AdminGroups'; Expression={
                                            ($memberOf | Where-Object { $AdminGroupSIDs -contains $_.SID.Value } | Select-Object -ExpandProperty Name) -join ", "
                                        }}
                                    $AdminUsers += $adminUser
                                    break # Genug, wenn wir eine Admin-Gruppe gefunden haben
                                }
                            }
                        }
                    }
                    
                    if ($AdminUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:DataGridResults.Columns.Clear()
                        $Global:DataGridResults.ItemsSource = $AdminUsers
                        Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $AdminUsers
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:DataGridResults.Columns.Clear()
                        Write-ADReportLog -Message "No administrator accounts found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading admin users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyCreatedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently created users (30 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1 
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenCreated", "Enabled", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object für Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue |
                        Select-Object DisplayName, SamAccountName, whenCreated, Enabled, mail |
                        Sort-Object whenCreated -Descending # Sortiere nach Erstellungsdatum absteigend
                    
                    if ($AllUsers) {
                        $CreatedDate = (Get-Date).AddDays(-30)
                        $RecentUsers = $AllUsers | Where-Object {
                            $_.whenCreated -and $_.whenCreated -gt $CreatedDate
                        }
                        
                        if ($RecentUsers -and $RecentUsers.Count -gt 0) {
                            Update-ADReportResults -Results $RecentUsers -StatusMessage "Recently created users loaded."
                            $Global:TextBlockStatus.Text = "Recently created users loaded. $($RecentUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently created users loaded. $($RecentUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently created users found."
                            $Global:TextBlockStatus.Text = "No users created in the last 30 days."
                            Write-ADReportLog -Message "No recently created users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently created users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickPasswordExpiringSoon.add_Click({
                Write-ADReportLog -Message "Loading users with password expiring soon (7 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires", "AccountExpirationDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Hole die Domain Password Policy
                    $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
                    $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
                    
                    # Lade aktivierte Benutzer deren Passwort ablaufen kann
                    $AllActiveUsers = Get-ADUser -Filter "PasswordNeverExpires -eq `$false -and Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllActiveUsers) {
                        $ExpiryThreshold = (Get-Date).AddDays(-($MaxPasswordAge - 7))
                        $UsersPasswordExpiring = $AllActiveUsers | Where-Object { 
                            $_.PasswordLastSet -and $_.PasswordLastSet -lt $ExpiryThreshold 
                        }
                        
                        if ($UsersPasswordExpiring -and $UsersPasswordExpiring.Count -gt 0) {
                            Update-ADReportResults -Results $UsersPasswordExpiring -StatusMessage "Users with password expiring soon loaded."
                            $Global:TextBlockStatus.Text = "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) record(s) found."
                            Write-ADReportLog -Message "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No users with password expiring soon found."
                            $Global:TextBlockStatus.Text = "No users with passwords expiring in the next 7 days."
                            Write-ADReportLog -Message "No users with password expiring soon found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with password expiring soon: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGroups.add_Click({
                Write-ADReportLog -Message "Loading all groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All groups loaded."
                    $Global:TextBlockStatus.Text = "All groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickSecurityGroups.add_Click({
                Write-ADReportLog -Message "Loading security groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Security groups loaded."
                    $Global:TextBlockStatus.Text = "Security groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickComputers.add_Click({
                Write-ADReportLog -Message "Loading all computers..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Computer"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All computers loaded."
                    $Global:TextBlockStatus.Text = "All computers loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveComputers.add_Click({
                Write-ADReportLog -Message "Loading inactive computers (90 days)..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "LastLogonDate", "Enabled", "OperatingSystem")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Computer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllComputers = Get-ADComputer -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllComputers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveComputers = $AllComputers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveComputers -StatusMessage "Inactive computers loaded."
                            $Global:TextBlockStatus.Text = "Inactive computers loaded. $($InactiveComputers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive computers found."
                            $Global:TextBlockStatus.Text = "No computers inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive computers found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query computers."
                        $Global:TextBlockStatus.Text = "Failed to query computers."
                        Write-ADReportLog -Message "Failed to query computers." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickExpiredPasswords.add_Click({
                Write-ADReportLog -Message "Loading users with expired passwords..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires", "AccountExpirationDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ExpiredPasswordUsers = Get-ExpiredPasswordUsers
                    if ($ExpiredPasswordUsers -and $ExpiredPasswordUsers.Count -gt 0) {
                        Update-ADReportResults -Results $ExpiredPasswordUsers
                        Write-ADReportLog -Message "Users with expired passwords loaded. $($ExpiredPasswordUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $ExpiredPasswordUsers
                        $Global:TextBlockStatus.Text = "Users with expired passwords loaded. $($ExpiredPasswordUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with expired passwords found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with expired passwords found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with expired passwords: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverLoggedOn.add_Click({
                Write-ADReportLog -Message "Loading users who never logged on..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "whenCreated", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "LastLogonDate -notlike '*'" -SelectedAttributes $QuickReportAttributes -ObjectType "User"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users who never logged on loaded."
                    $Global:TextBlockStatus.Text = "Users who never logged on loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users who never logged on loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users who never logged on: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyDeletedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently deleted users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenDeleted", "isDeleted", "lastKnownParent", "objectClass")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # FÃ¼r gelÃ¶schte Objekte mÃ¼ssen wir den Deleted Objects Container abfragen
                    try {
                        $deletedAfter = (Get-Date).AddDays(-30)
                        $Domain = Get-ADDomain
                        $DeletedObjectsContainer = "CN=Deleted Objects,$($Domain.DistinguishedName)"
                        
                        # Verwende Get-ADObject mit IncludeDeletedObjects
                        $DeletedUsers = Get-ADObject -Filter {(ObjectClass -eq "user") -and (whenDeleted -gt $deletedAfter)} `
                                                    -IncludeDeletedObjects `
                                                    -SearchBase $DeletedObjectsContainer `
                                                    -Properties DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent `
                                                    -ErrorAction SilentlyContinue | 
                                       Select-Object DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent, ObjectClass

                        if ($DeletedUsers) {
                            $Global:DataGridResults.ItemsSource = $DeletedUsers
                            Write-ADReportLog -Message "Recently deleted users loaded. $($DeletedUsers.Count) result(s) found." -Type Info
                            Update-ResultCounters -Results $DeletedUsers
                        } else {
                            $Global:DataGridResults.ItemsSource = $null
                            Write-ADReportLog -Message "No recently deleted users found." -Type Info
                            $Global:TextBlockStatus.Text = "No recently deleted users found in the last 30 days."
                        }
                    } catch {
                        # Fallback wenn keine Berechtigung fÃ¼r Deleted Objects
                        Write-ADReportLog -Message "Cannot access deleted objects. Requires appropriate permissions." -Type Warning
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:TextBlockStatus.Text = "Access to deleted objects denied. Administrator permissions required."
                    }
                } catch {
                    $ErrorMessage = "Error loading recently deleted users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyModifiedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently modified users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenChanged", "whenCreated", "Enabled", "modifyTimeStamp")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $ModifiedDate = (Get-Date).AddDays(-7)
                        $ModifiedUsers = $AllUsers | Where-Object { 
                            $_.whenChanged -and $_.whenChanged -gt $ModifiedDate 
                        }
                        
                        if ($ModifiedUsers -and $ModifiedUsers.Count -gt 0) {
                            Update-ADReportResults -Results $ModifiedUsers -StatusMessage "Recently modified users loaded."
                            $Global:TextBlockStatus.Text = "Recently modified users loaded. $($ModifiedUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently modified users loaded. $($ModifiedUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently modified users found."
                            $Global:TextBlockStatus.Text = "No users modified in the last 7 days."
                            Write-ADReportLog -Message "No recently modified users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently modified users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithoutManager.add_Click({
                Write-ADReportLog -Message "Loading users without manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Department", "Title", "Enabled", "manager")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Manager-Attribut unterstÃ¼tzt nur -eq und -ne Operatoren
                    # Lade alle aktivierten Benutzer und filtere dann nach fehlenden Managern
                    $AllEnabledUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllEnabledUsers) {
                        # Filtere Benutzer ohne Manager
                        $UsersWithoutManager = $AllEnabledUsers | Where-Object { [string]::IsNullOrWhiteSpace($_.manager) }
                        
                        if ($UsersWithoutManager.Count -gt 0) {
                            Update-ADReportResults -Results $UsersWithoutManager -StatusMessage "Users without manager loaded."
                            $Global:TextBlockStatus.Text = "Users without manager loaded. $($UsersWithoutManager.Count) record(s) found."
                            Write-ADReportLog -Message "Users without manager loaded. $($UsersWithoutManager.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "All users have a manager assigned."
                            $Global:TextBlockStatus.Text = "All users have a manager assigned."
                            Write-ADReportLog -Message "All users have a manager assigned." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No enabled users found."
                        $Global:TextBlockStatus.Text = "No enabled users found."
                        Write-ADReportLog -Message "No enabled users found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users without manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersMissingRequiredAttributes.add_Click({
                Write-ADReportLog -Message "Loading users missing required attributes..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer mit wichtigen Attributen
                    $AllUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties DisplayName, SamAccountName, mail, telephoneNumber, Department, Title, manager, Enabled
                    
                    # Filtere Benutzer mit fehlenden Attributen
                    $UsersWithMissingAttributes = @()
                    foreach ($user in $AllUsers) {
                        $missingAttributes = @()
                        
                        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { $missingAttributes += "DisplayName" }
                        if ([string]::IsNullOrWhiteSpace($user.mail)) { $missingAttributes += "Email" }
                        if ([string]::IsNullOrWhiteSpace($user.telephoneNumber)) { $missingAttributes += "Phone" }
                        if ([string]::IsNullOrWhiteSpace($user.Department)) { $missingAttributes += "Department" }
                        if ([string]::IsNullOrWhiteSpace($user.Title)) { $missingAttributes += "Title" }
                        if ([string]::IsNullOrWhiteSpace($user.manager)) { $missingAttributes += "Manager" }
                        
                        if ($missingAttributes.Count -gt 0) {
                            $UsersWithMissingAttributes += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                mail = $user.mail
                                Department = $user.Department
                                Title = $user.Title
                                Enabled = $user.Enabled
                                MissingAttributes = $missingAttributes -join ", "
                                MissingCount = $missingAttributes.Count
                            }
                        }
                    }
                    
                    if ($UsersWithMissingAttributes.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $UsersWithMissingAttributes | Sort-Object MissingCount -Descending
                        Write-ADReportLog -Message "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $UsersWithMissingAttributes
                        $Global:TextBlockStatus.Text = "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with missing required attributes found." -Type Info
                        $Global:TextBlockStatus.Text = "All active users have required attributes filled."
                    }
                } catch {
                    $ErrorMessage = "Error loading users missing required attributes: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersDuplicateLogonNames.add_Click({
                Write-ADReportLog -Message "Loading users with duplicate logon names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled, whenCreated
                    
                    # Gruppiere nach SamAccountName um Duplikate zu finden
                    $DuplicateSamAccounts = $AllUsers | Group-Object SamAccountName | Where-Object { $_.Count -gt 1 }
                    
                    # Gruppiere auch nach UserPrincipalName
                    $DuplicateUPNs = $AllUsers | Where-Object { $_.UserPrincipalName } | Group-Object UserPrincipalName | Where-Object { $_.Count -gt 1 }
                    
                    $DuplicateUsers = @()
                    
                    # Verarbeite SamAccountName Duplikate
                    foreach ($group in $DuplicateSamAccounts) {
                        foreach ($user in $group.Group) {
                            $DuplicateUsers += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                UserPrincipalName = $user.UserPrincipalName
                                Enabled = $user.Enabled
                                whenCreated = $user.whenCreated
                                DuplicateType = "SamAccountName"
                                DuplicateCount = $group.Count
                            }
                        }
                    }
                    
                    # Verarbeite UPN Duplikate
                    foreach ($group in $DuplicateUPNs) {
                        foreach ($user in $group.Group) {
                            # PrÃ¼fe ob dieser User nicht schon als SamAccountName Duplikat erfasst wurde
                            if (-not ($DuplicateUsers | Where-Object { $_.SamAccountName -eq $user.SamAccountName })) {
                                $DuplicateUsers += [PSCustomObject]@{
                                    DisplayName = $user.DisplayName
                                    SamAccountName = $user.SamAccountName
                                    UserPrincipalName = $user.UserPrincipalName
                                    Enabled = $user.Enabled
                                    whenCreated = $user.whenCreated
                                    DuplicateType = "UserPrincipalName"
                                    DuplicateCount = $group.Count
                                }
                            }
                        }
                    }
                    
                    if ($DuplicateUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $DuplicateUsers | Sort-Object SamAccountName
                        Write-ADReportLog -Message "Users with duplicate logon names loaded. $($DuplicateUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $DuplicateUsers
                        $Global:TextBlockStatus.Text = "Users with duplicate logon names loaded. $($DuplicateUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with duplicate logon names found." -Type Info
                        $Global:TextBlockStatus.Text = "No duplicate logon names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with duplicate logon names: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickOrphanedSIDsUsers.add_Click({
                Write-ADReportLog -Message "Loading orphaned SIDs (Foreign Security Principals)..." -Type Info
                try {
                    # Lade Foreign Security Principals
                    $Domain = Get-ADDomain
                    $ForeignSecurityPrincipals = Get-ADObject -Filter * -SearchBase "CN=ForeignSecurityPrincipals,$($Domain.DistinguishedName)" -Properties Name, ObjectSID, whenCreated, whenChanged
                    
                    $OrphanedSIDs = @()
                    
                    foreach ($fsp in $ForeignSecurityPrincipals) {
                        # Versuche den SID aufzulÃ¶sen
                        $resolved = $false
                        $resolvedName = "Unknown"
                        $sidString = $fsp.Name
                        
                        try {
                            # Versuche SID zu einem Namen aufzulÃ¶sen
                            $sid = New-Object System.Security.Principal.SecurityIdentifier($sidString)
                            $resolvedName = $sid.Translate([System.Security.Principal.NTAccount]).Value
                            $resolved = $true
                        } catch {
                            # SID konnte nicht aufgelÃ¶st werden - wahrscheinlich verwaist
                            $resolved = $false
                        }
                        
                        if (-not $resolved) {
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Orphaned"
                                ResolvedName = "Cannot resolve"
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        } else {
                            # Optional: Auch aufgelÃ¶ste SIDs anzeigen
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Active"
                                ResolvedName = $resolvedName
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        }
                    }
                    
                    # Filtere nur die verwaisten SIDs
                    $OrphanedOnly = $OrphanedSIDs | Where-Object { $_.Status -eq "Orphaned" }
                    
                    if ($OrphanedOnly.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OrphanedOnly
                        Write-ADReportLog -Message "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs found." -Type Info
                        Update-ResultCounters -Results $OrphanedOnly
                        $Global:TextBlockStatus.Text = "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No orphaned SIDs found. All $($OrphanedSIDs.Count) FSPs can be resolved." -Type Info
                        $Global:TextBlockStatus.Text = "No orphaned SIDs found. All Foreign Security Principals are valid."
                    }
                } catch {
                    $ErrorMessage = "Error loading orphaned SIDs: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickReversibleEncryption.add_Click({
                Write-ADReportLog -Message "Loading users with reversible encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $ReversibleUsers = Get-ReversibleEncryptionUsers
                    if ($ReversibleUsers -and $ReversibleUsers.Count -gt 0) {
                        Update-ADReportResults -Results $ReversibleUsers
                        Write-ADReportLog -Message "Users with reversible encryption loaded. $($ReversibleUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with reversible encryption loaded. $($ReversibleUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with reversible encryption found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with reversible encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with reversible encryption: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickKerberosDES.add_Click({
                Write-ADReportLog -Message "Loading users with Kerberos DES encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $DESUsers = Get-KerberosDESUsers
                    if ($DESUsers -and $DESUsers.Count -gt 0) {
                        Update-ADReportResults -Results $DESUsers
                        Write-ADReportLog -Message "Users with Kerberos DES loaded. $($DESUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with Kerberos DES loaded. $($DESUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with Kerberos DES found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Kerberos DES encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with Kerberos DES: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithSPN.add_Click({
                Write-ADReportLog -Message "Loading users with Service Principal Names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $SPNUsers = Get-UsersWithSPN
                    if ($SPNUsers -and $SPNUsers.Count -gt 0) {
                        Update-ADReportResults -Results $SPNUsers
                        Write-ADReportLog -Message "Users with SPN loaded. $($SPNUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with SPN loaded. $($SPNUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with SPN found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Service Principal Names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with SPN: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGuestAccountStatus.add_Click({
                Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $GuestStatus = Get-GuestAccountStatus
                    if ($GuestStatus -and $GuestStatus.Count -gt 0) {
                        Update-ADReportResults -Results $GuestStatus
                        Write-ADReportLog -Message "Guest account analysis completed. $($GuestStatus.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Guest account analysis completed. $($GuestStatus.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No Guest account information found." -Type Info
                        $Global:TextBlockStatus.Text = "No Guest account found."
                    }
                } catch {
                    $ErrorMessage = "Error analyzing Guest account: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByDepartment.add_Click({
                Write-ADReportLog -Message "Loading users by department..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByDept = Get-UsersByDepartment
                    if ($UsersByDept -and $UsersByDept.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByDept
                        Write-ADReportLog -Message "Users by department loaded. $($UsersByDept.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by department loaded. $($UsersByDept.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No department data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with department information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by department: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByManager.add_Click({
                Write-ADReportLog -Message "Loading users by manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByMgr = Get-UsersByManager
                    if ($UsersByMgr -and $UsersByMgr.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByMgr
                        Write-ADReportLog -Message "Users by manager loaded. $($UsersByMgr.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by manager loaded. $($UsersByMgr.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No manager data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with manager information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDistributionGroups.add_Click({
                Write-ADReportLog -Message "Loading distribution groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Distribution groups loaded."
                    $Global:TextBlockStatus.Text = "Distribution groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Distribution groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading distribution groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickWeakPasswordPolicy.add_Click({
                Write-ADReportLog -Message "Loading users with weak password policies..." -Type Info
                try {
                    $WeakPasswordUsers = Get-WeakPasswordPolicyUsers
                    if ($WeakPasswordUsers -and $WeakPasswordUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $WeakPasswordUsers
                        Write-ADReportLog -Message "Users with weak password policies loaded. $($WeakPasswordUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $WeakPasswordUsers
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with weak password policies found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with weak password policies: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickRiskyGroupMemberships.add_Click({
                Write-ADReportLog -Message "Loading risky group memberships..." -Type Info
                try {
                    $RiskyMemberships = Get-RiskyGroupMemberships
                    if ($RiskyMemberships -and $RiskyMemberships.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $RiskyMemberships
                        Write-ADReportLog -Message "Risky group memberships loaded. $($RiskyMemberships.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $RiskyMemberships
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No risky group memberships found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading risky group memberships: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickPrivilegedAccounts.add_Click({
                Write-ADReportLog -Message "Loading privileged accounts..." -Type Info
                try {
                    $PrivilegedAccounts = Get-PrivilegedAccounts
                    if ($PrivilegedAccounts -and $PrivilegedAccounts.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $PrivilegedAccounts
                        Write-ADReportLog -Message "Privileged accounts loaded. $($PrivilegedAccounts.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $PrivilegedAccounts
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No privileged accounts found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading privileged accounts: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickOUHierarchy.add_Click({
                Write-ADReportLog -Message "Loading OU hierarchy..." -Type Info
                try {
                    $OUHierarchy = Get-ADOUHierarchyReport
                    if ($OUHierarchy -and $OUHierarchy.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OUHierarchy
                        Write-ADReportLog -Message "OU hierarchy loaded. $($OUHierarchy.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $OUHierarchy
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No OU hierarchy information found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading OU hierarchy: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickSitesSubnets.add_Click({
                Write-ADReportLog -Message "Loading AD sites and subnets..." -Type Info
                try {
                    $SitesSubnets = @(Get-ADSitesAndSubnetsReport)
                    if ($SitesSubnets -and $SitesSubnets.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $SitesSubnets
                        Write-ADReportLog -Message "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $SitesSubnets
                        $Global:TextBlockStatus.Text = "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No AD sites and subnets information found." -Type Warning
                        $Global:TextBlockStatus.Text = "No AD sites and subnets information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading AD sites and subnets: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

    # --- AD-Health Event Handler ---
    $ButtonQuickFSMORoles.add_Click({
        Write-ADReportLog -Message "Loading FSMO role holders..." -Type Info
        try {
            $FSMORoles = Get-FSMORoles
            if ($FSMORoles -and $FSMORoles.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $FSMORoles
                Write-ADReportLog -Message "FSMO role holders loaded. $($FSMORoles.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FSMORoles
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No FSMO role information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading FSMO role holders: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDCStatus.add_Click({
        Write-ADReportLog -Message "Loading domain controller status..." -Type Info
        try {
            $DomainControllers = Get-DomainControllerStatus
            if ($DomainControllers -and $DomainControllers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DomainControllers
                Write-ADReportLog -Message "Domain controller status loaded. $($DomainControllers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $DomainControllers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No domain controllers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading domain controller status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickReplicationStatus.add_Click({
        Write-ADReportLog -Message "Loading AD replication status..." -Type Info
        try {
            $ReplicationStatus = Get-ReplicationStatus
            if ($ReplicationStatus -and $ReplicationStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ReplicationStatus
                Write-ADReportLog -Message "AD replication status loaded. $($ReplicationStatus.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReplicationStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No replication status information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading AD replication status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # --- Event Handler fÃ¼r neue erweiterte Reports ---
    
    # Organisationsstruktur-Reports
    if ($null -ne $Global:ButtonQuickDepartmentStats) {
        $ButtonQuickDepartmentStats.add_Click({
        Write-ADReportLog -Message "Loading department statistics..." -Type Info
        try {
            $DeptStats = Get-DepartmentStatistics
            if ($DeptStats -and $DeptStats.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptStats
                Write-ADReportLog -Message "Department statistics loaded. $($DeptStats.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptStats
                $Global:TextBlockStatus.Text = "Department statistics loaded. $($DeptStats.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department statistics found." -Type Warning
                $Global:TextBlockStatus.Text = "No department statistics found."
            }
        } catch {
            $ErrorMessage = "Error loading department statistics: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDepartmentSecurity) {
        $ButtonQuickDepartmentSecurity.add_Click({
        Write-ADReportLog -Message "Loading department security analysis..." -Type Info
        try {
            $DeptSecurity = Get-DepartmentSecurityRisks
            if ($DeptSecurity -and $DeptSecurity.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptSecurity
                Write-ADReportLog -Message "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptSecurity
                $Global:TextBlockStatus.Text = "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department security data found." -Type Warning
                $Global:TextBlockStatus.Text = "No department security data found."
            }
        } catch {
            $ErrorMessage = "Error loading department security analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    # Kerberos Security
    if ($null -ne $Global:ButtonQuickKerberoastable) {
        $ButtonQuickKerberoastable.add_Click({
        Write-ADReportLog -Message "Loading Kerberoastable accounts..." -Type Info
        try {
            $Kerberoastable = Get-KerberoastableAccounts
            if ($Kerberoastable -and $Kerberoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Kerberoastable
                Write-ADReportLog -Message "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $Kerberoastable
                $Global:TextBlockStatus.Text = "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) with SPNs found."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Kerberoastable accounts found." -Type Info
                $Global:TextBlockStatus.Text = "No Kerberoastable accounts (users with SPNs) found."
            }
        } catch {
            $ErrorMessage = "Error loading Kerberoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickASREPRoastable) {
        $ButtonQuickASREPRoastable.add_Click({
        Write-ADReportLog -Message "Loading ASREPRoastable accounts..." -Type Info
        try {
            $ASREPRoastable = Get-ASREPRoastableAccounts
            if ($ASREPRoastable -and $ASREPRoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ASREPRoastable
                Write-ADReportLog -Message "ASREPRoastable accounts loaded. $($ASREPRoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $ASREPRoastable
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No ASREPRoastable accounts found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading ASREPRoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDelegation) {
        $ButtonQuickDelegation.add_Click({
        Write-ADReportLog -Message "Loading delegation analysis..." -Type Info
        try {
            $Delegation = Get-DelegationAnalysis
            if ($Delegation -and $Delegation.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Delegation
                Write-ADReportLog -Message "Delegation analysis loaded. $($Delegation.Count) delegated object(s) found." -Type Info
                Update-ResultCounters -Results $Delegation
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No delegation settings found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading delegation analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    # Advanced Security
    if ($null -ne $Global:ButtonQuickDCSyncRights) {
        $ButtonQuickDCSyncRights.add_Click({
        Write-ADReportLog -Message "Loading DCSync rights analysis..." -Type Info
        try {
            $DCSyncRights = Get-DCSyncRights
            if ($DCSyncRights -and $DCSyncRights.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DCSyncRights
                Write-ADReportLog -Message "DCSync rights analysis loaded. $($DCSyncRights.Count) identities with DCSync found." -Type Info
                Update-ResultCounters -Results $DCSyncRights
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DCSync rights found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading DCSync rights: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    } else {
        Write-ADReportLog -Message "ButtonQuickDCSyncRights nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
    }

    # Advanced Security
    try {
        if ($null -ne $Global:ButtonQuickSchemaAdmins) {
            $ButtonQuickSchemaAdmins.add_Click({
        Write-ADReportLog -Message "Loading Schema Admin paths..." -Type Info
        try {
            $SchemaAdmins = Get-SchemaAdminPaths
            if ($SchemaAdmins -and $SchemaAdmins.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SchemaAdmins
                Write-ADReportLog -Message "Schema Admin paths loaded. $($SchemaAdmins.Count) path(s) found." -Type Info
                Update-ResultCounters -Results $SchemaAdmins
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Schema Admin paths found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading Schema Admin paths: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAdmins nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAdmins: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickCertificateAnalysis) {
            $ButtonQuickCertificateAnalysis.add_Click({
        Write-ADReportLog -Message "Loading certificate security analysis..." -Type Info
        try {
            $CertAnalysis = Get-CertificateSecurityAnalysis
            if ($CertAnalysis -and $CertAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $CertAnalysis
                Write-ADReportLog -Message "Certificate security analysis loaded. $($CertAnalysis.Count) finding(s)." -Type Info
                Update-ResultCounters -Results $CertAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No certificate security findings." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading certificate analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickCertificateAnalysis nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickCertificateAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Advanced Monitoring
    try {
        if ($null -ne $Global:ButtonQuickSYSVOLHealth) {
            $ButtonQuickSYSVOLHealth.add_Click({
        Write-ADReportLog -Message "Loading SYSVOL health check..." -Type Info
        try {
            $SYSVOLHealth = @(Get-SYSVOLHealthCheck)
            if ($SYSVOLHealth -and $SYSVOLHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SYSVOLHealth
                Write-ADReportLog -Message "SYSVOL health check loaded. $($SYSVOLHealth.Count) DC(s) checked." -Type Info
                Update-ResultCounters -Results $SYSVOLHealth
                $Global:TextBlockStatus.Text = "SYSVOL health check loaded. $($SYSVOLHealth.Count) domain controller(s) checked."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No SYSVOL health data found." -Type Warning
                $Global:TextBlockStatus.Text = "No SYSVOL health data found."
            }
        } catch {
            $ErrorMessage = "Error loading SYSVOL health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSYSVOLHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSYSVOLHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickDNSHealth) {
            $ButtonQuickDNSHealth.add_Click({
        Write-ADReportLog -Message "Loading DNS health analysis..." -Type Info
        try {
            $DNSHealth = Get-DNSHealthAnalysis
            if ($DNSHealth -and $DNSHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DNSHealth
                Write-ADReportLog -Message "DNS health analysis loaded. $($DNSHealth.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $DNSHealth
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DNS health data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading DNS health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickDNSHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickDNSHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickBackupStatus) {
            $ButtonQuickBackupStatus.add_Click({
        Write-ADReportLog -Message "Loading backup readiness status..." -Type Info
        try {
            $BackupStatus = Get-BackupReadinessStatus
            if ($BackupStatus -and $BackupStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $BackupStatus
                Write-ADReportLog -Message "Backup readiness status loaded. $($BackupStatus.Count) check(s) performed." -Type Info
                Update-ResultCounters -Results $BackupStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No backup status data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading backup status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickBackupStatus nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickBackupStatus: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Schema & Trusts
    try {
        if ($null -ne $Global:ButtonQuickSchemaAnalysis) {
            $ButtonQuickSchemaAnalysis.add_Click({
                Write-ADReportLog -Message "Loading schema analysis..." -Type Info
                try {
                    $SchemaAnalysis = Get-SchemaAnalysis
                    if ($SchemaAnalysis -and $SchemaAnalysis.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $SchemaAnalysis
                        Write-ADReportLog -Message "Schema analysis loaded. $($SchemaAnalysis.Count) item(s) found." -Type Info
                        Update-ResultCounters -Results $SchemaAnalysis
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No schema data found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading schema analysis: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAnalysis nicht gefunden - Funktion wird übersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickTrustRelationships) {
            $ButtonQuickTrustRelationships.add_Click({
        Write-ADReportLog -Message "Loading trust relationship analysis..." -Type Info
        try {
            $TrustAnalysis = Get-TrustRelationshipAnalysis
            if ($TrustAnalysis -and $TrustAnalysis.Count -gt 0) {
                Update-ADReportResults -Results $TrustAnalysis
                Write-ADReportLog -Message "Trust relationship analysis loaded. $($TrustAnalysis.Count) trust(s) analyzed." -Type Info
                Update-ResultCounters -Results $TrustAnalysis
                $Global:TextBlockStatus.Text = "Trust relationship analysis loaded. $($TrustAnalysis.Count) trust(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No trust relationships found." -Type Info
                $Global:TextBlockStatus.Text = "No trust relationships found."
            }
        } catch {
            $ErrorMessage = "Error loading trust analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            Update-ADReportResults -Results @()
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickTrustRelationships nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickTrustRelationships: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickQuotasLimits) {
            $ButtonQuickQuotasLimits.add_Click({
        Write-ADReportLog -Message "Loading quotas and limits analysis..." -Type Info
        try {
            $QuotasLimits = Get-QuotasAndLimits
            if ($QuotasLimits -and $QuotasLimits.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $QuotasLimits
                Write-ADReportLog -Message "Quotas and limits analysis loaded. $($QuotasLimits.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $QuotasLimits
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No quota data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading quotas and limits: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickQuotasLimits nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickQuotasLimits: $($_.Exception.Message)" -Type Warning -Terminal
    }

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information") | Out-Null
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in CSV export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV export canceled by user." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "Preparing HTML export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "No data available for export."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML File (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "Save HTML file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # HTML Design optimiert für DIN A4 Querformat
                $HtmlHead = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Active Directory Report</title>
<style>
  @media print {
    @page {
      size: A4 landscape;
      margin: 2cm;
    }
    body {
      width: 29.7cm;
      height: 21cm;
      margin: 0 auto;
    }
  }
  
  :root {
    --primary-color: #0078D4;
    --border-color: #E1E1E1;
  }
  
  body { 
    font-family: 'Segoe UI', Arial, sans-serif;
    margin: 0;
    padding: 15px;
    background: white;
    color: #333;
    line-height: 1.4;
    font-size: 10pt;
    max-width: 29.7cm;
    min-height: 21cm;
    margin: 0 auto;
    box-sizing: border-box;
  }
  
  .container {
    width: 100%;
    min-height: calc(21cm - 4cm); /* A4 height minus margins */
    margin: 0 auto;
    background: white;
    display: flex;
    flex-direction: column;
  }
  
  table { 
    border-collapse: collapse;
    width: 100%;
    margin: 15px 0;
    page-break-inside: auto;
  }
  
  tr { 
    page-break-inside: avoid;
    page-break-after: auto;
  }
  
  th, td { 
    padding: 6px 8px;
    text-align: left;
    border: 1px solid var(--border-color);
    font-size: 9pt;
    word-wrap: break-word;
    max-width: 250px;
  }
  
  th {
    background-color: var(--primary-color);
    color: white;
    font-weight: 600;
  }
  
  tr:nth-child(even) { 
    background-color: #F8F9FA;
  }
  
  h1 { 
    text-align: center;
    color: var(--primary-color);
    font-size: 16pt;
    margin: 10px 0 20px 0;
    padding-bottom: 10px;
    border-bottom: 2px solid var(--border-color);
  }
  
  .timestamp {
    text-align: right;
    color: #666;
    font-size: 8pt;
    margin: 10px 0;
  }

  .footer {
    text-align: center;
    border-top: 1px solid var(--border-color);
    padding-top: 10px;
    margin-top: auto;
    font-size: 8pt;
    color: #666;
    width: 100%;
    position: relative;
    bottom: 0;
  }

  @media print {
    .container {
      width: 100%;
      max-width: none;
    }
    
    table {
      font-size: 9pt;
    }
    
    th, td {
      padding: 4px 6px;
    }
  }
</style>
"@
                $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                $ReportTitle = "Active Directory Report"
                
                $HtmlBody = @"
<div class="container">
    <h1>$ReportTitle</h1>
    <div class="timestamp">Erstellt am $DateTimeNow mit AD Reporting Tool</div>
"@

                $HtmlFooter = @"
</div>
<div class="footer">
    <p>Copyright © $(Get-Date -Format "yyyy") |  Autor: Andreas Hepp | Webseite: <a href="https://www.phinit.de">www.phinit.de</a> / <a href="https://www.psscripts.de">www.psscripts.de</a></p>
</div>
"@

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body $HtmlBody | ForEach-Object {
                    $_ -replace '</body>', "$HtmlFooter</body>"
                } | Out-File -FilePath $FilePath -Encoding UTF8
                
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in HTML export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML export canceled by user." -Type Info
        }
    })

    # --- Event Handler fÃ¼r Attribute Selection Buttons ---
    if ($null -ne $Global:ButtonSelectAllAttributes) {
        $ButtonSelectAllAttributes.add_Click({
            Write-ADReportLog -Message "Selecting all attributes..." -Type Info
            try {
                # Select all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                Write-ADReportLog -Message "All attributes selected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error selecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    if ($null -ne $Global:ButtonSelectNoneAttributes) {
        $ButtonSelectNoneAttributes.add_Click({
            Write-ADReportLog -Message "Deselecting all attributes..." -Type Info
            try {
                # Deselect all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                Write-ADReportLog -Message "All attributes deselected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error deselecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    # --- Event Handler für Refresh und Copy Buttons ---
    if ($null -ne $Global:ButtonRefresh) {
        $ButtonRefresh.add_Click({
            Write-ADReportLog -Message "Setze Query-Fenster zurück..." -Type Info
            try {
                # Setze RadioButtons zurück
                if ($Global:RadioButtonUser -and $Global:RadioButtonUser.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonUser.IsChecked = $true
                }
                if ($Global:RadioButtonGroup -and $Global:RadioButtonGroup.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonGroup.IsChecked = $false
                }
                if ($Global:RadioButtonComputer -and $Global:RadioButtonComputer.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonComputer.IsChecked = $false
                }
                if ($Global:RadioButtonGroupMembership -and $Global:RadioButtonGroupMembership.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonGroupMembership.IsChecked = $false
                }

                # Leere alle ListBoxen
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }

                # Leere das DataGrid
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                }

                Write-ADReportLog -Message "Query-Fenster erfolgreich zurückgesetzt." -Type Info
            }
            catch {
                Write-ADReportLog -Message "Fehler beim Zurücksetzen des Query-Fensters: $($_.Exception.Message)" -Type Error
            }
        })
    }

    if ($null -ne $Global:ButtonCopy) {
        $ButtonCopy.add_Click({
            Write-ADReportLog -Message "Prüfe auf markierte Zeilen..." -Type Info
            try {
                # Prüfe ob Zeilen markiert sind
                $selectedItems = $Global:DataGridResults.SelectedItems
                
                if ($selectedItems.Count -gt 0) {
                    Write-ADReportLog -Message "Kopiere markierte Zeilen in die Zwischenablage..." -Type Info
                    
                    # Konvertiere ausgewählte Zeilen in tabulierte Textform
                    $clipboardText = ""
                    
                    # Header
                    $columns = $Global:DataGridResults.Columns | Where-Object {$_.Visibility -eq 'Visible'}
                    $clipboardText += $columns.Header -join "`t"
                    $clipboardText += "`r`n"
                    
                    # Nur markierte Zeilen
                    foreach ($item in $selectedItems) {
                        $rowData = @()
                        foreach ($column in $columns) {
                            $cellValue = $item.$($column.Header)
                            if ($null -eq $cellValue) { $cellValue = "" }
                            $rowData += $cellValue.ToString()
                        }
                        $clipboardText += $rowData -join "`t"
                        $clipboardText += "`r`n"
                    }
                    
                    # In Zwischenablage kopieren
                    [System.Windows.Clipboard]::SetText($clipboardText)
                    Write-ADReportLog -Message "Markierte Zeilen erfolgreich in die Zwischenablage kopiert." -Type Info
                }
                else {
                    [System.Windows.MessageBox]::Show(
                        "Bitte markieren Sie mindestens eine Zeile zum Kopieren.",
                        "Keine Auswahl",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    Write-ADReportLog -Message "Keine Zeilen markiert zum Kopieren." -Type Warning
                }
            }
            catch {
                Write-ADReportLog -Message "Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)" -Type Error
            }
        })
    }


    # --- Event Handler fÃ¼r Help and About Buttons ---
    if ($null -ne $Global:ButtonHelp) {
        $ButtonHelp.add_Click({
            Write-ADReportLog -Message "Showing help dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "AD Reporting Tool Help:`n`n" +
                "1. Select Object Type: Choose between Users, Groups, Computers, or Group Memberships`n`n" +
                "2. Set Filters: Configure search filters using attribute, operator, and value`n`n" +
                "3. Select Attributes: Choose which attributes to include in the report`n`n" +
                "4. Quick Reports: Use predefined reports for common AD queries`n`n" +
                "5. Export Options: Export results to CSV or HTML format`n`n" +
                "For more information, visit the documentation.",
                "Help",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    if ($null -ne $Global:ButtonAbout) {
        $ButtonAbout.add_Click({
            Write-ADReportLog -Message "Showing about dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "AD Reporting Tool Version 1.2" +
                "A comprehensive Active Directory reporting tool`n`n" +
                "Features:`n" +
                "â€¢ Advanced filtering and search capabilities`n" +
                "â€¢ Multiple export formats (CSV, HTML)`n" +
                "â€¢ Quick report templates for common scenarios`n" +
                "â€¢ Security and compliance reports`n" +
                "â€¢ Multi-language support`n`n" +
                "Â© 2024 Your Organization",
                "About AD Reporting Tool",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    # --- Event Handler fÃ¼r neue Quick Report Buttons ---
    
    # Event Handler fÃ¼r ButtonQuickStalePasswords
    if ($null -ne $Global:ButtonQuickStalePasswords) {
        $ButtonQuickStalePasswords.add_Click({
            Write-ADReportLog -Message "Loading users with stale passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $StalePasswords = Get-StalePasswords -Days 90
                if ($StalePasswords -and $StalePasswords.Count -gt 0) {
                    Update-ADReportResults -Results $StalePasswords
                    Write-ADReportLog -Message "Stale passwords loaded. $($StalePasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $StalePasswords
                    $Global:TextBlockStatus.Text = "Stale passwords loaded. $($StalePasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with stale passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with stale passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading stale passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickNeverChangingPasswords
    if ($null -ne $Global:ButtonQuickNeverChangingPasswords) {
        $ButtonQuickNeverChangingPasswords.add_Click({
            Write-ADReportLog -Message "Loading users with never changing passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $NeverChangingPasswords = Get-NeverChangingPasswords -Days 365
                if ($NeverChangingPasswords -and $NeverChangingPasswords.Count -gt 0) {
                    Update-ADReportResults -Results $NeverChangingPasswords
                    Write-ADReportLog -Message "Never changing passwords loaded. $($NeverChangingPasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NeverChangingPasswords
                    $Global:TextBlockStatus.Text = "Never changing passwords loaded. $($NeverChangingPasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with never changing passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with never changing passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading never changing passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickEmptyGroups
    if ($null -ne $Global:ButtonQuickEmptyGroups) {
        $ButtonQuickEmptyGroups.add_Click({
            Write-ADReportLog -Message "Loading empty groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $EmptyGroups = Get-EmptyGroups
                if ($EmptyGroups -and $EmptyGroups.Count -gt 0) {
                    Update-ADReportResults -Results $EmptyGroups
                    Write-ADReportLog -Message "Empty groups loaded. $($EmptyGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $EmptyGroups
                    $Global:TextBlockStatus.Text = "Empty groups loaded. $($EmptyGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No empty groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No empty groups found."
                }
            } catch {
                $ErrorMessage = "Error loading empty groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickNestedGroups
    if ($null -ne $Global:ButtonQuickNestedGroups) {
        $ButtonQuickNestedGroups.add_Click({
            Write-ADReportLog -Message "Loading nested groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $NestedGroups = Get-NestedGroups
                if ($NestedGroups -and $NestedGroups.Count -gt 0) {
                    Update-ADReportResults -Results $NestedGroups
                    Write-ADReportLog -Message "Nested groups loaded. $($NestedGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NestedGroups
                    $Global:TextBlockStatus.Text = "Nested groups loaded. $($NestedGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No nested groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No nested groups found."
                }
            } catch {
                $ErrorMessage = "Error loading nested groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickCircularGroups
    if ($null -ne $Global:ButtonQuickCircularGroups) {
        $ButtonQuickCircularGroups.add_Click({
            Write-ADReportLog -Message "Loading circular nested groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $CircularGroups = Get-CircularNestedGroups
                if ($CircularGroups -and $CircularGroups.Count -gt 0) {
                    Update-ADReportResults -Results $CircularGroups
                    Write-ADReportLog -Message "Circular nested groups loaded. $($CircularGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $CircularGroups
                    $Global:TextBlockStatus.Text = "Circular nested groups loaded. $($CircularGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No circular nested groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No circular nested groups found."
                }
            } catch {
                $ErrorMessage = "Error loading circular nested groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGroupsByTypeScope
    if ($null -ne $Global:ButtonQuickGroupsByTypeScope) {
        $ButtonQuickGroupsByTypeScope.add_Click({
            Write-ADReportLog -Message "Loading groups by type and scope..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $GroupsByTypeScope = Get-GroupsByTypeScope
                if ($GroupsByTypeScope -and $GroupsByTypeScope.Count -gt 0) {
                    Update-ADReportResults -Results $GroupsByTypeScope
                    Write-ADReportLog -Message "Groups by type and scope loaded. $($GroupsByTypeScope.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GroupsByTypeScope
                    $Global:TextBlockStatus.Text = "Groups by type and scope loaded. $($GroupsByTypeScope.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No groups found."
                }
            } catch {
                $ErrorMessage = "Error loading groups by type and scope: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDynamicDistGroups
    if ($null -ne $Global:ButtonQuickDynamicDistGroups) {
        $ButtonQuickDynamicDistGroups.add_Click({
            Write-ADReportLog -Message "Loading dynamic distribution groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $DynamicGroups = Get-DynamicDistributionGroups
                if ($DynamicGroups -and $DynamicGroups.Count -gt 0) {
                    Update-ADReportResults -Results $DynamicGroups
                    Write-ADReportLog -Message "Dynamic distribution groups loaded. $($DynamicGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $DynamicGroups
                    $Global:TextBlockStatus.Text = "Dynamic distribution groups loaded. $($DynamicGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No dynamic distribution groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No dynamic distribution groups found."
                }
            } catch {
                $ErrorMessage = "Error loading dynamic distribution groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickMailEnabledGroups
    if ($null -ne $Global:ButtonQuickMailEnabledGroups) {
        $ButtonQuickMailEnabledGroups.add_Click({
            Write-ADReportLog -Message "Loading mail-enabled groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $MailEnabledGroups = Get-MailEnabledGroups
                if ($MailEnabledGroups -and $MailEnabledGroups.Count -gt 0) {
                    Update-ADReportResults -Results $MailEnabledGroups
                    Write-ADReportLog -Message "Mail-enabled groups loaded. $($MailEnabledGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $MailEnabledGroups
                    $Global:TextBlockStatus.Text = "Mail-enabled groups loaded. $($MailEnabledGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No mail-enabled groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No mail-enabled groups found."
                }
            } catch {
                $ErrorMessage = "Error loading mail-enabled groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGroupsWithoutOwners
    if ($null -ne $Global:ButtonQuickGroupsWithoutOwners) {
        $ButtonQuickGroupsWithoutOwners.add_Click({
            Write-ADReportLog -Message "Loading groups without owners..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $GroupsWithoutOwners = Get-GroupsWithoutOwners
                if ($GroupsWithoutOwners -and $GroupsWithoutOwners.Count -gt 0) {
                    Update-ADReportResults -Results $GroupsWithoutOwners
                    Write-ADReportLog -Message "Groups without owners loaded. $($GroupsWithoutOwners.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GroupsWithoutOwners
                    $Global:TextBlockStatus.Text = "Groups without owners loaded. $($GroupsWithoutOwners.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No groups without owners found." -Type Info
                    $Global:TextBlockStatus.Text = "No groups without owners found."
                }
            } catch {
                $ErrorMessage = "Error loading groups without owners: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickLargeGroups
    if ($null -ne $Global:ButtonQuickLargeGroups) {
        $ButtonQuickLargeGroups.add_Click({
            Write-ADReportLog -Message "Loading large groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $LargeGroups = Get-LargeGroups -Threshold 100
                if ($LargeGroups -and $LargeGroups.Count -gt 0) {
                    Update-ADReportResults -Results $LargeGroups
                    Write-ADReportLog -Message "Large groups loaded. $($LargeGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $LargeGroups
                    $Global:TextBlockStatus.Text = "Large groups loaded. $($LargeGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No large groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No large groups found."
                }
            } catch {
                $ErrorMessage = "Error loading large groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickRecentlyModifiedGroups
    if ($null -ne $Global:ButtonQuickRecentlyModifiedGroups) {
        $ButtonQuickRecentlyModifiedGroups.add_Click({
            Write-ADReportLog -Message "Loading recently modified groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $RecentlyModifiedGroups = Get-RecentlyModifiedGroups -Days 30
                if ($RecentlyModifiedGroups -and $RecentlyModifiedGroups.Count -gt 0) {
                    Update-ADReportResults -Results $RecentlyModifiedGroups
                    Write-ADReportLog -Message "Recently modified groups loaded. $($RecentlyModifiedGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $RecentlyModifiedGroups
                    $Global:TextBlockStatus.Text = "Recently modified groups loaded. $($RecentlyModifiedGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No recently modified groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No recently modified groups found."
                }
            } catch {
                $ErrorMessage = "Error loading recently modified groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickOSSummary
    if ($null -ne $Global:ButtonQuickOSSummary) {
        $ButtonQuickOSSummary.add_Click({
            Write-ADReportLog -Message "Analyzing OS summary..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $OSSummary = Get-OSSummary
                if ($OSSummary -and $OSSummary.Count -gt 0) {
                    Update-ADReportResults -Results $OSSummary
                    Write-ADReportLog -Message "OS summary analysis completed. $($OSSummary.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $OSSummary
                    $Global:TextBlockStatus.Text = "OS summary analysis completed. $($OSSummary.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No OS summary data found." -Type Info
                    $Global:TextBlockStatus.Text = "No OS summary data found."
                }
            } catch {
                $ErrorMessage = "Error analyzing OS summary: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersByOSVersion
    if ($null -ne $Global:ButtonQuickComputersByOSVersion) {
        $ButtonQuickComputersByOSVersion.add_Click({
            Write-ADReportLog -Message "Analyzing computers by OS version..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $ComputersByOS = Get-ComputersByOSVersion
                if ($ComputersByOS -and $ComputersByOS.Count -gt 0) {
                    Update-ADReportResults -Results $ComputersByOS
                    Write-ADReportLog -Message "OS version analysis completed. $($ComputersByOS.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ComputersByOS
                    $Global:TextBlockStatus.Text = "OS version analysis completed. $($ComputersByOS.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computers found." -Type Info
                    $Global:TextBlockStatus.Text = "No computers found."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers by OS version: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickBitLockerStatus
    if ($null -ne $Global:ButtonQuickBitLockerStatus) {
        $ButtonQuickBitLockerStatus.add_Click({
            Write-ADReportLog -Message "Analyzing BitLocker status..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $BitLockerStatus = Get-BitLockerStatus
                if ($BitLockerStatus -and $BitLockerStatus.Count -gt 0) {
                    Update-ADReportResults -Results $BitLockerStatus
                    Write-ADReportLog -Message "BitLocker status analysis completed. $($BitLockerStatus.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $BitLockerStatus
                    $Global:TextBlockStatus.Text = "BitLocker status analysis completed. $($BitLockerStatus.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No BitLocker information found." -Type Info
                    $Global:TextBlockStatus.Text = "No BitLocker information found."
                }
            } catch {
                $ErrorMessage = "Error analyzing BitLocker status: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickStaleComputerPasswords
    if ($null -ne $Global:ButtonQuickStaleComputerPasswords) {
        $ButtonQuickStaleComputerPasswords.add_Click({
            Write-ADReportLog -Message "Analyzing stale computer passwords..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $StalePasswords = Get-StaleComputerPasswords
                if ($StalePasswords -and $StalePasswords.Count -gt 0) {
                    Update-ADReportResults -Results $StalePasswords
                    Write-ADReportLog -Message "Stale password analysis completed. $($StalePasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $StalePasswords
                    $Global:TextBlockStatus.Text = "Stale password analysis completed. $($StalePasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No stale computer passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No stale computer passwords found."
                }
            } catch {
                $ErrorMessage = "Error analyzing stale computer passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersNeverLoggedOn
    if ($null -ne $Global:ButtonQuickComputersNeverLoggedOn) {
        $ButtonQuickComputersNeverLoggedOn.add_Click({
            Write-ADReportLog -Message "Analyzing computers that never logged on..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $NeverLoggedOn = Get-ComputersNeverLoggedOn
                if ($NeverLoggedOn -and $NeverLoggedOn.Count -gt 0) {
                    Update-ADReportResults -Results $NeverLoggedOn
                    Write-ADReportLog -Message "Never logged on analysis completed. $($NeverLoggedOn.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NeverLoggedOn
                    $Global:TextBlockStatus.Text = "Never logged on analysis completed. $($NeverLoggedOn.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computers found that never logged on." -Type Info
                    $Global:TextBlockStatus.Text = "No computers found that never logged on."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers that never logged on: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDuplicateComputerNames
    if ($null -ne $Global:ButtonQuickDuplicateComputerNames) {
        $ButtonQuickDuplicateComputerNames.add_Click({
            Write-ADReportLog -Message "Analyzing duplicate computer names..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $DuplicateNames = Get-DuplicateComputerNames
                if ($DuplicateNames -and $DuplicateNames.Count -gt 0) {
                    Update-ADReportResults -Results $DuplicateNames
                    Write-ADReportLog -Message "Duplicate names analysis completed. $($DuplicateNames.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $DuplicateNames
                    $Global:TextBlockStatus.Text = "Duplicate names analysis completed. $($DuplicateNames.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No duplicate computer names found." -Type Info
                    $Global:TextBlockStatus.Text = "No duplicate computer names found."
                }
            } catch {
                $ErrorMessage = "Error analyzing duplicate computer names: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersByLocation
    if ($null -ne $Global:ButtonQuickComputersByLocation) {
        $ButtonQuickComputersByLocation.add_Click({
            Write-ADReportLog -Message "Analyzing computers by location..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $ComputersByLocation = Get-ComputersByLocation
                if ($ComputersByLocation -and $ComputersByLocation.Count -gt 0) {
                    Update-ADReportResults -Results $ComputersByLocation
                    Write-ADReportLog -Message "Location analysis completed. $($ComputersByLocation.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ComputersByLocation
                    $Global:TextBlockStatus.Text = "Location analysis completed. $($ComputersByLocation.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computer location data found." -Type Info
                    $Global:TextBlockStatus.Text = "No computer location data found."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers by location: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }
    
    # Event Handler fÃ¼r ButtonQuickServiceAccountsOverview
    if ($null -ne $Global:ButtonQuickServiceAccountsOverview) {
        $ButtonQuickServiceAccountsOverview.add_Click({
            Write-ADReportLog -Message "Loading service accounts overview..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ServiceAccounts = Get-ServiceAccountsOverview
                if ($ServiceAccounts -and $ServiceAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ServiceAccounts
                    Write-ADReportLog -Message "Service accounts loaded. $($ServiceAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ServiceAccounts
                    $Global:TextBlockStatus.Text = "Service accounts loaded. $($ServiceAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickManagedServiceAccounts
    if ($null -ne $Global:ButtonQuickManagedServiceAccounts) {
        $ButtonQuickManagedServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing managed service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ManagedAccounts = Get-ManagedServiceAccounts
                if ($ManagedAccounts -and $ManagedAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ManagedAccounts
                    Write-ADReportLog -Message "Managed service accounts loaded. $($ManagedAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ManagedAccounts
                    $Global:TextBlockStatus.Text = "Managed service accounts loaded. $($ManagedAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No managed service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No managed service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading managed service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickServiceAccountsSPN 
    if ($null -ne $Global:ButtonQuickServiceAccountsSPN) {
        $ButtonQuickServiceAccountsSPN.add_Click({
            Write-ADReportLog -Message "Analyzing service accounts with SPNs..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SPNAccounts = Get-ServiceAccountsSPN
                if ($SPNAccounts -and $SPNAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $SPNAccounts
                    Write-ADReportLog -Message "Service accounts with SPNs loaded. $($SPNAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SPNAccounts
                    $Global:TextBlockStatus.Text = "Service accounts with SPNs loaded. $($SPNAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts with SPNs found." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts with SPNs found."
                }
            } catch {
                $ErrorMessage = "Error loading service accounts with SPNs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickHighPrivServiceAccounts
    if ($null -ne $Global:ButtonQuickHighPrivServiceAccounts) {
        $ButtonQuickHighPrivServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing high privileged service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $HighPrivAccounts = Get-HighPrivServiceAccounts
                if ($HighPrivAccounts -and $HighPrivAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $HighPrivAccounts
                    Write-ADReportLog -Message "High privileged service accounts loaded. $($HighPrivAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $HighPrivAccounts
                    $Global:TextBlockStatus.Text = "High privileged service accounts loaded. $($HighPrivAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No high privileged service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No high privileged service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading high privileged service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickServiceAccountPasswordAge
    if ($null -ne $Global:ButtonQuickServiceAccountPasswordAge) {
        $ButtonQuickServiceAccountPasswordAge.add_Click({
            Write-ADReportLog -Message "Analyzing service account password ages..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $PasswordAgeAccounts = Get-ServiceAccountPasswordAge
                if ($PasswordAgeAccounts -and $PasswordAgeAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $PasswordAgeAccounts
                    Write-ADReportLog -Message "Service account password ages loaded. $($PasswordAgeAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PasswordAgeAccounts
                    $Global:TextBlockStatus.Text = "Service account password ages loaded. $($PasswordAgeAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts found for password age analysis." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts found for password age analysis."
                }
            } catch {
                $ErrorMessage = "Error analyzing service account password ages: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickUnusedServiceAccounts
    if ($null -ne $Global:ButtonQuickUnusedServiceAccounts) {
        $ButtonQuickUnusedServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing unused service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $UnusedAccounts = Get-UnusedServiceAccounts
                if ($UnusedAccounts -and $UnusedAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $UnusedAccounts
                    Write-ADReportLog -Message "Unused service accounts loaded. $($UnusedAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $UnusedAccounts
                    $Global:TextBlockStatus.Text = "Unused service accounts loaded. $($UnusedAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No unused service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No unused service accounts found."
                }
            } catch {
                $ErrorMessage = "Error analyzing unused service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickUnlinkedGPOs
    if ($null -ne $Global:ButtonQuickUnlinkedGPOs) {
        $ButtonQuickUnlinkedGPOs.add_Click({
            Write-ADReportLog -Message "Analyzing unlinked GPOs..." -Type Info
            try {
                $UnlinkedGPOs = Get-UnlinkedGPOs
                if ($UnlinkedGPOs -and $UnlinkedGPOs.Count -gt 0) {
                    Update-ADReportResults -Results $UnlinkedGPOs
                    Write-ADReportLog -Message "Unlinked GPOs loaded. $($UnlinkedGPOs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $UnlinkedGPOs
                    $Global:TextBlockStatus.Text = "Unlinked GPOs loaded. $($UnlinkedGPOs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No unlinked GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No unlinked GPOs found."
                }
            } catch {
                $ErrorMessage = "Error analyzing unlinked GPOs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickEmptyGPOs
    if ($null -ne $Global:ButtonQuickEmptyGPOs) {
        $ButtonQuickEmptyGPOs.add_Click({
            Write-ADReportLog -Message "Analyzing empty GPOs..." -Type Info
            try {
                $EmptyGPOs = Get-EmptyGPOs
                if ($EmptyGPOs -and $EmptyGPOs.Count -gt 0) {
                    Update-ADReportResults -Results $EmptyGPOs
                    Write-ADReportLog -Message "Empty GPOs loaded. $($EmptyGPOs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $EmptyGPOs
                    $Global:TextBlockStatus.Text = "Empty GPOs loaded. $($EmptyGPOs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No empty GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No empty GPOs found."
                }
            } catch {
                $ErrorMessage = "Error analyzing empty GPOs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGPOPermissions
    if ($null -ne $Global:ButtonQuickGPOPermissions) {
        $ButtonQuickGPOPermissions.add_Click({
            Write-ADReportLog -Message "Analyzing GPO permissions..." -Type Info
            try {
                $GPOPermissions = Get-GPOPermissions
                if ($GPOPermissions -and $GPOPermissions.Count -gt 0) {
                    Update-ADReportResults -Results $GPOPermissions
                    Write-ADReportLog -Message "GPO permissions loaded. $($GPOPermissions.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GPOPermissions
                    $Global:TextBlockStatus.Text = "GPO permissions loaded. $($GPOPermissions.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No GPO permissions found." -Type Info
                    $Global:TextBlockStatus.Text = "No GPO permissions found."
                }
            } catch {
                $ErrorMessage = "Error analyzing GPO permissions: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickPasswordPolicySummary
    if ($null -ne $Global:ButtonQuickPasswordPolicySummary) {
        $ButtonQuickPasswordPolicySummary.add_Click({
            Write-ADReportLog -Message "Loading password policy summary..." -Type Info
            try {
                $PasswordPolicies = Get-PasswordPolicySummary
                if ($PasswordPolicies -and $PasswordPolicies.Count -gt 0) {
                    Update-ADReportResults -Results $PasswordPolicies
                    Write-ADReportLog -Message "Password policies loaded. $($PasswordPolicies.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PasswordPolicies
                    $Global:TextBlockStatus.Text = "Password policies loaded. $($PasswordPolicies.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No password policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No password policies found."
                }
            } catch {
                $ErrorMessage = "Error loading password policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAccountLockoutPolicies
    if ($null -ne $Global:ButtonQuickAccountLockoutPolicies) {
        $ButtonQuickAccountLockoutPolicies.add_Click({
            Write-ADReportLog -Message "Loading account lockout policies..." -Type Info
            try {
                $LockoutPolicies = Get-AccountLockoutPolicies
                if ($LockoutPolicies -and $LockoutPolicies.Count -gt 0) {
                    Update-ADReportResults -Results $LockoutPolicies
                    Write-ADReportLog -Message "Account lockout policies loaded. $($LockoutPolicies.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $LockoutPolicies
                    $Global:TextBlockStatus.Text = "Account lockout policies loaded. $($LockoutPolicies.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No account lockout policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No account lockout policies found."
                }
            } catch {
                $ErrorMessage = "Error loading account lockout policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickFineGrainedPasswordPolicies
    if ($null -ne $Global:ButtonQuickFineGrainedPasswordPolicies) {
        $ButtonQuickFineGrainedPasswordPolicies.add_Click({
            Write-ADReportLog -Message "Loading fine-grained password policies..." -Type Info
            try {
                $FGPPs = Get-FineGrainedPasswordPolicies
                if ($FGPPs -and $FGPPs.Count -gt 0) {
                    Update-ADReportResults -Results $FGPPs
                    Write-ADReportLog -Message "Fine-grained password policies loaded. $($FGPPs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $FGPPs
                    $Global:TextBlockStatus.Text = "Fine-grained password policies loaded. $($FGPPs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No fine-grained password policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No fine-grained password policies found."
                }
            } catch {
                $ErrorMessage = "Error loading fine-grained password policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickACLAnalysis
    if ($null -ne $Global:ButtonQuickACLAnalysis) {
        $ButtonQuickACLAnalysis.add_Click({
            Write-ADReportLog -Message "Analyzing ACL permissions..." -Type Info
            try {
                $ACLAnalysis = Get-ACLAnalysis
                if ($ACLAnalysis -and $ACLAnalysis.Count -gt 0) {
                    Update-ADReportResults -Results $ACLAnalysis
                    Write-ADReportLog -Message "ACL analysis completed. $($ACLAnalysis.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ACLAnalysis
                    $Global:TextBlockStatus.Text = "ACL analysis completed. $($ACLAnalysis.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No ACL issues found." -Type Info
                    $Global:TextBlockStatus.Text = "No ACL issues found."
                }
            } catch {
                $ErrorMessage = "Error analyzing ACLs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickInheritanceBreaks 
    if ($null -ne $Global:ButtonQuickInheritanceBreaks) {
        $ButtonQuickInheritanceBreaks.add_Click({
            Write-ADReportLog -Message "Analyzing inheritance breaks..." -Type Info
            try {
                $InheritanceBreaks = Get-InheritanceBreaks
                if ($InheritanceBreaks -and $InheritanceBreaks.Count -gt 0) {
                    Update-ADReportResults -Results $InheritanceBreaks
                    Write-ADReportLog -Message "Inheritance break analysis completed. $($InheritanceBreaks.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $InheritanceBreaks
                    $Global:TextBlockStatus.Text = "Inheritance break analysis completed. $($InheritanceBreaks.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No inheritance breaks found." -Type Info
                    $Global:TextBlockStatus.Text = "No inheritance breaks found."
                }
            } catch {
                $ErrorMessage = "Error analyzing inheritance breaks: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAdminSDHolderObjects
    if ($null -ne $Global:ButtonQuickAdminSDHolderObjects) {
        $ButtonQuickAdminSDHolderObjects.add_Click({
            Write-ADReportLog -Message "Analyzing AdminSDHolder objects..." -Type Info
            try {
                $AdminSDHolderObjects = Get-AdminSDHolderObjects
                if ($AdminSDHolderObjects -and $AdminSDHolderObjects.Count -gt 0) {
                    Update-ADReportResults -Results $AdminSDHolderObjects
                    Write-ADReportLog -Message "AdminSDHolder analysis completed. $($AdminSDHolderObjects.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $AdminSDHolderObjects
                    $Global:TextBlockStatus.Text = "AdminSDHolder analysis completed. $($AdminSDHolderObjects.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No AdminSDHolder objects found." -Type Info
                    $Global:TextBlockStatus.Text = "No AdminSDHolder objects found."
                }
            } catch {
                $ErrorMessage = "Error analyzing AdminSDHolder objects: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAdvancedDelegation
    if ($null -ne $Global:ButtonQuickAdvancedDelegation) {
        $ButtonQuickAdvancedDelegation.add_Click({
            Write-ADReportLog -Message "Analyzing advanced delegations..." -Type Info
            try {
                $AdvancedDelegations = Get-AdvancedDelegations
                if ($AdvancedDelegations -and $AdvancedDelegations.Count -gt 0) {
                    Update-ADReportResults -Results $AdvancedDelegations
                    Write-ADReportLog -Message "Advanced delegation analysis completed. $($AdvancedDelegations.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $AdvancedDelegations
                    $Global:TextBlockStatus.Text = "Advanced delegation analysis completed. $($AdvancedDelegations.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No advanced delegations found." -Type Info
                    $Global:TextBlockStatus.Text = "No advanced delegations found."
                }
            } catch {
                $ErrorMessage = "Error analyzing advanced delegations: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSchemaPermissions
    if ($null -ne $Global:ButtonQuickSchemaPermissions) {
        $ButtonQuickSchemaPermissions.add_Click({
            Write-ADReportLog -Message "Analyzing schema permissions..." -Type Info
            try {
                $SchemaPermissions = Get-SchemaPermissions
                if ($SchemaPermissions -and $SchemaPermissions.Count -gt 0) {
                    Update-ADReportResults -Results $SchemaPermissions
                    Write-ADReportLog -Message "Schema permissions analysis completed. $($SchemaPermissions.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SchemaPermissions
                    $Global:TextBlockStatus.Text = "Schema permissions analysis completed. $($SchemaPermissions.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No schema permissions found." -Type Info
                    $Global:TextBlockStatus.Text = "No schema permissions found."
                }
            } catch {
                $ErrorMessage = "Error analyzing schema permissions: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickGPOOverview
    if ($null -ne $Global:ButtonQuickGPOOverview) {
        $ButtonQuickGPOOverview.add_Click({
            Write-ADReportLog -Message "Loading GPO overview..." -Type Info
            try {
                $GPOOverview = Get-GPOOverview
                if ($GPOOverview -and $GPOOverview.Count -gt 0) {
                    Update-ADReportResults -Results $GPOOverview
                    Write-ADReportLog -Message "GPO overview loaded. $($GPOOverview.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GPOOverview
                    $Global:TextBlockStatus.Text = "GPO overview loaded. $($GPOOverview.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No GPOs found."
                }
            } catch {
                $ErrorMessage = "Error loading GPO overview: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickHoneyTokens
    if ($null -ne $Global:ButtonQuickHoneyTokens) {
        $ButtonQuickHoneyTokens.add_Click({
            Write-ADReportLog -Message "Analyzing potential honey tokens..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $HoneyTokens = Get-HoneyTokens
                if ($HoneyTokens -and $HoneyTokens.Count -gt 0) {
                    Update-ADReportResults -Results $HoneyTokens
                    Write-ADReportLog -Message "Honey token analysis completed. $($HoneyTokens.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $HoneyTokens
                    $Global:TextBlockStatus.Text = "Honey token analysis completed. $($HoneyTokens.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No honey tokens detected." -Type Info
                    $Global:TextBlockStatus.Text = "No honey tokens detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing honey tokens: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickPrivilegeEscalation
    if ($null -ne $Global:ButtonQuickPrivilegeEscalation) {
        $ButtonQuickPrivilegeEscalation.add_Click({
            Write-ADReportLog -Message "Analyzing potential privilege escalation paths..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $PrivEscPaths = Get-PrivilegeEscalationPaths
                if ($PrivEscPaths -and $PrivEscPaths.Count -gt 0) {
                    Update-ADReportResults -Results $PrivEscPaths
                    Write-ADReportLog -Message "Privilege escalation analysis completed. $($PrivEscPaths.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PrivEscPaths
                    $Global:TextBlockStatus.Text = "Privilege escalation analysis completed. $($PrivEscPaths.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No privilege escalation paths detected." -Type Info
                    $Global:TextBlockStatus.Text = "No privilege escalation paths detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing privilege escalation paths: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickExposedCredentials
    if ($null -ne $Global:ButtonQuickExposedCredentials) {
        $ButtonQuickExposedCredentials.add_Click({
            Write-ADReportLog -Message "Analyzing exposed credentials..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ExposedCreds = Get-ExposedCredentials
                if ($ExposedCreds -and $ExposedCreds.Count -gt 0) {
                    Update-ADReportResults -Results $ExposedCreds
                    Write-ADReportLog -Message "Exposed credentials analysis completed. $($ExposedCreds.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ExposedCreds
                    $Global:TextBlockStatus.Text = "Exposed credentials analysis completed. $($ExposedCreds.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No exposed credentials detected." -Type Info
                    $Global:TextBlockStatus.Text = "No exposed credentials detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing exposed credentials: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSuspiciousLogons
    if ($null -ne $Global:ButtonQuickSuspiciousLogons) {
        $ButtonQuickSuspiciousLogons.add_Click({
            Write-ADReportLog -Message "Analyzing suspicious logon patterns..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SuspiciousLogons = Get-SuspiciousLogons
                if ($SuspiciousLogons -and $SuspiciousLogons.Count -gt 0) {
                    Update-ADReportResults -Results $SuspiciousLogons
                    Write-ADReportLog -Message "Suspicious logon analysis completed. $($SuspiciousLogons.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SuspiciousLogons
                    $Global:TextBlockStatus.Text = "Suspicious logon analysis completed. $($SuspiciousLogons.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No suspicious logon patterns detected." -Type Info
                    $Global:TextBlockStatus.Text = "No suspicious logon patterns detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing suspicious logons: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickForeignSecurityPrincipals
    if ($null -ne $Global:ButtonQuickForeignSecurityPrincipals) {
        $ButtonQuickForeignSecurityPrincipals.add_Click({
            Write-ADReportLog -Message "Analyzing foreign security principals..." -Type Info
            try {
                $ForeignPrincipals = Get-ForeignSecurityPrincipals
                if ($ForeignPrincipals -and $ForeignPrincipals.Count -gt 0) {
                    Update-ADReportResults -Results $ForeignPrincipals
                    Write-ADReportLog -Message "Foreign security principals analysis completed. $($ForeignPrincipals.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ForeignPrincipals
                    $Global:TextBlockStatus.Text = "Foreign security principals analysis completed. $($ForeignPrincipals.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No foreign security principals found." -Type Info
                    $Global:TextBlockStatus.Text = "No foreign security principals found."
                }
            } catch {
                $ErrorMessage = "Error analyzing foreign security principals: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSIDHistoryAbuse
    if ($null -ne $Global:ButtonQuickSIDHistoryAbuse) {
        $ButtonQuickSIDHistoryAbuse.add_Click({
            Write-ADReportLog -Message "Analyzing potential SID history abuse..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SIDHistoryAbuse = Get-SIDHistoryAbuse
                if ($SIDHistoryAbuse -and $SIDHistoryAbuse.Count -gt 0) {
                    Update-ADReportResults -Results $SIDHistoryAbuse
                    Write-ADReportLog -Message "SID history abuse analysis completed. $($SIDHistoryAbuse.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SIDHistoryAbuse
                    $Global:TextBlockStatus.Text = "SID history abuse analysis completed. $($SIDHistoryAbuse.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No potential SID history abuse detected." -Type Info
                    $Global:TextBlockStatus.Text = "No potential SID history abuse detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing SID history abuse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }
    
    # Event Handler fÃ¼r weitere Quick Report Buttons
    if ($null -ne $Global:ButtonQuickExpiringAccounts) {
        $ButtonQuickExpiringAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing expiring accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ExpiringAccounts = Get-ExpiringAccounts
                if ($ExpiringAccounts -and $ExpiringAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ExpiringAccounts
                    Write-ADReportLog -Message "Expiring accounts analysis completed. $($ExpiringAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ExpiringAccounts
                    $Global:TextBlockStatus.Text = "Expiring accounts analysis completed. $($ExpiringAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No expiring accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No expiring accounts found."
                }
            } catch {
                $ErrorMessage = "Error analyzing expiring accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickCompromiseIndicators
    if ($null -ne $Global:ButtonQuickCompromiseIndicators) {
        $ButtonQuickCompromiseIndicators.add_Click({
            Write-ADReportLog -Message "Analysiere Kompromittierungsindikatoren..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $CompromiseIndicators = Get-CompromiseIndicators
                if ($CompromiseIndicators -and $CompromiseIndicators.Count -gt 0) {
                    Update-ADReportResults -Results $CompromiseIndicators
                    Write-ADReportLog -Message "Kompromittierungsanalyse abgeschlossen. $($CompromiseIndicators.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $CompromiseIndicators
                    $Global:TextBlockStatus.Text = "Kompromittierungsanalyse abgeschlossen. $($CompromiseIndicators.Count) Datensatz/Datensätze gefunden."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Kompromittierungsindikatoren gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Kompromittierungsindikatoren gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Kompromittierungsanalyse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSecurityDashboard
    if ($null -ne $Global:ButtonQuickSecurityDashboard) {
        $ButtonQuickSecurityDashboard.add_Click({
            Write-ADReportLog -Message "Erstelle Sicherheitsbewertungs-Dashboard..." -Type Info
            try {
                $SecurityDashboard = Get-SecurityDashboard
                if ($SecurityDashboard -and $SecurityDashboard.Count -gt 0) {
                    Update-ADReportResults -Results $SecurityDashboard
                    Write-ADReportLog -Message "Sicherheitsbewertung abgeschlossen. $($SecurityDashboard.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $SecurityDashboard
                    $Global:TextBlockStatus.Text = "Sicherheitsbewertung abgeschlossen. $($SecurityDashboard.Count) Bewertungskategorie(n) analysiert."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Sicherheitsbewertung verfügbar." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Sicherheitsbewertung verfügbar."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Sicherheitsbewertung: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAuthProtocolAnalysis
    if ($null -ne $Global:ButtonQuickAuthProtocolAnalysis) {
        $ButtonQuickAuthProtocolAnalysis.add_Click({
            Write-ADReportLog -Message "Analysiere Authentifizierungsprotokolle..." -Type Info
            try {
                $AuthProtocolAnalysis = Get-AuthProtocolAnalysis
                if ($AuthProtocolAnalysis -and $AuthProtocolAnalysis.Count -gt 0) {
                    Update-ADReportResults -Results $AuthProtocolAnalysis
                    Write-ADReportLog -Message "Authentifizierungsprotokoll-Analyse abgeschlossen. $($AuthProtocolAnalysis.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $AuthProtocolAnalysis
                    $Global:TextBlockStatus.Text = "Authentifizierungsprotokoll-Analyse abgeschlossen. $($AuthProtocolAnalysis.Count) Datensatz/Datensätze gefunden."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Authentifizierungsprotokolldaten gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Authentifizierungsprotokolldaten gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Authentifizierungsprotokoll-Analyse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickFailedAuthPatterns
    if ($null -ne $Global:ButtonQuickFailedAuthPatterns) {
        $ButtonQuickFailedAuthPatterns.add_Click({
            Write-ADReportLog -Message "Analysiere fehlgeschlagene Authentifizierungsmuster..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $FailedAuthPatterns = Get-FailedAuthPatterns
                if ($FailedAuthPatterns -and $FailedAuthPatterns.Count -gt 0) {
                    Update-ADReportResults -Results $FailedAuthPatterns
                    Write-ADReportLog -Message "Analyse fehlgeschlagener Authentifizierungen abgeschlossen. $($FailedAuthPatterns.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $FailedAuthPatterns
                    $Global:TextBlockStatus.Text = "Analyse fehlgeschlagener Authentifizierungen abgeschlossen. $($FailedAuthPatterns.Count) verdächtige Muster gefunden."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine verdächtigen Authentifizierungsmuster gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine verdächtigen Authentifizierungsmuster gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Analyse fehlgeschlagener Authentifizierungen: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickConditionalAccessPolicies
    if ($null -ne $Global:ButtonQuickConditionalAccessPolicies) {
        $ButtonQuickConditionalAccessPolicies.add_Click({
            Write-ADReportLog -Message "Analysiere bedingte Zugriffsrichtlinien..." -Type Info
            try {
                $ConditionalAccessPolicies = Get-ConditionalAccessPolicies
                if ($ConditionalAccessPolicies -and $ConditionalAccessPolicies.Count -gt 0) {
                    Update-ADReportResults -Results $ConditionalAccessPolicies
                    Write-ADReportLog -Message "Analyse bedingter Zugriffsrichtlinien abgeschlossen. $($ConditionalAccessPolicies.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $ConditionalAccessPolicies
                    $Global:TextBlockStatus.Text = "Analyse bedingter Zugriffsrichtlinien abgeschlossen. $($ConditionalAccessPolicies.Count) Richtlinie(n) gefunden."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine bedingten Zugriffsrichtlinien gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine bedingten Zugriffsrichtlinien gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Analyse bedingter Zugriffsrichtlinien: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDSGVOCompliance
    if ($null -ne $Global:ButtonQuickDSGVOCompliance) {
        $ButtonQuickDSGVOCompliance.add_Click({
            Write-ADReportLog -Message "Erstelle DSGVO/GDPR Compliance Report..." -Type Info
            try {
                $DSGVOReport = Get-DSGVOComplianceReport
                if ($DSGVOReport -and $DSGVOReport.Count -gt 0) {
                    Update-ADReportResults -Results $DSGVOReport
                    Write-ADReportLog -Message "DSGVO Compliance Report abgeschlossen. $($DSGVOReport.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $DSGVOReport
                    $Global:TextBlockStatus.Text = "DSGVO Compliance Report abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine DSGVO-Daten gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine DSGVO-Daten gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler beim DSGVO Compliance Report: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickLoginTimeAnalysis
    if ($null -ne $Global:ButtonQuickLoginTimeAnalysis) {
        $ButtonQuickLoginTimeAnalysis.add_Click({
            Write-ADReportLog -Message "Analysiere Login-Zeiten und Muster..." -Type Info
            try {
                $LoginTimeAnalysis = Get-LoginTimeAnalysis
                if ($LoginTimeAnalysis -and $LoginTimeAnalysis.Count -gt 0) {
                    Update-ADReportResults -Results $LoginTimeAnalysis
                    Write-ADReportLog -Message "Login-Zeit Analyse abgeschlossen. $($LoginTimeAnalysis.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $LoginTimeAnalysis
                    $Global:TextBlockStatus.Text = "Login-Zeit Analyse abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine ungewöhnlichen Login-Zeiten gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine ungewöhnlichen Login-Zeiten gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Login-Zeit Analyse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickOrgChart
    if ($null -ne $Global:ButtonQuickOrgChart) {
        $ButtonQuickOrgChart.add_Click({
            Write-ADReportLog -Message "Erstelle Organigramm-Export..." -Type Info
            try {
                $OrgChart = Get-OrganizationChart
                if ($OrgChart -and $OrgChart.Count -gt 0) {
                    Update-ADReportResults -Results $OrgChart
                    Write-ADReportLog -Message "Organigramm-Export abgeschlossen. $($OrgChart.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $OrgChart
                    $Global:TextBlockStatus.Text = "Organigramm-Export abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Organisationsdaten gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Organisationsdaten gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler beim Organigramm-Export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGPOConflictAnalysis
    if ($null -ne $Global:ButtonQuickGPOConflictAnalysis) {
        $ButtonQuickGPOConflictAnalysis.add_Click({
            Write-ADReportLog -Message "Analysiere GPO-Konflikte..." -Type Info
            try {
                $GPOConflicts = Get-GPOConflictAnalysis
                if ($GPOConflicts -and $GPOConflicts.Count -gt 0) {
                    Update-ADReportResults -Results $GPOConflicts
                    Write-ADReportLog -Message "GPO-Konflikt Analyse abgeschlossen. $($GPOConflicts.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $GPOConflicts
                    $Global:TextBlockStatus.Text = "GPO-Konflikt Analyse abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine GPO-Konflikte gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine GPO-Konflikte gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der GPO-Konflikt Analyse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGPOSecurityBaseline
    if ($null -ne $Global:ButtonQuickGPOSecurityBaseline) {
        $ButtonQuickGPOSecurityBaseline.add_Click({
            Write-ADReportLog -Message "Vergleiche GPOs mit Security Baseline..." -Type Info
            try {
                $GPOBaseline = Get-GPOSecurityBaseline
                if ($GPOBaseline -and $GPOBaseline.Count -gt 0) {
                    Update-ADReportResults -Results $GPOBaseline
                    Write-ADReportLog -Message "GPO Security Baseline Vergleich abgeschlossen. $($GPOBaseline.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $GPOBaseline
                    $Global:TextBlockStatus.Text = "GPO Security Baseline Vergleich abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Baseline-Abweichungen gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Baseline-Abweichungen gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler beim GPO Security Baseline Vergleich: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAccountCreationTimeline
    if ($null -ne $Global:ButtonQuickAccountCreationTimeline) {
        $ButtonQuickAccountCreationTimeline.add_Click({
            Write-ADReportLog -Message "Erstelle Account-Erstellungs-Timeline..." -Type Info
            try {
                $AccountTimeline = Get-AccountCreationTimeline
                if ($AccountTimeline -and $AccountTimeline.Count -gt 0) {
                    Update-ADReportResults -Results $AccountTimeline
                    Write-ADReportLog -Message "Account-Erstellungs-Timeline abgeschlossen. $($AccountTimeline.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $AccountTimeline
                    $Global:TextBlockStatus.Text = "Account-Erstellungs-Timeline abgeschlossen."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine Account-Erstellungsdaten gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine Account-Erstellungsdaten gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Account-Erstellungs-Timeline: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDeletedObjectsRecovery
    if ($null -ne $Global:ButtonQuickDeletedObjectsRecovery) {
        $ButtonQuickDeletedObjectsRecovery.add_Click({
            Write-ADReportLog -Message "Suche gelöschte AD-Objekte..." -Type Info
            try {
                $DeletedObjects = Get-DeletedObjectsRecovery
                if ($DeletedObjects -and $DeletedObjects.Count -gt 0) {
                    Update-ADReportResults -Results $DeletedObjects
                    Write-ADReportLog -Message "Gelöschte Objekte gefunden. $($DeletedObjects.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $DeletedObjects
                    $Global:TextBlockStatus.Text = "Gelöschte Objekte gefunden."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine gelöschten Objekte gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine gelöschten Objekte gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der Suche nach gelöschten Objekten: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickProxyAddressChanges
    if ($null -ne $Global:ButtonQuickProxyAddressChanges) {
        $ButtonQuickProxyAddressChanges.add_Click({
            Write-ADReportLog -Message "Analysiere ProxyAddress-Änderungen..." -Type Info
            try {
                $ProxyChanges = Get-ProxyAddressChanges
                if ($ProxyChanges -and $ProxyChanges.Count -gt 0) {
                    Update-ADReportResults -Results $ProxyChanges
                    Write-ADReportLog -Message "ProxyAddress-Änderungen analysiert. $($ProxyChanges.Count) Ergebnis(se) gefunden." -Type Info
                    Update-ResultCounters -Results $ProxyChanges
                    $Global:TextBlockStatus.Text = "ProxyAddress-Änderungen analysiert."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "Keine ProxyAddress-Änderungen gefunden." -Type Info
                    $Global:TextBlockStatus.Text = "Keine ProxyAddress-Änderungen gefunden."
                }
            } catch {
                $ErrorMessage = "Fehler bei der ProxyAddress-Analyse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Fehler: $ErrorMessage"
            }
        })
    }

    # Definiere den ListBox Selection Changed Handler als Script Variable
    $script:ListBoxSelectionChangedHandler = $null
    
    # Fenster anzeigen
    [VOID]$Window.ShowDialog()
}

# Initialize Visualization Tabs wenn vorhanden
if ($null -ne $Global:ButtonRefreshTopology) {
    $ButtonRefreshTopology.Add_Click({
        $viewType = $Global:ComboBoxTopologyView.SelectedItem.Content
        $topology = Get-ADNetworkTopology -ViewType $viewType
        Draw-TopologyOnCanvas -Canvas $Global:CanvasNetworkTopology -Topology $topology
    })
}

if ($null -ne $Global:ButtonRefreshHeatMap) {
    $ButtonRefreshHeatMap.Add_Click({
        $metric = $Global:ComboBoxHeatMapMetric.SelectedItem.Content -replace ' Score', ''
        $grouping = $Global:ComboBoxHeatMapGrouping.SelectedItem.Content -replace 'By ', ''
        $heatMapData = Get-SecurityHeatMapData -MetricType $metric -GroupBy $grouping
        Draw-SecurityHeatMap -Container $Global:HeatMapContainer -HeatMapData $heatMapData
    })
}

if ($null -ne $Global:ListBoxAvailableFields) {
    Initialize-ReportBuilder -AvailableFieldsList $Global:ListBoxAvailableFields `
                            -ReportCanvas $Global:ReportBuilderCanvas
}

if ($null -ne $Global:ButtonPreviewReport) {
    $ButtonPreviewReport.Add_Click({
        $customReport = Build-CustomReport -ReportCanvas $Global:ReportBuilderCanvas
        Update-ADReportResults -Results $customReport
        $Global:MainTabControl.SelectedIndex = 0  # Wechsel zur Tabellenansicht
    })
}

# Funktion zur Visualisierung der Ergebnisse (Placeholder fÃ¼r zukÃ¼nftige Implementierung)
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$false)]
        $Results
    )
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    # Momentan wird nur ein Debug-Log-Eintrag erstellt
    Write-DebugLog "Update-ResultVisualization aufgerufen mit $($Results.Count) Ergebnissen" "Visualization"
}

# Platzhalter-Funktion fÃ¼r initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    Write-DebugLog "Initialize-ResultVisualization aufgerufen" "Visualization"
}

# --- Hauptlogik ---
Function Start-ADReportGUI {
    # Bereinige alle alten globalen UI-Variablen vor dem Start
    $UiVariables = @("Window", "ComboBoxFilterAttribute", "TextBoxFilterValue", "ListBoxSelectAttributes",
                  "ButtonQueryAD", "ButtonQuickAllUsers", "ButtonQuickDisabledUsers", "ButtonQuickLockedUsers",
                  "ButtonQuickGroups", "ButtonQuickSecurityGroups", "ButtonQuickDistributionGroups",
                  "ButtonQuickNeverExpire", "ButtonQuickInactiveUsers", "ButtonQuickAdminUsers",
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", 
                  "ButtonQuickWeakPasswordPolicy", "ButtonQuickRiskyGroupMemberships", "ButtonQuickPrivilegedAccounts",
                  "DataGridResults", "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
                  "ResultCountGrid", "UserCountText", "ComputerCountText", "GroupCountText")
    
    foreach ($var in $UiVariables) {
        Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
    }

    # Ruft Initialize-ADReportForm auf, welche die UI lÃ¤dt, Elemente zuweist und fÃ¼llt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Skriptstart ---
Start-ADReportGUI
# SIG # Begin signature block
# MIIoiQYJKoZIhvcNAQcCoIIoejCCKHYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCk92rJunwD5X/W
# LGaW63YTZdHMtWiERJ69TbxCNsBnGKCCILswggXJMIIEsaADAgECAhAbtY8lKt8j
# AEkoya49fu0nMA0GCSqGSIb3DQEBDAUAMH4xCzAJBgNVBAYTAlBMMSIwIAYDVQQK
# ExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEuMScwJQYDVQQLEx5DZXJ0dW0gQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkxIjAgBgNVBAMTGUNlcnR1bSBUcnVzdGVkIE5l
# dHdvcmsgQ0EwHhcNMjEwNTMxMDY0MzA2WhcNMjkwOTE3MDY0MzA2WjCBgDELMAkG
# A1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVjaG5vbG9naWVzIFMuQS4xJzAl
# BgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTEkMCIGA1UEAxMb
# Q2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEAvfl4+ObVgAxknYYblmRnPyI6HnUBfe/7XGeMycxca6mR5rlC
# 5SBLm9qbe7mZXdmbgEvXhEArJ9PoujC7Pgkap0mV7ytAJMKXx6fumyXvqAoAl4Va
# qp3cKcniNQfrcE1K1sGzVrihQTib0fsxf4/gX+GxPw+OFklg1waNGPmqJhCrKtPQ
# 0WeNG0a+RzDVLnLRxWPa52N5RH5LYySJhi40PylMUosqp8DikSiJucBb+R3Z5yet
# /5oCl8HGUJKbAiy9qbk0WQq/hEr/3/6zn+vZnuCYI+yma3cWKtvMrTscpIfcRnNe
# GWJoRVfkkIJCu0LW8GHgwaM9ZqNd9BjuiMmNF0UpmTJ1AjHuKSbIawLmtWJFfzcV
# WiNoidQ+3k4nsPBADLxNF8tNorMe0AZa3faTz1d1mfX6hhpneLO/lv403L3nUlbl
# s+V1e9dBkQXcXWnjlQ1DufyDljmVe2yAWk8TcsbXfSl6RLpSpCrVQUYJIP4ioLZb
# MI28iQzV13D4h1L92u+sUS4Hs07+0AnacO+Y+lbmbdu1V0vc5SwlFcieLnhO+Nqc
# noYsylfzGuXIkosagpZ6w7xQEmnYDlpGizrrJvojybawgb5CAKT41v4wLsfSRvbl
# jnX98sy50IdbzAYQYLuDNbdeZ95H7JlI8aShFf6tjGKOOVVPORa5sWOd/7cCAwEA
# AaOCAT4wggE6MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFLahVDkCw6A/joq8
# +tT4HKbROg79MB8GA1UdIwQYMBaAFAh2zcsH/yT2xc3tu5C84oQ3RnX3MA4GA1Ud
# DwEB/wQEAwIBBjAvBgNVHR8EKDAmMCSgIqAghh5odHRwOi8vY3JsLmNlcnR1bS5w
# bC9jdG5jYS5jcmwwawYIKwYBBQUHAQEEXzBdMCgGCCsGAQUFBzABhhxodHRwOi8v
# c3ViY2Eub2NzcC1jZXJ0dW0uY29tMDEGCCsGAQUFBzAChiVodHRwOi8vcmVwb3Np
# dG9yeS5jZXJ0dW0ucGwvY3RuY2EuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAmMCQG
# CCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcNAQEM
# BQADggEBAFHCoVgWIhCL/IYx1MIy01z4S6Ivaj5N+KsIHu3V6PrnCA3st8YeDrJ1
# BXqxC/rXdGoABh+kzqrya33YEcARCNQOTWHFOqj6seHjmOriY/1B9ZN9DbxdkjuR
# mmW60F9MvkyNaAMQFtXx0ASKhTP5N+dbLiZpQjy6zbzUeulNndrnQ/tjUoCFBMQl
# lVXwfqefAcVbKPjgzoZwpic7Ofs4LphTZSJ1Ldf23SIikZbr3WjtP6MZl9M7JYjs
# NhI9qX7OAo0FmpKnJ25FspxihjcNpDOO16hO0EoXQ0zF8ads0h5YbBRRfopUofbv
# n3l6XYGaFpAP4bvxSgD5+d2+7arszgowggaDMIIEa6ADAgECAhEAnpwE9lWotKcC
# bUmMbHiNqjANBgkqhkiG9w0BAQwFADBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMY
# QXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0gVGltZXN0
# YW1waW5nIDIwMjEgQ0EwHhcNMjUwMTA5MDg0MDQzWhcNMzYwMTA3MDg0MDQzWjBQ
# MQswCQYDVQQGEwJQTDEhMB8GA1UECgwYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MR4wHAYDVQQDDBVDZXJ0dW0gVGltZXN0YW1wIDIwMjUwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDHKV9n+Kwr3ZBF5UCLWOQ/NdbblAvQeGMjfCi/bibT
# 71hPkwKV4UvQt1MuOwoaUCYtsLhw8jrmOmoz2HoHKKzEpiS3A1rA3ssXUZMnSrbi
# iVpDj+5MtnbXSVEJKbccuHbmwcjl39N4W72zccoC/neKAuwO1DJ+9SO+YkHncRiV
# 95idWhxRAcDYv47hc9GEFZtTFxQXLbrL4N7N90BqLle3ayznzccEPQ+E6H6p00zE
# 9HUp++3bZTF4PfyPRnKCLc5ezAzEqqbbU5F/nujx69T1mm02jltlFXnTMF1vlake
# QXWYpGIjtrR7WP7tIMZnk78nrYSfeAp8le+/W/5+qr7tqQZufW9invsRTcfk7P+m
# nKjJLuSbwqgxelvCBryz9r51bT0561aR2c+joFygqW7n4FPCnMLOj40X4ot7wP2u
# 8kLRDVHbhsHq5SGLqr8DbFq14ws2ALS3tYa2GGiA7wX79rS5oDMnSY/xmJO5cupu
# SvqpylzO7jzcLOwWiqCrq05AXp51SRrj9xRt8KdZWpDdWhWmE8MFiFtmQ0AqODLJ
# Bn1hQAx3FvD/pte6pE1Bil0BOVC2Snbeq/3NylDwvDdAg/0CZRJsQIaydHswJwyY
# BlYUDyaQK2yUS57hobnYx/vStMvTB96ii4jGV3UkZh3GvwdDCsZkbJXaU8ATF/z6
# DwIDAQABo4IBUDCCAUwwdQYIKwYBBQUHAQEEaTBnMDsGCCsGAQUFBzAChi9odHRw
# Oi8vc3ViY2EucmVwb3NpdG9yeS5jZXJ0dW0ucGwvY3RzY2EyMDIxLmNlcjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAfBgNVHSMEGDAW
# gBS+VAIvv0Bsc0POrAklTp5DRBru4DAMBgNVHRMBAf8EAjAAMDkGA1UdHwQyMDAw
# LqAsoCqGKGh0dHA6Ly9zdWJjYS5jcmwuY2VydHVtLnBsL2N0c2NhMjAyMS5jcmww
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMCIGA1UdIAQb
# MBkwCAYGZ4EMAQQCMA0GCyqEaAGG9ncCBQELMB0GA1UdDgQWBBSBjAagKFP8AD/b
# fp5KwR8i7LISiTANBgkqhkiG9w0BAQwFAAOCAgEAmQ8ZDBvrBUPnaL87AYc4Jlmf
# H1ZP5yt65MtzYu8fbmsL3d3cvYs+Enbtfu9f2wMehzSyved3Rc59a04O8NN7plw4
# PXg71wfSE4MRFM1EuqL63zq9uTjm/9tA73r1aCdWmkprKp0aLoZolUN0qGcvr9+Q
# G8VIJVMcuSqFeEvRrLEKK2xVkMSdTTbDhseUjI4vN+BrXm5z45EA3aDpSiZQuoNd
# 4RFnDzddbgfcCQPaY2UyXqzNBjnuz6AyHnFzKtNlCevkMBgh4dIDt/0DGGDOaTEA
# WZtUEqK5AlHd0PBnd40Lnog4UATU3Bt6GHfeDmWEHFTjHKsmn9Q8wiGj906bVgL8
# 35tfEH9EgYDklqrOUxWxDf1cOA7ds/r8pIc2vjLQ9tOSkm9WXVbnTeLG3Q57frTg
# CvTObd/qf3UzE97nTNOU7vOMZEo41AgmhuEbGsyQIDM/V6fJQX1RnzzJNoqfTTkU
# zUoP2tlNHnNsjFo2YV+5yZcoaawmNWmR7TywUXG2/vFgJaG0bfEoodeeXp7A4I4H
# aDDpfRa7ypgJEPeTwHuBRJpj9N+1xtri+6BzHPwsAAvUJm58PGoVsteHAXwvpg4N
# VgvUk3BKbl7xFulWU1KHqH/sk7T0CFBQ5ohuKPmFf1oqAP4AO9a3Yg2wBMwEg1zP
# Oh6xbUXskzs9iSa9yGwwgga5MIIEoaADAgECAhEAmaOACiZVO2Wr3G6EprPqOTAN
# BgkqhkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8g
# VGVjaG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9u
# IEF1dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAy
# MB4XDTIxMDUxOTA1MzIxOFoXDTM2MDUxODA1MzIxOFowVjELMAkGA1UEBhMCUEwx
# ITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2Vy
# dHVtIENvZGUgU2lnbmluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAnSPPBDAjO8FGLOczcz5jXXp1ur5cTbq96y34vuTmflN4mSAfgLKT
# vggv24/rWiVGzGxT9YEASVMw1Aj8ewTS4IndU8s7VS5+djSoMcbvIKck6+hI1shs
# ylP4JyLvmxwLHtSworV9wmjhNd627h27a8RdrT1PH9ud0IF+njvMk2xqbNTIPsnW
# tw3E7DmDoUmDQiYi/ucJ42fcHqBkbbxYDB7SYOouu9Tj1yHIohzuC8KNqfcYf7Z4
# /iZgkBJ+UFNDcc6zokZ2uJIxWgPWXMEmhu1gMXgv8aGUsRdaCtVD2bSlbfsq7Biq
# ljjaCun+RJgTgFRCtsuAEw0pG9+FA+yQN9n/kZtMLK+Wo837Q4QOZgYqVWQ4x6cM
# 7/G0yswg1ElLlJj6NYKLw9EcBXE7TF3HybZtYvj9lDV2nT8mFSkcSkAExzd4prHw
# YjUXTeZIlVXqj+eaYqoMTpMrfh5MCAOIG5knN4Q/JHuurfTI5XDYO962WZayx7AC
# Ff5ydJpoEowSP07YaBiQ8nXpDkNrUA9g7qf/rCkKbWpQ5boufUnq1UiYPIAHlezf
# 4muJqxqIns/kqld6JVX8cixbd6PzkDpwZo4SlADaCi2JSplKShBSND36E/ENVv8u
# rPS0yOnpG4tIoBGxVCARPCg1BnyMJ4rBJAcOSnAWd18Jx5n858JSqPECAwEAAaOC
# AVUwggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFN10XUwA23ufoHTKsW73
# PMAywHDNMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB
# /wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcDAzAwBgNVHR8EKTAnMCWgI6Ahhh9o
# dHRwOi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAo
# BggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEF
# BQcwAoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYD
# VR0gBDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVt
# LnBsL0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAdYhYD+WPUCiaU58Q7EP89DttyZqG
# Yn2XRDhJkL6P+/T0IPZyxfxiXumYlARMgwRzLRUStJl490L94C9LGF3vjzzH8Jq3
# iR74BRlkO18J3zIdmCKQa5LyZ48IfICJTZVJeChDUyuQy6rGDxLUUAsO0eqeLNhL
# Vsgw6/zOfImNlARKn1FP7o0fTbj8ipNGxHBIutiRsWrhWM2f8pXdd3x2mbJCKKtl
# 2s42g9KUJHEIiLni9ByoqIUul4GblLQigO0ugh7bWRLDm0CdY9rNLqyA3ahe8Wlx
# VWkxyrQLjH8ItI17RdySaYayX3PhRSC4Am1/7mATwZWwSD+B7eMcZNhpn8zJ+6MT
# yE6YoEBSRVrs0zFFIHUR08Wk0ikSf+lIe5Iv6RY3/bFAEloMU+vUBfSouCReZwSL
# o8WdrDlPXtR0gicDnytO7eZ5827NS2x7gCBibESYkOh1/w1tVxTpV2Na3PR7nxYV
# lPu1JPoRZCbH86gc96UTvuWiOruWmyOEMLOGGniR+x+zPF/2DaGgK2W1eEJfo2qy
# rBNPvF7wuAyQfiFXLwvWHamoYtPZo0LHuH8X3n9C+xN4YaNjt2ywzOr+tKyEVAot
# nyU9vyEVOaIYMk3IeBrmFnn0gbKeTTyYeEEUz/Qwt4HOUBCrW602NCmvO1nm+/80
# nLy5r0AZvCQxaQ4wgga5MIIEoaADAgECAhEA5/9pxzs1zkuRJth0fGilhzANBgkq
# hkiG9w0BAQwFADCBgDELMAkGA1UEBhMCUEwxIjAgBgNVBAoTGVVuaXpldG8gVGVj
# aG5vbG9naWVzIFMuQS4xJzAlBgNVBAsTHkNlcnR1bSBDZXJ0aWZpY2F0aW9uIEF1
# dGhvcml0eTEkMCIGA1UEAxMbQ2VydHVtIFRydXN0ZWQgTmV0d29yayBDQSAyMB4X
# DTIxMDUxOTA1MzIwN1oXDTM2MDUxODA1MzIwN1owVjELMAkGA1UEBhMCUEwxITAf
# BgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVt
# IFRpbWVzdGFtcGluZyAyMDIxIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA6RIfBDXtuV16xaaVQb6KZX9Od9FtJXXTZo7b+GEof3+3g0ChWiKnO7R4
# +6MfrvLyLCWZa6GpFHjEt4t0/GiUQvnkLOBRdBqr5DOvlmTvJJs2X8ZmWgWJjC7P
# BZLYBWAs8sJl3kNXxBMX5XntjqWx1ZOuuXl0R4x+zGGSMzZ45dpvB8vLpQfZkfMC
# /1tL9KYyjU+htLH68dZJPtzhqLBVG+8ljZ1ZFilOKksS79epCeqFSeAUm2eMTGpO
# iS3gfLM6yvb8Bg6bxg5yglDGC9zbr4sB9ceIGRtCQF1N8dqTgM/dSViiUgJkcv5d
# LNJeWxGCqJYPgzKlYZTgDXfGIeZpEFmjBLwURP5ABsyKoFocMzdjrCiFbTvJn+bD
# 1kq78qZUgAQGGtd6zGJ88H4NPJ5Y2R4IargiWAmv8RyvWnHr/VA+2PrrK9eXe5q7
# M88YRdSTq9TKbqdnITUgZcjjm4ZUjteq8K331a4P0s2in0p3UubMEYa/G5w6jSWP
# UzchGLwWKYBfeSu6dIOC4LkeAPvmdZxSB1lWOb9HzVWZoM8Q/blaP4LWt6JxjkI9
# yQsYGMdCqwl7uMnPUIlcExS1mzXRxUowQref/EPaS7kYVaHHQrp4XB7nTEtQhkP0
# Z9Puz/n8zIFnUSnxDof4Yy650PAXSYmK2TcbyDoTNmmt8xAxzcMCAwEAAaOCAVUw
# ggFRMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFL5UAi+/QGxzQ86sCSVOnkNE
# Gu7gMB8GA1UdIwQYMBaAFLahVDkCw6A/joq8+tT4HKbROg79MA4GA1UdDwEB/wQE
# AwIBBjATBgNVHSUEDDAKBggrBgEFBQcDCDAwBgNVHR8EKTAnMCWgI6Ahhh9odHRw
# Oi8vY3JsLmNlcnR1bS5wbC9jdG5jYTIuY3JsMGwGCCsGAQUFBwEBBGAwXjAoBggr
# BgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNvbTAyBggrBgEFBQcw
# AoYmaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNhMi5jZXIwOQYDVR0g
# BDIwMDAuBgRVHSAAMCYwJAYIKwYBBQUHAgEWGGh0dHA6Ly93d3cuY2VydHVtLnBs
# L0NQUzANBgkqhkiG9w0BAQwFAAOCAgEAuJNZd8lMFf2UBwigp3qgLPBBk58BFCS3
# Q6aJDf3TISoytK0eal/JyCB88aUEd0wMNiEcNVMbK9j5Yht2whaknUE1G32k6uld
# 7wcxHmw67vUBY6pSp8QhdodY4SzRRaZWzyYlviUpyU4dXyhKhHSncYJfa1U75cXx
# Ce3sTp9uTBm3f8Bj8LkpjMUSVTtMJ6oEu5JqCYzRfc6nnoRUgwz/GVZFoOBGdrSE
# tDN7mZgcka/tS5MI47fALVvN5lZ2U8k7Dm/hTX8CWOw0uBZloZEW4HB0Xra3qE4q
# zzq/6M8gyoU/DE0k3+i7bYOrOk/7tPJg1sOhytOGUQ30PbG++0FfJioDuOFhj99b
# 151SqFlSaRQYz74y/P2XJP+cF19oqozmi0rRTkfyEJIvhIZ+M5XIFZttmVQgTxfp
# fJwMFFEoQrSrklOxpmSygppsUDJEoliC05vBLVQ+gMZyYaKvBJ4YxBMlKH5ZHkRd
# loRYlUDplk8GUa+OCMVhpDSQurU6K1ua5dmZftnvSSz2H96UrQDzA6DyiI1V3ejV
# tvn2azVAXg6NnjmuRZ+wa7Pxy0H3+V4K4rOTHlG3VYA6xfLsTunCz72T6Ot4+tkr
# DYOeaU1pPX1CBfYj6EW2+ELq46GP8KCNUQDirWLU4nOmgCat7vN0SD6RlwUiSsMe
# CiQDmZwgwrUwggbpMIIE0aADAgECAhBiOsZKIV2oSfsf25d4iu6HMA0GCSqGSIb3
# DQEBCwUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0YSBTeXN0
# ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcgMjAyMSBDQTAe
# Fw0yNTA3MzExMTM4MDhaFw0yNjA3MzExMTM4MDdaMIGOMQswCQYDVQQGEwJERTEb
# MBkGA1UECAwSQmFkZW4tV8O8cnR0ZW1iZXJnMRQwEgYDVQQHDAtCYWllcnNicm9u
# bjEeMBwGA1UECgwVT3BlbiBTb3VyY2UgRGV2ZWxvcGVyMSwwKgYDVQQDDCNPcGVu
# IFNvdXJjZSBEZXZlbG9wZXIsIEhlcHAgQW5kcmVhczCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOt2txKXx2UtfBNIw2kVihIAcgPkK3lp7np/qE0evLq2
# J/L5kx8m6dUY4WrrcXPSn1+W2/PVs/XBFV4fDfwczZnQ/hYzc8Ot5YxPKLx6hZxK
# C5v8LjNIZ3SRJvMbOpjzWoQH7MLIIj64n8mou+V0CMk8UElmU2d0nxBQyau1njQP
# CLvlfInu4tDndyp3P87V5bIdWw6MkZFhWDkILTYInYicYEkut5dN9hT02t/3rXu2
# 30DEZ6S1OQtm9loo8wzvwjRoVX3IxnfpCHGW8Z9ie9I9naMAOG2YpvpoUbLG3fL/
# B6JVNNR1mm/AYaqVMtAXJpRlqvbIZyepcG0YGB+kOQLdoQCWlIp3a14Z4kg6bU9C
# U1KNR4ueA+SqLNu0QGtgBAdTfqoWvyiaeyEogstBHglrZ39y/RW8OOa50pSleSRx
# SXiGW+yH+Ps5yrOopTQpKHy0kRincuJpYXgxGdGxxKHwuVJHKXL0nWScEku0C38p
# M9sYanIKncuF0Ed7RvyNqmPP5pt+p/0ZG+zLNu/Rce0LE5FjAIRtW2hFxmYMyohk
# afzyjCCCG0p2KFFT23CoUfXx59nCU+lyWx/iyDMV4sqrcvmZdPZF7lkaIb5B4PYP
# vFFE7enApz4Niycj1gPUFlx4qTcXHIbFLJDp0ry6MYelX+SiMHV7yDH/rnWXm5d3
# AgMBAAGjggF4MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1UdHwQ2MDQwMqAwoC6GLGh0
# dHA6Ly9jY3NjYTIwMjEuY3JsLmNlcnR1bS5wbC9jY3NjYTIwMjEuY3JsMHMGCCsG
# AQUFBwEBBGcwZTAsBggrBgEFBQcwAYYgaHR0cDovL2Njc2NhMjAyMS5vY3NwLWNl
# cnR1bS5jb20wNQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5w
# bC9jY3NjYTIwMjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA23ufoHTKsW73PMAywHDN
# MB0GA1UdDgQWBBQYl6R41hwxInb9JVvqbCTp9ILCcTBLBgNVHSAERDBCMAgGBmeB
# DAEEATA2BgsqhGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIBFhlodHRwczovL3d3dy5j
# ZXJ0dW0ucGwvQ1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIH
# gDANBgkqhkiG9w0BAQsFAAOCAgEAQ4guyo7zysB7MHMBOVKKY72rdY5hrlxPci8u
# 1RgBZ9ZDGFzhnUM7iIivieAeAYLVxP922V3ag9sDVNR+mzCmu1pWCgZyBbNXykue
# KJwOfE8VdpmC/F7637i8a7Pyq6qPbcfvLSqiXtVrT4NX4NIvODW3kIqf4nGwd0h3
# 1tuJVHLkdpGmT0q4TW0gAxnNoQ+lO8uNzCrtOBk+4e1/3CZXSDnjR8SUsHrHdhnm
# qkAnYb40vf69dfDR148tToUj872yYeBUEGUsQUDgJ6HSkMVpLQz/Nb3xy9qkY33M
# 7CBWKuBVwEcbGig/yj7CABhIrY1XwRddYQhEyozUS4mXNqXydAD6Ylt143qrECD2
# s3MDQBgP2sbRHdhVgzr9+n1iztXkPHpIlnnXPkZrt89E5iGL+1PtjETrhTkr7nxj
# yMFjrbmJ8W/XglwopUTCGfopDFPlzaoFf5rH/v3uzS24yb6+dwQrvCwFA9Y9ZHy2
# ITJx7/Ll6AxWt7Lz9JCJ5xRyYeRUHs6ycB8EuMPAKyGpzdGtjWv2rkTXbkIYUjkl
# FTpquXJBc/kO5L+Quu0a0uKn4ea16SkABy052XHQqd87cSJg3rGxsagi0IAfxGM6
# 08oupufSS/q9mpQPgkDuMJ8/zdre0st8OduAoG131W+XJ7mm0gIuh2zNmSIet5RD
# oa8THmwxggckMIIHIAIBATBqMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3Nl
# Y28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25p
# bmcgMjAyMSBDQQIQYjrGSiFdqEn7H9uXeIruhzANBglghkgBZQMEAgEFAKCBhDAY
# BgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEi
# BCDaNBgav/E8DPdhWCKajKuDqq0wb/hNQkNA0zFeAwIYPzANBgkqhkiG9w0BAQEF
# AASCAgAH2An5gVVqYCJ28cZicvtngNv9v/4UvVjPdsBAU0NX1PN+/pQbEUblTXug
# l9wtZHjv362cfvZpQc6Pxj+jPciHNDYHN9BFZARnqqX2MeYjsNfO26AOyQ0k/+QN
# FKPds2vU4vpoYaX0pETjjB2NppUrpLdvKZaY/u0gG7U908/YC2p4uN9IyEtibxxZ
# lJ5leTnRjc+s1KS2GC8wrm1Koe1tQQQtTmTSznsvKuMsBDKKwURjrgnm/5lTpUFM
# IElNZWjTyRAiYb0pPiJLz2+RtslS1zjuOz6m3aMK6KdcJmlEG2PVxUk3t0aBVBa3
# Gsq22dIJpapASxMXscovr9YpmlbJYFAQWePJ7YhRPrmUyQPHvue/wXdqmTpOyKcp
# oE+CT7R0ThzJ02XZaM+YjaygDE82Ie0JbWTjbM9W5rudsS2vvDxpXRwboTPJioni
# +/hefk/9+L/GEaaP080zO0q3wxdi+umICArQsqkszp6jqhZhkqGZU9Op04Yf2N5Q
# IxH9iigK+Zq8MWvHFjDUwMa7L5RS1Z0/RvjbTXuCPgZlsmNDWvdAE0Wj32VyX205
# J9ciprtNeG7+ei2gDc6wTRW0KDnLibgy7cpoiwSY1sQH9weQEFJfM95IperVfM5g
# ztNbAo5XoJ1PFFZbr+fgH4gYhYGeKkuMx/+Hs7BjU2vY94QkS6GCBAQwggQABgkq
# hkiG9w0BCQYxggPxMIID7QIBATBrMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhB
# c3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3Rh
# bXBpbmcgMjAyMSBDQQIRAJ6cBPZVqLSnAm1JjGx4jaowDQYJYIZIAWUDBAICBQCg
# ggFXMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcN
# MjUxMDI3MDcyNzIwWjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCDPodw1ne0rw8uJ
# D6Iw5dr3e1QPGm4rI93PF1ThjPqg1TA/BgkqhkiG9w0BCQQxMgQw4oh9KF8UH0DN
# kOijEB8Zl9lI/ALW6SrN4KTaYoxcGpcNdP2OjelGd60g+a95xmecMIGgBgsqhkiG
# 9w0BCRACDDGBkDCBjTCBijCBhwQUwyW4mxf8xQJgYc4rcXtFB92camowbzBapFgw
# VjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3RlbXMgUy5B
# LjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEAnpwE9lWo
# tKcCbUmMbHiNqjANBgkqhkiG9w0BAQEFAASCAgAe/t+FmfCEG1YmACqxYDifZ339
# y4U+MMYQf0YRciHLFut9jf9Os8bFeCuKvfrC+yA3XeYsZ3Cc3H9H9qlmWUWIqdKb
# 925kHV7KQH3rx1u2SC7Ria8u2Xs1mMx4QieGfvxw9E5SQSgbssWaJ99Uv9taIw+X
# Uj89vdM1JyBfi3M+1L3YeSpACqTvN7Pcedr2h/OeVVTsfWarnFF2+0WDqYqGLy81
# BOy22OowOGO7nmzDdSZxuSZMwy3ON/wW9wMeL46eeGOBFMCSMBr3MwK3VUxlKmQC
# 7Pr+Piwwn6AX9ODB+9Zg7Gj3n+lrTD85yXjJcR32YY3rryi6VwC1tN1JVZ0GjuT3
# v/PfR4mWeO7vd5IYfHKaREKoOEa25Os97jnAh8HXCSiaf22qxRk7vv45JC2w/8oH
# 1obKdpbwUlVDVQDu0PSw8dZxy/HEJ6wUvNeZDNPyQI9e+zyo6nUQ241vRM3A8/Sv
# dCvO7/FUMRFKH85d3QXJRH4QXrQWQc90nbv6LOk8LA6xQ+6dgeTkjSaOcfO8PUCv
# 4I2ZzfJ/jMpcqemN2meeDveJABsPSBEqb9fG+ZQ4KUMknLJ3oi6g2Rubq3dJJjJC
# sNXXBKZS9eQ8e2XOFdenETytLCVtLUAfKYrR7lBNqc5EqBUnsTUZecsjNgHfSdJw
# LiyycZ+/SU5bhHphGQ==
# SIG # End signature block
