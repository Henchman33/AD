# powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "C:\path\AD_Reporter_GUI_Ver7.ps1"
<#
.SYNOPSIS
    AD Reporter GUI v7 - Servers, DCs, Users, Groups with Credential Vault, LAPS, and Export.

.NOTES
    Author  : Stephen McKee - Server Administrator - IGT PLC
    Version : 7.0
    Requires: ActiveDirectory module. ImportExcel, PSWriteWord, CredentialManager are optional.
    Run in STA mode. Script auto-relaunches in STA if needed.
    Treat exported reports and LAPS values as sensitive.
#>

#Requires -Version 5.1

# ── STA Guard ──────────────────────────────────────────────────────────────────
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $psExe = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path $psExe)) { $psExe = 'pwsh' }
    $relaunchArgs = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    Start-Process -FilePath $psExe -ArgumentList $relaunchArgs
    exit
}

# ── Optional Module Installer ──────────────────────────────────────────────────
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not install module '$Name': $($_.Exception.Message)"
        }
    }
}
Ensure-Module -Name CredentialManager
Ensure-Module -Name ImportExcel
Ensure-Module -Name PSWriteWord

# ── WPF Assembly ───────────────────────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ══════════════════════════════════════════════════════════════════════════════
#  XAML  –  Dark Professional Theme
# ══════════════════════════════════════════════════════════════════════════════
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AD Reporter  v7  |  IGT PLC"
    Height="840" Width="1340"
    MinHeight="700" MinWidth="1100"
    WindowStartupLocation="CenterScreen"
    Background="#1E1E2E"
    Foreground="#CDD6F4">

  <Window.Resources>

    <!-- ── Colour Palette ──────────────────────────────────────── -->
    <SolidColorBrush x:Key="BgBase"        Color="#1E1E2E"/>
    <SolidColorBrush x:Key="BgSurface"     Color="#2A2A3E"/>
    <SolidColorBrush x:Key="BgElevated"    Color="#313149"/>
    <SolidColorBrush x:Key="BgInput"       Color="#24243A"/>
    <SolidColorBrush x:Key="BorderNormal"  Color="#45475A"/>
    <SolidColorBrush x:Key="BorderFocus"   Color="#89B4FA"/>
    <SolidColorBrush x:Key="AccentBlue"    Color="#89B4FA"/>
    <SolidColorBrush x:Key="AccentGreen"   Color="#A6E3A1"/>
    <SolidColorBrush x:Key="AccentRed"     Color="#F38BA8"/>
    <SolidColorBrush x:Key="AccentYellow"  Color="#F9E2AF"/>
    <SolidColorBrush x:Key="AccentPurple"  Color="#CBA6F7"/>
    <SolidColorBrush x:Key="TextPrimary"   Color="#CDD6F4"/>
    <SolidColorBrush x:Key="TextMuted"     Color="#6C7086"/>
    <SolidColorBrush x:Key="TextSubtle"    Color="#9399B2"/>

    <!-- ── Button Style ────────────────────────────────────────── -->
    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Background"         Value="#89B4FA"/>
      <Setter Property="Foreground"         Value="#1E1E2E"/>
      <Setter Property="FontWeight"         Value="SemiBold"/>
      <Setter Property="FontSize"           Value="12"/>
      <Setter Property="Padding"            Value="12,5"/>
      <Setter Property="BorderThickness"    Value="0"/>
      <Setter Property="Cursor"             Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bdr" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bdr" Property="Background" Value="#B4C7F8"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bdr" Property="Background" Value="#6A9CE8"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bdr" Property="Background" Value="#45475A"/>
                <Setter Property="Foreground" Value="#6C7086"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnSecondary" TargetType="Button">
      <Setter Property="Background"         Value="#313149"/>
      <Setter Property="Foreground"         Value="#CDD6F4"/>
      <Setter Property="FontSize"           Value="12"/>
      <Setter Property="Padding"            Value="10,5"/>
      <Setter Property="BorderBrush"        Value="#45475A"/>
      <Setter Property="BorderThickness"    Value="1"/>
      <Setter Property="Cursor"             Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bdr" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bdr" Property="Background"   Value="#3E3E5A"/>
                <Setter TargetName="bdr" Property="BorderBrush"  Value="#89B4FA"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bdr" Property="Background" Value="#2A2A3E"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnDanger" TargetType="Button" BasedOn="{StaticResource BtnSecondary}">
      <Setter Property="Foreground" Value="#F38BA8"/>
      <Setter Property="BorderBrush" Value="#F38BA8"/>
    </Style>

    <Style x:Key="BtnSuccess" TargetType="Button" BasedOn="{StaticResource BtnSecondary}">
      <Setter Property="Foreground"   Value="#A6E3A1"/>
      <Setter Property="BorderBrush"  Value="#A6E3A1"/>
    </Style>

    <!-- ── TextBox / ComboBox / PasswordBox Shared ─────────────── -->
    <Style TargetType="TextBox">
      <Setter Property="Background"       Value="#24243A"/>
      <Setter Property="Foreground"       Value="#CDD6F4"/>
      <Setter Property="CaretBrush"       Value="#89B4FA"/>
      <Setter Property="BorderBrush"      Value="#45475A"/>
      <Setter Property="BorderThickness"  Value="1"/>
      <Setter Property="Padding"          Value="6,4"/>
      <Setter Property="FontSize"         Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Style.Triggers>
        <Trigger Property="IsFocused" Value="True">
          <Setter Property="BorderBrush" Value="#89B4FA"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="PasswordBox">
      <Setter Property="Background"       Value="#24243A"/>
      <Setter Property="Foreground"       Value="#CDD6F4"/>
      <Setter Property="CaretBrush"       Value="#89B4FA"/>
      <Setter Property="BorderBrush"      Value="#45475A"/>
      <Setter Property="BorderThickness"  Value="1"/>
      <Setter Property="Padding"          Value="6,4"/>
      <Setter Property="FontSize"         Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Style.Triggers>
        <Trigger Property="IsFocused" Value="True">
          <Setter Property="BorderBrush" Value="#89B4FA"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Background"       Value="#24243A"/>
      <Setter Property="Foreground"       Value="#CDD6F4"/>
      <Setter Property="BorderBrush"      Value="#45475A"/>
      <Setter Property="BorderThickness"  Value="1"/>
      <Setter Property="Padding"          Value="6,4"/>
      <Setter Property="FontSize"         Value="12"/>
      <Setter Property="Height"           Value="28"/>
    </Style>

    <!-- ComboBoxItem -->
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background" Value="#2A2A3E"/>
      <Setter Property="Foreground" Value="#CDD6F4"/>
      <Setter Property="Padding"    Value="6,3"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="#3E3E5A"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- ── CheckBox ────────────────────────────────────────────── -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground"  Value="#CDD6F4"/>
      <Setter Property="FontSize"    Value="12"/>
      <Setter Property="Margin"      Value="0,3,0,3"/>
    </Style>

    <!-- ── Separator ───────────────────────────────────────────── -->
    <Style TargetType="Separator">
      <Setter Property="Background" Value="#45475A"/>
      <Setter Property="Margin"     Value="0,6,0,6"/>
    </Style>

    <!-- ── GroupBox ────────────────────────────────────────────── -->
    <Style TargetType="GroupBox">
      <Setter Property="Foreground"      Value="#89B4FA"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="BorderBrush"     Value="#45475A"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="8,6"/>
      <Setter Property="Margin"          Value="0,0,0,8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GroupBox">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Grid.RowSpan="2"
                      Background="#2A2A3E"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}"
                      CornerRadius="6"/>
              <Border Grid.Row="0" Margin="10,0,0,0" Background="#2A2A3E" Padding="4,0">
                <ContentPresenter ContentSource="Header"
                                  TextBlock.Foreground="#89B4FA"
                                  TextBlock.FontWeight="SemiBold"
                                  TextBlock.FontSize="11"/>
              </Border>
              <ContentPresenter Grid.Row="1"
                                Margin="{TemplateBinding Padding}"
                                ContentSource="Content"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ── DataGrid ─────────────────────────────────────────────── -->
    <Style TargetType="DataGrid">
      <Setter Property="Background"              Value="#1E1E2E"/>
      <Setter Property="Foreground"              Value="#CDD6F4"/>
      <Setter Property="GridLinesVisibility"     Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#313149"/>
      <Setter Property="BorderBrush"             Value="#45475A"/>
      <Setter Property="BorderThickness"         Value="1"/>
      <Setter Property="RowBackground"           Value="#1E1E2E"/>
      <Setter Property="AlternatingRowBackground" Value="#242436"/>
      <Setter Property="ColumnHeaderHeight"      Value="30"/>
      <Setter Property="FontSize"                Value="12"/>
    </Style>

    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background"  Value="#313149"/>
      <Setter Property="Foreground"  Value="#89B4FA"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Padding"     Value="8,0"/>
      <Setter Property="BorderBrush" Value="#45475A"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>

    <Style TargetType="DataGridRow">
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#3E4A6A"/>
          <Setter Property="Foreground" Value="#CDD6F4"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#2E2E46"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="DataGridCell">
      <Setter Property="Padding"          Value="6,2"/>
      <Setter Property="BorderThickness"  Value="0"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
          <Setter Property="Foreground" Value="#CDD6F4"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- ── ScrollBar ─────────────────────────────────────────────── -->
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="#2A2A3E"/>
    </Style>

    <!-- ── Label (section header) ───────────────────────────────── -->
    <Style x:Key="SectionLabel" TargetType="TextBlock">
      <Setter Property="Foreground"  Value="#89B4FA"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="FontSize"    Value="11"/>
      <Setter Property="Margin"      Value="0,0,0,4"/>
    </Style>

    <Style x:Key="MutedLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#6C7086"/>
      <Setter Property="FontSize"   Value="11"/>
    </Style>

    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="Foreground"          Value="#9399B2"/>
      <Setter Property="FontSize"            Value="12"/>
      <Setter Property="VerticalAlignment"   Value="Center"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="52"/>   <!-- Title / Ribbon bar -->
      <RowDefinition Height="52"/>   <!-- Toolbar row        -->
      <RowDefinition Height="*"/>    <!-- Main content       -->
      <RowDefinition Height="26"/>   <!-- Status bar         -->
    </Grid.RowDefinitions>

    <!-- ══ ROW 0 – Title Banner ════════════════════════════════════════════ -->
    <Border Grid.Row="0" Background="#181825" BorderBrush="#45475A" BorderThickness="0,0,0,1">
      <Grid Margin="16,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Icon + Title -->
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="32" Height="32" CornerRadius="6" Background="#89B4FA" Margin="0,0,10,0">
            <TextBlock Text="AD" Foreground="#1E1E2E" FontWeight="Bold" FontSize="13"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="AD Reporter" Foreground="#CDD6F4" FontSize="16" FontWeight="Bold"/>
            <TextBlock Text="Active Directory Query &amp; Export Tool" Foreground="#6C7086" FontSize="10"/>
          </StackPanel>
        </StackPanel>

        <!-- Version badge -->
        <Border Grid.Column="2" Background="#313149" CornerRadius="4" Padding="10,4" VerticalAlignment="Center">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="v7.0  " Foreground="#CBA6F7" FontWeight="Bold" FontSize="11"/>
            <TextBlock Text="|  IGT PLC" Foreground="#6C7086" FontSize="11"/>
          </StackPanel>
        </Border>
      </Grid>
    </Border>

    <!-- ══ ROW 1 – Query Toolbar ════════════════════════════════════════════ -->
    <Border Grid.Row="1" Background="#242436" BorderBrush="#45475A" BorderThickness="0,0,0,1">
      <Grid Margin="12,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>  <!-- Object type label  -->
          <ColumnDefinition Width="190"/>   <!-- Object type combo  -->
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>  <!-- SearchBase label   -->
          <ColumnDefinition Width="420"/>   <!-- SearchBase textbox -->
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>  <!-- Preset label       -->
          <ColumnDefinition Width="260"/>   <!-- Preset combo       -->
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>  <!-- Run button         -->
          <ColumnDefinition Width="*"/>     <!-- Spacer             -->
          <ColumnDefinition Width="Auto"/>  <!-- Refresh Presets    -->
        </Grid.ColumnDefinitions>

        <TextBlock Grid.Column="0" Text="Object:" Style="{StaticResource FieldLabel}" Margin="0,0,6,0"/>
        <ComboBox  x:Name="cbObject" Grid.Column="1" SelectedIndex="0" VerticalAlignment="Center">
          <ComboBoxItem>Servers</ComboBoxItem>
          <ComboBoxItem>DomainControllers</ComboBoxItem>
          <ComboBoxItem>Users</ComboBoxItem>
          <ComboBoxItem>Groups</ComboBoxItem>
          <ComboBoxItem>SecurityGroups</ComboBoxItem>
          <ComboBoxItem>ServiceAccounts</ComboBoxItem>
          <ComboBoxItem>ManagedServiceAccounts</ComboBoxItem>
        </ComboBox>

        <TextBlock Grid.Column="3" Text="Search Base:" Style="{StaticResource FieldLabel}" Margin="0,0,6,0"/>
        <TextBox   x:Name="tbSearchBase" Grid.Column="4" VerticalAlignment="Center"
                   ToolTip="DN of the OU to search, e.g. OU=Servers,DC=myigt,DC=com"/>

        <TextBlock Grid.Column="6" Text="Preset:" Style="{StaticResource FieldLabel}" Margin="0,0,6,0"/>
        <ComboBox  x:Name="cbPreset" Grid.Column="7" VerticalAlignment="Center"/>

        <Button x:Name="btnRun" Grid.Column="9" Content="▶  Run Query"
                Style="{StaticResource BtnPrimary}" Width="110" Height="30" VerticalAlignment="Center"/>

        <Button x:Name="btnRefreshPresets" Grid.Column="11" Content="↺  Refresh"
                Style="{StaticResource BtnSecondary}" Height="28" Padding="10,4" VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <!-- ══ ROW 2 – Main Content (Left + Right panels) ═══════════════════════ -->
    <Grid Grid.Row="2" Margin="10,8,10,4">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="360"/>
      </Grid.ColumnDefinitions>

      <!-- ── LEFT PANEL ──────────────────────────────────────────── -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>  <!-- Command Preview group  -->
          <RowDefinition Height="Auto"/>  <!-- DataGrid toolbar row   -->
          <RowDefinition Height="*"/>     <!-- DataGrid               -->
        </Grid.RowDefinitions>

        <!-- Command Preview GroupBox -->
        <GroupBox Grid.Row="0" Header="  Command Preview" Margin="0,0,0,6">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="110"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBox x:Name="tbCommandPreview"
                     AcceptsReturn="True"
                     VerticalScrollBarVisibility="Auto"
                     TextWrapping="Wrap"
                     IsReadOnly="True"
                     Background="#1A1A2A"
                     FontFamily="Consolas"
                     FontSize="11"
                     Foreground="#A6E3A1"
                     BorderThickness="0"
                     Padding="6,4"/>
            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0">
              <Button x:Name="btnEditPreview" Content="✏  Edit Preview"
                      Style="{StaticResource BtnSecondary}" Height="26" Padding="8,3" FontSize="11"/>
            </StackPanel>
          </Grid>
        </GroupBox>

        <!-- DataGrid Toolbar -->
        <Border Grid.Row="1" Background="#242436" CornerRadius="4"
                BorderBrush="#45475A" BorderThickness="1" Padding="8,5" Margin="0,0,0,6">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="8"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="8"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <Button x:Name="btnMaskToggle"     Grid.Column="0"
                    Content="🔒  Mask LAPS (On)"
                    Style="{StaticResource BtnSecondary}" Height="26" Padding="8,3" FontSize="11"/>
            <Button x:Name="btnRevealSelected" Grid.Column="2"
                    Content="👁  Reveal Selected"
                    Style="{StaticResource BtnDanger}"   Height="26" Padding="8,3" FontSize="11"/>
            <Button x:Name="btnRefreshGrid"    Grid.Column="4"
                    Content="⟳  Refresh Grid"
                    Style="{StaticResource BtnSecondary}" Height="26" Padding="8,3" FontSize="11"/>

            <TextBlock Grid.Column="5"
                       Text="  Select rows then Reveal to fetch LAPS values"
                       Style="{StaticResource MutedLabel}"
                       VerticalAlignment="Center"/>

            <!-- Result count badge -->
            <Border Grid.Column="6" Background="#313149" CornerRadius="4" Padding="8,3" VerticalAlignment="Center">
              <TextBlock x:Name="tbResultCount" Text="0 rows" Foreground="#CBA6F7" FontSize="11" FontWeight="SemiBold"/>
            </Border>
          </Grid>
        </Border>

        <!-- Results DataGrid -->
        <Border Grid.Row="2" BorderBrush="#45475A" BorderThickness="1" CornerRadius="4">
          <DataGrid x:Name="dgResults"
                    AutoGenerateColumns="True"
                    SelectionMode="Extended"
                    CanUserAddRows="False"
                    CanUserDeleteRows="False"
                    IsReadOnly="True"
                    BorderThickness="0"/>
        </Border>
      </Grid>

      <!-- ── RIGHT PANEL ─────────────────────────────────────────── -->
      <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
        <StackPanel>

          <!-- Sign-In GroupBox -->
          <GroupBox Header="  Sign In / Credential Management">
            <StackPanel>

              <Grid Margin="0,0,0,5">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="72"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Domain" Style="{StaticResource FieldLabel}"/>
                <ComboBox x:Name="cbLoginDomain" Grid.Column="1" IsEditable="True"
                          ToolTip="Active Directory domain FQDN or NetBIOS name"/>
              </Grid>

              <Grid Margin="0,0,0,5">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="72"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Username" Style="{StaticResource FieldLabel}"/>
                <TextBox x:Name="tbLoginUser" Grid.Column="1"
                         ToolTip="DOMAIN\username or UPN format"/>
              </Grid>

              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="72"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Password" Style="{StaticResource FieldLabel}"/>
                <PasswordBox x:Name="pbLoginPass" Grid.Column="1"/>
              </Grid>

              <!-- Options row -->
              <WrapPanel Margin="0,0,0,6">
                <CheckBox x:Name="chkSaveCred"      Content="Save to session"                   IsChecked="True"/>
                <CheckBox x:Name="chkPersistVault"  Content="Persist to Credential Manager"     IsChecked="False" Margin="8,3,0,3"/>
              </WrapPanel>

              <!-- Action buttons -->
              <Grid Margin="0,0,0,4">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="6"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="btnSignIn"   Grid.Column="0" Content="🔑  Sign In / Test"
                        Style="{StaticResource BtnPrimary}" Height="30"/>
                <Button x:Name="btnUseSaved" Grid.Column="2" Content="📋  Use Saved"
                        Style="{StaticResource BtnSecondary}" Height="30"/>
              </Grid>

              <!-- Status -->
              <Border Background="#1A1A2A" CornerRadius="4" Padding="6,4" Margin="0,2,0,6">
                <TextBlock x:Name="tbSignInStatus" Text="Enter credentials above and click Sign In"
                           Foreground="#6C7086" TextWrapping="Wrap" FontSize="11"/>
              </Border>

              <!-- Saved credentials -->
              <TextBlock Text="Saved Credentials  (session + vault)" Style="{StaticResource SectionLabel}"/>
              <ComboBox x:Name="cbCredentials" Margin="0,0,0,0"
                        ToolTip="Select a saved credential to use for queries"/>
            </StackPanel>
          </GroupBox>

          <!-- Export GroupBox -->
          <GroupBox Header="  Export Options">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0">
                <CheckBox x:Name="cbCsv"  Content="CSV"  IsChecked="True"/>
                <CheckBox x:Name="cbHtml" Content="HTML" IsChecked="True"/>
                <CheckBox x:Name="cbXlsx" Content="XLSX (ImportExcel)"/>
              </StackPanel>
              <StackPanel Grid.Column="1">
                <CheckBox x:Name="cbDocx" Content="DOCX (PSWriteWord)"/>
                <CheckBox x:Name="cbPdf"  Content="PDF (wkhtmltopdf)"/>
              </StackPanel>
            </Grid>
          </GroupBox>

          <!-- Max Results + Export button -->
          <GroupBox Header="  Query &amp; Export Settings">
            <StackPanel>
              <Grid Margin="0,0,0,6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="6"/>
                  <ColumnDefinition Width="110"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Text="Max Results" Style="{StaticResource FieldLabel}" Margin="0,0,0,3"/>
                  <TextBox x:Name="tbMax" Text="1000" Height="28"/>
                </StackPanel>
                <Button x:Name="btnExport" Grid.Column="2"
                        Content="💾  Export"
                        Style="{StaticResource BtnSuccess}"
                        Height="28" VerticalAlignment="Bottom"/>
              </Grid>

              <TextBlock Style="{StaticResource MutedLabel}"
                         TextWrapping="Wrap"
                         Text="Exports are saved to your Desktop with timestamp prefix ADReport_*."/>
            </StackPanel>
          </GroupBox>

          <!-- LAPS Audit GroupBox -->
          <GroupBox Header="  LAPS Audit">
            <StackPanel>
              <Border Background="#1A1A2A" CornerRadius="4" Padding="8,6">
                <StackPanel>
                  <TextBlock Foreground="#F9E2AF" FontWeight="SemiBold" FontSize="11"
                             Text="⚠  LAPS Reveal Logging"/>
                  <TextBlock Foreground="#9399B2" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"
                             Text="All LAPS reveal actions are logged to:"/>
                  <TextBlock Foreground="#CBA6F7" FontSize="10" FontFamily="Consolas"
                             TextWrapping="Wrap" Margin="0,3,0,0"
                             Text="C:\SecureLogs\ADReports\LAPS_Audit.csv"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </GroupBox>

        </StackPanel>
      </ScrollViewer>
    </Grid>

    <!-- ══ ROW 3 – Status Bar ═══════════════════════════════════════════════ -->
    <Border Grid.Row="3" Background="#181825" BorderBrush="#45475A" BorderThickness="0,1,0,0">
      <Grid Margin="12,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse x:Name="elStatus" Width="8" Height="8" Fill="#A6E3A1" Margin="0,0,6,0"/>
          <TextBlock Text="Status:" Foreground="#6C7086" FontSize="11" Margin="0,0,4,0"/>
          <TextBlock x:Name="tbStatus" Text="Ready" Foreground="#9399B2" FontSize="11"/>
        </StackPanel>

        <TextBlock Grid.Column="2" VerticalAlignment="Center"
                   Foreground="#45475A" FontSize="10"
                   Text="AD Reporter v7.0  |  Stephen McKee  |  IGT PLC"/>
      </Grid>
    </Border>

  </Grid>
</Window>
"@

# ══════════════════════════════════════════════════════════════════════════════
#  XAML LOAD
# ══════════════════════════════════════════════════════════════════════════════
try {
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "XAML load failed:`n$($_.Exception.Message)",
        "AD Reporter – Fatal Error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    throw
}

# ══════════════════════════════════════════════════════════════════════════════
#  CONTROL BINDINGS
# ══════════════════════════════════════════════════════════════════════════════
$cbObject           = $window.FindName('cbObject')
$tbSearchBase       = $window.FindName('tbSearchBase')
$btnRun             = $window.FindName('btnRun')
$btnRefreshPresets  = $window.FindName('btnRefreshPresets')
$cbPreset           = $window.FindName('cbPreset')
$tbCommandPreview   = $window.FindName('tbCommandPreview')
$btnEditPreview     = $window.FindName('btnEditPreview')
$dgResults          = $window.FindName('dgResults')
$btnMaskToggle      = $window.FindName('btnMaskToggle')
$btnRevealSelected  = $window.FindName('btnRevealSelected')
$btnRefreshGrid     = $window.FindName('btnRefreshGrid')
$tbStatus           = $window.FindName('tbStatus')
$tbResultCount      = $window.FindName('tbResultCount')
$elStatus           = $window.FindName('elStatus')

$cbLoginDomain      = $window.FindName('cbLoginDomain')
$tbLoginUser        = $window.FindName('tbLoginUser')
$pbLoginPass        = $window.FindName('pbLoginPass')
$chkSaveCred        = $window.FindName('chkSaveCred')
$chkPersistVault    = $window.FindName('chkPersistVault')
$btnSignIn          = $window.FindName('btnSignIn')
$btnUseSaved        = $window.FindName('btnUseSaved')
$tbSignInStatus     = $window.FindName('tbSignInStatus')
$cbCredentials      = $window.FindName('cbCredentials')

$btnExport          = $window.FindName('btnExport')
$cbCsv              = $window.FindName('cbCsv')
$cbHtml             = $window.FindName('cbHtml')
$cbXlsx             = $window.FindName('cbXlsx')
$cbDocx             = $window.FindName('cbDocx')
$cbPdf              = $window.FindName('cbPdf')
$tbMax              = $window.FindName('tbMax')

# ══════════════════════════════════════════════════════════════════════════════
#  RUNTIME STATE
# ══════════════════════════════════════════════════════════════════════════════
$CredentialStore   = @{}
$global:LAPSCache  = @{}
$global:MaskLAPS   = $true
$global:results    = $null

# ── Status helpers ─────────────────────────────────────────────────────────────
function Set-StatusReady {
    $tbStatus.Text = "Ready"
    $elStatus.Fill = [System.Windows.Media.Brushes]::LightGreen
}
function Set-StatusBusy { param($msg = "Working…")
    $tbStatus.Text = $msg
    $elStatus.Fill = [System.Windows.Media.Brushes]::Orange
}
function Set-StatusError { param($msg = "Error")
    $tbStatus.Text = $msg
    $elStatus.Fill = [System.Windows.Media.Brushes]::Tomato
}

# ══════════════════════════════════════════════════════════════════════════════
#  LAPS AUDIT
# ══════════════════════════════════════════════════════════════════════════════
$global:AuditPath = "C:\SecureLogs\ADReports\LAPS_Audit.csv"
if (-not (Test-Path (Split-Path $global:AuditPath))) {
    New-Item -ItemType Directory -Path (Split-Path $global:AuditPath) -Force | Out-Null
}
if (-not (Test-Path $global:AuditPath)) {
    "Timestamp,Operator,CredentialLabel,Target,Action,Result" |
        Out-File -FilePath $global:AuditPath -Encoding UTF8
}

function Write-LAPSAudit {
    param($operator, $credLabel, $target, $action, $result)
    $ts   = (Get-Date).ToString("o")
    $line = "$ts,$operator,$credLabel,$target,$action,$result"
    Add-Content -Path $global:AuditPath -Value $line
}

# ══════════════════════════════════════════════════════════════════════════════
#  CREDENTIAL MANAGER HELPERS
# ══════════════════════════════════════════════════════════════════════════════
Import-Module CredentialManager -ErrorAction SilentlyContinue

function Save-CredToVault {
    param($label, [System.Management.Automation.PSCredential]$cred)
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password))
        New-StoredCredential -Target $label -UserName $cred.UserName -Password $plain -Persist LocalMachine | Out-Null
        return $true
    }
    catch { return $false }
}

function Get-CredFromVault {
    param($label)
    try {
        $s = Get-StoredCredential -Target $label -ErrorAction Stop
        if ($s) {
            $secure = ConvertTo-SecureString $s.Password -AsPlainText -Force
            return New-Object System.Management.Automation.PSCredential ($s.UserName, $secure)
        }
    }
    catch { }
    return $null
}

function Remove-CredFromVault {
    param($label)
    try { Remove-StoredCredential -Target $label -ErrorAction Stop; return $true }
    catch { return $false }
}

function Load-CredsFromVault {
    $vaultList = @()
    try {
        $out = cmdkey /list 2>$null
        if ($out) {
            foreach ($line in $out) {
                if ($line -match 'Target: (.+)') {
                    $target = $Matches[1].Trim()
                    if ($target -like 'ADReporter::*') {
                        $cred = Get-CredFromVault -label $target
                        if ($cred) { $vaultList += @{ Label = $target; Cred = $cred } }
                    }
                }
            }
        }
    }
    catch { }
    return $vaultList
}

function Refresh-CredCombo {
    $cbCredentials.Items.Clear()
    foreach ($k in $CredentialStore.Keys) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $k
        $cbCredentials.Items.Add($item) | Out-Null
    }
    $vaults = Load-CredsFromVault
    foreach ($v in $vaults) {
        $label = $v.Label
        if (-not $CredentialStore.ContainsKey($label)) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $label
            $cbCredentials.Items.Add($item) | Out-Null
        }
    }
    if ($cbCredentials.Items.Count -gt 0) { $cbCredentials.SelectedIndex = 0 }
}

function Get-LoginCredential {
    $user = $tbLoginUser.Text.Trim()
    $pass = $pbLoginPass.Password
    if (-not $user -or -not $pass) { return $null }
    $secure = ConvertTo-SecureString -String $pass -AsPlainText -Force
    return New-Object System.Management.Automation.PSCredential ($user, $secure)
}

function Populate-LoginDomainCombo {
    param($default)
    $cbLoginDomain.Items.Clear()
    foreach ($k in $CredentialStore.Keys) {
        if ($k -match '^ADReporter::(?<dom>[^ ]+)\s*-\s*') {
            $dom = $Matches['dom']
            if (-not ($cbLoginDomain.Items | Where-Object { $_.Content -eq $dom })) {
                $item = New-Object System.Windows.Controls.ComboBoxItem
                $item.Content = $dom
                $cbLoginDomain.Items.Add($item) | Out-Null
            }
        }
    }
    if ($default) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $default
        $cbLoginDomain.Items.Add($item) | Out-Null
    }
    if ($cbLoginDomain.Items.Count -gt 0) { $cbLoginDomain.SelectedIndex = 0 }
}

# ══════════════════════════════════════════════════════════════════════════════
#  PRESETS
# ══════════════════════════════════════════════════════════════════════════════
$Presets = @{}

# ── Servers ────────────────────────────────────────────────────────────────────
$Presets['Servers'] = @(
    @{
        Label   = 'Inventory: Name, OS, IPv4, Owner, Description, OU, Enabled, LastLogon, LAPS'
        Preview = 'Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase "<SearchBase>" -Properties Name,OperatingSystem,IPv4Address,ManagedBy,Description,Enabled,LastLogonDate,OperatingSystemVersion,DistinguishedName,ms-Mcs-AdmPwd | Select Name,OperatingSystem,IPv4Address,@{n="Owner";e={ if ($_.ManagedBy) { (Get-ADUser -Identity $_.ManagedBy -EA SilentlyContinue).SamAccountName } else { $null }}},Description,@{n="OU";e={ ($_.DistinguishedName -split ",",2)[1] }},Enabled,LastLogonDate,@{n="LAPSPassword";e={ if ($_."ms-Mcs-AdmPwd") { $_."ms-Mcs-AdmPwd" } else { "No pw set" } }}'
        ScriptBlock = {
            param($sb,$max,$cred)
            $props = 'Name','OperatingSystem','IPv4Address','ManagedBy','Description','Enabled','LastLogonDate','OperatingSystemVersion','DistinguishedName','ms-Mcs-AdmPwd'
            if ($cred) {
                $computers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -Properties $props -ResultSetSize $max -Credential $cred
            }
            else {
                $computers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -Properties $props -ResultSetSize $max
            }
            $out = foreach ($c in $computers) {
                $owner = $null
                if ($c.ManagedBy) {
                    try {
                        if ($cred) { $owner = (Get-ADUser -Identity $c.ManagedBy -Credential $cred -ErrorAction Stop).SamAccountName }
                        else       { $owner = (Get-ADUser -Identity $c.ManagedBy -ErrorAction Stop).SamAccountName }
                    }
                    catch { $owner = $c.ManagedBy }
                }
                $ou   = if ($c.DistinguishedName) { ($c.DistinguishedName -split ",", 2)[1] } else { $null }
                $laps = 'No pw set'
                try {
                    if ($c.'ms-Mcs-AdmPwd') { $laps = $c.'ms-Mcs-AdmPwd' }
                    else {
                        $tmp = if ($cred) {
                            Get-ADComputer -Identity $c.DistinguishedName -Properties 'ms-Mcs-AdmPwd' -Credential $cred -ErrorAction SilentlyContinue
                        }
                        else {
                            Get-ADComputer -Identity $c.DistinguishedName -Properties 'ms-Mcs-AdmPwd' -ErrorAction SilentlyContinue
                        }
                        if ($tmp -and $tmp.'ms-Mcs-AdmPwd') { $laps = $tmp.'ms-Mcs-AdmPwd' } else { $laps = 'No pw set' }
                    }
                }
                catch { $laps = 'No access / not set' }
                [PSCustomObject]@{
                    Name            = $c.Name
                    OperatingSystem = $c.OperatingSystem
                    IPv4Address     = $c.IPv4Address
                    Owner           = $owner
                    Description     = $c.Description
                    OU              = $ou
                    Enabled         = if ($null -ne $c.Enabled) { $c.Enabled } else { $true }
                    OSVersion       = $c.OperatingSystemVersion
                    LastLogonDate   = $c.LastLogonDate
                    LAPSPassword    = $laps
                }
            }
            return $out
        }
    },
    @{
        Label   = 'LAPS Password (if permitted) – Name, LAPSPassword'
        Preview = 'Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase "<SearchBase>" -Properties ms-Mcs-AdmPwd | Select Name,@{n="LAPSPassword";e={ if ($_."ms-Mcs-AdmPwd") { $_."ms-Mcs-AdmPwd" } else { "No pw set" } }}'
        ScriptBlock = {
            param($sb,$max,$cred)
            $computers = if ($cred) {
                Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -Properties 'ms-Mcs-AdmPwd' -ResultSetSize $max -Credential $cred
            }
            else {
                Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -Properties 'ms-Mcs-AdmPwd' -ResultSetSize $max
            }
            foreach ($c in $computers) {
                $laps = 'No pw set'
                try { if ($c.'ms-Mcs-AdmPwd') { $laps = $c.'ms-Mcs-AdmPwd' } }
                catch { $laps = 'No access / not set' }
                [PSCustomObject]@{ Name = $c.Name; LAPSPassword = $laps }
            }
        }
    },
    @{
        Label   = 'Reboot Pending and Last Boot Time (CIM)'
        Preview = 'Invoke-Command -ComputerName (Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase "<SearchBase>" -ResultSetSize 200 | Select -Expand Name) -ScriptBlock { Get-CimInstance Win32_OperatingSystem | Select CSName,LastBootUpTime } -EA SilentlyContinue'
        ScriptBlock = {
            param($sb,$max,$cred)
            $names = if ($cred) {
                Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -ResultSetSize $max -Credential $cred | Select-Object -ExpandProperty Name
            }
            else {
                Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -SearchBase $sb -ResultSetSize $max | Select-Object -ExpandProperty Name
            }
            foreach ($n in $names) {
                try {
                    $os = if ($cred) { Get-CimInstance -ComputerName $n -ClassName Win32_OperatingSystem -ErrorAction Stop }
                          else       { Get-CimInstance -ComputerName $n -ClassName Win32_OperatingSystem -ErrorAction Stop }
                    $rebootPending = $false
                    try {
                        $rebootPending = if ($cred) {
                            Invoke-Command -ComputerName $n -Credential $cred -ScriptBlock {
                                Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
                            } -ErrorAction SilentlyContinue
                        }
                        else {
                            Test-Path "\\$n\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
                        }
                    }
                    catch { }
                    [PSCustomObject]@{ Name = $n; LastBootUpTime = $os.LastBootUpTime; RebootPending = $rebootPending }
                }
                catch { [PSCustomObject]@{ Name = $n; LastBootUpTime = $null; RebootPending = $null } }
            }
        }
    }
)

# ── Domain Controllers ─────────────────────────────────────────────────────────
$Presets['DomainControllers'] = @(
    @{
        Label   = 'DC Replication Partners and Last Sync'
        Preview = 'Get-ADDomainController -Filter * | ForEach-Object { Get-ADReplicationPartnerMetadata -Target $_.Name -Scope Domain } | Select Server,Partner,LastReplicationSuccess'
        ScriptBlock = {
            param($sb,$max,$cred)
            $dcs = if ($cred) { Get-ADDomainController -Filter * -Credential $cred | Select-Object -ExpandProperty Name }
                   else       { Get-ADDomainController -Filter * | Select-Object -ExpandProperty Name }
            foreach ($dc in $dcs) {
                try {
                    $meta = if ($cred) { Get-ADReplicationPartnerMetadata -Target $dc -Scope Domain -Credential $cred -ErrorAction Stop }
                            else       { Get-ADReplicationPartnerMetadata -Target $dc -Scope Domain -ErrorAction Stop }
                    foreach ($m in $meta) {
                        [PSCustomObject]@{
                            DC                     = $dc
                            Partner                = $m.Partner
                            LastReplicationSuccess = $m.LastReplicationSuccess
                            LastAttempt            = $m.LastAttempt
                            ConsecutiveFailures    = $m.ConsecutiveFailureCount
                        }
                    }
                }
                catch { [PSCustomObject]@{ DC = $dc; Partner = $null; LastReplicationSuccess = $null; LastAttempt = $null; ConsecutiveFailures = $null } }
            }
        }
    },
    @{
        Label   = 'Replication Failures Summary'
        Preview = 'Get-ADReplicationFailure -Scope Site | Select Server,FirstFailureTime,FailureCount,Partner'
        ScriptBlock = {
            param($sb,$max,$cred)
            if ($cred) { Get-ADReplicationFailure -Scope Site -Credential $cred -ErrorAction SilentlyContinue | Select-Object Server,FirstFailureTime,FailureCount,Partner }
            else       { Get-ADReplicationFailure -Scope Site -ErrorAction SilentlyContinue | Select-Object Server,FirstFailureTime,FailureCount,Partner }
        }
    },
    @{
        Label   = 'DC Health Quick Check'
        Preview = 'Get-ADDomainController -Filter * | Select Name,IsGlobalCatalog,OperatingSystem,IPv4Address,Site'
        ScriptBlock = {
            param($sb,$max,$cred)
            if ($cred) { Get-ADDomainController -Filter * -Credential $cred | Select-Object Name,IsGlobalCatalog,OperatingSystem,IPv4Address,Site }
            else       { Get-ADDomainController -Filter * | Select-Object Name,IsGlobalCatalog,OperatingSystem,IPv4Address,Site }
        }
    }
)

# ── Users ─────────────────────────────────────────────────────────────────────
$Presets['Users'] = @(
    @{
        Label   = 'Basic: Name, UPN, Email, Enabled, LastLogon, PwdLastSet'
        Preview = 'Get-ADUser -Filter * -SearchBase "<SearchBase>" -Properties DisplayName,UserPrincipalName,Mail,Enabled,LastLogonDate,PasswordLastSet | Select DisplayName,UserPrincipalName,Mail,Enabled,LastLogonDate,PasswordLastSet'
        ScriptBlock = {
            param($sb,$max,$cred)
            $props = 'DisplayName','UserPrincipalName','Mail','Enabled','LastLogonDate','PasswordLastSet'
            if ($cred) { Get-ADUser -Filter * -SearchBase $sb -Properties $props -ResultSetSize $max -Credential $cred | Select-Object $props }
            else       { Get-ADUser -Filter * -SearchBase $sb -Properties $props -ResultSetSize $max | Select-Object $props }
        }
    },
    @{
        Label   = 'Privileged Accounts (Domain Admins, Enterprise Admins)'
        Preview = 'Get-ADGroupMember "Domain Admins" -Recursive | Select Name,SamAccountName,DistinguishedName'
        ScriptBlock = {
            param($sb,$max,$cred)
            $groups = @('Domain Admins','Enterprise Admins')
            foreach ($g in $groups) {
                try {
                    if ($cred) { Get-ADGroupMember -Identity $g -Recursive -Credential $cred -ErrorAction Stop | Select-Object @{n='Group';e={$g}},Name,SamAccountName,DistinguishedName }
                    else       { Get-ADGroupMember -Identity $g -Recursive -ErrorAction Stop | Select-Object @{n='Group';e={$g}},Name,SamAccountName,DistinguishedName }
                }
                catch { [PSCustomObject]@{ Group = $g; Name = $null; SamAccountName = $null; DistinguishedName = $null } }
            }
        }
    },
    @{
        Label   = 'Service-like Users (SPN or machine$ account)'
        Preview = 'Get-ADUser -Filter {ServicePrincipalName -like "*" -or SamAccountName -like "*$"} -SearchBase "<SearchBase>" -Properties SamAccountName,ServicePrincipalName,Enabled,Description | Select SamAccountName,ServicePrincipalName,Enabled,Description'
        ScriptBlock = {
            param($sb,$max,$cred)
            if ($cred) { Get-ADUser -Filter {ServicePrincipalName -like "*" -or SamAccountName -like "*$"} -SearchBase $sb -Properties SamAccountName,ServicePrincipalName,Enabled,Description -ResultSetSize $max -Credential $cred | Select-Object SamAccountName,ServicePrincipalName,Enabled,Description }
            else       { Get-ADUser -Filter {ServicePrincipalName -like "*" -or SamAccountName -like "*$"} -SearchBase $sb -Properties SamAccountName,ServicePrincipalName,Enabled,Description -ResultSetSize $max | Select-Object SamAccountName,ServicePrincipalName,Enabled,Description }
        }
    }
)

# ── Groups ────────────────────────────────────────────────────────────────────
$Presets['Groups'] = @(
    @{
        Label   = 'Groups: Name, Description, MembersCount, OU'
        Preview = 'Get-ADGroup -Filter * -SearchBase "<SearchBase>" -Properties Description,DistinguishedName | Select Name,Description,@{n="MembersCount";e={(Get-ADGroupMember $_ -EA SilentlyContinue).Count}},@{n="OU";e={($_.DistinguishedName -split ",",2)[1]}}'
        ScriptBlock = {
            param($sb,$max,$cred)
            $groups = if ($cred) { Get-ADGroup -Filter * -SearchBase $sb -Properties Description,DistinguishedName -ResultSetSize $max -Credential $cred }
                      else       { Get-ADGroup -Filter * -SearchBase $sb -Properties Description,DistinguishedName -ResultSetSize $max }
            foreach ($g in $groups) {
                $count = 0
                try {
                    if ($cred) { $count = (Get-ADGroupMember -Identity $g -Credential $cred -ErrorAction Stop).Count }
                    else       { $count = (Get-ADGroupMember -Identity $g -ErrorAction Stop).Count }
                }
                catch { $count = 0 }
                [PSCustomObject]@{
                    Name         = $g.Name
                    Description  = $g.Description
                    MembersCount = $count
                    OU           = if ($g.DistinguishedName) { ($g.DistinguishedName -split ",",2)[1] } else { $null }
                }
            }
        }
    }
)

# ── Security Groups ────────────────────────────────────────────────────────────
$Presets['SecurityGroups'] = @(
    @{
        Label   = 'Security Groups: Name, Scope, MembersCount, OU'
        Preview = 'Get-ADGroup -Filter {GroupCategory -eq "Security"} -SearchBase "<SearchBase>" -Properties GroupScope,Description,DistinguishedName | Select Name,GroupScope,Description,MembersCount,OU'
        ScriptBlock = {
            param($sb,$max,$cred)
            $groups = if ($cred) { Get-ADGroup -Filter {GroupCategory -eq "Security"} -SearchBase $sb -Properties GroupScope,Description,DistinguishedName -ResultSetSize $max -Credential $cred }
                      else       { Get-ADGroup -Filter {GroupCategory -eq "Security"} -SearchBase $sb -Properties GroupScope,Description,DistinguishedName -ResultSetSize $max }
            foreach ($g in $groups) {
                $count = 0
                try {
                    if ($cred) { $count = (Get-ADGroupMember -Identity $g -Credential $cred -ErrorAction Stop).Count }
                    else       { $count = (Get-ADGroupMember -Identity $g -ErrorAction Stop).Count }
                }
                catch { $count = 0 }
                [PSCustomObject]@{
                    Name         = $g.Name
                    GroupScope   = $g.GroupScope
                    Description  = $g.Description
                    MembersCount = $count
                    OU           = if ($g.DistinguishedName) { ($g.DistinguishedName -split ",",2)[1] } else { $null }
                }
            }
        }
    }
)

# ── Service Accounts ──────────────────────────────────────────────────────────
$Presets['ServiceAccounts'] = @(
    @{
        Label   = 'Accounts with ServicePrincipalName (SPN)'
        Preview = 'Get-ADUser -Filter {ServicePrincipalName -like "*"} -SearchBase "<SearchBase>" -Properties SamAccountName,ServicePrincipalName,Description,Enabled | Select SamAccountName,ServicePrincipalName,Enabled,Description'
        ScriptBlock = {
            param($sb,$max,$cred)
            if ($cred) { Get-ADUser -Filter {ServicePrincipalName -like "*"} -SearchBase $sb -Properties SamAccountName,ServicePrincipalName,Description,Enabled -ResultSetSize $max -Credential $cred | Select-Object SamAccountName,ServicePrincipalName,Enabled,Description }
            else       { Get-ADUser -Filter {ServicePrincipalName -like "*"} -SearchBase $sb -Properties SamAccountName,ServicePrincipalName,Description,Enabled -ResultSetSize $max | Select-Object SamAccountName,ServicePrincipalName,Enabled,Description }
        }
    }
)

# ── Managed Service Accounts ──────────────────────────────────────────────────
$Presets['ManagedServiceAccounts'] = @(
    @{
        Label   = 'Managed Service Accounts (Get-ADServiceAccount)'
        Preview = 'Get-ADServiceAccount -Filter * -SearchBase "<SearchBase>" | Select Name,SamAccountName,PrincipalsAllowedToRetrieveManagedPassword,DistinguishedName'
        ScriptBlock = {
            param($sb,$max,$cred)
            try {
                if ($cred) { Get-ADServiceAccount -Filter * -SearchBase $sb -ResultSetSize $max -Credential $cred | Select-Object Name,SamAccountName,PrincipalsAllowedToRetrieveManagedPassword,DistinguishedName }
                else       { Get-ADServiceAccount -Filter * -SearchBase $sb -ResultSetSize $max | Select-Object Name,SamAccountName,PrincipalsAllowedToRetrieveManagedPassword,DistinguishedName }
            }
            catch { @() }
        }
    }
)

# ══════════════════════════════════════════════════════════════════════════════
#  PRESET / UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
function Populate-Presets {
    param($type)
    $cbPreset.Items.Clear()
    if (-not $Presets.ContainsKey($type)) { return }
    foreach ($p in $Presets[$type]) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $p.Label
        $cbPreset.Items.Add($item) | Out-Null
    }
    if ($cbPreset.Items.Count -gt 0) { $cbPreset.SelectedIndex = 0 }
}

function Apply-LAPSMasking {
    param($rows)
    if (-not $rows) { return }
    foreach ($r in $rows) {
        if ($r.PSObject.Properties.Match('LAPSPassword')) {
            $key = ($r.Name) -as [string]
            if (-not $global:LAPSCache.ContainsKey($key)) { $global:LAPSCache[$key] = $r.LAPSPassword }
            $r.LAPSPassword = if ($global:MaskLAPS) { '*****' } else { $global:LAPSCache[$key] }
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  UI EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

# Object type changes → reload presets
$cbObject.Add_SelectionChanged({
    $type = $cbObject.SelectedItem.Content
    Populate-Presets -type $type
    if ($Presets.ContainsKey($type) -and $Presets[$type].Count -gt 0) {
        $tbCommandPreview.Text = ($Presets[$type][0].Preview -replace '<SearchBase>', $tbSearchBase.Text)
    }
    else { $tbCommandPreview.Text = '' }
})

# Preset changes → update preview
$cbPreset.Add_SelectionChanged({
    $type = $cbObject.SelectedItem.Content
    $idx  = $cbPreset.SelectedIndex
    if ($idx -ge 0) {
        $tbCommandPreview.Text = $Presets[$type][$idx].Preview -replace '<SearchBase>', $tbSearchBase.Text
    }
})

$btnRefreshPresets.Add_Click({ Populate-Presets -type $cbObject.SelectedItem.Content })

$btnEditPreview.Add_Click({
    $tbCommandPreview.IsReadOnly = -not $tbCommandPreview.IsReadOnly
    $btnEditPreview.Content = if ($tbCommandPreview.IsReadOnly) { '✏  Edit Preview' } else { '🔒  Lock Preview' }
})

# ── Run ────────────────────────────────────────────────────────────────────────
$btnRun.Add_Click({
    $type  = $cbObject.SelectedItem.Content
    $idx   = $cbPreset.SelectedIndex
    $sb    = if ($tbSearchBase.Text) { $tbSearchBase.Text } else { $null }
    $max   = [int]($tbMax.Text -as [int]); if (-not $max) { $max = 1000 }

    $selectedCred = $null
    if ($cbCredentials.SelectedItem) {
        $label = $cbCredentials.SelectedItem.Content
        $selectedCred = if ($CredentialStore.ContainsKey($label)) { $CredentialStore[$label] }
                        else { Get-CredFromVault -label $label }
    }

    try {
        Set-StatusBusy "Querying AD…"
        $btnRun.IsEnabled = $false
        $tbResultCount.Text = "…"

        if ($idx -ge 0 -and $Presets[$type][$idx].ScriptBlock) {
            $script  = $Presets[$type][$idx].ScriptBlock
            $results = & $script $sb $max $selectedCred
        }
        else {
            $cmd = $tbCommandPreview.Text -replace '<SearchBase>', $sb
            Set-Variable -Name selectedCred -Value $selectedCred -Scope Script
            $results = Invoke-Expression $cmd
        }

        $global:results = $results
        Apply-LAPSMasking -rows $global:results
        $dgResults.ItemsSource = $global:results

        $count = if ($global:results) { @($global:results).Count } else { 0 }
        $tbResultCount.Text = "$count row$(if($count -ne 1){'s'})"
        Set-StatusReady
    }
    catch {
        Set-StatusError "Query failed"
        [System.Windows.MessageBox]::Show(
            "Error running query:`n$($_.Exception.Message)",
            "AD Reporter – Query Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
    }
    finally {
        $btnRun.IsEnabled = $true
    }
})

# ── Mask Toggle ────────────────────────────────────────────────────────────────
$btnMaskToggle.Add_Click({
    $global:MaskLAPS = -not $global:MaskLAPS
    $btnMaskToggle.Content = if ($global:MaskLAPS) { '🔒  Mask LAPS (On)' } else { '🔓  Mask LAPS (Off)' }
    if ($global:results) { Apply-LAPSMasking -rows $global:results; $dgResults.Items.Refresh() }
})

# ── Reveal Selected ────────────────────────────────────────────────────────────
$btnRevealSelected.Add_Click({
    $sel = $dgResults.SelectedItems
    if (-not $sel -or $sel.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Select one or more rows to reveal LAPS.", "AD Reporter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    $selectedCred = $null
    if ($cbCredentials.SelectedItem) {
        $label = $cbCredentials.SelectedItem.Content
        $selectedCred = if ($CredentialStore.ContainsKey($label)) { $CredentialStore[$label] }
                        else { Get-CredFromVault -label $label }
    }
    foreach ($item in $sel) {
        if (-not $item.PSObject.Properties.Match('LAPSPassword')) { continue }
        try {
            $tmp = if ($selectedCred) { Get-ADComputer -Identity $item.Name -Properties 'ms-Mcs-AdmPwd' -Credential $selectedCred -ErrorAction Stop }
                   else               { Get-ADComputer -Identity $item.Name -Properties 'ms-Mcs-AdmPwd' -ErrorAction Stop }
            $val = if ($tmp -and $tmp.'ms-Mcs-AdmPwd') { $tmp.'ms-Mcs-AdmPwd' } else { 'No pw set' }
        }
        catch { $val = 'No access / not set' }
        $global:LAPSCache[$item.Name] = $val
        $item.LAPSPassword = $val
        $operator  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $credLabel = if ($cbCredentials.SelectedItem) { $cbCredentials.SelectedItem.Content } else { 'Local' }
        Write-LAPSAudit -operator $operator -credLabel $credLabel -target $item.Name -action 'Reveal' -result ($val -replace ',',';')
    }
    $dgResults.Items.Refresh()
})

# ── Refresh Grid ───────────────────────────────────────────────────────────────
$btnRefreshGrid.Add_Click({
    if ($global:results) { Apply-LAPSMasking -rows $global:results; $dgResults.Items.Refresh() }
})

# ── Sign In ────────────────────────────────────────────────────────────────────
$btnSignIn.Add_Click({
    $tbSignInStatus.Text       = "Testing credential…"
    $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::Gray
    $domain = if ($cbLoginDomain.Text) { $cbLoginDomain.Text.Trim() } else { $null }
    $cred   = Get-LoginCredential
    if (-not $cred) { $tbSignInStatus.Text = "Enter username and password."; return }
    try {
        if ($domain) {
            $dc = Get-ADDomainController -Filter * -Server $domain -Credential $cred -ErrorAction Stop | Select-Object -First 1
            if ($dc) {
                $tbSignInStatus.Text       = "Sign-in OK  (DC: $($dc.HostName))"
                $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
            }
            else {
                $tbSignInStatus.Text       = "Authenticated but no DC found for '$domain'."
                $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::Orange
            }
        }
        else {
            Get-ADUser -Identity $cred.UserName -Credential $cred -ErrorAction Stop | Out-Null
            $tbSignInStatus.Text       = "Sign-in successful."
            $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
        }

        if ($chkSaveCred.IsChecked) {
            $label = if ($domain) { "ADReporter::$domain - $($cred.UserName)" } else { "ADReporter::$($cred.UserName)" }
            $CredentialStore[$label] = $cred
            Refresh-CredCombo
            Populate-LoginDomainCombo -default $env:USERDOMAIN
            $tbSignInStatus.Text += "  Saved to session."
        }
        if ($chkPersistVault.IsChecked) {
            $label = if ($domain) { "ADReporter::$domain - $($cred.UserName)" } else { "ADReporter::$($cred.UserName)" }
            $ok = Save-CredToVault -label $label -cred $cred
            $tbSignInStatus.Text += if ($ok) { "  Persisted to vault." } else { "  Vault persist failed." }
        }
    }
    catch {
        $tbSignInStatus.Text       = "Sign-in failed: $($_.Exception.Message)"
        $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::Tomato
    }
})

# ── Use Saved Credential ───────────────────────────────────────────────────────
$btnUseSaved.Add_Click({
    $sel = $cbCredentials.SelectedItem
    if (-not $sel) {
        [System.Windows.MessageBox]::Show("Select a saved credential first.", "AD Reporter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    $label = $sel.Content
    if (-not $CredentialStore.ContainsKey($label) -and $label -like 'ADReporter::*') {
        $cred = Get-CredFromVault -label $label
        if ($cred) { $CredentialStore[$label] = $cred }
    }
    if ($CredentialStore.ContainsKey($label)) {
        $cred = $CredentialStore[$label]
        if ($label -match '^ADReporter::(?<dom>[^ ]+)\s*-\s*(?<user>.+)$') { $cbLoginDomain.Text = $Matches['dom'] }
        $tbLoginUser.Text  = $cred.UserName
        $pbLoginPass.Password = ''
        $tbSignInStatus.Text       = "Loaded: '$label'"
        $tbSignInStatus.Foreground = [System.Windows.Media.Brushes]::Gray
    }
    else {
        [System.Windows.MessageBox]::Show("Could not load credential from vault.", "AD Reporter – Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT
# ══════════════════════════════════════════════════════════════════════════════
function Build-HtmlReport {
    param($data, $title, $path)
    $data | ConvertTo-Html -Title $title -PreContent "<h2>$title</h2>" | Out-File -FilePath $path -Encoding UTF8
}

function Export-Docx-PSWriteWord {
    param($data, $path)
    try {
        Import-Module PSWriteWord -ErrorAction Stop
        $doc  = New-WordDocument -Path $path -Force
        $html = $data | ConvertTo-Html -Fragment
        Add-WordText -WordDocument $doc -Text $html -IsHtml
        Close-WordDocument -WordDocument $doc -Save
    }
    catch { throw "PSWriteWord export failed: $($_.Exception.Message)" }
}

$btnExport.Add_Click({
    if (-not $global:results) {
        [System.Windows.MessageBox]::Show("No results to export. Run a query first.", "AD Reporter", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    $ts   = Get-Date -Format yyyyMMdd_HHmm
    $base = "$env:USERPROFILE\Desktop\ADReport_$ts"
    try {
        if ($cbCsv.IsChecked)  { $global:results | Export-Csv "$base.csv" -NoTypeInformation -Force }
        if ($cbHtml.IsChecked) { Build-HtmlReport -data $global:results -title "AD Report $ts" -path "$base.html" }
        if ($cbXlsx.IsChecked) { Import-Module ImportExcel -ErrorAction Stop; $global:results | Export-Excel -Path "$base.xlsx" -AutoSize -Force }
        if ($cbDocx.IsChecked) { Export-Docx-PSWriteWord -data $global:results -path "$base.docx" }
        if ($cbPdf.IsChecked)  {
            $htmlPath = "$base`_pdf_source.html"
            Build-HtmlReport -data $global:results -title "AD Report $ts" -path $htmlPath
            # Requires wkhtmltopdf in PATH:  & wkhtmltopdf $htmlPath "$base.pdf"
        }
        [System.Windows.MessageBox]::Show(
            "Export complete.`nFiles saved to Desktop:`n$base.*",
            "AD Reporter – Export",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Export error:`n$($_.Exception.Message)",
            "AD Reporter – Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  INITIALISE AND SHOW
# ══════════════════════════════════════════════════════════════════════════════
Populate-Presets        -type 'Servers'
Populate-LoginDomainCombo -default $env:USERDOMAIN
Refresh-CredCombo

# Seed command preview for default selection
$tbCommandPreview.Text = $Presets['Servers'][0].Preview -replace '<SearchBase>', ''

$window.ShowDialog() | Out-Null
