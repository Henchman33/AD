<#
.SYNOPSIS
    Enterprise Active Directory Server Inventory & Health Report

.DESCRIPTION
    Recursively scans ALL domains in the Active Directory Forest
    and inventories all Windows Server systems.

    Features:
    - Multi-domain forest scanning
    - Recursive OU discovery
    - Ping / online status
    - OS version and build number
    - Last Logon Timestamp
    - Domain Controller identification
    - Stale server detection
    - CSV Export
    - Excel Export (.xlsx)
    - Executive HTML Dashboard
    - PowerShell GUI Interface

.OUTPUTS
    Desktop\All_Servers_Report\

.REQUIREMENTS
    - RSAT ActiveDirectory Module
    - ImportExcel PowerShell Module
    - PowerShell 5.1+
    - Recommended: Run as Administrator

.AUTHOR
    ChatGPT Enterprise AD Inventory Edition
#>

# ============================================
# LOAD ASSEMBLIES
# ============================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================
# GUI FORM
# ============================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Enterprise AD Server Inventory Tool"
$form.Size = New-Object System.Drawing.Size(600,420)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

# ============================================
# TITLE
# ============================================

$Title = New-Object System.Windows.Forms.Label
$Title.Location = New-Object System.Drawing.Point(20,20)
$Title.Size = New-Object System.Drawing.Size(520,30)
$Title.Text = "Enterprise Active Directory Server Inventory"
$Title.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($Title)

# ============================================
# CHECKBOXES
# ============================================

$options = @(
    "Ping Servers / Online Detection",
    "Collect OS Version & Build",
    "Collect Last Logon Timestamp",
    "Detect Domain Controllers",
    "Detect Stale Servers",
    "Export Excel Report",
    "Export HTML Executive Report",
    "Multi-Domain Forest Scan"
)

$CheckBoxes = @()
$y = 70

foreach ($option in $options)
{
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Location = New-Object System.Drawing.Point(30,$y)
    $cb.Size = New-Object System.Drawing.Size(300,25)
    $cb.Text = $option
    $cb.Checked = $true

    $form.Controls.Add($cb)
    $CheckBoxes += $cb

    $y += 30
}

# ============================================
# RUN BUTTON
# ============================================

$RunButton = New-Object System.Windows.Forms.Button
$RunButton.Location = New-Object System.Drawing.Point(30,330)
$RunButton.Size = New-Object System.Drawing.Size(150,40)
$RunButton.Text = "Run Inventory"
$RunButton.BackColor = "#0078D7"
$RunButton.ForeColor = "White"
$form.Controls.Add($RunButton)

# ============================================
# STATUS BOX
# ============================================

$StatusBox = New-Object System.Windows.Forms.TextBox
$StatusBox.Location = New-Object System.Drawing.Point(350,70)
$StatusBox.Size = New-Object System.Drawing.Size(210,250)
$StatusBox.Multiline = $true
$StatusBox.ScrollBars = "Vertical"
$form.Controls.Add($StatusBox)

# ============================================
# RUN LOGIC
# ============================================

$RunButton.Add_Click({

    $StatusBox.AppendText("Starting inventory scan...`r`n")

    Try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    Catch {
        [System.Windows.Forms.MessageBox]::Show("ActiveDirectory module missing.")
        return
    }

    # ============================================
    # INSTALL IMPORTEXCEL IF MISSING
    # ============================================

    if ($CheckBoxes[5].Checked)
    {
        if (!(Get-Module -ListAvailable -Name ImportExcel))
        {
            $StatusBox.AppendText("Installing ImportExcel module...`r`n")

            Try {
                Install-Module ImportExcel -Force -Scope CurrentUser -AllowClobber
            }
            Catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to install ImportExcel module.")
                return
            }
        }

        Import-Module ImportExcel
    }

    # ============================================
    # REPORT DIRECTORY
    # ============================================

    $Desktop = [Environment]::GetFolderPath("Desktop")
    $ReportFolder = Join-Path $Desktop "All_Servers_Report"

    if (!(Test-Path $ReportFolder))
    {
        New-Item -ItemType Directory -Path $ReportFolder | Out-Null
    }

    $TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"

    $CSVReport  = Join-Path $ReportFolder "Enterprise_Server_Report_$TimeStamp.csv"
    $HTMLReport = Join-Path $ReportFolder "Enterprise_Server_Report_$TimeStamp.html"
    $ExcelReport = Join-Path $ReportFolder "Enterprise_Server_Report_$TimeStamp.xlsx"

    # ============================================
    # GET FOREST DOMAINS
    # ============================================

    $Forest = Get-ADForest
    $Domains = $Forest.Domains

    $AllResults = @()

    foreach ($Domain in $Domains)
    {
        $StatusBox.AppendText("Scanning Domain: $Domain`r`n")

        Try {

            $Servers = Get-ADComputer `
                -Server $Domain `
                -Filter {
                    OperatingSystem -like "*Server*"
                } `
                -Properties *
        }
        Catch {
            $StatusBox.AppendText("Failed to scan domain: $Domain`r`n")
            continue
        }

        foreach ($Server in $Servers)
        {
            $StatusBox.AppendText("Processing: $($Server.Name)`r`n")

            # ============================================
            # PING STATUS
            # ============================================

            $Online = "Unknown"

            if ($CheckBoxes[0].Checked)
            {
                Try {
                    if (Test-Connection $Server.Name -Count 1 -Quiet)
                    {
                        $Online = "Online"
                    }
                    else
                    {
                        $Online = "Offline"
                    }
                }
                Catch {
                    $Online = "Offline"
                }
            }

            # ============================================
            # LAST LOGON
            # ============================================

            $LastLogon = $null

            if ($CheckBoxes[2].Checked)
            {
                if ($Server.LastLogonDate)
                {
                    $LastLogon = $Server.LastLogonDate
                }
            }

            # ============================================
            # STALE DETECTION
            # ============================================

            $Stale = "No"

            if ($CheckBoxes[4].Checked)
            {
                if ($LastLogon)
                {
                    $DaysOld = (New-TimeSpan -Start $LastLogon -End (Get-Date)).Days

                    if ($DaysOld -gt 90)
                    {
                        $Stale = "YES - Over 90 Days"
                    }
                }
                else
                {
                    $Stale = "Unknown"
                }
            }

            # ============================================
            # DOMAIN CONTROLLER DETECTION
            # ============================================

            $IsDC = "No"

            if ($CheckBoxes[3].Checked)
            {
                if ($Server.PrimaryGroupID -eq 516)
                {
                    $IsDC = "YES"
                }
            }

            # ============================================
            # OU EXTRACTION
            # ============================================

            $OU = ($Server.DistinguishedName -replace '^CN=.*?,','')

            # ============================================
            # BUILD OBJECT
            # ============================================

            $Object = [PSCustomObject]@{

                ServerName         = $Server.Name
                IPv4Address        = $Server.IPv4Address
                OnlineStatus       = $Online
                DomainFQDN         = $Domain
                OperatingSystem    = $Server.OperatingSystem
                OSVersion          = $Server.OperatingSystemVersion
                BuildNumber        = $Server.OperatingSystemHotfix
                DomainController   = $IsDC
                LastLogonDate      = $LastLogon
                StaleServer        = $Stale
                OrganizationalUnit = $OU
                DNSHostName        = $Server.DNSHostName
                Enabled            = $Server.Enabled
                Created            = $Server.Created
            }

            $AllResults += $Object
        }
    }

    # ============================================
    # EXPORT CSV
    # ============================================

    $AllResults |
        Sort-Object ServerName |
        Export-Csv $CSVReport -NoTypeInformation -Encoding UTF8

    $StatusBox.AppendText("CSV Export Complete`r`n")

    # ============================================
    # EXPORT EXCEL
    # ============================================

    if ($CheckBoxes[5].Checked)
    {
        $AllResults |
            Export-Excel `
                -Path $ExcelReport `
                -WorksheetName "Servers" `
                -AutoSize `
                -BoldTopRow `
                -FreezeTopRow `
                -TableName "EnterpriseServers"

        $StatusBox.AppendText("Excel Export Complete`r`n")
    }

    # ============================================
    # HTML REPORT
    # ============================================

    if ($CheckBoxes[6].Checked)
    {
        $ServerCount = $AllResults.Count
        $OnlineCount = ($AllResults | Where-Object {$_.OnlineStatus -eq "Online"}).Count
        $OfflineCount = ($AllResults | Where-Object {$_.OnlineStatus -eq "Offline"}).Count
        $DCCount = ($AllResults | Where-Object {$_.DomainController -eq "YES"}).Count
        $StaleCount = ($AllResults | Where-Object {$_.StaleServer -like "YES*"}).Count

        $HTMLHeader = @"
<html>
<head>
<title>Enterprise Server Inventory</title>

<style>

body {
    font-family: Segoe UI;
    background-color: #f5f5f5;
    margin: 20px;
}

h1 {
    color: #003366;
}

.summary {
    background-color: white;
    padding: 20px;
    margin-bottom: 20px;
    border: 1px solid #dcdcdc;
}

table {
    border-collapse: collapse;
    width: 100%;
    background-color: white;
}

th {
    background-color: #003366;
    color: white;
    padding: 10px;
    border: 1px solid #cccccc;
}

td {
    padding: 8px;
    border: 1px solid #cccccc;
}

tr:nth-child(even) {
    background-color: #f2f2f2;
}

</style>
</head>
<body>

<h1>Enterprise Active Directory Server Inventory</h1>

<div class='summary'>

<h2>Executive Summary</h2>

<p><strong>Total Servers:</strong> $ServerCount</p>
<p><strong>Online Servers:</strong> $OnlineCount</p>
<p><strong>Offline Servers:</strong> $OfflineCount</p>
<p><strong>Domain Controllers:</strong> $DCCount</p>
<p><strong>Stale Servers:</strong> $StaleCount</p>
<p><strong>Generated:</strong> $(Get-Date)</p>

</div>

"@

        $HTMLFooter = @"
</body>
</html>
"@

        $AllResults |
            ConvertTo-Html `
                -Property ServerName,
                          IPv4Address,
                          OnlineStatus,
                          DomainFQDN,
                          OperatingSystem,
                          OSVersion,
                          DomainController,
                          LastLogonDate,
                          StaleServer,
                          OrganizationalUnit `
                -Head $HTMLHeader `
                -PostContent $HTMLFooter |
            Out-File $HTMLReport -Encoding UTF8

        $StatusBox.AppendText("HTML Report Complete`r`n")
    }

    # ============================================
    # COMPLETE
    # ============================================

    $StatusBox.AppendText("Inventory Complete.`r`n")
    $StatusBox.AppendText("Opening report folder...`r`n")

    Invoke-Item $ReportFolder

})

# ============================================
# SHOW FORM
# ============================================

[void]$form.ShowDialog()
