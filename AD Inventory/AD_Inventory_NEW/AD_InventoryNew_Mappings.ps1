#Requires -Version 5.1
<#
.SYNOPSIS
    Enterprise Active Directory Inventory & Mapping Tool

.DESCRIPTION
    Performs a deep Active Directory inventory across the forest:
    - Forest / Domain discovery
    - OU and container analysis
    - Users
    - Computers
    - Domain Controllers
    - Security & Distribution Groups
    - Group Memberships
    - GPOs
    - Sites & Subnets
    - Trusts
    - Tiered administration container detection
    - HTML, CSV, Excel exports
    - Visio-compatible mapping export (DGML + Graphviz DOT)

.NOTES
    Recommended:
    - Run as Domain Admin / Enterprise Admin
    - Run from management workstation with RSAT installed
    - PowerShell 5.1+
#>

Import-Module ActiveDirectory -ErrorAction Stop

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------

$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Desktop = [Environment]::GetFolderPath("Desktop")

$RootFolder = Join-Path $Desktop "AD_Inventory_$TimeStamp"

$HtmlFolder  = Join-Path $RootFolder "HTML"
$CsvFolder   = Join-Path $RootFolder "CSV"
$ExcelFolder = Join-Path $RootFolder "EXCEL"
$MapFolder   = Join-Path $RootFolder "MAPS"
$LogFolder   = Join-Path $RootFolder "LOGS"

$Folders = @(
    $RootFolder,
    $HtmlFolder,
    $CsvFolder,
    $ExcelFolder,
    $MapFolder,
    $LogFolder
)

foreach ($Folder in $Folders) {
    if (!(Test-Path $Folder)) {
        New-Item -Path $Folder -ItemType Directory -Force | Out-Null
    }
}

$LogFile = Join-Path $LogFolder "AD_Inventory.log"

# ------------------------------------------------------------
# LOGGING
# ------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $Line = "[$Time] [$Level] $Message"

    Write-Host $Line

    Add-Content -Path $LogFile -Value $Line
}

# ------------------------------------------------------------
# INSTALL IMPORTEXCEL IF MISSING
# ------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {

    Write-Log "ImportExcel module not found. Attempting install..." "WARN"

    try {
        Install-Module ImportExcel -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
        Write-Log "ImportExcel module installed successfully." "SUCCESS"
    }
    catch {
        Write-Log "Failed to install ImportExcel module. Excel export may fail." "ERROR"
    }
}

Import-Module ImportExcel -ErrorAction SilentlyContinue

# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------

function Export-DataSet {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [object]$Data
    )

    try {

        $CsvPath = Join-Path $CsvFolder "$Name.csv"
        $HtmlPath = Join-Path $HtmlFolder "$Name.html"

        $Data | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        $Data | ConvertTo-Html -Title $Name |
            Out-File -FilePath $HtmlPath -Encoding UTF8

        if (Get-Module ImportExcel) {

            $ExcelPath = Join-Path $ExcelFolder "AD_Inventory.xlsx"

            $Data | Export-Excel `
                -Path $ExcelPath `
                -WorksheetName $Name `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow `
                -AutoFilter
        }

        Write-Log "Exported dataset: $Name" "SUCCESS"
    }
    catch {
        Write-Log "Failed exporting dataset $Name : $_" "ERROR"
    }
}

function Get-ContainerPurpose {

    param(
        [string]$DistinguishedName,
        [string]$Name
    )

    $Combined = "$DistinguishedName $Name"

    switch -Regex ($Combined) {

        "Domain Controllers" { return "Domain Controllers" }

        "Tier.?0|Privileged|Admins" { return "Tier 0 Administrative" }

        "Tier.?1|Server" { return "Tier 1 Servers" }

        "Tier.?2|Workstation|Desktop" { return "Tier 2 Workstations" }

        "User" { return "User Container" }

        "Group" { return "Security Groups" }

        "Service" { return "Service Accounts" }

        "OU=Servers" { return "Server OU" }

        default { return "General Container" }
    }
}

# ------------------------------------------------------------
# FOREST INFORMATION
# ------------------------------------------------------------

Write-Log "Collecting forest information..."

$Forest = Get-ADForest
$Domains = $Forest.Domains

$ForestInfo = [PSCustomObject]@{
    ForestName             = $Forest.Name
    RootDomain             = $Forest.RootDomain
    ForestMode             = $Forest.ForestMode
    Domains                = ($Forest.Domains -join ", ")
    Sites                  = ($Forest.Sites -join ", ")
    GlobalCatalogs         = ($Forest.GlobalCatalogs -join ", ")
    ApplicationPartitions  = ($Forest.ApplicationPartitions -join ", ")
}

Export-DataSet -Name "ForestInformation" -Data $ForestInfo

# ------------------------------------------------------------
# DOMAIN INFORMATION
# ------------------------------------------------------------

$DomainInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Processing domain: $DomainName"

    try {

        $Domain = Get-ADDomain -Identity $DomainName

        $DomainInventory += [PSCustomObject]@{
            DomainName             = $Domain.DNSRoot
            NetBIOSName            = $Domain.NetBIOSName
            DomainMode             = $Domain.DomainMode
            ParentDomain           = $Domain.ParentDomain
            PDCEmulator            = $Domain.PDCEmulator
            RIDMaster              = $Domain.RIDMaster
            InfrastructureMaster   = $Domain.InfrastructureMaster
        }
    }
    catch {
        Write-Log "Failed domain inventory for $DomainName : $_" "ERROR"
    }
}

Export-DataSet -Name "Domains" -Data $DomainInventory

# ------------------------------------------------------------
# OU INVENTORY
# ------------------------------------------------------------

$OUInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting OUs from $DomainName"

    try {

        $OUs = Get-ADOrganizationalUnit `
            -Filter * `
            -Server $DomainName `
            -Properties *

        foreach ($OU in $OUs) {

            $OUInventory += [PSCustomObject]@{
                Name                    = $OU.Name
                DistinguishedName       = $OU.DistinguishedName
                Description             = $OU.Description
                ManagedBy               = $OU.ManagedBy
                Created                 = $OU.Created
                Modified                = $OU.Modified
                ProtectedFromDeletion   = $OU.ProtectedFromAccidentalDeletion
                ContainerPurpose        = Get-ContainerPurpose `
                                            -DistinguishedName $OU.DistinguishedName `
                                            -Name $OU.Name
            }
        }
    }
    catch {
        Write-Log "Failed OU inventory for $DomainName : $_" "ERROR"
    }
}

Export-DataSet -Name "OrganizationalUnits" -Data $OUInventory

# ------------------------------------------------------------
# USERS
# ------------------------------------------------------------

$UsersInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting users from $DomainName"

    try {

        $Users = Get-ADUser `
            -Filter * `
            -Server $DomainName `
            -Properties *

        foreach ($User in $Users) {

            $UsersInventory += [PSCustomObject]@{
                Name                = $User.Name
                SamAccountName      = $User.SamAccountName
                Enabled             = $User.Enabled
                UserPrincipalName   = $User.UserPrincipalName
                EmailAddress        = $User.Mail
                Department          = $User.Department
                Title               = $User.Title
                LastLogonDate       = $User.LastLogonDate
                PasswordLastSet     = $User.PasswordLastSet
                DistinguishedName   = $User.DistinguishedName
            }
        }
    }
    catch {
        Write-Log "Failed user inventory for $DomainName : $_" "ERROR"
    }
}

Export-DataSet -Name "Users" -Data $UsersInventory

# ------------------------------------------------------------
# COMPUTERS
# ------------------------------------------------------------

$ComputerInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting computers from $DomainName"

    try {

        $Computers = Get-ADComputer `
            -Filter * `
            -Server $DomainName `
            -Properties *

        foreach ($Computer in $Computers) {

            $Type = "Workstation"

            if ($Computer.OperatingSystem -match "Server") {
                $Type = "Server"
            }

            if ($Computer.PrimaryGroupID -eq 516) {
                $Type = "Domain Controller"
            }

            $ComputerInventory += [PSCustomObject]@{
                Name                    = $Computer.Name
                DNSHostName             = $Computer.DNSHostName
                OperatingSystem         = $Computer.OperatingSystem
                OperatingSystemVersion  = $Computer.OperatingSystemVersion
                Enabled                 = $Computer.Enabled
                LastLogonDate           = $Computer.LastLogonDate
                Type                    = $Type
                DistinguishedName       = $Computer.DistinguishedName
            }
        }
    }
    catch {
        Write-Log "Failed computer inventory for $DomainName : $_" "ERROR"
    }
}

Export-DataSet -Name "Computers" -Data $ComputerInventory

# ------------------------------------------------------------
# GROUPS
# ------------------------------------------------------------

$GroupInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting groups from $DomainName"

    try {

        $Groups = Get-ADGroup `
            -Filter * `
            -Server $DomainName `
            -Properties *

        foreach ($Group in $Groups) {

            $GroupInventory += [PSCustomObject]@{
                Name                = $Group.Name
                GroupCategory       = $Group.GroupCategory
                GroupScope          = $Group.GroupScope
                ManagedBy           = $Group.ManagedBy
                Description         = $Group.Description
                DistinguishedName   = $Group.DistinguishedName
            }
        }
    }
    catch {
        Write-Log "Failed group inventory for $DomainName : $_" "ERROR"
    }
}

Export-DataSet -Name "Groups" -Data $GroupInventory

# ------------------------------------------------------------
# GROUP MEMBERSHIP
# ------------------------------------------------------------

$MembershipInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting group memberships from $DomainName"

    try {

        $Groups = Get-ADGroup -Filter * -Server $DomainName

        foreach ($Group in $Groups) {

            try {

                $Members = Get-ADGroupMember -Identity $Group -Recursive

                foreach ($Member in $Members) {

                    $MembershipInventory += [PSCustomObject]@{
                        GroupName   = $Group.Name
                        MemberName  = $Member.Name
                        ObjectClass = $Member.objectClass
                    }
                }
            }
            catch {
                Write-Log "Failed membership collection for group $($Group.Name)" "WARN"
            }
        }
    }
    catch {
        Write-Log "Failed group membership inventory." "ERROR"
    }
}

Export-DataSet -Name "GroupMemberships" -Data $MembershipInventory

# ------------------------------------------------------------
# DOMAIN CONTROLLERS
# ------------------------------------------------------------

$DCInventory = @()

foreach ($DomainName in $Domains) {

    Write-Log "Collecting domain controllers from $DomainName"

    try {

        $DCs = Get-ADDomainController `
            -Filter * `
            -Server $DomainName

        foreach ($DC in $DCs) {

            $DCInventory += [PSCustomObject]@{
                HostName        = $DC.HostName
                Site            = $DC.Site
                IPv4Address     = $DC.IPv4Address
                OperatingSystem = $DC.OperatingSystem
                IsGlobalCatalog = $DC.IsGlobalCatalog
                Forest          = $DC.Forest
                Domain          = $DC.Domain
            }
        }
    }
    catch {
        Write-Log "Failed domain controller inventory." "ERROR"
    }
}

Export-DataSet -Name "DomainControllers" -Data $DCInventory

# ------------------------------------------------------------
# GPOS
# ------------------------------------------------------------

try {

    Import-Module GroupPolicy -ErrorAction Stop

    $GPOInventory = Get-GPO -All | ForEach-Object {

        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            Owner       = $_.Owner
            CreationTime= $_.CreationTime
            ModificationTime = $_.ModificationTime
            Id          = $_.Id
        }
    }

    Export-DataSet -Name "GPOs" -Data $GPOInventory
}
catch {
    Write-Log "GroupPolicy module unavailable." "WARN"
}

# ------------------------------------------------------------
# TRUSTS
# ------------------------------------------------------------

$TrustInventory = @()

foreach ($DomainName in $Domains) {

    try {

        $Trusts = Get-ADTrust -Filter * -Server $DomainName

        foreach ($Trust in $Trusts) {

            $TrustInventory += [PSCustomObject]@{
                Name            = $Trust.Name
                Direction       = $Trust.Direction
                TrustType       = $Trust.TrustType
                ForestTransitive= $Trust.ForestTransitive
                IntraForest     = $Trust.IntraForest
            }
        }
    }
    catch {
        Write-Log "Failed trust inventory." "WARN"
    }
}

Export-DataSet -Name "Trusts" -Data $TrustInventory

# ------------------------------------------------------------
# VISIO / MAPPING EXPORTS
# ------------------------------------------------------------

Write-Log "Building AD mapping files..."

# DGML (Visual Studio / Visio compatible)
$DGMLPath = Join-Path $MapFolder "AD_Map.dgml"

$DGML = @()
$DGML += '<?xml version="1.0" encoding="utf-8"?>'
$DGML += '<DirectedGraph xmlns="http://schemas.microsoft.com/vs/2009/dgml">'
$DGML += '<Nodes>'

foreach ($OU in $OUInventory) {
    $DGML += "<Node Id='$($OU.Name)' Label='$($OU.Name)' Category='OU' />"
}

foreach ($Computer in $ComputerInventory) {
    $DGML += "<Node Id='$($Computer.Name)' Label='$($Computer.Name)' Category='$($Computer.Type)' />"
}

foreach ($User in $UsersInventory) {
    $DGML += "<Node Id='$($User.SamAccountName)' Label='$($User.SamAccountName)' Category='User' />"
}

$DGML += '</Nodes>'
$DGML += '<Links>'

foreach ($Computer in $ComputerInventory) {

    $ParentOU = ($Computer.DistinguishedName -split ",",2)[1]

    $DGML += "<Link Source='$ParentOU' Target='$($Computer.Name)' />"
}

foreach ($User in $UsersInventory) {

    $ParentOU = ($User.DistinguishedName -split ",",2)[1]

    $DGML += "<Link Source='$ParentOU' Target='$($User.SamAccountName)' />"
}

$DGML += '</Links>'
$DGML += '</DirectedGraph>'

$DGML | Out-File -FilePath $DGMLPath -Encoding UTF8

# Graphviz DOT
$DOTPath = Join-Path $MapFolder "AD_Map.dot"

$DOT = @()
$DOT += 'digraph ActiveDirectory {'

foreach ($Computer in $ComputerInventory) {

    $ParentOU = ($Computer.DistinguishedName -split ",",2)[1]

    $DOT += "`"$ParentOU`" -> `"$($Computer.Name)`";"
}

foreach ($User in $UsersInventory) {

    $ParentOU = ($User.DistinguishedName -split ",",2)[1]

    $DOT += "`"$ParentOU`" -> `"$($User.SamAccountName)`";"
}

$DOT += '}'

$DOT | Out-File -FilePath $DOTPath -Encoding UTF8

Write-Log "Mapping files generated." "SUCCESS"

# ------------------------------------------------------------
# SUMMARY HTML REPORT
# ------------------------------------------------------------

$SummaryPath = Join-Path $HtmlFolder "AD_Summary.html"

$SummaryHtml = @"
<html>
<head>
<title>AD Inventory Summary</title>
<style>
body {
    font-family: Arial;
    margin: 20px;
}
table {
    border-collapse: collapse;
    width: 80%;
}
th, td {
    border: 1px solid black;
    padding: 8px;
}
th {
    background-color: #2f4f4f;
    color: white;
}
</style>
</head>
<body>

<h1>Active Directory Inventory Summary</h1>

<table>
<tr><th>Category</th><th>Count</th></tr>
<tr><td>Domains</td><td>$($DomainInventory.Count)</td></tr>
<tr><td>Organizational Units</td><td>$($OUInventory.Count)</td></tr>
<tr><td>Users</td><td>$($UsersInventory.Count)</td></tr>
<tr><td>Computers</td><td>$($ComputerInventory.Count)</td></tr>
<tr><td>Groups</td><td>$($GroupInventory.Count)</td></tr>
<tr><td>Domain Controllers</td><td>$($DCInventory.Count)</td></tr>
<tr><td>Trusts</td><td>$($TrustInventory.Count)</td></tr>
</table>

<h2>Export Locations</h2>

<ul>
<li>CSV: $CsvFolder</li>
<li>HTML: $HtmlFolder</li>
<li>Excel: $ExcelFolder</li>
<li>Maps: $MapFolder</li>
</ul>

</body>
</html>
"@

$SummaryHtml | Out-File -FilePath $SummaryPath -Encoding UTF8

# ------------------------------------------------------------
# COMPLETE
# ------------------------------------------------------------

Write-Log "AD inventory completed successfully." "SUCCESS"

Write-Host ""
Write-Host "====================================================="
Write-Host " Active Directory Inventory Completed"
Write-Host "====================================================="
Write-Host " Root Folder : $RootFolder"
Write-Host " HTML Report : $SummaryPath"
Write-Host " Excel File  : $(Join-Path $ExcelFolder 'AD_Inventory.xlsx')"
Write-Host " DGML Map    : $DGMLPath"
Write-Host " DOT Map     : $DOTPath"
Write-Host "====================================================="
