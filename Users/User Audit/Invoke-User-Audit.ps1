<#
.SYNOPSIS
    Active Directory User Security Export Script
.DESCRIPTION
    Queries all Active Directory users across specified domains and exports detailed
    user information including account status, password details, last logon, and
    Tier 0 / Tier 1 administrative classification.
.NOTES
    Requires Active Directory module. Run on a Domain Controller with appropriate
    admin credentials. Outputs to Desktop\DUMPSEC\YYYY-MM-DD_HH-MM-SS\
.EXAMPLE
    .\ADUserSecurityExport.ps1
#>

#Requires -Modules ActiveDirectory - DS Result

# ============================================================
# 1. CONFIGURATION
# ============================================================

# List of domains to query (adjust as needed)
$Domains = @(
    "MYIGT.COM"     # Replace with your primary domain
    # "child.yourdomain.com"  # Add additional domains as needed
)

# Output directory: User's Desktop\DUMPSEC\DateTimeFolder
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$TimeStamp   = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputRoot  = Join-Path -Path $DesktopPath -ChildPath "DUMPSEC"
$OutputDir   = Join-Path -Path $OutputRoot -ChildPath $TimeStamp

# Create output directory
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# File paths
$CsvPath   = Join-Path -Path $OutputDir -ChildPath "AD_User_Report.csv"
$HtmlPath  = Join-Path -Path $OutputDir -ChildPath "AD_User_Report.html"
$XlsxPath  = Join-Path -Path $OutputDir -ChildPath "AD_User_Report.xlsx"

# ============================================================
# 2. DEFINE ALL REQUIRED AD ATTRIBUTES
# ============================================================

$AdProperties = @(
    # Core identity
    'DisplayName',
    'CN',
    'SamAccountName',
    'Name',
    'GivenName',
    'Surname',
    'DistinguishedName',
    'UserPrincipalName',
    
    # Account status & dates
    'Enabled',
    'AccountExpirationDate',
    'accountExpires',
    'LastLogonDate',
    'lastLogonTimestamp',
    'lastLogon',
    'PasswordLastSet',
    'pwdLastSet',
    'PasswordExpired',
    'PasswordNeverExpires',
    'LockedOut',
    'LockoutTime',
    'Created',
    'whenCreated',
    'Modified',
    'whenChanged',
    'Description',
    
    # Email & aliases
    'EmailAddress',
    'mail',
    'mailNickname',
    'proxyAddresses',
    
    # Extension attributes (Ext8, Ext9, Ext10, Ext13, Ext14, Ext15)
    'extensionAttribute8',
    'extensionAttribute9',
    'extensionAttribute10',
    'extensionAttribute13',
    'extensionAttribute14',
    'extensionAttribute15',
    
    # Organizational
    'Department',
    'Company',
    'Title',
    'Office',
    'Manager',
    
    # Security / Tier detection
    'adminCount',
    'MemberOf',
    'PrimaryGroup',
    'SID',
    'ObjectGUID',
    
    # Other useful attributes
    'StreetAddress',
    'PostalCode',
    'State',
    'Country',
    'TelephoneNumber',
    'MobilePhone',
    'Fax',
    'HomePage',
    'ScriptPath',
    'ProfilePath',
    'HomeDirectory',
    'HomeDrive',
    'LogonWorkstations',
    'AccountNotDelegated',
    'AllowReversiblePasswordEncryption',
    'CannotChangePassword',
    'DoesNotRequirePreAuth',
    'HomedirRequired',
    'UserAccountControl'
)

# ============================================================
# 3. TIER CLASSIFICATION HELPERS
# ============================================================

# Tier 0 groups (highly privileged) - expand as needed
$Tier0Groups = @(
    'Domain Admins',
    'Enterprise Admins',
    'Schema Admins',
    'Administrators',
    'Domain Controllers',
    'Group Policy Creator Owners'
)

# Tier 1 groups (server/desktop admin) - expand as needed
$Tier1Groups = @(
    'Server Operators',
    'Account Operators',
    'Backup Operators',
    'Print Operators',
    'Cert Publishers'
)

function Get-UserTier {
    param(
        [string]$SamAccountName,
        [array]$MemberOf,
        [int]$AdminCount
    )
    
    # If adminCount = 1, likely a protected admin account (Tier 0)
    if ($AdminCount -eq 1) {
        return "Tier 0"
    }
    
    # Check group memberships
    $isTier0 = $false
    $isTier1 = $false
    
    foreach ($group in $MemberOf) {
        $groupName = ($group -split ',')[0] -replace '^CN=',''
        if ($Tier0Groups -contains $groupName) {
            $isTier0 = $true
        }
        if ($Tier1Groups -contains $groupName) {
            $isTier1 = $true
        }
    }
    
    if ($isTier0) { return "Tier 0" }
    if ($isTier1) { return "Tier 1" }
    
    return "Standard"
}

# ============================================================
# 4. QUERY ACTIVE DIRECTORY USERS
# ============================================================

Write-Host "Querying Active Directory users across $($Domains.Count) domain(s)..." -ForegroundColor Cyan

$AllUsers = @()

foreach ($Domain in $Domains) {
    Write-Host "  Processing domain: $Domain" -ForegroundColor Yellow
    
    try {
        $DomainUsers = Get-ADUser -Filter * -Server $Domain -Properties $AdProperties -ErrorAction Stop
        
        Write-Host "    Found $($DomainUsers.Count) users in $Domain" -ForegroundColor Green
        
        $AllUsers += $DomainUsers
    }
    catch {
        Write-Host "    ERROR querying domain $Domain : $_" -ForegroundColor Red
        continue
    }
}

Write-Host "Total users across all domains: $($AllUsers.Count)" -ForegroundColor Cyan

# ============================================================
# 5. PROCESS AND TRANSFORM DATA
# ============================================================

Write-Host "Processing user data..." -ForegroundColor Cyan

$ReportData = @()

$userCount = $AllUsers.Count
$processed = 0

foreach ($user in $AllUsers) {
    $processed++
    if ($processed % 100 -eq 0) {
        Write-Progress -Activity "Processing Users" -Status "Processed $processed of $userCount" -PercentComplete (($processed / $userCount) * 100)
    }
    
    # Calculate days since last logon
    $DaysSinceLastLogon = $null
    if ($user.LastLogonDate) {
        $DaysSinceLastLogon = [math]::Round((New-TimeSpan -Start $user.LastLogonDate -End (Get-Date)).TotalDays, 2)
    }
    elseif ($user.lastLogon) {
        try {
            $lastLogonDate = [datetime]::FromFileTime($user.lastLogon)
            if ($lastLogonDate -gt [datetime]::FromFileTime(0)) {
                $DaysSinceLastLogon = [math]::Round((New-TimeSpan -Start $lastLogonDate -End (Get-Date)).TotalDays, 2)
            }
        }
        catch { $DaysSinceLastLogon = $null }
    }
    
    # Days since password last set
    $DaysSincePasswordSet = $null
    if ($user.PasswordLastSet) {
        $DaysSincePasswordSet = [math]::Round((New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).TotalDays, 2)
    }
    
    # Password expiry date and days until expiry
    $PasswordExpiryDate = $null
    $PasswordExpiresInDays = $null
    if ($user.PasswordLastSet -and -not $user.PasswordNeverExpires) {
        # Default AD password max age is 42 days if not overridden by Fine-Grained Password Policy
        # This is a simplified calculation - for production, query msDS-UserPasswordExpiryTimeComputed
        try {
            $maxPwdAge = (Get-ADDefaultDomainPasswordPolicy -Server $user.DistinguishedName).MaxPasswordAge.Days
            $PasswordExpiryDate = $user.PasswordLastSet.AddDays($maxPwdAge)
            $PasswordExpiresInDays = [math]::Round((New-TimeSpan -Start (Get-Date) -End $PasswordExpiryDate).TotalDays, 2)
        }
        catch {
            $PasswordExpiryDate = $null
            $PasswordExpiresInDays = $null
        }
    }
    
    # Extract domain name from DistinguishedName
    $DomainName = ($user.DistinguishedName -split ',') | Where-Object { $_ -like 'DC=*' } | ForEach-Object { ($_ -replace 'DC=','') -join '.' }
    
    # Extract OU from DistinguishedName
    $LogonToOU = ($user.DistinguishedName -split ',') | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_ -replace 'OU=','' } | Join-String -Separator '\'
    if (-not $LogonToOU) { $LogonToOU = "Domain Root" }
    
    # Alias (mailNickname)
    $Alias = $user.mailNickname
    
    # Email (prefer EmailAddress, fallback to mail)
    $Email = if ($user.EmailAddress) { $user.EmailAddress } else { $user.mail }
    
    # Account status (enabled/disabled)
    $AccountStatus = if ($user.Enabled) { "Enabled" } else { "Disabled" }
    
    # Password status
    $PasswordStatus = "Valid"
    if ($user.PasswordExpired -eq $true) { $PasswordStatus = "Expired" }
    if ($user.PasswordNeverExpires -eq $true) { $PasswordStatus = "Never Expires" }
    if (-not $user.PasswordLastSet) { $PasswordStatus = "Never Set" }
    
    # Determine Tier
    $Tier = Get-UserTier -SamAccountName $user.SamAccountName -MemberOf $user.MemberOf -AdminCount $user.adminCount
    
    # Build report object
    $ReportItem = [PSCustomObject]@{
        'Display Name'             = $user.DisplayName
        'Common Name'              = $user.CN
        'SAM Account Name'         = $user.SamAccountName
        'Domain Name'              = $DomainName
        'First Name'               = $user.GivenName
        'Last Name'                = $user.Surname
        'Full Name'                = $user.Name
        'Account Expiry Date'      = if ($user.AccountExpirationDate) { $user.AccountExpirationDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Account Status'           = $AccountStatus
        'Status (enabled/disabled)'= $AccountStatus
        'Days Since Last Logon'    = $DaysSinceLastLogon
        'Last Logon Date'          = if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Last Logon Timestamp'     = if ($user.lastLogonTimestamp) { [datetime]::FromFileTime($user.lastLogonTimestamp).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Days Since Password Set'  = $DaysSincePasswordSet
        'Password Last Set'        = if ($user.PasswordLastSet) { $user.PasswordLastSet.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Password Expires In (Days)' = $PasswordExpiresInDays
        'Password Expiry Date'     = if ($PasswordExpiryDate) { $PasswordExpiryDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Password Never Expires'   = $user.PasswordNeverExpires
        'Password Status'          = $PasswordStatus
        'Account Locked'           = $user.LockedOut
        'Description'              = $user.Description
        'Distinguished Name'       = $user.DistinguishedName
        'E-mail'                   = $Email
        'Alias'                    = $Alias
        'Ext8'                     = $user.extensionAttribute8
        'Ext9'                     = $user.extensionAttribute9
        'Ext10'                    = $user.extensionAttribute10
        'Ext13'                    = $user.extensionAttribute13
        'Ext14'                    = $user.extensionAttribute14
        'Ext15'                    = $user.extensionAttribute15
        'Logon To OU'              = $LogonToOU
        'When Changed'             = if ($user.whenChanged) { $user.whenChanged.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'When Created'             = if ($user.whenCreated) { $user.whenCreated.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Created'                  = if ($user.Created) { $user.Created.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'Modified'                 = if ($user.Modified) { $user.Modified.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        'UserPrincipalName'        = $user.UserPrincipalName
        'Department'               = $user.Department
        'Company'                  = $user.Company
        'Title'                    = $user.Title
        'Office'                   = $user.Office
        'Manager'                  = if ($user.Manager) { ($user.Manager -split ',')[0] -replace '^CN=','' } else { '' }
        'Tier'                     = $Tier
        'adminCount'               = $user.adminCount
        'SID'                      = $user.SID
        'ObjectGUID'               = $user.ObjectGUID
    }
    
    $ReportData += $ReportItem
}

Write-Progress -Activity "Processing Users" -Completed

# ============================================================
# 6. EXPORT TO CSV
# ============================================================

Write-Host "Exporting to CSV..." -ForegroundColor Cyan
$ReportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# ============================================================
# 7. EXPORT TO HTML
# ============================================================

Write-Host "Exporting to HTML..." -ForegroundColor Cyan

# Build HTML report with embedded CSS for readability
$HtmlHead = @"
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 12px; margin: 20px; }
    h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
    .timestamp { color: #7f8c8d; font-size: 14px; }
    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
    th { background-color: #2c3e50; color: white; padding: 8px 6px; text-align: left; border: 1px solid #34495e; }
    td { padding: 6px; border: 1px solid #bdc3c7; }
    tr:nth-child(even) { background-color: #f8f9fa; }
    tr:hover { background-color: #eaf2f8; }
    .status-enabled { color: #27ae60; font-weight: bold; }
    .status-disabled { color: #e74c3c; font-weight: bold; }
    .tier0 { background-color: #ff6b6b !important; color: white; font-weight: bold; }
    .tier1 { background-color: #feca57 !important; font-weight: bold; }
    .expired { color: #e74c3c; font-weight: bold; }
    .never-expires { color: #3498db; font-weight: bold; }
    .locked { color: #e74c3c; font-weight: bold; }
</style>
"@

$HtmlTitle = "Active Directory User Security Report - $TimeStamp"

# Select a subset of key columns for HTML to keep it readable
$HtmlColumns = @(
    'Display Name',
    'SAM Account Name',
    'Domain Name',
    'Account Status',
    'Days Since Last Logon',
    'Last Logon Date',
    'Days Since Password Set',
    'Password Expires In (Days)',
    'Password Status',
    'Account Locked',
    'Tier',
    'Description',
    'Department',
    'Title',
    'Logon To OU',
    'When Created'
)

$HtmlData = $ReportData | Select-Object -Property $HtmlColumns

# Generate HTML
$HtmlContent = $HtmlData | ConvertTo-Html -Title $HtmlTitle -Head $HtmlHead -PreContent "<h1>Active Directory User Security Report</h1><p class='timestamp'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Total Users: $($ReportData.Count)</p>" | Out-String

# Write HTML to file
$HtmlContent | Out-File -FilePath $HtmlPath -Encoding UTF8

# ============================================================
# 8. EXPORT TO XLSX
# ============================================================

Write-Host "Exporting to XLSX..." -ForegroundColor Cyan

# Check if ImportExcel module is available
$ImportExcelAvailable = Get-Module -Name ImportExcel -ListAvailable

if ($ImportExcelAvailable) {
    # Use ImportExcel module (recommended - no Excel required)
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
        $ReportData | Export-Excel -Path $XlsxPath -AutoSize -AutoFilter -FreezeTopRow -WorksheetName 'AD_Users'
        Write-Host "XLSX exported using ImportExcel module." -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module found but Export-Excel command not available. Falling back to COM." -ForegroundColor Yellow
        $ImportExcelAvailable = $false
    }
}

if (-not $ImportExcelAvailable) {
    # Fallback: Use COM object (requires Excel installed)
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Workbook = $Excel.Workbooks.Add()
        $Worksheet = $Workbook.Worksheets.Item(1)
        $Worksheet.Name = "AD_Users"
        
        # Add headers
        $Properties = $ReportData[0].PSObject.Properties.Name
        for ($i = 0; $i -lt $Properties.Count; $i++) {
            $Worksheet.Cells.Item(1, $i + 1) = $Properties[$i]
            $Worksheet.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        
        # Add data
        $row = 2
        foreach ($item in $ReportData) {
            for ($col = 0; $col -lt $Properties.Count; $col++) {
                $value = $item.($Properties[$col])
                if ($value -is [DateTime]) {
                    $value = $value.ToString('yyyy-MM-dd HH:mm:ss')
                }
                $Worksheet.Cells.Item($row, $col + 1) = $value
            }
            $row++
        }
        
        # Auto-fit columns
        $UsedRange = $Worksheet.UsedRange
        $UsedRange.EntireColumn.AutoFit() | Out-Null
        
        # Save and close
        $Workbook.SaveAs($XlsxPath, 51)  # 51 = xlOpenXMLWorkbook
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "XLSX exported using Excel COM object." -ForegroundColor Green
    }
    catch {
        Write-Host "WARNING: Could not export to XLSX. COM object failed: $_" -ForegroundColor Red
        Write-Host "Consider installing the ImportExcel module: Install-Module -Name ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
    }
}

# ============================================================
# 9. SUMMARY
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "           EXPORT COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output Directory: $OutputDir" -ForegroundColor White
Write-Host ""
Write-Host "Files Created:" -ForegroundColor Yellow
Write-Host "  CSV   : $CsvPath" -ForegroundColor White
Write-Host "  HTML  : $HtmlPath" -ForegroundColor White
if (Test-Path $XlsxPath) {
    Write-Host "  XLSX  : $XlsxPath" -ForegroundColor White
} else {
    Write-Host "  XLSX  : NOT CREATED (see warning above)" -ForegroundColor Red
}
Write-Host ""
Write-Host "Total Users Exported: $($ReportData.Count)" -ForegroundColor Cyan
Write-Host ""

# Display Tier breakdown
$TierBreakdown = $ReportData | Group-Object -Property 'Tier' | Select-Object Name, Count
Write-Host "Tier Breakdown:" -ForegroundColor Yellow
foreach ($tier in $TierBreakdown) {
    Write-Host "  $($tier.Name): $($tier.Count)" -ForegroundColor White
}

Write-Host ""
Write-Host "Script completed successfully!" -ForegroundColor Green
