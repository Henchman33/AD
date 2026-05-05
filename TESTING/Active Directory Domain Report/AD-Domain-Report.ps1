#Requires -Version 5.1
#Requires -Modules ActiveDirectory, GroupPolicy

<#
.SYNOPSIS
    Comprehensive Active Directory Domain Documentation Script.

.DESCRIPTION
    Recursively documents the entire Active Directory domain including:
    Domain/Forest info, Domain Controllers, FSMO roles, Users, Groups,
    Computers, OUs, GPOs, Sites & Subnets, DNS Zones, Password Policies,
    Trusts, and Replication. Exports to CSV files and a modern executive
    HTML report saved to the user's Desktop.

.NOTES
    Author  : Stephen McKee - Server Administrator - IGT PLC
    Version : 1.0
    Requires: ActiveDirectory RSAT module, GroupPolicy module
              Run from a domain-joined machine with Domain Admin or equivalent read rights.

.EXAMPLE
    .\Export-ADDomainReport.ps1
#>

[CmdletBinding()]
param()

# ─────────────────────────────────────────────────────────────────────────────
#  INITIALISATION
# ─────────────────────────────────────────────────────────────────────────────

$ScriptVersion  = "1.0"
$Author         = "Stephen McKee - Server Administrator - IGT PLC"
$RunDateTime    = Get-Date
$RunDateDisplay = $RunDateTime.ToString("dddd dd MMMM yyyy  HH:mm:ss")
$RunDateFile    = $RunDateTime.ToString("yyyy-MM-dd_HH-mm-ss")
$ReportTitle    = "Active Directory Domain Report"

# Output folder on user Desktop
$DesktopPath    = [System.Environment]::GetFolderPath('Desktop')
$OutputFolder   = Join-Path $DesktopPath "Active Directory Domain"
$CsvFolder      = Join-Path $OutputFolder "CSV"
$HtmlReportPath = Join-Path $OutputFolder "$ReportTitle $RunDateFile.html"

# Create folders
foreach ($folder in @($OutputFolder, $CsvFolder)) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  $ReportTitle" -ForegroundColor White
Write-Host "  Author  : $Author" -ForegroundColor Gray
Write-Host "  Started : $RunDateDisplay" -ForegroundColor Gray
Write-Host "================================================================`n" -ForegroundColor Cyan

# Progress helper
function Write-Step {
    param([string]$Message, [int]$Step, [int]$Total)
    $pct = [int](($Step / $Total) * 100)
    Write-Progress -Activity "Documenting Active Directory" -Status $Message -PercentComplete $pct
    Write-Host "  [$Step/$Total] $Message" -ForegroundColor Yellow
}

# Safe CSV export helper
function Export-SafeCsv {
    param([object[]]$Data, [string]$Path, [string]$Label)
    try {
        if ($Data -and $Data.Count -gt 0) {
            $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
            Write-Host "         -> CSV: $(Split-Path $Path -Leaf)  ($($Data.Count) records)" -ForegroundColor DarkGreen
        } else {
            Write-Host "         -> CSV: $(Split-Path $Path -Leaf)  (no data)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Warning "Failed to export CSV for '$Label': $_"
    }
}

$TotalSteps = 14
$Step = 0

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 1 – DOMAIN & FOREST OVERVIEW
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Domain & Forest Information" $Step $TotalSteps

try {
    $Domain  = Get-ADDomain -ErrorAction Stop
    $Forest  = Get-ADForest -ErrorAction Stop

    $DomainInfo = [PSCustomObject]@{
        'Domain Name (FQDN)'          = $Domain.DNSRoot
        'NetBIOS Name'                = $Domain.NetBIOSName
        'Distinguished Name'          = $Domain.DistinguishedName
        'Domain Mode'                 = $Domain.DomainMode
        'Forest Name'                 = $Forest.Name
        'Forest Mode'                 = $Forest.ForestMode
        'Root Domain'                 = $Forest.RootDomain
        'Schema Master'               = $Forest.SchemaMaster
        'Domain Naming Master'        = $Forest.DomainNamingMaster
        'PDC Emulator'                = $Domain.PDCEmulator
        'RID Master'                  = $Domain.RIDMaster
        'Infrastructure Master'       = $Domain.InfrastructureMaster
        'Domains in Forest'           = ($Forest.Domains -join "; ")
        'Global Catalogs'             = ($Forest.GlobalCatalogs -join "; ")
        'Sites'                       = ($Forest.Sites -join "; ")
        'UPN Suffixes'                = ($Forest.UPNSuffixes -join "; ")
        'SPN Suffixes'                = ($Forest.SPNSuffixes -join "; ")
        'Child Domains'               = ($Domain.ChildDomains -join "; ")
        'Replica Directory Servers'   = ($Domain.ReplicaDirectoryServers -join "; ")
    }

    Export-SafeCsv @($DomainInfo) (Join-Path $CsvFolder "01-Domain-Forest-Overview.csv") "Domain Overview"
} catch {
    Write-Warning "Domain/Forest collection failed: $_"
    $DomainInfo = $null
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 2 – DOMAIN CONTROLLERS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Domain Controllers" $Step $TotalSteps

try {
    $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop | ForEach-Object {
        [PSCustomObject]@{
            'Name'               = $_.Name
            'FQDN'               = $_.HostName
            'IPv4 Address'       = $_.IPv4Address
            'Site'               = $_.Site
            'OS'                 = $_.OperatingSystem
            'OS Version'         = $_.OperatingSystemVersion
            'Is Global Catalog'  = $_.IsGlobalCatalog
            'Is Read-Only (RODC)'= $_.IsReadOnly
            'Enabled'            = $_.Enabled
            'Distinguished Name' = $_.ComputerObjectDN
        }
    } | Sort-Object Name

    Export-SafeCsv $DomainControllers (Join-Path $CsvFolder "02-Domain-Controllers.csv") "Domain Controllers"
} catch {
    Write-Warning "Domain Controller collection failed: $_"
    $DomainControllers = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 3 – AD SITES & SUBNETS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting AD Sites & Subnets" $Step $TotalSteps

try {
    $Sites = Get-ADReplicationSite -Filter * -Properties Description, Location -ErrorAction Stop |
    Select-Object @{n='Site Name';e={$_.Name}},
                  @{n='Description';e={$_.Description}},
                  @{n='Location';e={$_.Location}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Site Name'

    $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Site -ErrorAction Stop |
    Select-Object @{n='Subnet';e={$_.Name}},
                  @{n='Site';e={$_.Site -replace '^CN=([^,]+).*','$1'}},
                  @{n='Description';e={$_.Description}},
                  @{n='Location';e={$_.Location}} |
    Sort-Object 'Subnet'

    $SiteLinks = Get-ADReplicationSiteLink -Filter * -Properties Cost, ReplicationFrequencyInMinutes, SitesIncluded -ErrorAction Stop |
    Select-Object @{n='Site Link Name';e={$_.Name}},
                  @{n='Cost';e={$_.Cost}},
                  @{n='Replication Frequency (min)';e={$_.ReplicationFrequencyInMinutes}},
                  @{n='Sites Included';e={($_.SitesIncluded -replace 'CN=([^,]+).*','$1') -join '; '}} |
    Sort-Object 'Site Link Name'

    Export-SafeCsv $Sites   (Join-Path $CsvFolder "03-AD-Sites.csv") "AD Sites"
    Export-SafeCsv $Subnets (Join-Path $CsvFolder "03-AD-Subnets.csv") "AD Subnets"
    Export-SafeCsv $SiteLinks (Join-Path $CsvFolder "03-AD-SiteLinks.csv") "AD Site Links"
} catch {
    Write-Warning "Sites/Subnets collection failed: $_"
    $Sites = @(); $Subnets = @(); $SiteLinks = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 4 – ORGANISATIONAL UNITS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Organisational Units" $Step $TotalSteps

try {
    $OUs = Get-ADOrganizationalUnit -Filter * -Properties Description, ManagedBy, ProtectedFromAccidentalDeletion -ErrorAction Stop |
    Select-Object @{n='OU Name';e={$_.Name}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}},
                  @{n='Description';e={$_.Description}},
                  @{n='Managed By';e={
                      if ($_.ManagedBy) {
                          try { (Get-ADObject $_.ManagedBy -ErrorAction SilentlyContinue).Name } catch { $_.ManagedBy }
                      } else { '' }
                  }},
                  @{n='Protected from Accidental Deletion';e={$_.ProtectedFromAccidentalDeletion}} |
    Sort-Object 'Distinguished Name'

    Export-SafeCsv $OUs (Join-Path $CsvFolder "04-Organisational-Units.csv") "OUs"
} catch {
    Write-Warning "OU collection failed: $_"
    $OUs = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 5 – USER ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting User Accounts (all properties)" $Step $TotalSteps

try {
    $UserProps = @(
        'SamAccountName','GivenName','Surname','DisplayName','EmailAddress',
        'Department','Title','Description','Company','Office',
        'TelephoneNumber','Mobile','StreetAddress','City','State','Country',
        'Enabled','LockedOut','PasswordNeverExpires','PasswordExpired',
        'PasswordLastSet','LastLogonDate','Created','Modified',
        'Manager','MemberOf','DistinguishedName','UserPrincipalName',
        'AccountExpirationDate','BadLogonCount','CannotChangePassword',
        'SmartcardLogonRequired','TrustedForDelegation','ServicePrincipalNames'
    )

    $AllUsers = Get-ADUser -Filter * -Properties $UserProps -ErrorAction Stop |
    Select-Object @{n='SAM Account';e={$_.SamAccountName}},
                  @{n='First Name';e={$_.GivenName}},
                  @{n='Last Name';e={$_.Surname}},
                  @{n='Display Name';e={$_.DisplayName}},
                  @{n='UPN';e={$_.UserPrincipalName}},
                  @{n='Email';e={$_.EmailAddress}},
                  @{n='Department';e={$_.Department}},
                  @{n='Job Title';e={$_.Title}},
                  @{n='Company';e={$_.Company}},
                  @{n='Office';e={$_.Office}},
                  @{n='Phone';e={$_.TelephoneNumber}},
                  @{n='Mobile';e={$_.Mobile}},
                  @{n='Description';e={$_.Description}},
                  @{n='Enabled';e={$_.Enabled}},
                  @{n='Locked Out';e={$_.LockedOut}},
                  @{n='Password Never Expires';e={$_.PasswordNeverExpires}},
                  @{n='Password Expired';e={$_.PasswordExpired}},
                  @{n='Cannot Change Password';e={$_.CannotChangePassword}},
                  @{n='Smartcard Required';e={$_.SmartcardLogonRequired}},
                  @{n='Trusted for Delegation';e={$_.TrustedForDelegation}},
                  @{n='Password Last Set';e={$_.PasswordLastSet}},
                  @{n='Last Logon Date';e={$_.LastLogonDate}},
                  @{n='Account Expiration';e={$_.AccountExpirationDate}},
                  @{n='Bad Logon Count';e={$_.BadLogonCount}},
                  @{n='Created';e={$_.Created}},
                  @{n='Last Modified';e={$_.Modified}},
                  @{n='Manager';e={
                      if ($_.Manager) {
                          try { (Get-ADUser $_.Manager -ErrorAction SilentlyContinue).DisplayName } catch { $_.Manager }
                      } else { '' }
                  }},
                  @{n='Group Memberships';e={
                      ($_.MemberOf | ForEach-Object {
                          ($_ -split ',')[0] -replace '^CN=',''
                      }) -join '; '
                  }},
                  @{n='SPNs';e={$_.ServicePrincipalNames -join '; '}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'SAM Account'

    # Subsets for quick analysis
    $EnabledUsers   = $AllUsers | Where-Object { $_.Enabled -eq $true }
    $DisabledUsers  = $AllUsers | Where-Object { $_.Enabled -eq $false }
    $LockedUsers    = $AllUsers | Where-Object { $_.'Locked Out' -eq $true }
    $PwdNeverExp    = $AllUsers | Where-Object { $_.'Password Never Expires' -eq $true -and $_.Enabled -eq $true }
    $Cutoff90       = (Get-Date).AddDays(-90)
    $InactiveUsers  = $EnabledUsers | Where-Object { $_.'Last Logon Date' -and [datetime]$_.'Last Logon Date' -lt $Cutoff90 }
    $NeverLoggedIn  = $AllUsers    | Where-Object { -not $_.'Last Logon Date' }

    Export-SafeCsv $AllUsers      (Join-Path $CsvFolder "05-All-Users.csv")           "All Users"
    Export-SafeCsv $EnabledUsers  (Join-Path $CsvFolder "05-Enabled-Users.csv")       "Enabled Users"
    Export-SafeCsv $DisabledUsers (Join-Path $CsvFolder "05-Disabled-Users.csv")      "Disabled Users"
    Export-SafeCsv $LockedUsers   (Join-Path $CsvFolder "05-Locked-Users.csv")        "Locked Users"
    Export-SafeCsv $PwdNeverExp   (Join-Path $CsvFolder "05-Password-Never-Expires.csv") "Pwd Never Expires"
    Export-SafeCsv $InactiveUsers (Join-Path $CsvFolder "05-Inactive-Users-90days.csv")  "Inactive 90 days"
    Export-SafeCsv $NeverLoggedIn (Join-Path $CsvFolder "05-Never-Logged-In.csv")     "Never Logged In"
} catch {
    Write-Warning "User collection failed: $_"
    $AllUsers = @(); $EnabledUsers = @(); $DisabledUsers = @()
    $LockedUsers = @(); $PwdNeverExp = @(); $InactiveUsers = @(); $NeverLoggedIn = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 6 – SECURITY GROUPS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Security Groups & Memberships" $Step $TotalSteps

try {
    $AllGroups = Get-ADGroup -Filter * -Properties Description, ManagedBy, Members, MemberOf, Created, Modified -ErrorAction Stop |
    Select-Object @{n='Group Name';e={$_.Name}},
                  @{n='SAM Account';e={$_.SamAccountName}},
                  @{n='Category';e={$_.GroupCategory}},
                  @{n='Scope';e={$_.GroupScope}},
                  @{n='Description';e={$_.Description}},
                  @{n='Member Count';e={$_.Members.Count}},
                  @{n='Managed By';e={
                      if ($_.ManagedBy) {
                          try { (Get-ADObject $_.ManagedBy -ErrorAction SilentlyContinue).Name } catch { $_.ManagedBy }
                      } else { '' }
                  }},
                  @{n='Created';e={$_.Created}},
                  @{n='Last Modified';e={$_.Modified}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Group Name'

    # Privileged group memberships (recursive)
    $PrivGroups = @(
        "Domain Admins","Enterprise Admins","Schema Admins","Administrators",
        "Account Operators","Backup Operators","Server Operators",
        "Group Policy Creator Owners","Print Operators","Remote Desktop Users",
        "Network Configuration Operators","DnsAdmins"
    )

    $PrivMembers = foreach ($g in $PrivGroups) {
        try {
            Get-ADGroupMember -Identity $g -Recursive -ErrorAction SilentlyContinue |
            ForEach-Object {
                [PSCustomObject]@{
                    'Privileged Group'    = $g
                    'Member Name'         = $_.Name
                    'SAM Account'         = $_.SamAccountName
                    'Object Type'         = $_.objectClass
                    'Distinguished Name'  = $_.distinguishedName
                }
            }
        } catch {}
    }

    Export-SafeCsv $AllGroups   (Join-Path $CsvFolder "06-All-Groups.csv")              "All Groups"
    Export-SafeCsv $PrivMembers (Join-Path $CsvFolder "06-Privileged-Group-Members.csv") "Privileged Groups"
} catch {
    Write-Warning "Group collection failed: $_"
    $AllGroups = @(); $PrivMembers = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 7 – COMPUTER ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Computer Accounts" $Step $TotalSteps

try {
    $CompProps = @(
        'Name','DNSHostName','OperatingSystem','OperatingSystemVersion',
        'IPv4Address','Enabled','LastLogonDate','Created','Modified',
        'Description','MemberOf','DistinguishedName','TrustedForDelegation',
        'ServicePrincipalNames','Location','ManagedBy','SID'
    )

    $AllComputers = Get-ADComputer -Filter * -Properties $CompProps -ErrorAction Stop |
    Select-Object @{n='Computer Name';e={$_.Name}},
                  @{n='DNS Hostname';e={$_.DNSHostName}},
                  @{n='Operating System';e={$_.OperatingSystem}},
                  @{n='OS Version';e={$_.OperatingSystemVersion}},
                  @{n='IPv4 Address';e={$_.IPv4Address}},
                  @{n='Enabled';e={$_.Enabled}},
                  @{n='Last Logon Date';e={$_.LastLogonDate}},
                  @{n='Description';e={$_.Description}},
                  @{n='Location';e={$_.Location}},
                  @{n='Managed By';e={
                      if ($_.ManagedBy) {
                          try { (Get-ADObject $_.ManagedBy -ErrorAction SilentlyContinue).Name } catch { $_.ManagedBy }
                      } else { '' }
                  }},
                  @{n='Trusted for Delegation';e={$_.TrustedForDelegation}},
                  @{n='SPNs';e={$_.ServicePrincipalNames -join '; '}},
                  @{n='Group Memberships';e={
                      ($_.MemberOf | ForEach-Object {
                          ($_ -split ',')[0] -replace '^CN=',''
                      }) -join '; '
                  }},
                  @{n='SID';e={$_.SID}},
                  @{n='Created';e={$_.Created}},
                  @{n='Last Modified';e={$_.Modified}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Operating System','Computer Name'

    $Servers       = $AllComputers | Where-Object { $_.'Operating System' -match 'Server' }
    $Workstations  = $AllComputers | Where-Object { $_.'Operating System' -notmatch 'Server' }
    $DisabledComps = $AllComputers | Where-Object { $_.Enabled -eq $false }
    $Cutoff90c     = (Get-Date).AddDays(-90)
    $StaleComps    = $AllComputers | Where-Object { $_.Enabled -eq $true -and $_.'Last Logon Date' -and [datetime]$_.'Last Logon Date' -lt $Cutoff90c }

    Export-SafeCsv $AllComputers  (Join-Path $CsvFolder "07-All-Computers.csv")       "All Computers"
    Export-SafeCsv $Servers       (Join-Path $CsvFolder "07-Servers.csv")             "Servers"
    Export-SafeCsv $Workstations  (Join-Path $CsvFolder "07-Workstations.csv")        "Workstations"
    Export-SafeCsv $DisabledComps (Join-Path $CsvFolder "07-Disabled-Computers.csv")  "Disabled Computers"
    Export-SafeCsv $StaleComps    (Join-Path $CsvFolder "07-Stale-Computers-90days.csv") "Stale Computers"
} catch {
    Write-Warning "Computer collection failed: $_"
    $AllComputers = @(); $Servers = @(); $Workstations = @(); $DisabledComps = @(); $StaleComps = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 8 – GROUP POLICY OBJECTS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Group Policy Objects" $Step $TotalSteps

try {
    $AllGPOs = Get-GPO -All -ErrorAction Stop

    $GPOReport = foreach ($gpo in $AllGPOs) {
        $links = @()
        try {
            $xml = [xml]($gpo | Get-GPOReport -ReportType XML -ErrorAction SilentlyContinue)
            $links = $xml.GPO.LinksTo | ForEach-Object { $_.SOMPath }
        } catch {}

        [PSCustomObject]@{
            'GPO Name'              = $gpo.DisplayName
            'GPO ID'                = $gpo.Id
            'Status'                = $gpo.GpoStatus
            'Owner'                 = $gpo.Owner
            'Domain'                = $gpo.DomainName
            'Creation Time'         = $gpo.CreationTime
            'Modification Time'     = $gpo.ModificationTime
            'Computer Enabled'      = ($gpo.Computer.Enabled)
            'User Enabled'          = ($gpo.User.Enabled)
            'WMI Filter'            = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { '' }
            'Linked To'             = ($links -join "; ")
            'Link Count'            = $links.Count
        }
    }

    $GPOReport = $GPOReport | Sort-Object 'GPO Name'
    Export-SafeCsv $GPOReport (Join-Path $CsvFolder "08-Group-Policy-Objects.csv") "GPOs"

    # Unlinked GPOs
    $UnlinkedGPOs = $GPOReport | Where-Object { $_.'Link Count' -eq 0 }
    Export-SafeCsv $UnlinkedGPOs (Join-Path $CsvFolder "08-Unlinked-GPOs.csv") "Unlinked GPOs"
} catch {
    Write-Warning "GPO collection failed: $_"
    $GPOReport = @(); $UnlinkedGPOs = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 9 – PASSWORD POLICIES (DEFAULT + FINE-GRAINED)
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Password Policies" $Step $TotalSteps

try {
    # Default Domain Password Policy
    $DDPPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
    $DefaultPolicy = [PSCustomObject]@{
        'Policy Type'                  = 'Default Domain Policy'
        'Min Password Length'          = $DDPPolicy.MinPasswordLength
        'Password History Count'       = $DDPPolicy.PasswordHistoryCount
        'Complexity Enabled'           = $DDPPolicy.ComplexityEnabled
        'Reversible Encryption'        = $DDPPolicy.ReversibleEncryptionEnabled
        'Max Password Age'             = $DDPPolicy.MaxPasswordAge.Days
        'Min Password Age (days)'      = $DDPPolicy.MinPasswordAge.Days
        'Lockout Threshold'            = $DDPPolicy.LockoutThreshold
        'Lockout Duration (mins)'      = $DDPPolicy.LockoutDuration.TotalMinutes
        'Lockout Observation Window'   = $DDPPolicy.LockoutObservationWindow.TotalMinutes
    }

    # Fine-Grained Password Policies
    $PSOs = Get-ADFineGrainedPasswordPolicy -Filter * -Properties AppliesTo -ErrorAction SilentlyContinue |
    ForEach-Object {
        $pso = $_
        $appliesToNames = ($pso.AppliesTo | ForEach-Object {
            try { (Get-ADObject $_ -ErrorAction SilentlyContinue).Name } catch { $_ }
        }) -join '; '

        [PSCustomObject]@{
            'Policy Type'                  = 'Fine-Grained PSO'
            'PSO Name'                     = $pso.Name
            'Precedence'                   = $pso.Precedence
            'Applies To'                   = $appliesToNames
            'Min Password Length'          = $pso.MinPasswordLength
            'Password History Count'       = $pso.PasswordHistoryCount
            'Complexity Enabled'           = $pso.ComplexityEnabled
            'Reversible Encryption'        = $pso.ReversibleEncryptionEnabled
            'Max Password Age (days)'      = $pso.MaxPasswordAge.Days
            'Min Password Age (days)'      = $pso.MinPasswordAge.Days
            'Lockout Threshold'            = $pso.LockoutThreshold
            'Lockout Duration (mins)'      = $pso.LockoutDuration.TotalMinutes
            'Lockout Observation Window'   = $pso.LockoutObservationWindow.TotalMinutes
        }
    } | Sort-Object Precedence

    Export-SafeCsv @($DefaultPolicy) (Join-Path $CsvFolder "09-Default-Password-Policy.csv") "Default Policy"
    Export-SafeCsv $PSOs             (Join-Path $CsvFolder "09-Fine-Grained-Password-Policies.csv") "PSOs"
} catch {
    Write-Warning "Password policy collection failed: $_"
    $DefaultPolicy = $null; $PSOs = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 10 – DNS ZONES
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting DNS Zones" $Step $TotalSteps

try {
    $DnsZones = Get-ADObject -SearchBase "CN=MicrosoftDNS,DC=DomainDnsZones,$($Domain.DistinguishedName)" `
                -Filter { ObjectClass -eq "dnsZone" } -Properties Name, DistinguishedName `
                -ErrorAction SilentlyContinue |
    Select-Object @{n='Zone Name';e={$_.Name}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Zone Name'

    if (-not $DnsZones) {
        # Fallback: try DnsServer module if available
        if (Get-Module -ListAvailable -Name DnsServer -ErrorAction SilentlyContinue) {
            Import-Module DnsServer -ErrorAction SilentlyContinue
            $DnsZones = Get-DnsServerZone -ErrorAction SilentlyContinue |
            Select-Object @{n='Zone Name';e={$_.ZoneName}},
                          @{n='Zone Type';e={$_.ZoneType}},
                          @{n='Is Auto-Created';e={$_.IsAutoCreated}},
                          @{n='Is DS Integrated';e={$_.IsDsIntegrated}},
                          @{n='Is Reverse Lookup';e={$_.IsReverseLookupZone}},
                          @{n='Replication Scope';e={$_.ReplicationScope}} |
            Sort-Object 'Zone Name'
        }
    }

    Export-SafeCsv $DnsZones (Join-Path $CsvFolder "10-DNS-Zones.csv") "DNS Zones"
} catch {
    Write-Warning "DNS zone collection failed: $_"
    $DnsZones = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 11 – TRUSTS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Domain Trusts" $Step $TotalSteps

try {
    $Trusts = Get-ADTrust -Filter * -ErrorAction SilentlyContinue |
    Select-Object @{n='Trust Name';e={$_.Name}},
                  @{n='Direction';e={$_.Direction}},
                  @{n='Trust Type';e={$_.TrustType}},
                  @{n='SID Filtering';e={$_.SIDFilteringQuarantined}},
                  @{n='Selective Auth';e={$_.SelectiveAuthentication}},
                  @{n='Forest Transitive';e={$_.ForestTransitive}},
                  @{n='Intra-Forest';e={$_.IntraForest}},
                  @{n='Within Forest';e={$_.IsTreeRoot}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Trust Name'

    Export-SafeCsv $Trusts (Join-Path $CsvFolder "11-Domain-Trusts.csv") "Trusts"
} catch {
    Write-Warning "Trust collection failed: $_"
    $Trusts = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 12 – AD REPLICATION
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Replication Status" $Step $TotalSteps

try {
    $ReplPartners = Get-ADReplicationPartnerMetadata -Target * -ErrorAction SilentlyContinue |
    Select-Object @{n='Server';e={$_.Server -replace '\.\S+$'}},
                  @{n='Partner';e={$_.Partner -replace 'CN=NTDS Settings,CN=([^,]+).*','$1'}},
                  @{n='Partition';e={$_.Partition}},
                  @{n='Last Replication Success';e={$_.LastReplicationSuccess}},
                  @{n='Last Replication Attempt';e={$_.LastReplicationAttempt}},
                  @{n='Last Replication Result';e={$_.LastReplicationResult}},
                  @{n='Consecutive Failures';e={$_.ConsecutiveReplicationFailures}} |
    Sort-Object Server, Partner

    Export-SafeCsv $ReplPartners (Join-Path $CsvFolder "12-Replication-Status.csv") "Replication"
} catch {
    Write-Warning "Replication status collection failed: $_"
    $ReplPartners = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 13 – SERVICE ACCOUNTS & MANAGED SERVICE ACCOUNTS
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Service & Managed Service Accounts" $Step $TotalSteps

try {
    # gMSA accounts
    $gMSA = Get-ADServiceAccount -Filter * -Properties Description, Created, Modified, PrincipalsAllowedToRetrieveManagedPassword -ErrorAction SilentlyContinue |
    Select-Object @{n='Account Name';e={$_.Name}},
                  @{n='SAM Account';e={$_.SamAccountName}},
                  @{n='Account Type';e={if($_.ObjectClass -eq 'msDS-GroupManagedServiceAccount'){'gMSA'}else{'MSA'}}},
                  @{n='Description';e={$_.Description}},
                  @{n='Enabled';e={$_.Enabled}},
                  @{n='Principals Allowed to Retrieve Password';e={
                      ($_.PrincipalsAllowedToRetrieveManagedPassword | ForEach-Object {
                          try { (Get-ADObject $_ -ErrorAction SilentlyContinue).Name } catch { $_ }
                      }) -join '; '
                  }},
                  @{n='Created';e={$_.Created}},
                  @{n='Last Modified';e={$_.Modified}},
                  @{n='Distinguished Name';e={$_.DistinguishedName}} |
    Sort-Object 'Account Name'

    # Users with SPNs (potential service accounts)
    $SPNUsers = $AllUsers | Where-Object { $_.SPNs -ne '' } |
    Select-Object 'SAM Account','Display Name','Department','Enabled','Password Never Expires','Last Logon Date','SPNs','Distinguished Name'

    Export-SafeCsv $gMSA     (Join-Path $CsvFolder "13-Managed-Service-Accounts.csv") "MSA/gMSA"
    Export-SafeCsv $SPNUsers (Join-Path $CsvFolder "13-Users-With-SPNs.csv")          "SPN Users"
} catch {
    Write-Warning "Service account collection failed: $_"
    $gMSA = @(); $SPNUsers = @()
}

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 14 – SCHEMA & OPTIONAL FEATURES
# ─────────────────────────────────────────────────────────────────────────────
$Step++
Write-Step "Collecting Schema & Optional Features" $Step $TotalSteps

try {
    $SchemaInfo = Get-ADObject (Get-ADRootDSE).schemaNamingContext -Properties objectVersion -ErrorAction SilentlyContinue
    $SchemaVersion = $SchemaInfo.objectVersion

    $OptFeatures = Get-ADOptionalFeature -Filter * -ErrorAction SilentlyContinue |
    Select-Object @{n='Feature Name';e={$_.Name}},
                  @{n='Feature Scope';e={$_.FeatureScope}},
                  @{n='Enabled Scopes';e={$_.EnabledScopes -join '; '}},
                  @{n='Required Forest Mode';e={$_.RequiredForestMode}},
                  @{n='Required Domain Mode';e={$_.RequiredDomainMode}},
                  @{n='Is Disabled By Default';e={$_.IsDisabledByDefault}}

    Export-SafeCsv $OptFeatures (Join-Path $CsvFolder "14-Optional-Features.csv") "Optional Features"
} catch {
    Write-Warning "Schema/features collection failed: $_"
    $SchemaVersion = "Unknown"; $OptFeatures = @()
}

Write-Progress -Activity "Documenting Active Directory" -Completed
Write-Host "`n  [OK] Data collection complete. Generating reports...`n" -ForegroundColor Green

# ─────────────────────────────────────────────────────────────────────────────
#  HELPER – BUILD HTML TABLE
# ─────────────────────────────────────────────────────────────────────────────
function ConvertTo-HtmlTable {
    param(
        [object[]]$Data,
        [string]$TableId
    )
    if (-not $Data -or $Data.Count -eq 0) {
        return '<p class="no-data">No data found for this section.</p>'
    }

    $props = ($Data[0] | Get-Member -MemberType NoteProperty,Property | Select-Object -ExpandProperty Name)

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrapper'><table id='$TableId'>")
    [void]$sb.Append("<thead><tr>")
    foreach ($p in $props) {
        [void]$sb.Append("<th>$p</th>")
    }
    [void]$sb.Append("</tr></thead><tbody>")

    foreach ($row in $Data) {
        [void]$sb.Append("<tr>")
        foreach ($p in $props) {
            $val = $row.$p
            if ($null -eq $val) { $val = '' }

            # Colour-code boolean values
            $cell = switch ($val.ToString().ToLower()) {
                'true'    { "<td><span class='badge badge-green'>True</span></td>" }
                'false'   { "<td><span class='badge badge-red'>False</span></td>" }
                default   { "<td>$([System.Web.HttpUtility]::HtmlEncode($val.ToString()))</td>" }
            }
            [void]$sb.Append($cell)
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

# Load HttpUtility for HTML encoding
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD STAT CARDS
# ─────────────────────────────────────────────────────────────────────────────
$domainName       = if ($DomainInfo) { $DomainInfo.'Domain Name (FQDN)' } else { 'N/A' }
$forestMode       = if ($DomainInfo) { $DomainInfo.'Forest Mode' } else { 'N/A' }
$domainMode       = if ($DomainInfo) { $DomainInfo.'Domain Mode' } else { 'N/A' }
$schemaVer        = if ($SchemaVersion) { $SchemaVersion } else { 'N/A' }
$totalUsers       = if ($AllUsers)    { $AllUsers.Count }    else { 0 }
$totalEnabled     = if ($EnabledUsers) { $EnabledUsers.Count } else { 0 }
$totalDisabled    = if ($DisabledUsers){ $DisabledUsers.Count } else { 0 }
$totalLocked      = if ($LockedUsers)  { $LockedUsers.Count }  else { 0 }
$totalPwdNE       = if ($PwdNeverExp)  { $PwdNeverExp.Count }  else { 0 }
$totalInactive    = if ($InactiveUsers){ $InactiveUsers.Count } else { 0 }
$totalComputers   = if ($AllComputers) { $AllComputers.Count } else { 0 }
$totalServers     = if ($Servers)      { $Servers.Count }      else { 0 }
$totalWorkstations= if ($Workstations) { $Workstations.Count } else { 0 }
$totalGroups      = if ($AllGroups)    { $AllGroups.Count }    else { 0 }
$totalGPOs        = if ($GPOReport)    { $GPOReport.Count }    else { 0 }
$totalDCs         = if ($DomainControllers){ $DomainControllers.Count } else { 0 }
$totalOUs         = if ($OUs)          { $OUs.Count }          else { 0 }
$totalSites       = if ($Sites)        { $Sites.Count }        else { 0 }
$totalTrusts      = if ($Trusts)       { $Trusts.Count }       else { 0 }
$totalUnlinkedGPO = if ($UnlinkedGPOs) { $UnlinkedGPOs.Count } else { 0 }
$totalStaleComp   = if ($StaleComps)   { $StaleComps.Count }   else { 0 }
$totalPrivMembers = if ($PrivMembers)  { $PrivMembers.Count }  else { 0 }

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD DOMAIN INFO TABLE
# ─────────────────────────────────────────────────────────────────────────────
function Build-KVTable {
    param([object]$Obj)
    if (-not $Obj) { return '<p class="no-data">No data available.</p>' }
    $props = $Obj | Get-Member -MemberType NoteProperty,Property | Select-Object -ExpandProperty Name
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<table class='kv-table'>")
    foreach ($p in $props) {
        $val = $Obj.$p
        if ($null -eq $val) { $val = '' }
        [void]$sb.Append("<tr><th>$p</th><td>$([System.Web.HttpUtility]::HtmlEncode($val.ToString()))</td></tr>")
    }
    [void]$sb.Append("</table>")
    return $sb.ToString()
}

# ─────────────────────────────────────────────────────────────────────────────
#  GENERATE HTML REPORT
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "  Generating HTML report..." -ForegroundColor Cyan

$DomainInfoKV       = Build-KVTable $DomainInfo
$DefaultPolicyKV    = Build-KVTable $DefaultPolicy

$T_DCs         = ConvertTo-HtmlTable $DomainControllers   "tbl-dc"
$T_Sites       = ConvertTo-HtmlTable $Sites               "tbl-sites"
$T_Subnets     = ConvertTo-HtmlTable $Subnets             "tbl-subnets"
$T_SiteLinks   = ConvertTo-HtmlTable $SiteLinks           "tbl-sitelinks"
$T_OUs         = ConvertTo-HtmlTable $OUs                 "tbl-ous"
$T_AllUsers    = ConvertTo-HtmlTable $AllUsers            "tbl-allusers"
$T_PwdNE       = ConvertTo-HtmlTable $PwdNeverExp         "tbl-pwdne"
$T_Inactive    = ConvertTo-HtmlTable $InactiveUsers       "tbl-inactive"
$T_Locked      = ConvertTo-HtmlTable $LockedUsers         "tbl-locked"
$T_NeverLogin  = ConvertTo-HtmlTable $NeverLoggedIn       "tbl-neverlogin"
$T_Groups      = ConvertTo-HtmlTable $AllGroups           "tbl-groups"
$T_PrivMem     = ConvertTo-HtmlTable $PrivMembers         "tbl-privmem"
$T_Computers   = ConvertTo-HtmlTable $AllComputers        "tbl-computers"
$T_Servers     = ConvertTo-HtmlTable $Servers             "tbl-servers"
$T_StaleComps  = ConvertTo-HtmlTable $StaleComps          "tbl-stalecomp"
$T_GPOs        = ConvertTo-HtmlTable $GPOReport           "tbl-gpos"
$T_UnlinkGPOs  = ConvertTo-HtmlTable $UnlinkedGPOs        "tbl-unlinkgpo"
$T_PSOs        = ConvertTo-HtmlTable $PSOs                "tbl-psos"
$T_DNS         = ConvertTo-HtmlTable $DnsZones            "tbl-dns"
$T_Trusts      = ConvertTo-HtmlTable $Trusts              "tbl-trusts"
$T_Repl        = ConvertTo-HtmlTable $ReplPartners        "tbl-repl"
$T_MSA         = ConvertTo-HtmlTable $gMSA                "tbl-msa"
$T_SPNUsers    = ConvertTo-HtmlTable $SPNUsers            "tbl-spnusers"
$T_OptFeat     = ConvertTo-HtmlTable $OptFeatures         "tbl-optfeat"

$HtmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$ReportTitle – $RunDateFile</title>
<style>
  :root {
    --bg:         #0d1117;
    --bg-card:    #161b22;
    --bg-header:  #0d1117;
    --border:     #30363d;
    --accent:     #388bfd;
    --accent2:    #1f6feb;
    --text:       #e6edf3;
    --text-muted: #8b949e;
    --green:      #2ea043;
    --green-bg:   #0d2818;
    --red:        #f85149;
    --red-bg:     #3d0c0c;
    --amber:      #d29922;
    --amber-bg:   #2c2000;
    --radius:     8px;
    --radius-lg:  12px;
    --shadow:     0 4px 24px rgba(0,0,0,.5);
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body {
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 14px;
    line-height: 1.6;
  }

  /* ── TOP NAV ── */
  .top-nav {
    position: sticky; top: 0; z-index: 100;
    background: rgba(13,17,23,.92);
    backdrop-filter: blur(12px);
    border-bottom: 1px solid var(--border);
    padding: 10px 32px;
    display: flex; align-items: center; gap: 16px;
  }
  .top-nav h1 { font-size: 15px; font-weight: 600; color: var(--text); white-space: nowrap; }
  .top-nav .domain-badge {
    background: var(--accent2); color: #fff;
    padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600;
  }
  .search-wrap { flex: 1; max-width: 420px; position: relative; }
  .search-wrap input {
    width: 100%; background: var(--bg-card); border: 1px solid var(--border);
    color: var(--text); padding: 7px 12px 7px 36px; border-radius: 6px; font-size: 13px;
    outline: none; transition: border-color .2s;
  }
  .search-wrap input:focus { border-color: var(--accent); }
  .search-wrap .search-icon {
    position: absolute; left: 11px; top: 50%; transform: translateY(-50%);
    color: var(--text-muted); pointer-events: none;
  }
  .btn-collapse-all {
    background: var(--bg-card); border: 1px solid var(--border);
    color: var(--text-muted); padding: 6px 14px; border-radius: 6px;
    cursor: pointer; font-size: 12px; white-space: nowrap; transition: all .2s;
  }
  .btn-collapse-all:hover { border-color: var(--accent); color: var(--text); }

  /* ── HERO HEADER ── */
  .hero {
    background: linear-gradient(135deg, #0d1117 0%, #161b22 50%, #0d1117 100%);
    border-bottom: 1px solid var(--border);
    padding: 48px 48px 36px;
    position: relative; overflow: hidden;
  }
  .hero::before {
    content: ''; position: absolute; inset: 0;
    background: radial-gradient(ellipse 80% 60% at 60% 40%, rgba(56,139,253,.08), transparent);
    pointer-events: none;
  }
  .hero-logo {
    display: flex; align-items: center; gap: 16px; margin-bottom: 20px;
  }
  .hero-logo .logo-icon {
    width: 52px; height: 52px;
    background: linear-gradient(135deg, var(--accent2), #388bfd);
    border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
    font-size: 26px; flex-shrink: 0; box-shadow: 0 0 20px rgba(56,139,253,.35);
  }
  .hero h2 { font-size: 28px; font-weight: 700; letter-spacing: -.5px; }
  .hero .subtitle { color: var(--text-muted); font-size: 14px; margin-top: 4px; }
  .hero-meta {
    display: flex; gap: 32px; flex-wrap: wrap; margin-top: 24px;
    padding-top: 20px; border-top: 1px solid var(--border);
  }
  .hero-meta .meta-item { display: flex; flex-direction: column; gap: 2px; }
  .hero-meta .meta-label { font-size: 11px; text-transform: uppercase; letter-spacing: .8px; color: var(--text-muted); }
  .hero-meta .meta-value { font-size: 14px; font-weight: 600; color: var(--text); }

  /* ── STAT CARDS ── */
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
    gap: 12px;
    padding: 28px 48px;
  }
  .stat-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 18px 20px;
    transition: border-color .2s, transform .15s;
  }
  .stat-card:hover { border-color: var(--accent); transform: translateY(-2px); }
  .stat-card .stat-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: .6px; }
  .stat-card .stat-value { font-size: 28px; font-weight: 700; margin: 4px 0 2px; line-height: 1; }
  .stat-card .stat-sub { font-size: 11px; color: var(--text-muted); }
  .stat-card.accent  .stat-value { color: var(--accent); }
  .stat-card.green   .stat-value { color: var(--green); }
  .stat-card.red     .stat-value { color: var(--red); }
  .stat-card.amber   .stat-value { color: var(--amber); }

  /* ── MAIN CONTENT ── */
  .container { padding: 0 48px 60px; }

  /* ── SECTION CARD ── */
  .section-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    margin-bottom: 16px;
    overflow: hidden;
    box-shadow: var(--shadow);
  }
  .section-header {
    display: flex; align-items: center; gap: 12px;
    padding: 16px 20px; cursor: pointer;
    user-select: none;
    transition: background .15s;
    border-bottom: 1px solid transparent;
  }
  .section-header:hover { background: rgba(56,139,253,.06); }
  .section-header.open { border-bottom-color: var(--border); }
  .section-icon { font-size: 18px; width: 32px; text-align: center; }
  .section-title { font-size: 15px; font-weight: 600; flex: 1; }
  .section-count {
    font-size: 11px; font-weight: 600;
    background: rgba(56,139,253,.15); color: var(--accent);
    padding: 2px 9px; border-radius: 20px;
  }
  .section-count.warn {
    background: rgba(210,153,34,.15); color: var(--amber);
  }
  .section-count.danger {
    background: rgba(248,81,73,.15); color: var(--red);
  }
  .chevron { color: var(--text-muted); transition: transform .25s; font-size: 12px; }
  .section-header.open .chevron { transform: rotate(90deg); }
  .section-body { display: none; padding: 20px; }
  .section-body.open { display: block; }

  /* subsection */
  .subsection { margin-top: 20px; }
  .subsection-title {
    font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: .8px;
    color: var(--text-muted); margin-bottom: 10px;
    display: flex; align-items: center; gap: 8px;
  }
  .subsection-title::after {
    content: ''; flex: 1; height: 1px; background: var(--border);
  }

  /* ── TABLES ── */
  .table-wrapper { overflow-x: auto; border-radius: var(--radius); border: 1px solid var(--border); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead { background: #1c2128; }
  thead th {
    padding: 10px 14px; text-align: left; font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: .6px; color: var(--text-muted);
    border-bottom: 1px solid var(--border); white-space: nowrap;
  }
  tbody tr { transition: background .1s; }
  tbody tr:nth-child(even) { background: rgba(255,255,255,.02); }
  tbody tr:hover { background: rgba(56,139,253,.08); }
  tbody td {
    padding: 9px 14px; border-bottom: 1px solid rgba(48,54,61,.5);
    color: var(--text); vertical-align: top; max-width: 320px;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }
  tbody tr:last-child td { border-bottom: none; }

  /* KV table */
  .kv-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .kv-table th {
    width: 260px; padding: 9px 14px; text-align: left; font-weight: 600;
    color: var(--text-muted); font-size: 12px;
    background: rgba(255,255,255,.02); border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }
  .kv-table td {
    padding: 9px 14px; border-bottom: 1px solid var(--border);
    color: var(--text); word-break: break-all;
  }
  .kv-table tr:last-child th, .kv-table tr:last-child td { border-bottom: none; }

  /* ── BADGES ── */
  .badge {
    display: inline-block; padding: 2px 8px; border-radius: 20px;
    font-size: 11px; font-weight: 600;
  }
  .badge-green { background: var(--green-bg); color: var(--green); }
  .badge-red   { background: var(--red-bg);   color: var(--red); }

  /* ── ALERT BOXES ── */
  .alert {
    padding: 12px 16px; border-radius: var(--radius); margin-bottom: 16px;
    display: flex; align-items: flex-start; gap: 10px; font-size: 13px;
  }
  .alert.warn  { background: var(--amber-bg); border: 1px solid rgba(210,153,34,.3); color: var(--amber); }
  .alert.info  { background: rgba(56,139,253,.08); border: 1px solid rgba(56,139,253,.25); color: var(--accent); }

  /* ── MISC ── */
  .no-data { color: var(--text-muted); font-style: italic; padding: 20px 0; }
  .highlight { background: rgba(255,255,0,.25); border-radius: 2px; }

  /* ── FOOTER ── */
  .footer {
    text-align: center; padding: 32px;
    border-top: 1px solid var(--border);
    color: var(--text-muted); font-size: 12px; margin-top: 20px;
  }
  .footer strong { color: var(--text); }

  @media print {
    .top-nav { display: none; }
    .section-body { display: block !important; }
    body { background: #fff; color: #000; }
  }
</style>
</head>
<body>

<!-- ── TOP NAVIGATION ── -->
<nav class="top-nav">
  <h1>&#x1F4BB; $ReportTitle</h1>
  <span class="domain-badge">$domainName</span>
  <div class="search-wrap">
    <span class="search-icon">&#128269;</span>
    <input type="text" id="globalSearch" placeholder="Search all tables…" oninput="globalSearch(this.value)">
  </div>
  <button class="btn-collapse-all" onclick="toggleAll()">&#x25B6; Expand / Collapse All</button>
</nav>

<!-- ── HERO ── -->
<header class="hero">
  <div class="hero-logo">
    <div class="logo-icon">&#x1F3C1;</div>
    <div>
      <h2>$ReportTitle</h2>
      <div class="subtitle">Comprehensive Active Directory Domain Documentation</div>
    </div>
  </div>
  <div class="hero-meta">
    <div class="meta-item"><span class="meta-label">Domain</span><span class="meta-value">$domainName</span></div>
    <div class="meta-item"><span class="meta-label">Forest Mode</span><span class="meta-value">$forestMode</span></div>
    <div class="meta-item"><span class="meta-label">Domain Mode</span><span class="meta-value">$domainMode</span></div>
    <div class="meta-item"><span class="meta-label">Schema Version</span><span class="meta-value">$schemaVer</span></div>
    <div class="meta-item"><span class="meta-label">Generated</span><span class="meta-value">$RunDateDisplay</span></div>
    <div class="meta-item"><span class="meta-label">Author</span><span class="meta-value">$Author</span></div>
    <div class="meta-item"><span class="meta-label">Script Version</span><span class="meta-value">v$ScriptVersion</span></div>
  </div>
</header>

<!-- ── STATS GRID ── -->
<section class="stats-grid">
  <div class="stat-card accent"><div class="stat-label">Domain Controllers</div><div class="stat-value">$totalDCs</div><div class="stat-sub">Active DCs</div></div>
  <div class="stat-card green"><div class="stat-label">Enabled Users</div><div class="stat-value">$totalEnabled</div><div class="stat-sub">of $totalUsers total</div></div>
  <div class="stat-card"><div class="stat-label">Disabled Users</div><div class="stat-value">$totalDisabled</div><div class="stat-sub">User accounts</div></div>
  <div class="stat-card $(if($totalLocked -gt 0){'red'}else{'green'})"><div class="stat-label">Locked Accounts</div><div class="stat-value">$totalLocked</div><div class="stat-sub">Currently locked</div></div>
  <div class="stat-card $(if($totalPwdNE -gt 0){'amber'}else{'green'})"><div class="stat-label">Pwd Never Expires</div><div class="stat-value">$totalPwdNE</div><div class="stat-sub">Enabled users</div></div>
  <div class="stat-card $(if($totalInactive -gt 0){'amber'}else{'green'})"><div class="stat-label">Inactive Users</div><div class="stat-value">$totalInactive</div><div class="stat-sub">90+ days no logon</div></div>
  <div class="stat-card accent"><div class="stat-label">Total Computers</div><div class="stat-value">$totalComputers</div><div class="stat-sub">$totalServers servers / $totalWorkstations workstations</div></div>
  <div class="stat-card $(if($totalStaleComp -gt 0){'amber'}else{'green'})"><div class="stat-label">Stale Computers</div><div class="stat-value">$totalStaleComp</div><div class="stat-sub">90+ days no logon</div></div>
  <div class="stat-card"><div class="stat-label">Security Groups</div><div class="stat-value">$totalGroups</div><div class="stat-sub">All group types</div></div>
  <div class="stat-card $(if($totalPrivMembers -gt 0){'amber'}else{'green'})"><div class="stat-label">Privileged Members</div><div class="stat-value">$totalPrivMembers</div><div class="stat-sub">Across all priv groups</div></div>
  <div class="stat-card accent"><div class="stat-label">Group Policies</div><div class="stat-value">$totalGPOs</div><div class="stat-sub">$totalUnlinkedGPO unlinked</div></div>
  <div class="stat-card"><div class="stat-label">Sites</div><div class="stat-value">$totalSites</div><div class="stat-sub">AD Sites</div></div>
  <div class="stat-card"><div class="stat-label">OUs</div><div class="stat-value">$totalOUs</div><div class="stat-sub">Organisational Units</div></div>
  <div class="stat-card $(if($totalTrusts -gt 0){'accent'}else{'green'})"><div class="stat-label">Domain Trusts</div><div class="stat-value">$totalTrusts</div><div class="stat-sub">Active trusts</div></div>
</section>

<!-- ── MAIN CONTENT ── -->
<main class="container">

<!-- ═══ 1. DOMAIN & FOREST ═══ -->
<div class="section-card" id="sec-domain">
  <div class="section-header open" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F30D;</span>
    <span class="section-title">Domain &amp; Forest Overview</span>
    <span class="section-count">FSMO Roles &amp; Forest Structure</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body open">
    $DomainInfoKV
  </div>
</div>

<!-- ═══ 2. DOMAIN CONTROLLERS ═══ -->
<div class="section-card" id="sec-dcs">
  <div class="section-header open" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F5A5;&#xFE0F;</span>
    <span class="section-title">Domain Controllers</span>
    <span class="section-count">$totalDCs DCs</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body open">
    $T_DCs
  </div>
</div>

<!-- ═══ 3. SITES & SUBNETS ═══ -->
<div class="section-card" id="sec-sites">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F4CD;</span>
    <span class="section-title">Sites &amp; Subnets</span>
    <span class="section-count">$totalSites Sites</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    <div class="subsection"><div class="subsection-title">AD Sites</div>$T_Sites</div>
    <div class="subsection"><div class="subsection-title">Subnets</div>$T_Subnets</div>
    <div class="subsection"><div class="subsection-title">Site Links</div>$T_SiteLinks</div>
  </div>
</div>

<!-- ═══ 4. ORGANISATIONAL UNITS ═══ -->
<div class="section-card" id="sec-ous">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F4C2;</span>
    <span class="section-title">Organisational Units (OUs)</span>
    <span class="section-count">$totalOUs OUs</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $T_OUs
  </div>
</div>

<!-- ═══ 5. USER ACCOUNTS ═══ -->
<div class="section-card" id="sec-users">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F465;</span>
    <span class="section-title">User Accounts</span>
    <span class="section-count">$totalUsers Total / $totalEnabled Enabled</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">

    $(if ($totalLocked -gt 0) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalLocked + ' user account(s) are currently locked out.</div>' })
    $(if ($totalPwdNE -gt 0) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalPwdNE + ' enabled user(s) have Password Never Expires set – review for compliance.</div>' })
    $(if ($totalInactive -gt 0) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalInactive + ' enabled user(s) have not logged in for 90+ days – consider reviewing or disabling.</div>' })

    <div class="subsection"><div class="subsection-title">All Users ($totalUsers)</div>$T_AllUsers</div>
    <div class="subsection"><div class="subsection-title">Locked Out Accounts ($totalLocked)</div>$T_Locked</div>
    <div class="subsection"><div class="subsection-title">Password Never Expires ($totalPwdNE)</div>$T_PwdNE</div>
    <div class="subsection"><div class="subsection-title">Inactive Users – 90+ Days ($totalInactive)</div>$T_Inactive</div>
    <div class="subsection"><div class="subsection-title">Never Logged In</div>$T_NeverLogin</div>
  </div>
</div>

<!-- ═══ 6. SECURITY GROUPS ═══ -->
<div class="section-card" id="sec-groups">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F512;</span>
    <span class="section-title">Security Groups</span>
    <span class="section-count">$totalGroups Groups / $totalPrivMembers Privileged Members</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $(if ($totalPrivMembers -gt 5) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalPrivMembers + ' members found in privileged groups. Review for least-privilege compliance.</div>' })
    <div class="subsection"><div class="subsection-title">All Groups ($totalGroups)</div>$T_Groups</div>
    <div class="subsection"><div class="subsection-title">Privileged Group Members (Recursive)</div>$T_PrivMem</div>
  </div>
</div>

<!-- ═══ 7. COMPUTERS ═══ -->
<div class="section-card" id="sec-computers">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F4BB;</span>
    <span class="section-title">Computer Accounts</span>
    <span class="section-count">$totalComputers Total ($totalServers Servers / $totalWorkstations Workstations)</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $(if ($totalStaleComp -gt 0) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalStaleComp + ' enabled computer(s) have not contacted AD in 90+ days.</div>' })
    <div class="subsection"><div class="subsection-title">Servers ($totalServers)</div>$T_Servers</div>
    <div class="subsection"><div class="subsection-title">All Computers ($totalComputers)</div>$T_Computers</div>
    <div class="subsection"><div class="subsection-title">Stale Computers – 90+ Days ($totalStaleComp)</div>$T_StaleComps</div>
  </div>
</div>

<!-- ═══ 8. GROUP POLICY ═══ -->
<div class="section-card" id="sec-gpo">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F4DC;</span>
    <span class="section-title">Group Policy Objects (GPOs)</span>
    <span class="section-count">$totalGPOs GPOs / $totalUnlinkedGPO Unlinked</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $(if ($totalUnlinkedGPO -gt 0) { '<div class="alert warn">&#x26A0;&#xFE0F; ' + $totalUnlinkedGPO + ' GPO(s) are not linked to any OU – consider reviewing or removing.</div>' })
    <div class="subsection"><div class="subsection-title">All GPOs ($totalGPOs)</div>$T_GPOs</div>
    <div class="subsection"><div class="subsection-title">Unlinked GPOs ($totalUnlinkedGPO)</div>$T_UnlinkGPOs</div>
  </div>
</div>

<!-- ═══ 9. PASSWORD POLICIES ═══ -->
<div class="section-card" id="sec-pwdpol">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F511;</span>
    <span class="section-title">Password Policies</span>
    <span class="section-count">Default + Fine-Grained PSOs</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    <div class="subsection"><div class="subsection-title">Default Domain Password Policy</div>$DefaultPolicyKV</div>
    <div class="subsection"><div class="subsection-title">Fine-Grained Password Policies (PSOs)</div>$T_PSOs</div>
  </div>
</div>

<!-- ═══ 10. DNS ZONES ═══ -->
<div class="section-card" id="sec-dns">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F310;</span>
    <span class="section-title">DNS Zones</span>
    <span class="section-count">AD-Integrated Zones</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $T_DNS
  </div>
</div>

<!-- ═══ 11. TRUSTS ═══ -->
<div class="section-card" id="sec-trusts">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F91D;</span>
    <span class="section-title">Domain Trusts</span>
    <span class="section-count">$totalTrusts Trust(s)</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $T_Trusts
  </div>
</div>

<!-- ═══ 12. REPLICATION ═══ -->
<div class="section-card" id="sec-repl">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F504;</span>
    <span class="section-title">AD Replication Status</span>
    <span class="section-count">Partner Metadata</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    $T_Repl
  </div>
</div>

<!-- ═══ 13. SERVICE ACCOUNTS ═══ -->
<div class="section-card" id="sec-svc">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x2699;&#xFE0F;</span>
    <span class="section-title">Service &amp; Managed Service Accounts</span>
    <span class="section-count">MSA / gMSA &amp; SPN Users</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    <div class="subsection"><div class="subsection-title">Managed Service Accounts (MSA / gMSA)</div>$T_MSA</div>
    <div class="subsection"><div class="subsection-title">User Accounts with Service Principal Names (SPNs)</div>$T_SPNUsers</div>
  </div>
</div>

<!-- ═══ 14. OPTIONAL FEATURES ═══ -->
<div class="section-card" id="sec-features">
  <div class="section-header" onclick="toggleSection(this)">
    <span class="section-icon">&#x1F9E9;</span>
    <span class="section-title">Optional Features &amp; Schema</span>
    <span class="section-count">Schema v$schemaVer</span>
    <span class="chevron">&#9654;</span>
  </div>
  <div class="section-body">
    <div class="alert info">&#x2139;&#xFE0F; AD Schema Version $schemaVer — use this to identify the Windows Server version that last extended the schema.</div>
    $T_OptFeat
  </div>
</div>

</main>

<!-- ── FOOTER ── -->
<footer class="footer">
  <p>Generated by <strong>$Author</strong> &nbsp;|&nbsp; $RunDateDisplay &nbsp;|&nbsp; Script v$ScriptVersion</p>
  <p style="margin-top:6px; color:#555">This report is confidential and intended for authorised personnel only.</p>
</footer>

<script>
// ── Section collapse / expand ──
function toggleSection(header) {
  header.classList.toggle('open');
  const body = header.nextElementSibling;
  body.classList.toggle('open');
}

// ── Expand / collapse all ──
let allExpanded = true;
function toggleAll() {
  const headers = document.querySelectorAll('.section-header');
  allExpanded = !allExpanded;
  headers.forEach(h => {
    const body = h.nextElementSibling;
    if (allExpanded) {
      h.classList.add('open');
      body.classList.add('open');
    } else {
      h.classList.remove('open');
      body.classList.remove('open');
    }
  });
}

// ── Global search ──
function globalSearch(query) {
  const q = query.toLowerCase().trim();
  // Remove previous highlights
  document.querySelectorAll('.highlight').forEach(el => {
    el.outerHTML = el.innerHTML;
  });

  if (!q) {
    // Restore all rows
    document.querySelectorAll('tbody tr').forEach(r => r.style.display = '');
    return;
  }

  document.querySelectorAll('tbody tr').forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(q) ? '' : 'none';

    if (text.includes(q)) {
      // Expand parent section
      const card = row.closest('.section-card');
      if (card) {
        const header = card.querySelector('.section-header');
        const body   = card.querySelector('.section-body');
        if (header) header.classList.add('open');
        if (body)   body.classList.add('open');
      }
      // Highlight matching cells
<script>

  row.querySelectorAll('td').forEach(cell => {
    if (cell.textContent.toLowerCase().includes(q)) {
      const re = new RegExp('(' + q.replace(/[.*+?^\`${}()|[\]\\]/g,'\\$&') + ')', 'gi');
      cell.innerHTML = cell.innerHTML.replace(re, '<span class="highlight">$</span>');
    }
  });
</script>
'@

// Allow pressing Escape to clear search
document.getElementById('globalSearch').addEventListener('keydown', function(e) {
  if (e.key === 'Escape') { this.value = ''; globalSearch(''); }
});
</script>
</body>
</html>
"@

# Write HTML file
try {
    [System.IO.File]::WriteAllText($HtmlReportPath, $HtmlContent, [System.Text.Encoding]::UTF8)
    Write-Host "  [OK] HTML report saved: $HtmlReportPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write HTML report: $_"
}

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

$Elapsed = (Get-Date) - $RunDateTime

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Output Folder : $OutputFolder" -ForegroundColor White
Write-Host "  HTML Report   : $(Split-Path $HtmlReportPath -Leaf)" -ForegroundColor White
Write-Host "  CSV Folder    : $CsvFolder" -ForegroundColor White
Write-Host ""
Write-Host "  Domain Summary:" -ForegroundColor Gray
Write-Host "    Domain Controllers : $totalDCs" -ForegroundColor White
Write-Host "    Total Users        : $totalUsers  (Enabled: $totalEnabled | Disabled: $totalDisabled)" -ForegroundColor White
Write-Host "    Locked Accounts    : $totalLocked" -ForegroundColor $(if($totalLocked -gt 0){'Yellow'}else{'White'})
Write-Host "    Pwd Never Expires  : $totalPwdNE" -ForegroundColor $(if($totalPwdNE -gt 0){'Yellow'}else{'White'})
Write-Host "    Inactive Users     : $totalInactive  (90+ days)" -ForegroundColor $(if($totalInactive -gt 0){'Yellow'}else{'White'})
Write-Host "    Total Computers    : $totalComputers  (Servers: $totalServers | Workstations: $totalWorkstations)" -ForegroundColor White
Write-Host "    Stale Computers    : $totalStaleComp  (90+ days)" -ForegroundColor $(if($totalStaleComp -gt 0){'Yellow'}else{'White'})
Write-Host "    Security Groups    : $totalGroups" -ForegroundColor White
Write-Host "    GPOs               : $totalGPOs  (Unlinked: $totalUnlinkedGPO)" -ForegroundColor $(if($totalUnlinkedGPO -gt 0){'Yellow'}else{'White'})
Write-Host "    Sites              : $totalSites" -ForegroundColor White
Write-Host "    OUs                : $totalOUs" -ForegroundColor White
Write-Host "    Domain Trusts      : $totalTrusts" -ForegroundColor White
Write-Host ""
Write-Host "  Completed in  : $([math]::Round($Elapsed.TotalMinutes,1)) minutes" -ForegroundColor Gray
Write-Host "================================================================`n" -ForegroundColor Cyan

# Open the output folder automatically
Start-Process explorer.exe -ArgumentList $OutputFolder
