<#
.SYNOPSIS
    Active Directory / DNS / DHCP / GPO / Exchange Inventory (HTML + CSV)
.DESCRIPTION
    Collects forest, domain, FSMO, DC, DNS, DHCP, GPO, privileged groups, and Exchange info.
    Outputs CSVs + full styled HTML report.
#>

#region Setup
$OutputPath = "$env:USERPROFILE\Desktop\AD_Inventory_$(Get-Date -Format yyyyMMdd_HHmmss)"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

$global:FullReportBuilder = New-Object System.Text.StringBuilder

function Add-ContentReport {
    param(
        [Parameter(Mandatory=$true)]
        $html,
        [switch]$LineBreak
    )
    # join arrays safely into single string
    if ($html -is [System.Array]) { $html = ($html -join "`r`n") }
    [void]$global:FullReportBuilder.AppendLine($html)
    if ($LineBreak) { [void]$global:FullReportBuilder.AppendLine("<br/>") }
}

$global:ReportHeaderHTML = @"
<html>
<head>
<title>Active Directory Inventory Report</title>
<style>
body { font-family: Segoe UI, Arial; background-color: #f8f8f8; color: #333; }
h1,h2,h3 { color: #003366; }
table { border-collapse: collapse; width: 98%; margin: 10px 0; }
th,td { border:1px solid #ccc; padding:5px 8px; }
th { background-color:#004080; color:white; }
tr:nth-child(even){background-color:#f2f2f2;}
tr:hover{background-color:#e6f3ff;}
</style>
</head>
<body>
<h1>Active Directory Inventory Report</h1>
<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@

$global:ReportFooterHTML = "</body></html>"

function Import-ModuleSafe {
    param([string]$Name)
    try {
        if (!(Get-Module -ListAvailable -Name $Name)) {
            Write-Verbose "Module $Name not found."
            return $false
        }
        Import-Module -Name $Name -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "Failed to import module ${Name}: $_"
        return $false
    }
}
#endregion

#region Import Modules
Import-ModuleSafe ActiveDirectory | Out-Null
Import-ModuleSafe DnsServer | Out-Null
Import-ModuleSafe DhcpServer | Out-Null
Import-ModuleSafe GroupPolicy | Out-Null
#endregion

#region Forest / Domain
Write-Output "Collecting forest and domain information..."
try {
    $forest = Get-ADForest
    $domain = Get-ADDomain
    $forestObj = [PSCustomObject]@{
        ForestRootDomain     = $forest.RootDomain
        ForestFunctionalLevel = $forest.ForestMode
        ForestDomains        = ($forest.Domains -join ", ")
        ADRecycleBinEnabled  = $forest.RecycleBinEnabled
    }
    $forestObj | Export-Csv "$OutputPath\forest_info.csv" -NoTypeInformation
    Add-ContentReport "<h2>Forest Information</h2>"
    Add-ContentReport (($forestObj | ConvertTo-Html -Fragment) -join "`r`n")

    $domainObj = [PSCustomObject]@{
        DomainName           = $domain.DNSRoot
        NetBIOS_Name         = $domain.NetBIOSName
        DomainFunctionalLevel = $domain.DomainMode
    }
    $domainObj | Export-Csv "$OutputPath\domain_info.csv" -NoTypeInformation
    Add-ContentReport "<h2>Domain Information</h2>"
    Add-ContentReport (($domainObj | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "Forest/Domain info collection failed: $_" }
#endregion

#region FSMO
Write-Output "Collecting FSMO roles..."
try {
    $fsmo = [PSCustomObject]@{
        DomainNamingMaster  = (Get-ADForest).DomainNamingMaster
        SchemaMaster        = (Get-ADForest).SchemaMaster
        PDCEmulator         = (Get-ADDomain).PDCEmulator
        RIDMaster           = (Get-ADDomain).RIDMaster
        InfrastructureMaster = (Get-ADDomain).InfrastructureMaster
    }
    $fsmo | Export-Csv "$OutputPath\fsmo_roles.csv" -NoTypeInformation
    Add-ContentReport "<h2>FSMO Roles</h2>"
    Add-ContentReport (($fsmo | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "FSMO info collection failed: $_" }
#endregion

#region Domain Controllers
Write-Output "Collecting Domain Controllers..."
try {
    $dcList = Get-ADDomainController -Filter *
    $dcOutput = foreach ($dc in $dcList) {
        [PSCustomObject]@{
            Domain = $dc.Domain
            Forest = $dc.Forest
            Name   = $dc.HostName
            IPv4Address = $dc.IPv4Address
            IsGlobalCatalog = $dc.IsGlobalCatalog
            IsReadOnly = $dc.IsReadOnly
            OperatingSystem = $dc.OperatingSystem
            OSVersion = $dc.OperatingSystemVersion
            Site = $dc.Site
        }
    }
    $dcOutput | Export-Csv "$OutputPath\domain_controllers.csv" -NoTypeInformation
    Add-ContentReport "<h2>Domain Controllers</h2>"
    Add-ContentReport (($dcOutput | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "DC collection failed: $_" }
#endregion

#region DNS
Write-Output "Collecting DNS information..."
try {
    if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
        $zones = Get-DnsServerZone -ErrorAction SilentlyContinue
        $dnsRecords = @()
        foreach ($zone in $zones) {
            $records = Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ErrorAction SilentlyContinue
            foreach ($r in $records) {
                $dnsRecords += [PSCustomObject]@{
                    Zone = $zone.ZoneName
                    Host = $r.HostName
                    Type = $r.RecordType
                    Data = ($r.RecordData | Out-String).Trim()
                }
            }
        }
        $dnsRecords | Export-Csv "$OutputPath\dns_records.csv" -NoTypeInformation
        $forwarders = Get-DnsServerForwarder -ErrorAction SilentlyContinue |
            Select-Object IPAddress, Timeout, UseRecursion
        $forwarders | Export-Csv "$OutputPath\dns_forwarders.csv" -NoTypeInformation
        Add-ContentReport "<h2>DNS Zones and Records</h2>"
        Add-ContentReport (($dnsRecords | ConvertTo-Html -Fragment) -join "`r`n")
        Add-ContentReport "<h3>Forwarders</h3>"
        Add-ContentReport (($forwarders | ConvertTo-Html -Fragment) -join "`r`n")
    }
    else {
        Add-ContentReport "<h2>DNS Information</h2><p>DNS cmdlets not available.</p>"
    }
}
catch { Write-Warning "DNS info collection failed: $_" }
#endregion

#region DHCP
Write-Output "Collecting DHCP information..."
try {
    $dhcpOut = @()
    if (Get-Command Get-DhcpServerInDC -ErrorAction SilentlyContinue) {
        $servers = Get-DhcpServerInDC -ErrorAction SilentlyContinue
        foreach ($s in $servers) {
            $dhcpOut += [PSCustomObject]@{
                Name = $s.DnsName
                IPAddress = $s.IPAddress
            }
        }
    } else {
        $local = Get-WindowsFeature -Name DHCP -ErrorAction SilentlyContinue
        if ($local.Installed) {
            $ip = (Get-NetIPAddress -AddressFamily IPv4 |
                Where-Object { $_.IPAddress -notmatch '^169\.254' -and $_.InterfaceAlias -notmatch 'Loopback' } |
                Select-Object -First 1 -ExpandProperty IPAddress)
            $dhcpOut += [PSCustomObject]@{ Name = $env:COMPUTERNAME; IPAddress = $ip }
        }
    }
    $dhcpOut | Export-Csv "$OutputPath\dhcp_servers.csv" -NoTypeInformation
    Add-ContentReport "<h2>DHCP Servers</h2>"
    Add-ContentReport (($dhcpOut | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "DHCP info collection failed: $_" }
#endregion

#region AD Sites
Write-Output "Collecting AD Sites..."
try {
    $sites = Get-ADReplicationSite -Filter * | Select-Object Name
    $links = Get-ADReplicationSiteLink -Filter * | 
        Select-Object Name, Cost, ReplicationFrequencyInMinutes, SitesIncluded
    $sites | Export-Csv "$OutputPath\ad_sites.csv" -NoTypeInformation
    $links | Export-Csv "$OutputPath\ad_sitelinks.csv" -NoTypeInformation
    Add-ContentReport "<h2>AD Sites</h2>"
    Add-ContentReport (($sites | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h2>AD Site Links</h2>"
    Add-ContentReport (($links | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "AD Sites collection failed: $_" }
#endregion

#region GPOs
Write-Output "Collecting GPOs..."
try {
    $gpos = Get-GPO -All -ErrorAction SilentlyContinue |
        Select-Object DisplayName, DomainName, GpoStatus, ModificationTime
    $gpos | Export-Csv "$OutputPath\gpos.csv" -NoTypeInformation
    Add-ContentReport "<h2>Group Policy Objects</h2>"
    Add-ContentReport (($gpos | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "GPO collection failed: $_" }
#endregion

#region Privileged Groups
Write-Output "Collecting privileged groups..."
try {
    $ent = Get-ADGroupMember "Enterprise Admins" -Recursive | Select-Object Name, SamAccountName
    $dom = Get-ADGroupMember "Domain Admins" -Recursive | Select-Object Name, SamAccountName
    $schema = Get-ADGroupMember "Schema Admins" -Recursive | Select-Object Name, SamAccountName
    $pwdNever = Get-ADUser -Filter { PasswordNeverExpires -eq $true } -Properties PasswordNeverExpires |
        Select-Object Name, SamAccountName
    $ent | Export-Csv "$OutputPath\enterprise_admins.csv" -NoTypeInformation
    $dom | Export-Csv "$OutputPath\domain_admins.csv" -NoTypeInformation
    $schema | Export-Csv "$OutputPath\schema_admins.csv" -NoTypeInformation
    $pwdNever | Export-Csv "$OutputPath\password_never_expires.csv" -NoTypeInformation
    Add-ContentReport "<h2>Privileged Accounts</h2>"
    Add-ContentReport "<h3>Enterprise Admins</h3>"
    Add-ContentReport (($ent | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Domain Admins</h3>"
    Add-ContentReport (($dom | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Schema Admins</h3>"
    Add-ContentReport (($schema | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Password Never Expires</h3>"
    Add-ContentReport (($pwdNever | ConvertTo-Html -Fragment) -join "`r`n")
}
catch { Write-Warning "Privileged group collection failed: $_" }
#endregion

#region Exchange
Write-Output "Collecting Exchange information..."
try {
    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
        $exSrv = Get-ExchangeServer | Select-Object Name, Edition, AdminDisplayVersion
        $orgAdmins = Get-ADGroupMember "Organization Management" -Recursive | Select Name, SamAccountName
        $exSrv | Export-Csv "$OutputPath\exchange_servers.csv" -NoTypeInformation
        $orgAdmins | Export-Csv "$OutputPath\exchange_org_admins.csv" -NoTypeInformation
        Add-ContentReport "<h2>Exchange Information</h2>"
        Add-ContentReport (($exSrv | ConvertTo-Html -Fragment) -join "`r`n")
        Add-ContentReport "<h3>Organization Management Members</h3>"
        Add-ContentReport (($orgAdmins | ConvertTo-Html -Fragment) -join "`r`n")
    } else {
        Add-ContentReport "<h2>Exchange Information</h2><p>Exchange cmdlets not found.</p>"
    }
}
catch { Write-Warning "Exchange collection failed: $_" }
#endregion

#region Finalize
Write-Output "Finalizing report..."
$reportFile = Join-Path $OutputPath 'FullReport.html'
$final = $global:ReportHeaderHTML + $global:FullReportBuilder.ToString() + $global:ReportFooterHTML
$final | Out-File -FilePath $reportFile -Encoding UTF8
Write-Output "Report saved to: $reportFile"
#endregion
