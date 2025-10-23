<#
.SYNOPSIS
    Active Directory / DNS / DHCP / GPO / Exchange Inventory (HTML + CSV)
.DESCRIPTION
    Collects forest, domain, FSMO, DC, DNS, DHCP, GPO, privileged groups, and Exchange info.
    Outputs CSVs + a fully styled, collapsible HTML report.
       ***PowerShell5ADScript***
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
    if ($html -is [System.Array]) { $html = ($html -join "`r`n") }
    [void]$global:FullReportBuilder.AppendLine($html)
    if ($LineBreak) { [void]$global:FullReportBuilder.AppendLine("<br/>") }
}

$global:ReportHeaderHTML = @"
<html>
<head>
<title>Active Directory Inventory Report</title>
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8f8f8; color: #333; margin: 20px; }
h1,h2,h3 { color: #003366; }
table { border-collapse: collapse; width: 98%; margin: 10px 0; }
th,td { border:1px solid #ccc; padding:5px 8px; font-size: 13px; }
th { background-color:#004080; color:white; text-align:left; }
tr:nth-child(even){background-color:#f2f2f2;}
tr:hover{background-color:#e6f3ff;}
details { background: #ffffff; border: 1px solid #ccc; border-radius: 6px; margin: 8px 0; padding: 8px; }
summary { font-weight: bold; cursor: pointer; font-size: 16px; color: #004080; }
summary:hover { color: #0078d7; }
</style>
</head>
<body>
<h1>Active Directory Inventory Report</h1>
<h2>Author: Steve McKee IGTPLC<h2>
<p><b>Generated:</b> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
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
    Add-ContentReport "<details open><summary>Forest Information</summary>"
    Add-ContentReport (($forestObj | ConvertTo-Html -Fragment) -join "`r`n")

    $domainObj = [PSCustomObject]@{
        DomainName           = $domain.DNSRoot
        NetBIOS_Name         = $domain.NetBIOSName
        DomainFunctionalLevel = $domain.DomainMode
    }
    $domainObj | Export-Csv "$OutputPath\domain_info.csv" -NoTypeInformation
    Add-ContentReport "<h3>Domain Information</h3>"
    Add-ContentReport (($domainObj | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
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
    Add-ContentReport "<details><summary>FSMO Roles</summary>"
    Add-ContentReport (($fsmo | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
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
    Add-ContentReport "<details><summary>Domain Controllers</summary>"
    Add-ContentReport (($dcOutput | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
catch { Write-Warning "DC collection failed: $_" }
#endregion

#region DNS
Write-Output "Collecting DNS information..."
try {
    if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
        $zones = Get-DnsServerZone -ErrorAction SilentlyContinue
        $dnsRecords = @()
        $serversUsed = @()

        foreach ($zone in $zones) {
            $records = Get-DnsServerResourceRecord -ZoneName $zone.ZoneName -ErrorAction SilentlyContinue
            foreach ($r in $records) {
                $serversUsed += $r.PSComputerName
                $data = switch ($r.RecordType) {
                    'A'      { if ($r.RecordData) { $r.RecordData.IPv4Address.ToString() } else { '' } }
                    'AAAA'   { if ($r.RecordData) { $r.RecordData.IPv6Address.ToString() } else { '' } }
                    'CNAME'  { if ($r.RecordData) { $r.RecordData.HostNameAlias } else { '' } }
                    'NS'     { if ($r.RecordData) { $r.RecordData.NameServer } else { '' } }
                    'SRV'    { if ($r.RecordData) { "$($r.RecordData.DomainNameTarget):$($r.RecordData.Port)" } else { '' } }
                    'PTR'    { if ($r.RecordData) { $r.RecordData.PtrDomainName } else { '' } }
                    default  { ($r.RecordData | Out-String).Trim() }
                }

                $data = $data -replace "PSComputerName\s*:.*", ""
                $data = $data.Trim()

                $dnsRecords += [PSCustomObject]@{
                    Zone   = $zone.ZoneName
                    Host   = $r.HostName
                    Type   = $r.RecordType
                    Data   = $data
                    Server = $r.PSComputerName
                }
            }
        }

        $dnsRecords | Export-Csv "$OutputPath\dns_records.csv" -NoTypeInformation
        $forwarders = Get-DnsServerForwarder -ErrorAction SilentlyContinue |
            Select-Object @{Name='IPAddress';Expression={$_.IPAddress.IPAddressToString}}, Timeout, UseRecursion
        $forwarders | Export-Csv "$OutputPath\dns_forwarders.csv" -NoTypeInformation

        $serversUsed = $serversUsed | Sort-Object -Unique

        Add-ContentReport "<details><summary>DNS Zones and Records</summary>"
        if ($serversUsed.Count -gt 1) {
            Add-ContentReport (($dnsRecords | ConvertTo-Html -Fragment) -join "`r`n")
        } else {
            # Hide "Server" column if only one server
            $dnsRecordsView = $dnsRecords | Select-Object Zone,Host,Type,Data
            Add-ContentReport (($dnsRecordsView | ConvertTo-Html -Fragment) -join "`r`n")
        }
        Add-ContentReport "<h3>DNS Forwarders</h3>"
        Add-ContentReport (($forwarders | ConvertTo-Html -Fragment) -join "`r`n")
        Add-ContentReport "</details>"
    }
    else {
        Add-ContentReport "<details><summary>DNS Information</summary><p>DNS cmdlets not available.</p></details>"
    }
}
catch { Write-Warning "DNS info collection failed: $_" }
#endregion

#region DHCP
Write-Output "Collecting DHCP information..."
try {
    $dhcpOut = @()

    # Primary: use Get-DhcpServerInDC (authorized DHCP servers in AD) if available
    if (Get-Command Get-DhcpServerInDC -ErrorAction SilentlyContinue) {
        try {
            $servers = Get-DhcpServerInDC -ErrorAction Stop
        } catch {
            Write-Warning "Get-DhcpServerInDC failed or returned nothing: $_"
            $servers = $null
        }

        if ($servers -and $servers.Count -gt 0) {
            foreach ($s in $servers) {
                # Different versions/properties may exist; be defensive.
                $name = $s.DnsName
                if (-not $name) { $name = $s.ServerId -as [string] }
                $ip = $null
                if ($s.PSObject.Properties['IPAddress']) { $ip = $s.IPAddress }
                elseif ($s.PSObject.Properties['IpAddress']) { $ip = $s.IpAddress }
                elseif ($s.PSObject.Properties['ServerIpAddress']) { $ip = $s.ServerIpAddress }
                $dhcpOut += [PSCustomObject]@{
                    Name = $name
                    IPAddress = $ip
                }
            }
        }
    }

    # Fallback 1: If the DhcpServer module/cmdlets not available or returned nothing, try to read AD DhcpRoot container
    if (($dhcpOut.Count -eq 0) -and (Get-Command Get-ADObject -ErrorAction SilentlyContinue)) {
        try {
            $cfg = (Get-ADRootDSE).ConfigurationNamingContext
            $dhcpRootDN = "CN=DhcpRoot,CN=NetServices,CN=Services,$cfg"
            $dhcpRoot = Get-ADObject -Identity $dhcpRootDN -ErrorAction SilentlyContinue
            if ($dhcpRoot) {
                # Get children of DhcpRoot — typically each authorized server is represented under it
                $children = Get-ADObject -SearchBase $dhcpRootDN -Filter * -SearchScope OneLevel -Properties * -ErrorAction SilentlyContinue
                foreach ($c in $children) {
                    # Best-effort mapping; show name and DN so admin can inspect
                    $displayName = $c.Name
                    $dn = $c.DistinguishedName
                    $dhcpOut += [PSCustomObject]@{
                        Name = $displayName
                        IPAddress = ($c.IPAddress -join ', ')  # may be empty
                        DistinguishedName = $dn
                    }
                }
            }
        } catch {
            Write-Warning "AD DhcpRoot lookup failed: $_"
        }
    }

    # Fallback 2: As a last resort, search for computers with 'dhcp' in name/description (useful for small networks)
    if ($dhcpOut.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
        try {
            $candidates = Get-ADComputer -Filter { Name -like "*dhcp*" -or Description -like "*dhcp*" } -Properties Name,IPv4Address,Description -ErrorAction SilentlyContinue
            foreach ($c in $candidates) {
                $dhcpOut += [PSCustomObject]@{
                    Name = $c.Name
                    IPAddress = ($c.IPv4Address -join ', ')
                    Description = $c.Description
                }
            }
        } catch {
            Write-Warning "AD computer fallback for DHCP failed: $_"
        }
    }

    # If still nothing, add a note
    if ($dhcpOut.Count -eq 0) {
        $dhcpOut += [PSCustomObject]@{
            Name = '<No DHCP servers found>'
            IPAddress = ''
            Note = 'No DHCP servers discovered by cmdlets or AD fallbacks. Ensure DHCP server is authorized in AD and you have permissions.'
        }
    }

    $dhcpOut | Export-Csv "$OutputPath\dhcp_servers.csv" -NoTypeInformation
    Add-ContentReport "<details><summary>DHCP Servers</summary>"
    Add-ContentReport (($dhcpOut | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
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
    Add-ContentReport "<details><summary>AD Sites and Site Links</summary>"
    Add-ContentReport "<h3>Sites</h3>"
    Add-ContentReport (($sites | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Site Links</h3>"
    Add-ContentReport (($links | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
catch { Write-Warning "AD Sites collection failed: $_" }
#endregion

#region GPOs
Write-Output "Collecting GPOs..."
try {
    $gpos = Get-GPO -All -ErrorAction SilentlyContinue |
        Select-Object DisplayName, DomainName, GpoStatus, ModificationTime
    $gpos | Export-Csv "$OutputPath\gpos.csv" -NoTypeInformation
    Add-ContentReport "<details><summary>Group Policy Objects</summary>"
    Add-ContentReport (($gpos | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
catch { Write-Warning "GPO collection failed: $_" }
#endregion

#region Privileged Groups
Write-Output "Collecting privileged groups..."
try {
    $ent = Get-ADGroupMember "Enterprise Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name, SamAccountName
    $dom = Get-ADGroupMember "Domain Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name, SamAccountName
    $schema = Get-ADGroupMember "Schema Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name, SamAccountName
    $pwdNever = Get-ADUser -Filter { PasswordNeverExpires -eq $true } -Properties PasswordNeverExpires |
        Select-Object Name, SamAccountName
    $ent | Export-Csv "$OutputPath\enterprise_admins.csv" -NoTypeInformation
    $dom | Export-Csv "$OutputPath\domain_admins.csv" -NoTypeInformation
    $schema | Export-Csv "$OutputPath\schema_admins.csv" -NoTypeInformation
    $pwdNever | Export-Csv "$OutputPath\password_never_expires.csv" -NoTypeInformation
    Add-ContentReport "<details><summary>Privileged Accounts</summary>"
    Add-ContentReport "<h3>Enterprise Admins</h3>"
    Add-ContentReport (($ent | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Domain Admins</h3>"
    Add-ContentReport (($dom | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Schema Admins</h3>"
    Add-ContentReport (($schema | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "<h3>Password Never Expires</h3>"
    Add-ContentReport (($pwdNever | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
catch { Write-Warning "Privileged group collection failed: $_" }
#endregion

#region Exchange
Write-Output "Collecting Exchange information..."
try {
    $exSrv = @()
    $orgAdmins = @()

    # Primary: native Exchange cmdlet (requires Exchange Management Shell / remote PS session)
    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
        try {
            $exSrv = Get-ExchangeServer -ErrorAction Stop | Select-Object Name, Edition, AdminDisplayVersion
            # Organization Management group members
            $orgAdmins = Get-ADGroupMember "Organization Management" -Recursive -ErrorAction SilentlyContinue | Select Name, SamAccountName
        } catch {
            Write-Warning "Get-ExchangeServer failed: $_"
            $exSrv = @()
        }
    }

    # Fallback A: search AD for computer objects with msExchVersion attribute (typical Exchange server attribute)
    if ($exSrv.Count -eq 0 -and (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
        try {
            $exComputers = Get-ADComputer -LDAPFilter "(msExchVersion=*)" -Properties msExchVersion,OperatingSystem -ErrorAction SilentlyContinue
            if ($exComputers -and $exComputers.Count -gt 0) {
                foreach ($c in $exComputers) {
                    $exSrv += [PSCustomObject]@{
                        Name = $c.Name
                        Edition = $null
                        AdminDisplayVersion = $c.msExchVersion
                        OperatingSystem = $c.OperatingSystem
                    }
                }
            }
        } catch {
            Write-Warning "LDAP search for msExchVersion failed: $_"
        }
    }

    # Fallback B: find members of common Exchange groups (Exchange Servers / Exchange Enterprise Servers)
    if ($exSrv.Count -eq 0 -and (Get-Command Get-ADGroup -ErrorAction SilentlyContinue)) {
        $possibleGroups = @("Exchange Servers","Exchange Enterprise Servers","Microsoft Exchange Security Group","Exchange Servers (Default Domain Group)")
        foreach ($g in $possibleGroups) {
            try {
                $grp = Get-ADGroup -Filter "Name -eq '$g'" -ErrorAction SilentlyContinue
                if ($grp) {
                    $members = Get-ADGroupMember -Identity $grp -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'computer' }
                    foreach ($m in $members) {
                        $name = $m.Name
                        # avoid duplicates
                        if (-not ($exSrv | Where-Object { $_.Name -eq $name })) {
                            $exSrv += [PSCustomObject]@{ Name = $name; Edition = $null; AdminDisplayVersion = $null }
                        }
                    }
                }
            } catch { } # ignore and continue
        }
    }

    # If we still have nothing, return a helpful note
    if ($exSrv.Count -eq 0) {
        Add-ContentReport "<details><summary>Exchange Information</summary><p><b>Exchange cmdlets not found and no Exchange server objects discovered by AD fallbacks.</b></p><p>To get detailed Exchange info, run this script from an Exchange Management Shell or enable remote Exchange PowerShell, or ensure Exchange tools are installed on this machine.</p></details>"
        # Write minimal CSV so user sees something
        $nullObj = [PSCustomObject]@{ Name = '<No Exchange servers found>'; Edition = ''; AdminDisplayVersion = ''; Note = 'Native Exchange cmdlets not available and AD fallbacks did not find Exchange objects.' }
        $nullObj | Export-Csv "$OutputPath\exchange_servers.csv" -NoTypeInformation
    } else {
        $exSrv | Export-Csv "$OutputPath\exchange_servers.csv" -NoTypeInformation
        if ($orgAdmins.Count -gt 0) { $orgAdmins | Export-Csv "$OutputPath\exchange_org_admins.csv" -NoTypeInformation }
        Add-ContentReport "<details><summary>Exchange Information</summary>"
        Add-ContentReport (($exSrv | ConvertTo-Html -Fragment) -join "`r`n")
        if ($orgAdmins.Count -gt 0) {
            Add-ContentReport "<h3>Organization Management Members</h3>"
            Add-ContentReport (($orgAdmins | ConvertTo-Html -Fragment) -join "`r`n")
        }
        Add-ContentReport "</details>"
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
#region GPOs - NTLM Configuration
Write-Output "Analyzing GPOs for NTLM configuration..."
try {
    $ntlmGpoData = @()

    $allGpos = Get-GPO -All -ErrorAction SilentlyContinue
    foreach ($gpo in $allGpos) {
        try {
            # Export GPO report as XML (fast, complete)
            $reportXml = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
            if (-not $reportXml) { continue }

            [xml]$xml = $reportXml

            # Look for any policy or registry entry mentioning NTLM
            $matches = @()
            $matches += $xml.GPO.Computer.ExtensionData.Extension.Policy | Where-Object {
                $_.Name -match 'NTLM' -or $_.KeyName -match 'NTLM' -or $_.DisplayName -match 'NTLM'
            }
            $matches += $xml.GPO.User.ExtensionData.Extension.Policy | Where-Object {
                $_.Name -match 'NTLM' -or $_.KeyName -match 'NTLM' -or $_.DisplayName -match 'NTLM'
            }

            if ($matches.Count -gt 0) {
                # Gather link info (OUs)
                $links = @()
                $inherit = Get-GPInheritance -Domain $gpo.DomainName -Target "DC=$($gpo.DomainName -replace '\.',',DC=')" -ErrorAction SilentlyContinue
                # If Get-GPInheritance doesn’t yield all OU links, fall back to ADSI search for gPLink attribute
                try {
                    $linkedOUs = Get-ADOrganizationalUnit -Filter * -Properties gPLink | Where-Object { $_.gPLink -match $gpo.Id.Guid }
                    if ($linkedOUs) {
                        $links = $linkedOUs | Select-Object -ExpandProperty DistinguishedName
                    }
                } catch { }

                # Create output rows
                foreach ($m in $matches) {
                    $ntlmGpoData += [PSCustomObject]@{
                        GPOName   = $gpo.DisplayName
                        GPOEnabled = $gpo.GpoStatus
                        SettingName = $m.Name
                        RegistryKey = $m.KeyName
                        Value      = $m.State
                        LinkedOUs  = ($links -join '; ')
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed NTLM analysis for $($gpo.DisplayName): $_"
        }
    }

    if ($ntlmGpoData.Count -eq 0) {
        $ntlmGpoData += [PSCustomObject]@{
            GPOName = '<No NTLM-related settings found>'
            GPOEnabled = ''
            SettingName = ''
            RegistryKey = ''
            Value = ''
            LinkedOUs = ''
        }
    }

    # Export to CSV
    $ntlmGpoData | Export-Csv "$OutputPath\gpos_ntlm.csv" -NoTypeInformation

    # Add to HTML report
    Add-ContentReport "<details><summary>NTLM Configuration GPOs</summary>"
    Add-Co

