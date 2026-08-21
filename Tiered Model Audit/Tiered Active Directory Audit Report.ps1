<#
.SYNOPSIS
    Audits cross-tier membership in a Microsoft Active Directory tiered model (Tier 0 / Tier 1).
.DESCRIPTION
    Checks that no user, group, or computer from Tier 0 is a member of a Tier 1 group,
    and vice versa. Detects group nesting across tiers.
    Uses OU=Admin as the search base to match your AD structure.
    Outputs CSV and HTML reports to the user's Desktop.
.NOTES
    Requires: ActiveDirectory module, read access to Admin OU.
#>

#Requires -Modules ActiveDirectory

# ---- Configuration ----
$domainDN = (Get-ADDomain).DistinguishedName
$adminOU = "OU=Admin,$domainDN"
Write-Host "Using search base: $adminOU" -ForegroundColor Cyan

# Define the OU name suffixes for each tier (prefix T0- or T1-)
$tierOUSuffixes = @('T0-Accounts', 'T0-Admin Workstations', 'T0-Groups', 'T0-Servers', 'T0-Service Accounts')
$groupsOUNames = @('T0-Groups', 'T1-Groups')

# ---- Helper: get OUs by prefix and suffix list ----
function Get-OUByPrefix {
    param(
        [string]$prefix,          # "T0-" or "T1-"
        [string[]]$suffixes,
        [string]$searchBase
    )
    $dnList = @()
    foreach ($suffix in $suffixes) {
        $ouName = "$prefix$suffix"
        $ous = Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $searchBase -ErrorAction SilentlyContinue
        if ($ous) {
            foreach ($ou in $ous) {
                $dnList += $ou.DistinguishedName
                Write-Host "  Found OU: $($ou.DistinguishedName)" -ForegroundColor Gray
            }
        } else {
            Write-Warning "OU '$ouName' not found under '$searchBase'"
        }
    }
    return $dnList
}

# ---- Locate all tier OUs under Admin ----
$tierOUDNs = @{}
$tierOUDNs[0] = Get-OUByPrefix -prefix "T0-" -suffixes $tierOUSuffixes -searchBase $adminOU
$tierOUDNs[1] = Get-OUByPrefix -prefix "T1-" -suffixes $tierOUSuffixes -searchBase $adminOU

if ($tierOUDNs[0].Count -eq 0) { Write-Warning "No Tier 0 OUs found." }
if ($tierOUDNs[1].Count -eq 0) { Write-Warning "No Tier 1 OUs found." }

# ---- Build object tier map (users, groups, computers under each tier OU) ----
$objectTierMap = @{}
Write-Host "`nEnumerating objects under Tier OUs..." -ForegroundColor Yellow
foreach ($tier in $tierOUDNs.Keys) {
    foreach ($ouDN in $tierOUDNs[$tier]) {
        Write-Host "  Scanning OU: $ouDN (Tier $tier)" -ForegroundColor Gray
        $objects = Get-ADObject -Filter "ObjectClass -eq 'user' -or ObjectClass -eq 'group' -or ObjectClass -eq 'computer'" `
            -SearchBase $ouDN -SearchScope Subtree -Properties CanonicalName, ObjectClass, DistinguishedName -ErrorAction SilentlyContinue
        foreach ($obj in $objects) {
            if (-not $objectTierMap.ContainsKey($obj.DistinguishedName)) {
                $objectTierMap[$obj.DistinguishedName] = @{
                    Tier = $tier
                    ObjectClass = $obj.ObjectClass
                }
            }
        }
    }
}
Write-Host "Collected $($objectTierMap.Count) objects under tier OUs." -ForegroundColor Green

# ---- Locate T0-Groups and T1-Groups OUs under Admin ----
$groupsOUSearch = @()
foreach ($groupsName in $groupsOUNames) {
    $foundOUs = Get-ADOrganizationalUnit -Filter "Name -eq '$groupsName'" -SearchBase $adminOU -ErrorAction SilentlyContinue
    if ($foundOUs) {
        foreach ($ou in $foundOUs) {
            # Determine tier from the name prefix
            if ($groupsName -like "T0-*") { $tier = 0 }
            elseif ($groupsName -like "T1-*") { $tier = 1 }
            else { $tier = $null }
            if ($tier -ne $null) {
                $groupsOUSearch += [PSCustomObject]@{
                    DN   = $ou.DistinguishedName
                    Tier = $tier
                }
                Write-Host "  Found groups OU: $($ou.DistinguishedName) (Tier $tier)" -ForegroundColor Gray
            }
        }
    } else {
        Write-Warning "OU '$groupsName' not found under '$adminOU'"
    }
}

if ($groupsOUSearch.Count -eq 0) {
    Write-Error "No T0-Groups or T1-Groups OUs found. Audit cannot proceed."
    exit
}

# ---- Retrieve all groups from those OUs (including sub-OUs) ----
$groupsToAudit = @()
foreach ($groupsOU in $groupsOUSearch) {
    Write-Host "`nFetching groups from $($groupsOU.DN) (Tier $($groupsOU.Tier))..." -ForegroundColor Gray
    $groups = Get-ADGroup -Filter * -SearchBase $groupsOU.DN -SearchScope Subtree -Properties DistinguishedName, Name -ErrorAction SilentlyContinue
    foreach ($g in $groups) {
        $tierInfo = $objectTierMap[$g.DistinguishedName]
        $groupsToAudit += [PSCustomObject]@{
            DN   = $g.DistinguishedName
            Name = $g.Name
            Tier = if ($tierInfo) { $tierInfo.Tier } else { $groupsOU.Tier }
        }
    }
}
Write-Host "`nFound $($groupsToAudit.Count) groups to audit." -ForegroundColor Green

# ---- Audit each group's members ----
$violations = @()
$totalGroups = $groupsToAudit.Count
$processed = 0

foreach ($group in $groupsToAudit) {
    $processed++
    Write-Progress -Activity "Auditing group memberships" -Status "Processing $($group.Name)" -PercentComplete (($processed / $totalGroups) * 100)
    
    $members = Get-ADGroupMember -Identity $group.DN -ErrorAction SilentlyContinue
    foreach ($member in $members) {
        $memberDN = $member.DistinguishedName
        $memberTierInfo = $objectTierMap[$memberDN]
        if ($memberTierInfo) {
            $memberTier = $memberTierInfo.Tier
            $memberType = $memberTierInfo.ObjectClass
        } else {
            $memberTier = $null
            $memberType = $member.objectClass   # user/group/computer
        }

        if ($memberTier -ne $null -and $memberTier -ne $group.Tier) {
            $violations += [PSCustomObject]@{
                GroupName        = $group.Name
                GroupTier        = $group.Tier
                GroupDN          = $group.DN
                MemberName       = $member.Name
                MemberType       = $memberType
                MemberDN         = $memberDN
                MemberTier       = $memberTier
                Issue            = "Cross-tier membership (Tier $($group.Tier) group contains Tier $memberTier object)"
            }
        }
    }
}
Write-Progress -Activity "Auditing group memberships" -Completed

# ---- Generate Reports ----
$reportDate = Get-Date -Format "yyyy-MM-dd_HHmm"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$reportFolder = Join-Path $desktopPath "Tiered_Audit_$reportDate"
New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null

# CSV
$csvPath = Join-Path $reportFolder "Violations.csv"
$violations | Export-Csv -Path $csvPath -NoTypeInformation

# HTML
$htmlPath = Join-Path $reportFolder "Violations.html"
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Tiered AD Audit Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2c3e50; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #34495e; color: white; padding: 8px; text-align: left; }
        td { padding: 6px; border: 1px solid #ddd; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .summary { background-color: #ecf0f1; padding: 10px; border-radius: 5px; }
        .violation { background-color: #e74c3c; color: white; padding: 3px 8px; border-radius: 3px; }
        .tier0 { color: #2980b9; }
        .tier1 { color: #e67e22; }
    </style>
</head>
<body>
    <h1>Tiered Active Directory Audit Report</h1>
    <div class="summary">
        <p><strong>Audit Run:</strong> $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <p><strong>Total Groups Audited:</strong> $($groupsToAudit.Count)</p>
        <p><strong>Cross-tier Violations Found:</strong> <span class="violation">$($violations.Count)</span></p>
    </div>
"@

$htmlFooter = @"
</body>
</html>
"@

if ($violations.Count -gt 0) {
    $tableRows = $violations | ForEach-Object {
        @"
        <tr>
            <td>$($_.GroupName)</td>
            <td class="tier$($_.GroupTier)">Tier $($_.GroupTier)</td>
            <td>$($_.MemberName)</td>
            <td>$($_.MemberType)</td>
            <td class="tier$($_.MemberTier)">Tier $($_.MemberTier)</td>
            <td>$($_.Issue)</td>
        </tr>
"@
    }
    $tableHtml = @"
    <table>
        <thead>
            <tr>
                <th>Group Name</th>
                <th>Group Tier</th>
                <th>Member Name</th>
                <th>Member Type</th>
                <th>Member Tier</th>
                <th>Issue</th>
            </tr>
        </thead>
        <tbody>
$tableRows
        </tbody>
    </table>
"@
} else {
    $tableHtml = "<p style='color:green; font-weight:bold;'>No cross-tier violations found.</p>"
}

$htmlContent = $htmlHeader + $tableHtml + $htmlFooter
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8

# ---- Final Output ----
$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host "`nAudit completed in $($duration.TotalSeconds) seconds." -ForegroundColor Cyan
Write-Host "Reports saved to: $reportFolder" -ForegroundColor Green
Write-Host "  CSV: $csvPath" -ForegroundColor Green
Write-Host "  HTML: $htmlPath" -ForegroundColor Green

# Optionally open the HTML report
# Start-Process $htmlPath
