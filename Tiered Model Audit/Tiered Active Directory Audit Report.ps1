<#
.SYNOPSIS
    Audits cross-tier memberships recursively across ALL groups under Tier 0 and Tier 1 OUs.
.DESCRIPTION
    Finds the root "Tier 0" and "Tier 1" OUs under OU=Admin.
    Scans all users, groups, and computers under these roots to build tier assignments.
    Audits EVERY group under these roots for cross-tier members, including nested group memberships.
    Outputs CSV and HTML reports to the user's Desktop.
.NOTES
    Requires: ActiveDirectory module, read access to Admin OU.
#>

#Requires -Modules ActiveDirectory

# ---- Configuration ----
$domainDN = (Get-ADDomain).DistinguishedName
$adminOU = "OU=Admin,$domainDN"
Write-Host "Using search base: $adminOU" -ForegroundColor Cyan

# ---- Locate the Tier 0 and Tier 1 root OUs ----
$tier0Root = Get-ADOrganizationalUnit -Filter "Name -eq 'Tier 0'" -SearchBase $adminOU -ErrorAction Stop
$tier1Root = Get-ADOrganizationalUnit -Filter "Name -eq 'Tier 1'" -SearchBase $adminOU -ErrorAction Stop

if (-not $tier0Root -or -not $tier1Root) {
    Write-Error "Could not find 'Tier 0' or 'Tier 1' OUs under '$adminOU'. Exiting."
    exit
}

$tierRoots = @(
    @{ DN = $tier0Root.DistinguishedName; Tier = 0 },
    @{ DN = $tier1Root.DistinguishedName; Tier = 1 }
)

Write-Host "Found Tier 0 root: $($tier0Root.DistinguishedName)" -ForegroundColor Green
Write-Host "Found Tier 1 root: $($tier1Root.DistinguishedName)" -ForegroundColor Green

# ---- Build object tier map (all users, groups, computers under both roots) ----
$objectTierMap = @{}
Write-Host "`nEnumerating all objects under Tier roots..." -ForegroundColor Yellow

foreach ($tier in $tierRoots) {
    Write-Host "  Scanning: $($tier.DN) (Tier $($tier.Tier))" -ForegroundColor Gray
    $objects = Get-ADObject -Filter "ObjectClass -eq 'user' -or ObjectClass -eq 'group' -or ObjectClass -eq 'computer'" `
        -SearchBase $tier.DN -SearchScope Subtree -Properties ObjectClass, DistinguishedName -ErrorAction SilentlyContinue
    
    foreach ($obj in $objects) {
        if (-not $objectTierMap.ContainsKey($obj.DistinguishedName)) {
            $objectTierMap[$obj.DistinguishedName] = @{
                Tier = $tier.Tier
                ObjectClass = $obj.ObjectClass
            }
        }
    }
}
Write-Host "Collected $($objectTierMap.Count) objects under Tier roots." -ForegroundColor Green

# ---- Get ALL groups under both Tier roots ----
$groupsToAudit = @()
Write-Host "`nFetching all groups under Tier roots..." -ForegroundColor Yellow

foreach ($tier in $tierRoots) {
    Write-Host "  Fetching groups from: $($tier.DN) (Tier $($tier.Tier))" -ForegroundColor Gray
    $groups = Get-ADGroup -Filter * -SearchBase $tier.DN -SearchScope Subtree -Properties DistinguishedName, Name -ErrorAction SilentlyContinue
    
    foreach ($g in $groups) {
        $tierInfo = $objectTierMap[$g.DistinguishedName]
        $groupsToAudit += [PSCustomObject]@{
            DN   = $g.DistinguishedName
            Name = $g.Name
            Tier = if ($tierInfo) { $tierInfo.Tier } else { $tier.Tier }
        }
    }
}
Write-Host "Found $($groupsToAudit.Count) groups to audit." -ForegroundColor Green

# ---- Audit each group's members recursively ----
$violations = @()
$totalGroups = $groupsToAudit.Count
$processed = 0

Write-Host "`nAuditing group memberships recursively (nested groups included)..." -ForegroundColor Yellow

foreach ($group in $groupsToAudit) {
    $processed++
    Write-Progress -Activity "Auditing group memberships (recursive)" -Status "Processing $($group.Name)" -PercentComplete (($processed / $totalGroups) * 100)
    
    # Get all members recursively (users, groups, computers)
    $members = Get-ADGroupMember -Identity $group.DN -Recursive -ErrorAction SilentlyContinue
    
    # Use a hashtable to avoid duplicate violation entries for the same member
    $reportedMembers = @{}
    
    foreach ($member in $members) {
        $memberDN = $member.DistinguishedName
        # Skip if already reported for this group (avoids duplicates from multiple paths)
        if ($reportedMembers.ContainsKey($memberDN)) { continue }
        $reportedMembers[$memberDN] = $true
        
        $memberTierInfo = $objectTierMap[$memberDN]
        if ($memberTierInfo) {
            $memberTier = $memberTierInfo.Tier
            $memberType = $memberTierInfo.ObjectClass
        } else {
            # Member is outside the tiered structure – ignore (or optionally flag? We'll ignore)
            continue
        }

        # Cross-tier violation
        if ($memberTier -ne $group.Tier) {
            $violations += [PSCustomObject]@{
                GroupName        = $group.Name
                GroupTier        = $group.Tier
                GroupDN          = $group.DN
                MemberName       = $member.Name
                MemberType       = $memberType
                MemberDN         = $memberDN
                MemberTier       = $memberTier
                Issue            = "Cross-tier membership (recursive) – Tier $($group.Tier) group contains Tier $memberTier object"
            }
        }
    }
}
Write-Progress -Activity "Auditing group memberships (recursive)" -Completed

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
    <title>Tiered AD Audit Report (Recursive)</title>
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
        .no-violations { color: green; font-weight: bold; }
    </style>
</head>
<body>
    <h1>Tiered Active Directory Audit Report (Recursive)</h1>
    <div class="summary">
        <p><strong>Audit Run:</strong> $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <p><strong>Total Groups Audited:</strong> $($groupsToAudit.Count)</p>
        <p><strong>Cross-tier Violations Found:</strong> <span class="violation">$($violations.Count)</span></p>
        <p><em>Note: Nested group memberships are included recursively.</em></p>
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
    $tableHtml = "<p class='no-violations'>✅ No cross-tier violations found (recursive check).</p>"
}

$htmlContent = $htmlHeader + $tableHtml + $htmlFooter
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8

# ---- Final Output ----
$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host "`n✅ Audit completed in $($duration.TotalSeconds) seconds." -ForegroundColor Cyan
Write-Host "Reports saved to: $reportFolder" -ForegroundColor Green
Write-Host "  CSV: $csvPath" -ForegroundColor Green
Write-Host "  HTML: $htmlPath" -ForegroundColor Green

# Optionally open the HTML report
# Start-Process $htmlPath
