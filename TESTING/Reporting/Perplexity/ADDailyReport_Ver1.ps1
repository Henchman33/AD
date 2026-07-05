param(
    [string]$OutputFolder = "C:\Reports\AD",
    [int]$DaysBack = 7,
    [string[]]$PrivilegedGroups = @(
        "Domain Admins",
        "Enterprise Admins",
        "Schema Admins",
        "Administrators",
        "Account Operators",
        "Server Operators",
        "Backup Operators",
        "Print Operators"
    )
)

Import-Module ActiveDirectory -ErrorAction Stop

$start = (Get-Date).AddDays(-$DaysBack)
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$reportDir = Join-Path $OutputFolder $stamp
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$dcList = Get-ADDomainController -Filter * | Sort-Object HostName
$forest = Get-ADForest
$domain = Get-ADDomain

$repAdminSummary = & repadmin /replsummary 2>&1
$repAdminShowRepl = & repadmin /showrepl * /csv 2>&1

$dcHealth = foreach ($dc in $dcList) {
    $ping = Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue
    $failures = Get-ADReplicationFailure -Target $dc.HostName -ErrorAction SilentlyContinue
    $partners = Get-ADReplicationPartnerMetadata -Target $dc.HostName -Scope Server -ErrorAction SilentlyContinue

    [pscustomobject]@{
        DomainController = $dc.HostName
        Site             = $dc.Site
        IPv4Address      = $dc.IPv4Address
        Reachable        = $ping
        ReplicationFailures = @($failures).Count
        LastPartnerUpdate = ($partners | Sort-Object LastReplicationSuccess -Descending | Select-Object -First 1).LastReplicationSuccess
    }
}

$timeSkew = foreach ($dc in $dcList) {
    try {
        $w32tm = & w32tm /monitor /computers:$($dc.HostName) 2>&1
        [pscustomobject]@{
            DomainController = $dc.HostName
            RawOutput = ($w32tm -join "`n")
        }
    } catch {}
}

$eventIds = 4727,4728,4729,4730,4735,4737,4754,4755,4756,4757,4758
$groupChanges = foreach ($dc in $dcList) {
    try {
        Get-WinEvent -ComputerName $dc.HostName -FilterHashtable @{
            LogName   = 'Security'
            Id        = $eventIds
            StartTime = $start
        } -ErrorAction Stop | ForEach-Object {
            $xml = [xml]$_.ToXml()
            $data = @{}
            foreach ($node in $xml.Event.EventData.Data) { $data[$node.Name] = $node.'#text' }

            [pscustomobject]@{
                TimeCreated = $_.TimeCreated
                DC          = $dc.HostName
                EventId     = $_.Id
                GroupName   = $data['TargetUserName']
                Subject     = $data['SubjectUserName']
                Member      = $data['MemberName']
                MemberSid   = $data['MemberSid']
                ChangeType  = switch ($_.Id) {
                    4727 { 'Global group created' }
                    4728 { 'Member added to global group' }
                    4729 { 'Member removed from global group' }
                    4730 { 'Global group deleted' }
                    4735 { 'Global group changed' }
                    4737 { 'Global group changed' }
                    4754 { 'Universal group created' }
                    4755 { 'Universal group changed' }
                    4756 { 'Member added to universal group' }
                    4757 { 'Member removed from universal group' }
                    4758 { 'Universal group deleted' }
                    default { 'Other' }
                }
            }
        }
    } catch {}
}

$privChanges = $groupChanges | Where-Object {
    $g = $_.GroupName
    $PrivilegedGroups -contains $g -or ($g -match 'Tier\s*0|Tier0|Privileged')
} | Sort-Object TimeCreated -Descending

$summary = [pscustomobject]@{
    GeneratedOn            = Get-Date
    Forest                 = $forest.Name
    Domain                 = $domain.DNSRoot
    DomainControllers      = $dcList.Count
    ReplicationErrorDCs    = @($dcHealth | Where-Object { $_.ReplicationFailures -gt 0 }).Count
    GroupChangesLast7Days  = @($groupChanges).Count
    PrivGroupChangesLast7D = @($privChanges).Count
}

$summary | Export-Csv (Join-Path $reportDir "summary.csv") -NoTypeInformation
$dcHealth | Export-Csv (Join-Path $reportDir "dc-health.csv") -NoTypeInformation
$groupChanges | Sort-Object TimeCreated -Descending | Export-Csv (Join-Path $reportDir "group-changes-last7days.csv") -NoTypeInformation
$privChanges | Export-Csv (Join-Path $reportDir "privileged-group-changes.csv") -NoTypeInformation
$timeSkew | Export-Csv (Join-Path $reportDir "time-monitor.csv") -NoTypeInformation
$repAdminSummary | Out-File (Join-Path $reportDir "repadmin-replsummary.txt")
$repAdminShowRepl | Out-File (Join-Path $reportDir "repadmin-showrepl.csv")

$body = @"
AD Daily Report

Forest: $($summary.Forest)
Domain: $($summary.Domain)
DCs: $($summary.DomainControllers)
DCs with replication failures: $($summary.ReplicationErrorDCs)
Group changes last $DaysBack days: $($summary.GroupChangesLast7Days)
Privileged group changes last $DaysBack days: $($summary.PrivGroupChangesLast7D)
Report folder: $reportDir
"@

$body | Out-File (Join-Path $reportDir "email-summary.txt")
