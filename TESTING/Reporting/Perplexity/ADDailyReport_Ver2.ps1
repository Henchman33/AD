param(
    [int]$DaysBack = 7,
    [string[]]$Tier0Groups = @(
        'Domain Admins',
        'Enterprise Admins',
        'Schema Admins',
        'Administrators'
    ),
    [string]$OutputFolder = "C:\Reports\Tier0",
    [string]$SmtpServer = "smtp.contoso.com",
    [int]$SmtpPort = 25,
    [string]$MailFrom = "ad-report@contoso.com",
    [string[]]$MailTo = @("security@contoso.com"),
    [string]$MailSubjectPrefix = "Tier 0 AD Change Report"
)

Import-Module ActiveDirectory -ErrorAction Stop

$start = (Get-Date).AddDays(-$DaysBack)
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$reportDir = Join-Path $OutputFolder $stamp
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null

$eventIds = 4727,4728,4729,4730,4731,4732,4733,4734,4735,4737,4754,4755,4756,4757,4758

function Convert-AdEvent {
    param($Event)

    $xml = [xml]$Event.ToXml()
    $data = @{}
    foreach ($node in $xml.Event.EventData.Data) {
        $data[$node.Name] = $node.'#text'
    }

    [pscustomobject]@{
        TimeCreated = $Event.TimeCreated
        DC          = $Event.MachineName
        EventId     = $Event.Id
        GroupName   = $data['TargetUserName']
        Actor       = $data['SubjectUserName']
        Member      = $data['MemberName']
        ChangeType  = switch ($Event.Id) {
            4727 { 'Global group created' }
            4728 { 'Member added to global group' }
            4729 { 'Member removed from global group' }
            4730 { 'Global group deleted' }
            4731 { 'Local group created' }
            4732 { 'Member added to local group' }
            4733 { 'Member removed from local group' }
            4734 { 'Local group deleted' }
            4735 { 'Local group changed' }
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

$events = Get-WinEvent -FilterHashtable @{
    LogName   = 'Security'
    Id        = $eventIds
    StartTime = $start
} -ErrorAction Stop | ForEach-Object { Convert-AdEvent $_ }

$tier0Changes = $events | Where-Object { $_.GroupName -in $Tier0Groups } | Sort-Object TimeCreated -Descending

Write-Host ""
Write-Host "Tier 0 Active Directory Changes" -ForegroundColor Cyan
Write-Host "Period: $start to $(Get-Date)"
Write-Host "Groups: $($Tier0Groups -join ', ')"
Write-Host "Found: $($tier0Changes.Count) event(s)"
Write-Host ""

if ($tier0Changes.Count -gt 0) {
    $tier0Changes |
        Select-Object TimeCreated, DC, EventId, GroupName, Actor, Member, ChangeType |
        Format-Table -AutoSize | Out-String | Write-Host
} else {
    Write-Host "No Tier 0 group changes found in the selected period."
}

$summary = [pscustomobject]@{
    GeneratedOn       = Get-Date
    WindowStart       = $start
    WindowEnd         = Get-Date
    Tier0GroupCount   = $Tier0Groups.Count
    EventCount        = $tier0Changes.Count
    AddRemoveCount    = @($tier0Changes | Where-Object { $_.EventId -in 4728,4729,4732,4733,4756,4757 }).Count
    CreateDeleteCount = @($tier0Changes | Where-Object { $_.EventId -in 4727,4730,4734,4754,4758 }).Count
}

$csvPath = Join-Path $reportDir "tier0-changes.csv"
$htmlPath = Join-Path $reportDir "tier0-changes.html"

$tier0Changes | Export-Csv $csvPath -NoTypeInformation

$style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 10pt; color: #222; }
h1, h2, h3 { color: #1f4e79; }
table { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
th, td { border: 1px solid #999; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background: #d9e2f3; }
tr:nth-child(even) { background: #f8f9fb; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; background: #eef4ff; }
</style>
"@

$summaryHtml = @"
<h1>Tier 0 Active Directory Change Report</h1>
<p><span class="badge">Generated</span> $($summary.GeneratedOn)</p>
<p><span class="badge">Window</span> $($summary.WindowStart) to $($summary.WindowEnd)</p>
<p><span class="badge">Tier 0 Groups</span> $($summary.Tier0GroupCount)</p>
<p><span class="badge">Events Found</span> $($summary.EventCount)</p>
<p><span class="badge">Add/Remove</span> $($summary.AddRemoveCount)</p>
<p><span class="badge">Create/Delete</span> $($summary.CreateDeleteCount)</p>
<h2>Events</h2>
"@

$body = $tier0Changes |
    Select-Object TimeCreated, DC, EventId, GroupName, Actor, Member, ChangeType |
    ConvertTo-Html -Head $style -PreContent $summaryHtml -Title "Tier 0 Active Directory Change Report"

$body | Out-File $htmlPath -Encoding UTF8

$mailParams = @{
    From       = $MailFrom
    To         = $MailTo -join ','
    Subject    = "$MailSubjectPrefix - $((Get-Date).ToString('yyyy-MM-dd'))"
    Body       = (Get-Content $htmlPath -Raw)
    BodyAsHtml  = $true
    SmtpServer = $SmtpServer
    Port       = $SmtpPort
}

Send-MailMessage @mailParams
Write-Host "HTML report saved to: $htmlPath"
Write-Host "CSV saved to: $csvPath"
Write-Host "Email sent to: $($MailTo -join ', ')"
