<#
.SYNOPSIS
    Tiered AD group-change report.

.DESCRIPTION
    - Monitors Tier 0 and Tier 1 privileged/operator groups.
    - Collects Security Group Management events from domain controllers.
    - Displays events in the console.
    - Writes an HTML and CSV report.
    - Opens the completed HTML report in the default browser.
    - Sends an HTML report by SMTP when SendEmail is set to $true.
    - Generates a layout/status report even if AD services, RSAT,
      security-log access, or DC discovery are unavailable.
#>

[CmdletBinding()]
param(
    [int]$DaysBack = 7,

    # A single DC can be supplied for testing, such as:
    # -DomainControllers @('DC01.contoso.com')
    [string[]]$DomainControllers = @(),

    [string]$OutputFolder = "C:\Reports\Tiered-AD",

    # Set to $true only after configuring the SMTP settings below.
    [bool]$SendEmail = $false,

    [string]$SmtpServer = "smtp.contoso.com",
    [int]$SmtpPort = 25,
    [bool]$UseSsl = $false,

    [string]$MailFrom = "ad-report@contoso.com",
    [string[]]$MailTo = @(
        "security@contoso.com"
    ),

    [string]$MailSubjectPrefix = "Tiered AD Group Change Report"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------
# Monitored groups
# Add or remove groups here. Keys must exactly match the AD group Name.
# ---------------------------------------------------------------------
$MonitoredGroups = [ordered]@{
    "Domain Admins"       = "Tier 0"
    "Enterprise Admins"   = "Tier 0"
    "Schema Admins"       = "Tier 0"
    "Administrators"      = "Tier 0"
    "Tier 0 Operators"    = "Tier 0"
    "Tier 1 Operators"    = "Tier 1"
}

$MonitoredGroupNames = @($MonitoredGroups.Keys)

# Security group management events:
# 4728/4729 = global group add/remove
# 4732/4733 = local group add/remove
# 4756/4757 = universal group add/remove
# Other listed IDs capture group create/change/delete activity.
$EventIds = @(
    4727, 4728, 4729, 4730,
    4731, 4732, 4733, 4734,
    4735, 4737,
    4754, 4755, 4756, 4757, 4758
)

$RunStart = Get-Date
$WindowStart = $RunStart.AddDays(-$DaysBack)
$TimeStamp = $RunStart.ToString("yyyyMMdd-HHmmss")
$ReportFolder = Join-Path -Path $OutputFolder -ChildPath $TimeStamp
$HtmlPath = Join-Path -Path $ReportFolder -ChildPath "Tiered-AD-Group-Change-Report.html"
$CsvPath = Join-Path -Path $ReportFolder -ChildPath "Tiered-AD-Group-Change-Report.csv"

$CollectionErrors = New-Object System.Collections.Generic.List[string]
$AllEvents = New-Object System.Collections.Generic.List[object]
$QueriedDCs = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------
function ConvertTo-HtmlEncoded {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-ChangeType {
    param(
        [int]$EventId
    )

    switch ($EventId) {
        4727 { return "Global group created" }
        4728 { return "Member added to global group" }
        4729 { return "Member removed from global group" }
        4730 { return "Global group deleted" }
        4731 { return "Local group created" }
        4732 { return "Member added to local group" }
        4733 { return "Member removed from local group" }
        4734 { return "Local group deleted" }
        4735 { return "Local group changed" }
        4737 { return "Global group changed" }
        4754 { return "Universal group created" }
        4755 { return "Universal group changed" }
        4756 { return "Member added to universal group" }
        4757 { return "Member removed from universal group" }
        4758 { return "Universal group deleted" }
        default { return "Other group-management event" }
    }
}

function Convert-GroupEvent {
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Eventing.Reader.EventRecord]$Event,

        [Parameter(Mandatory)]
        [string]$SourceDC
    )

    $Xml = [xml]$Event.ToXml()
    $Data = @{}

    foreach ($Node in $Xml.Event.EventData.Data) {
        $Data[$Node.Name] = $Node.'#text'
    }

    $GroupName = $Data["TargetUserName"]
    $Tier = $MonitoredGroups[$GroupName]

    [pscustomobject]@{
        TimeCreated = $Event.TimeCreated
        Tier        = $Tier
        DC          = $SourceDC
        EventId     = $Event.Id
        GroupName   = $GroupName
        Actor       = $Data["SubjectUserName"]
        ActorDomain = $Data["SubjectDomainName"]
        Member      = $Data["MemberName"]
        MemberSid   = $Data["MemberSid"]
        ChangeType  = Get-ChangeType -EventId $Event.Id
    }
}

function New-EmptyReportRow {
    param(
        [string]$Message
    )

    return @"
<tr>
    <td colspan="8" class="empty">$([System.Net.WebUtility]::HtmlEncode($Message))</td>
</tr>
"@
}

# ---------------------------------------------------------------------
# Ensure report output exists before querying AD.
# This guarantees an HTML report can be built even if collection fails.
# ---------------------------------------------------------------------
try {
    New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null
}
catch {
    throw "Could not create report folder '$ReportFolder'. $($_.Exception.Message)"
}

# ---------------------------------------------------------------------
# Discover domain controllers, unless supplied manually.
# ---------------------------------------------------------------------
if ($DomainControllers.Count -eq 0) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $DomainControllers = @(
            Get-ADDomainController -Filter * -ErrorAction Stop |
            Sort-Object HostName |
            Select-Object -ExpandProperty HostName
        )

        if ($DomainControllers.Count -eq 0) {
            $CollectionErrors.Add("Active Directory returned no domain controllers.")
        }
    }
    catch {
        $CollectionErrors.Add(
            "Unable to discover domain controllers through the ActiveDirectory module. " +
            "The report layout will still be generated. Error: $($_.Exception.Message)"
        )
    }
}

if ($DomainControllers.Count -eq 0) {
    $CollectionErrors.Add(
        "No domain controllers were queried. To test collection without AD discovery, " +
        "run the script with -DomainControllers @('YourDC.FQDN')."
    )
}

# ---------------------------------------------------------------------
# Query Security logs from each domain controller.
# ---------------------------------------------------------------------
foreach ($DC in $DomainControllers) {
    try {
        Write-Host "Querying Security log on $DC..." -ForegroundColor Cyan

        $RawEvents = Get-WinEvent -ComputerName $DC -FilterHashtable @{
            LogName   = "Security"
            Id        = $EventIds
            StartTime = $WindowStart
        } -ErrorAction Stop

        $QueriedDCs.Add($DC)

        foreach ($Event in $RawEvents) {
            try {
                $ConvertedEvent = Convert-GroupEvent -Event $Event -SourceDC $DC

                if ($ConvertedEvent.GroupName -in $MonitoredGroupNames) {
                    $AllEvents.Add($ConvertedEvent)
                }
            }
            catch {
                $CollectionErrors.Add(
                    "Could not parse Security event $($Event.Id) from $DC. " +
                    "Error: $($_.Exception.Message)"
                )
            }
        }
    }
    catch {
        $CollectionErrors.Add(
            "Could not query the Security log on $DC. " +
            "Verify DC connectivity, Event Log service, firewall access, and permission to read the Security log. " +
            "Error: $($_.Exception.Message)"
        )
    }
}

$Changes = @(
    $AllEvents |
    Sort-Object TimeCreated -Descending
)

$MembershipChanges = @(
    $Changes | Where-Object { $_.EventId -in 4728, 4729, 4732, 4733, 4756, 4757 }
)

$Tier0Changes = @(
    $Changes | Where-Object { $_.Tier -eq "Tier 0" }
)

$Tier1Changes = @(
    $Changes | Where-Object { $_.Tier -eq "Tier 1" }
)

$RunEnd = Get-Date

# ---------------------------------------------------------------------
# Console output
# ---------------------------------------------------------------------
Write-Host ""
Write-Host "Tiered Active Directory Group Change Report" -ForegroundColor Cyan
Write-Host "Report window: $WindowStart through $RunEnd"
Write-Host "Monitored groups: $($MonitoredGroupNames -join ', ')"
Write-Host "Domain controllers queried: $($QueriedDCs.Count)"
Write-Host "Changes found: $($Changes.Count)"
Write-Host "Tier 0 changes: $($Tier0Changes.Count)"
Write-Host "Tier 1 changes: $($Tier1Changes.Count)"
Write-Host ""

if ($Changes.Count -gt 0) {
    $Changes |
        Select-Object TimeCreated, Tier, DC, EventId, GroupName, Actor, Member, ChangeType |
        Format-Table -AutoSize |
        Out-String |
        Write-Host
}
else {
    Write-Host "No matching group changes were found, or no events could be collected." -ForegroundColor Yellow
}

if ($CollectionErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "Collection warnings/errors:" -ForegroundColor Yellow

    foreach ($CollectionError in $CollectionErrors) {
        Write-Host " - $CollectionError" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------
# CSV report
# ---------------------------------------------------------------------
if ($Changes.Count -gt 0) {
    $Changes |
        Select-Object TimeCreated, Tier, DC, EventId, GroupName, Actor, ActorDomain, Member, MemberSid, ChangeType |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
}
else {
    @(
        [pscustomobject]@{
            TimeCreated = ""
            Tier        = ""
            DC          = ""
            EventId     = ""
            GroupName   = ""
            Actor       = ""
            ActorDomain = ""
            Member      = ""
            MemberSid   = ""
            ChangeType  = "No matching group changes found, or collection was unavailable."
        }
    ) |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
}

# ---------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------
$StatusClass = if ($CollectionErrors.Count -eq 0) { "status-good" } else { "status-warning" }
$StatusText = if ($CollectionErrors.Count -eq 0) {
    "Collection completed without reported errors."
}
else {
    "Collection completed with $($CollectionErrors.Count) warning(s) or error(s). Review the details below."
}

$EventRows = if ($Changes.Count -gt 0) {
    foreach ($Change in $Changes) {
        $TierClass = if ($Change.Tier -eq "Tier 0") { "tier0" } else { "tier1" }

        @"
<tr>
    <td>$((ConvertTo-HtmlEncoded $Change.TimeCreated))</td>
    <td><span class="$TierClass">$((ConvertTo-HtmlEncoded $Change.Tier))</span></td>
    <td>$((ConvertTo-HtmlEncoded $Change.DC))</td>
    <td>$((ConvertTo-HtmlEncoded $Change.EventId))</td>
    <td>$((ConvertTo-HtmlEncoded $Change.GroupName))</td>
    <td>$((ConvertTo-HtmlEncoded $Change.Actor))</td>
    <td>$((ConvertTo-HtmlEncoded $Change.Member))</td>
    <td>$((ConvertTo-HtmlEncoded $Change.ChangeType))</td>
</tr>
"@
    }
}
else {
    New-EmptyReportRow -Message "No monitored group changes were found in the selected time window, or event collection was unavailable."
}

$ErrorRows = if ($CollectionErrors.Count -gt 0) {
    foreach ($CollectionError in $CollectionErrors) {
        "<li>$([System.Net.WebUtility]::HtmlEncode($CollectionError))</li>"
    }
}
else {
    "<li>No collection errors reported.</li>"
}

$MonitoredGroupRows = foreach ($GroupName in $MonitoredGroupNames) {
    $Tier = $MonitoredGroups[$GroupName]
    $TierClass = if ($Tier -eq "Tier 0") { "tier0" } else { "tier1" }

    @"
<tr>
    <td>$([System.Net.WebUtility]::HtmlEncode($GroupName))</td>
    <td><span class="$TierClass">$([System.Net.WebUtility]::HtmlEncode($Tier))</span></td>
</tr>
"@
}

$Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Tiered AD Group Change Report</title>
<style>
    body {
        background: #f4f7fb;
        color: #1f2937;
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 14px;
        line-height: 1.45;
        margin: 0;
        padding: 24px;
    }

    .container {
        background: #ffffff;
        border: 1px solid #d7deea;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.08);
        margin: auto;
        max-width: 1500px;
        padding: 28px;
    }

    h1 {
        color: #0f3d6e;
        font-size: 26px;
        margin: 0 0 8px 0;
    }

    h2 {
        border-bottom: 2px solid #d9e6f5;
        color: #0f3d6e;
        font-size: 19px;
        margin-top: 30px;
        padding-bottom: 7px;
    }

    .subtitle {
        color: #5f6b7a;
        margin: 0 0 18px 0;
    }

    .status {
        border-radius: 6px;
        font-weight: 600;
        margin: 16px 0;
        padding: 12px 14px;
    }

    .status-good {
        background: #eaf7ef;
        border: 1px solid #b8e0c4;
        color: #155724;
    }

    .status-warning {
        background: #fff4df;
        border: 1px solid #efd18f;
        color: #7a4f01;
    }

    .cards {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin: 18px 0 24px 0;
    }

    .card {
        background: #f8fbff;
        border: 1px solid #d8e5f3;
        border-radius: 7px;
        min-width: 150px;
        padding: 12px 16px;
    }

    .card-label {
        color: #5f6b7a;
        display: block;
        font-size: 12px;
        text-transform: uppercase;
    }

    .card-value {
        color: #0f3d6e;
        display: block;
        font-size: 23px;
        font-weight: 700;
        margin-top: 2px;
    }

    table {
        border-collapse: collapse;
        margin-top: 12px;
        width: 100%;
    }

    th, td {
        border: 1px solid #d6deea;
        padding: 9px 10px;
        text-align: left;
        vertical-align: top;
    }

    th {
        background: #0f3d6e;
        color: #ffffff;
        font-weight: 600;
    }

    tr:nth-child(even) {
        background: #f8fbff;
    }

    .tier0, .tier1 {
        border-radius: 12px;
        display: inline-block;
        font-size: 12px;
        font-weight: 700;
        padding: 3px 9px;
        white-space: nowrap;
    }

    .tier0 {
        background: #ffe2e5;
        color: #9b1c31;
    }

    .tier1 {
        background: #e0efff;
        color: #095b9c;
    }

    .empty {
        color: #5f6b7a;
        font-style: italic;
        padding: 22px;
        text-align: center;
    }

    ul {
        margin-top: 8px;
        padding-left: 22px;
    }

    .footer {
        color: #6b7280;
        font-size: 12px;
        margin-top: 26px;
    }
</style>
</head>
<body>
<div class="container">
    <h1>Tiered Active Directory Group Change Report</h1>
    <p class="subtitle">Tier 0 and Tier 1 privileged group monitoring</p>

    <div class="status $StatusClass">$([System.Net.WebUtility]::HtmlEncode($StatusText))</div>

    <div class="cards">
        <div class="card">
            <span class="card-label">Report generated</span>
            <span class="card-value">$($RunEnd.ToString("yyyy-MM-dd HH:mm:ss"))</span>
        </div>
        <div class="card">
            <span class="card-label">Window start</span>
            <span class="card-value">$($WindowStart.ToString("yyyy-MM-dd HH:mm:ss"))</span>
        </div>
        <div class="card">
            <span class="card-label">DCs queried</span>
            <span class="card-value">$($QueriedDCs.Count)</span>
        </div>
        <div class="card">
            <span class="card-label">Total changes</span>
            <span class="card-value">$($Changes.Count)</span>
        </div>
        <div class="card">
            <span class="card-label">Tier 0 changes</span>
            <span class="card-value">$($Tier0Changes.Count)</span>
        </div>
        <div class="card">
            <span class="card-label">Tier 1 changes</span>
            <span class="card-value">$($Tier1Changes.Count)</span>
        </div>
        <div class="card">
            <span class="card-label">Membership changes</span>
            <span class="card-value">$($MembershipChanges.Count)</span>
        </div>
    </div>

    <h2>Monitored Groups</h2>
    <table>
        <thead>
            <tr>
                <th>Group Name</th>
                <th>Assigned Tier</th>
            </tr>
        </thead>
        <tbody>
            $($MonitoredGroupRows -join "`n")
        </tbody>
    </table>

    <h2>Detected Changes</h2>
    <table>
        <thead>
            <tr>
                <th>Time</th>
                <th>Tier</th>
                <th>Domain Controller</th>
                <th>Event ID</th>
                <th>Group</th>
                <th>Actor</th>
                <th>Member</th>
                <th>Change Type</th>
            </tr>
        </thead>
        <tbody>
            $($EventRows -join "`n")
        </tbody>
    </table>

    <h2>Collection Status</h2>
    <ul>
        $($ErrorRows -join "`n")
    </ul>

    <p class="footer">
        HTML report: $([System.Net.WebUtility]::HtmlEncode($HtmlPath))<br>
        CSV report: $([System.Net.WebUtility]::HtmlEncode($CsvPath))
    </p>
</div>
</body>
</html>
"@

$Html | Out-File -FilePath $HtmlPath -Encoding UTF8

# ---------------------------------------------------------------------
# SMTP delivery
# Disabled by default. Configure settings and set $SendEmail = $true.
# ---------------------------------------------------------------------
if ($SendEmail) {
    try {
        $MailParams = @{
            From       = $MailFrom
            To         = $MailTo
            Subject    = "$MailSubjectPrefix - $($RunEnd.ToString('yyyy-MM-dd'))"
            Body       = $Html
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            Port       = $SmtpPort
            UseSsl     = $UseSsl
            ErrorAction = "Stop"
        }

        Send-MailMessage @MailParams
        Write-Host "SMTP email sent to: $($MailTo -join ', ')" -ForegroundColor Green
    }
    catch {
        $CollectionErrors.Add("SMTP email delivery failed. Error: $($_.Exception.Message)")
        Write-Warning "SMTP email delivery failed: $($_.Exception.Message)"
    }
}
else {
    Write-Host "SMTP email is disabled. Set `$SendEmail = `$true after configuring SMTP settings." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
# Always open the generated HTML report in the default browser.
# ---------------------------------------------------------------------
try {
    Start-Process -FilePath $HtmlPath
    Write-Host "Opened HTML report in the default browser: $HtmlPath" -ForegroundColor Green
}
catch {
    Write-Warning "Could not open the HTML report automatically. Open it manually: $HtmlPath"
}

Write-Host "CSV report saved to: $CsvPath" -ForegroundColor Green