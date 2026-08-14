<#
.SYNOPSIS
    Retrieves recent shutdown/restart events from the local server, including who initiated them and why.
.DESCRIPTION
    Queries the System event log for Event ID 1074 (user/process initiated shutdown/restart) and
    Event ID 6008 (unexpected shutdown). Extracts the user, time, shutdown type, reason code, and
    any comment provided.
.PARAMETER MaxEvents
    Maximum number of events to retrieve. Default is 10.
.PARAMETER DaysBack
    Number of days to look back. Default is 7.
.EXAMPLE
    .\Get-ServerRebootHistory.ps1
    .\Get-ServerRebootHistory.ps1 -MaxEvents 20 -DaysBack 14
#>

param(
    [int]$MaxEvents = 50,
    [int]$DaysBack = 20
)

$StartTime = (Get-Date).AddDays(-$DaysBack)
$LogName = 'System'

# Event IDs to check:
# 1074 - User or process initiated shutdown/restart (includes who and why)
# 6008 - Unexpected shutdown (e.g., power loss, crash)
$EventIDs = @(1074, 6008)

Write-Host "Querying events from the last $DaysBack day(s)..." -ForegroundColor Cyan

try {
    $Events = Get-WinEvent -FilterHashtable @{
        LogName   = $LogName
        ID        = $EventIDs
        StartTime = $StartTime
    } -MaxEvents $MaxEvents -ErrorAction Stop
}
catch {
    Write-Host "Error retrieving events: $_" -ForegroundColor Red
    exit 1
}

if ($Events.Count -eq 0) {
    Write-Host "No shutdown/restart events found in the last $DaysBack day(s)." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($Events.Count) event(s)." -ForegroundColor Green
Write-Host ""

# Process and display each event
$Results = @()

foreach ($Event in $Events) {
    $EventData = [PSCustomObject]@{
        Time         = $Event.TimeCreated
        EventID      = $Event.Id
        User         = $null
        ShutdownType = $null
        ReasonCode   = $null
        Comment      = $null
        Message      = $Event.Message
    }

    if ($Event.Id -eq 1074) {
        # Event 1074: Properties array contains user, type, reason code, comment, etc.
        # Typical order: [0]=Process, [1]=ProcessPath, [2]=User, [3]=ShutdownType, [4]=ReasonCode, [5]=Comment
        try {
            $props = $Event.Properties
            if ($props.Count -ge 6) {
                $EventData.User = $props[2].Value
                $EventData.ShutdownType = $props[4].Value
                $EventData.ReasonCode = $props[5].Value
                $EventData.Comment = $props[6].Value
            }
        }
        catch {
            # Fallback: parse from message if property extraction fails
            if ($Event.Message -match "user:\s*(.+?)(?:\r|\n|$)") {
                $EventData.User = $Matches[1].Trim()
            }
        }
    }
    elseif ($Event.Id -eq 6008) {
        # Unexpected shutdown - no user info available, just mark as such
        $EventData.User = "N/A (unexpected/power loss)"
        $EventData.ShutdownType = "Unexpected"
        $EventData.ReasonCode = "N/A"
        $EventData.Comment = "System shut down unexpectedly"
    }

    $Results += $EventData
}

# Display results in a formatted table
$Results | Sort-Object Time -Descending | Format-Table -AutoSize -Property @(
    @{Name="Time"; Expression={$_.Time.ToString("yyyy-MM-dd HH:mm:ss")}},
    @{Name="EventID"; Expression={$_.EventID}},
    @{Name="User"; Expression={$_.User}},
    @{Name="Type"; Expression={$_.ShutdownType}},
    @{Name="ReasonCode"; Expression={$_.ReasonCode}},
    @{Name="Comment"; Expression={$_.Comment}}
)

# Optional: Export to CSV for further analysis
# $Results | Export-Csv -Path "RebootHistory_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
