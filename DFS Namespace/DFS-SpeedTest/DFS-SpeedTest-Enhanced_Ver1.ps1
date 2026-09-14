# Install-Module DFSN -Force
# Install-Module DFSR -Force

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

$form = New-Object System.Windows.Forms.Form
$form.Text = "DFS / SMB Enterprise Performance Assessment Tool"
$form.Size = New-Object System.Drawing.Size(1400,900)
$form.StartPosition = "CenterScreen"

$txtSource = New-Object Windows.Forms.TextBox
$txtSource.Location = 20,20
$txtSource.Size = 600,25

$txtDestination = New-Object Windows.Forms.TextBox
$txtDestination.Location = 20,60
$txtDestination.Size = 600,25

$txtServer = New-Object Windows.Forms.TextBox
$txtServer.Location = 20,100
$txtServer.Size = 300,25

$txtServer.Text = $env:COMPUTERNAME

$form.Controls.AddRange(@(
    $txtSource,
    $txtDestination,
    $txtServer
))

$cmbSize = New-Object Windows.Forms.ComboBox
$cmbSize.Location = 650,20

@(
    "100MB",
    "1GB",
    "5GB",
    "10GB",
    "Custom"
) | ForEach-Object{
    $cmbSize.Items.Add($_)
}

$cmbSize.SelectedIndex=1

$form.Controls.Add($cmbSize)

$txtResults = New-Object Windows.Forms.RichTextBox

$txtResults.Location = 20,500
$txtResults.Size = 1320,300

$form.Controls.Add($txtResults)

function Write-Log{
    param($Message)

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $txtResults.AppendText(
        "$TimeStamp : $Message`r`n"
    )
}

$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart

$chart.Width = 1300
$chart.Height = 250
$chart.Location = New-Object Drawing.Point(20,220)

$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chart.ChartAreas.Add($chartArea)

$series = New-Object System.Windows.Forms.DataVisualization.Charting.Series

$series.ChartType = "Line"
$series.Name = "Transfer Speed MBps"

$chart.Series.Add($series)

$form.Controls.Add($chart)

function Test-PingLatency {

param($Target)

Write-Log "Starting Ping Test"

$result = Test-Connection $Target -Count 20

$avg =
($result | Measure-Object ResponseTime -Average).Average

$min =
($result | Measure-Object ResponseTime -Minimum).Minimum

$max =
($result | Measure-Object ResponseTime -Maximum).Maximum

$jitter =
(($result |
Select -Expand ResponseTime |
Measure-Object -StandardDeviation).StandardDeviation)

Write-Log "Ping Average: $avg ms"
Write-Log "Ping Min: $min ms"
Write-Log "Ping Max: $max ms"
Write-Log "Jitter: $jitter"
}

function Test-SMB {

param($Target)

Write-Log "Testing SMB"

$result = Test-NetConnection `
    -ComputerName $Target `
    -Port 445

Write-Log "445 Reachable: $($result.TcpTestSucceeded)"

Get-SmbConnection |
Out-String |
ForEach-Object{

    Write-Log $_

}
}

function Get-DFSReferralInfo {

Write-Log "Collecting DFS Referral Information"

dfsutil /pktinfo |
Out-String |
ForEach-Object{

    Write-Log $_

}
}

function Get-DFSRBacklogInfo {

param(
$Source,
$Destination
)

try {

dfsrdiag backlog `
    /rgname:* `
    /rfname:* `
    /sendingmember:$Source `
    /receivingmember:$Destination |
Out-String |
ForEach-Object{

     Write-Log $_

}

}
catch {

Write-Log $_

}
}

function Get-RemoteDiagnostics {

param($Server)

Write-Log "Remote Diagnostics: $Server"

Invoke-Command `
    -ComputerName $Server `
    -ScriptBlock {

        Get-Volume

        Get-PhysicalDisk

        Get-NetAdapterStatistics

        Get-SmbServerConfiguration

    } |
Out-String |
ForEach-Object{

   $_

} |
ForEach-Object{

   Write-Log $_

}
}

function Get-NetworkHealth {

$stats = Get-NetAdapterStatistics

foreach($adapter in $stats){

Write-Log "Adapter: $($adapter.Name)"

Write-Log "In Errors: $($adapter.ReceivedPacketErrors)"
Write-Log "Out Errors: $($adapter.OutboundPacketErrors)"

Write-Log "Discards: $($adapter.ReceivedDiscardedPackets)"

}
}

function Get-DiskLatency {

Get-Counter `
 '\PhysicalDisk(*)\Avg. Disk sec/Transfer' |
Select -Expand CounterSamples |
ForEach-Object{

Write-Log "$($_.Path) = $($_.CookedValue)"

}

}

function Start-TransferTest {

param(
$Source,
$Destination,
$SizeMB
)

$tempFile = Join-Path $Source "dfs_speed_test.bin"

fsutil file createnew `
    $tempFile `
    ($SizeMB * 1MB)

$destFile = Join-Path $Destination "dfs_speed_test.bin"

$sw = [Diagnostics.Stopwatch]::StartNew()

$job = Start-Job {

param($src,$dst)

robocopy `
  (Split-Path $src) `
  (Split-Path $dst) `
  (Split-Path $src -Leaf) `
  /MT:64 `
  /R:0 `
  /W:0 `
  /NFL `
  /NDL `
  /NP

} -ArgumentList $tempFile,$destFile

while($job.State -eq "Running"){

    if(Test-Path $destFile){

        $size =
            (Get-Item $destFile).Length

        $elapsed =
            :Max(
                $sw.Elapsed.TotalSeconds,
                1
            )

        $mbps =
            ($size/1MB)/$elapsed

        $chart.Series[0].Points.AddY(
            :Round(
                $mbps,
                2
            )
        )

        [Windows.Forms.Application]::DoEvents()

    }

    Start-Sleep 1

}

$sw.Stop()

Receive-Job $job

$totalSecs =
$sw.Elapsed.TotalSeconds

$speedMBps =
$SizeMB / $totalSecs

$speedMbps =
($SizeMB * 8)/$totalSecs

Write-Log "Transfer Complete"

Write-Log "Time: $totalSecs"

Write-Log "MB/s: $speedMBps"

Write-Log "Mbps: $speedMbps"

Remove-Item $tempFile -Force
Remove-Item $destFile -Force
}

function Export-HTMLReport {

$file =
"$env:USERPROFILE\Desktop\DFS_Report_$(Get-Date -f yyyyMMdd_HHmmss).html"

$html = @"

<html>
<head>
<title>DFS Performance Report</title>
</head>

<body>

<h1>DFS Assessment Report</h1>

<pre>

$($txtResults.Text)

</pre>

</body>

</html>

"@

$html |
Out-File $file

Write-Log "HTML Report Saved: $file"

}


function Export-CSVReport {

$file =
"$env:USERPROFILE\Desktop\DFS_Report.csv"

$txtResults.Lines |
ForEach-Object{

[PSCustomObject]@{

Result = $_

}

} |
Export-Csv `
$file `
-NoTypeInformation

Write-Log "CSV Saved: $file"

}
$btnFullAssessment = New-Object Windows.Forms.Button

$btnFullAssessment.Text =
"Run Full Assessment"

$btnFullAssessment.Location =
New-Object Drawing.Point(20,150)

$btnFullAssessment.Add_Click({

Write-Log "Starting Enterprise DFS Assessment"

Test-PingLatency $txtServer.Text

Test-SMB $txtServer.Text

Get-NetworkHealth

Get-DiskLatency

Get-DFSReferralInfo

Get-RemoteDiagnostics $txtServer.Text

Start-TransferTest `
    $txtSource.Text `
    $txtDestination.Text `
    1024

Export-CSVReport

Export-HTMLReport

Write-Log "Assessment Complete"

})

$form.Controls.Add($btnFullAssessment)
``
