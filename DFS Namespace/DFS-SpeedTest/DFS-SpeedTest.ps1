Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "DFS File Transfer Speed Test"
$form.Size = New-Object System.Drawing.Size(650,500)
$form.StartPosition = "CenterScreen"

# Source
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Folder:"
$lblSource.Location = New-Object Drawing.Point(10,20)
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object Drawing.Point(120,18)
$txtSource.Size = New-Object Drawing.Size(480,20)
$form.Controls.Add($txtSource)

# Destination
$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination Folder:"
$lblDest.Location = New-Object Drawing.Point(10,60)
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object Drawing.Point(120,58)
$txtDest.Size = New-Object Drawing.Size(480,20)
$form.Controls.Add($txtDest)

# File Size
$lblSize = New-Object System.Windows.Forms.Label
$lblSize.Text = "Test File Size (MB):"
$lblSize.Location = New-Object Drawing.Point(10,100)
$form.Controls.Add($lblSize)

$txtSize = New-Object System.Windows.Forms.TextBox
$txtSize.Location = New-Object Drawing.Point(120,98)
$txtSize.Text = "1024"
$form.Controls.Add($txtSize)

# Ping Target
$lblPing = New-Object System.Windows.Forms.Label
$lblPing.Text = "Ping Target:"
$lblPing.Location = New-Object Drawing.Point(10,140)
$form.Controls.Add($lblPing)

$txtPing = New-Object System.Windows.Forms.TextBox
$txtPing.Location = New-Object Drawing.Point(120,138)
$txtPing.Size = New-Object Drawing.Size(200,20)
$form.Controls.Add($txtPing)

# Results
$txtResults = New-Object System.Windows.Forms.TextBox
$txtResults.Location = New-Object Drawing.Point(10,220)
$txtResults.Size = New-Object Drawing.Size(600,220)
$txtResults.Multiline = $true
$txtResults.ScrollBars = "Vertical"
$form.Controls.Add($txtResults)

function Write-Log {
    param($Text)
    $txtResults.AppendText("$Text`r`n")
}

# Ping Test
$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Text = "Ping Test"
$btnPing.Location = New-Object Drawing.Point(350,135)

$btnPing.Add_Click({
    try {
        $result = Test-Connection $txtPing.Text -Count 4
        $avg = ($result | Measure-Object ResponseTime -Average).Average
        Write-Log "Average Ping: $(:Round($avg,2)) ms"
    }
    catch {
        Write-Log "Ping failed."
    }
})

$form.Controls.Add($btnPing)

# Transfer Test
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Run Transfer Test"
$btnTest.Location = New-Object Drawing.Point(10,180)

$btnTest.Add_Click({

    $source = $txtSource.Text
    $dest = $txtDest.Text
    $sizeMB = [int]$txtSize.Text

    $testFile = Join-Path $source "DFSSpeedTest.dat"

    Write-Log ""
    Write-Log "Creating $sizeMB MB test file..."

    fsutil file createnew $testFile ($sizeMB * 1MB) | Out-Null

    $destFile = Join-Path $dest "DFSSpeedTest.dat"

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Copy-Item $testFile $destFile -Force

    $sw.Stop()

    $seconds = $sw.Elapsed.TotalSeconds

    $mbps = $sizeMB / $seconds
    $networkMbps = ($sizeMB * 8) / $seconds

    Write-Log "Transfer Complete"
    Write-Log "Time: $(:Round($seconds,2)) seconds"
    Write-Log "Speed: $([math]::Round($mbps,2))"
    Write-Log "Network Rate: $([math]::RoundrkMbps,2)) Mbps"

    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Remove-Item $destFile -Force -ErrorAction SilentlyContinue
})

$form.Controls.Add($btnTest)

# Export
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Results"
$btnExport.Location = New-Object Drawing.Point(150,180)

$btnExport.Add_Click({

    $file = "$env:USERPROFILE\Desktop\DFS_Speed_Test.csv"

    $txtResults.Lines |
        ForEach-Object {
            [PSCustomObject]@{
                Result = $_
            }
        } |
        Export-Csv $file -NoTypeInformation

    [System.Windows.Forms.MessageBox]::Show(
        "Saved to $file"
    )
})

$form.Controls.Add($btnExport)

$form.ShowDialog()
