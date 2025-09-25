<#
.Version 4 - Written By Steve McKee - IGT Systems Admin II

Instructions:
1. Copy and Paste into an Elevate PowerShell ISE session on a domain controller
2. Make sure there is a c:\Temp\GPOS folder before running
3. All files will be extracted to c:\Temp\GPOS
4. Wait for finish, All GPO's that have a Firewall built-in will be the file called - AllGPOFirewallRules.csv
5. There is a file called - file:///C:/Temp/GPOS/AllGPOFirewallRules.html which has All GPO's in the Domain, and below that is the Firewall GPO's in HTML format for your viewing pleasure.


.SYNOPSIS
  Scan domain GPOs for firewall settings, export per-GPO XMLs, produce consolidated CSV + HTML,
  and export effective GroupPolicy-based firewall rules on the local machine (CSV + HTML).


.OUTPUT
  C:\Temp\GPOS\Individual\*.xml
  C:\Temp\GPOS\All-GPO-FirewallRules.csv
  C:\Temp\GPOS\All-GPO-FirewallRules.html
  C:\Temp\GPOS\EffectiveFirewallRules.csv
  C:\Temp\GPOS\EffectiveFirewallRules.html

.NOTES
  Run as Administrator. Requires GroupPolicy module (RSAT).
#>

# --- Paths ---
$BasePath      = "C:\Temp\GPOS"
$XmlPath       = Join-Path $BasePath "Individual"
$MasterCsv     = Join-Path $BasePath "AllGPOFirewallRules.csv"
$MasterHtml    = Join-Path $BasePath "AllGPOFirewallRules.html"
$EffectiveCsv  = Join-Path $BasePath "EffectiveFirewallRules.csv"
$EffectiveHtml = Join-Path $BasePath "EffectiveFirewallRules.html"

# create folders
foreach ($p in @($BasePath, $XmlPath)) {
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory | Out-Null }
}

# load module
Import-Module GroupPolicy -ErrorAction Stop

# container for results
$masterResults = [System.Collections.Generic.List[PSObject]]::new()

Write-Host "Enumerating GPOs..."
$gpos = Get-GPO -All

foreach ($gpo in $gpos) {
    try {
        $reportXml = Get-GPOReport -Guid $gpo.Id -ReportType Xml
        $fileName  = ($gpo.DisplayName -replace '[\\/:*?"<>|]', '_') + ".xml"
        $fullXmlPath = Join-Path $XmlPath $fileName
        $reportXml | Out-File -FilePath $fullXmlPath -Encoding UTF8

        # parse xml
        $xml = [xml]$reportXml

        # find links (where GPO is linked)
        $linkSummary = "NotFound/Unknown"
        try {
            $links = $xml.SelectNodes("//LinksTo//Link")
            if ($links -and $links.Count -gt 0) {
                $linkItems = @()
                foreach ($ln in $links) {
                    $inner = $ln.InnerText.Trim()
                    if ($inner) { $linkItems += $inner } else {
                        $attrStr = ""
                        foreach ($a in $ln.Attributes) { $attrStr += ($a.Name + "=" + $a.Value + "; ") }
                        if ($attrStr) { $linkItems += $attrStr.TrimEnd() }
                    }
                }
                if ($linkItems.Count -gt 0) { $linkSummary = ($linkItems -join " | ") }
            }
        } catch {}

        $found = $false

        # Strategy A: structured Policy nodes
        $policyNodes = @()
        try { $policyNodes = $xml.SelectNodes("//Policy") } catch {}

        if ($policyNodes -and $policyNodes.Count -gt 0) {
            foreach ($pnode in $policyNodes) {
                $nodeText = $pnode.OuterXml
                if ($nodeText -match "(?i)firewall" -or $nodeText -match "(?i)netfirewallrule" -or $nodeText -match "(?i)Windows Defender Firewall") {
                    $found = $true
                    $displayName = $null
                    try { $nameChild = $pnode.SelectSingleNode("Name"); if ($nameChild) { $displayName = $nameChild.'#text' } } catch {}
                    if (-not $displayName) { $displayName = ($pnode.Name) }

                    $key = ""; $value = ""; $state = ""
                    try {
                        $kNode = $pnode.SelectSingleNode("Key");    if ($kNode) { $key = $kNode.'#text' }
                        $vNode = $pnode.SelectSingleNode("Value");  if ($vNode) { $value = $vNode.'#text' }
                        $sNode = $pnode.SelectSingleNode("State");  if ($sNode) { $state = $sNode.'#text' }
                    } catch {}

                    $masterResults.Add([PSCustomObject]@{
                        Timestamp  = (Get-Date).ToString("u")
                        GPOName    = $gpo.DisplayName
                        GPOId      = $gpo.Id.Guid
                        LinkTarget = $linkSummary
                        RuleNode   = ($displayName -replace '\r|\n',' ')
                        Key        = $key
                        Value      = $value
                        State      = $state
                        SourceFile = $fullXmlPath
                    })
                }
            }
        }

        # Strategy B: raw NetFirewallRule fragments
        if (-not $found) {
            $netMatches = [regex]::Matches($reportXml, "(?is)<NetFirewallRule.*?>.*?</NetFirewallRule>")
            if ($netMatches.Count -gt 0) {
                foreach ($m in $netMatches) {
                    $nfxml = $m.Value
                    $ruleName = ""

                    $nameMatch = [regex]::Match($nfxml, "(?is)<Name>(.*?)</Name>")
                    if ($nameMatch.Success) { $ruleName = $nameMatch.Groups[1].Value.Trim() }
                    else {
                        $dispMatch = [regex]::Match($nfxml, '(?i)DisplayName="([^"]+)"')
                        if ($dispMatch.Success) { $ruleName = $dispMatch.Groups[1].Value.Trim() }
                    }

                    $action = ""
                    $actMatch = [regex]::Match($nfxml, "(?is)<Action>(.*?)</Action>")
                    if ($actMatch.Success) { $action = $actMatch.Groups[1].Value.Trim() }

                    $profile = ""
                    $profMatch = [regex]::Match($nfxml, "(?is)<Profile>(.*?)</Profile>")
                    if ($profMatch.Success) { $profile = $profMatch.Groups[1].Value.Trim() }

                    $masterResults.Add([PSCustomObject]@{
                        Timestamp  = (Get-Date).ToString("u")
                        GPOName    = $gpo.DisplayName
                        GPOId      = $gpo.Id.Guid
                        LinkTarget = $linkSummary
                        RuleNode   = $ruleName
                        Key        = "NetFirewallRule-XML"
                        Value      = ($action + " / " + $profile).Trim()
                        State      = "Imported"
                        SourceFile = $fullXmlPath
                    })
                }
                $found = $true
            } elseif ($reportXml -match "(?i)Windows Defender Firewall|Windows Firewall") {
                # generic text match (least precise)
                $snippet = ($reportXml -replace "<.*?>"," ") -replace "\s{2,}"," "
                $masterResults.Add([PSCustomObject]@{
                    Timestamp  = (Get-Date).ToString("u")
                    GPOName    = $gpo.DisplayName
                    GPOId      = $gpo.Id.Guid
                    LinkTarget = $linkSummary
                    RuleNode   = "FirewallPolicy (text-match)"
                    Key        = ""
                    Value      = ($snippet.Substring(0,[math]::Min($snippet.Length,200)))
                    State      = "TextMatch"
                    SourceFile = $fullXmlPath
                })
                $found = $true
            }
        }

        if ($found) { Write-Host "Firewall-related policy found in GPO: $($gpo.DisplayName)" }
        else { Write-Host "No firewall settings in GPO: $($gpo.DisplayName)" }

    } catch {
        Write-Warning "Failed to process GPO '$($gpo.DisplayName)': $_"
    }
}

# --- Export consolidated CSV ---
if ($masterResults.Count -gt 0) {
    $masterResults | Sort-Object GPOName, RuleNode | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Master CSV: $MasterCsv"
} else {
    Write-Host "No firewall-related settings found in any GPOs."
}

# --- Build HTML report ---
$style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
h1 { color: #003366; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
a { color: #0066CC; text-decoration: none; }
a:hover { text-decoration: underline; }
.small { font-size: 11px; color: #666; }
</style>
"@

$header = "<h1>GPO Firewall Rules Report</h1>"
$summary = "<p class='small'>Generated: $(Get-Date -Format 'u') — Individual GPO XMLs are in <code>$XmlPath</code></p>"
$linkListHtml = "<h2>GPO XML files</h2><ul>"
Get-ChildItem -Path $XmlPath -Filter *.xml -ErrorAction SilentlyContinue | ForEach-Object {
    $full = $_.FullName.Replace('\','/')
    $linkListHtml += "<li><a href='file:///$full'>$(($_.Name))</a></li>"
}
$linkListHtml += "</ul>"

if ($masterResults.Count -gt 0) {
    $tableObj = $masterResults | Select-Object Timestamp, GPOName, GPOId, LinkTarget, RuleNode, Key, Value, State, SourceFile
    $tableFragment = $tableObj | ConvertTo-Html -Fragment

    $fullHtml = "<html><head>$style<title>GPO Firewall Rules Report</title></head><body>$header$summary$linkListHtml$tableFragment</body></html>"
    $fullHtml | Out-File -FilePath $MasterHtml -Encoding UTF8
    Write-Host "HTML report: $MasterHtml"
}

# --- Export effective firewall rules applied via Group Policy on local machine ---
try {
    $effective = Get-NetFirewallRule -ErrorAction Stop | Where-Object { $_.PolicyStoreSourceType -eq "GroupPolicy" } |
                 Select-Object DisplayName, Name, Direction, Action, Enabled, Profile, PolicyStoreSource

    if ($effective.Count -gt 0) {
        $effective | Export-Csv -Path $EffectiveCsv -NoTypeInformation -Encoding UTF8

        $effFragment = $effective | ConvertTo-Html -Fragment
        $effHtml = "<html><head>$style<title>Effective GPO Firewall Rules</title></head><body><h1>Effective GPO Firewall Rules (Local Machine)</h1>$effFragment</body></html>"
        $effHtml | Out-File -FilePath $EffectiveHtml -Encoding UTF8

        Write-Host "Effective CSV: $EffectiveCsv"
        Write-Host "Effective HTML: $EffectiveHtml"
    } else {
        Write-Host "No GroupPolicy-sourced firewall rules found on this machine."
    }
} catch {
    Write-Warning "Failed to enumerate/export effective firewall rules: $_"
}

Write-Host "`nFinished. Files are in: $BasePath"
