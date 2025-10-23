<#
.SYNOPSIS
    Active Directory / DNS / DHCP / GPO / Exchange Inventory (HTML + CSV)
    + NTLM-focused GPO detection and optional remote verification of NTLM-related registry settings.
.DESCRIPTION
    Collects forest, domain, FSMO, DC, DNS, DHCP, GPO, privileged groups, Exchange info,
    and a dedicated NTLM GPO section (settings + links + enabled state).
    Optional: verifies effective registry settings on specified machines.
    Author: Steve McKee - stevemckee@outlook.com
#>

#region Configuration
# Output location
$OutputPath = "$env:USERPROFILE\Desktop\AD_Inventory_$(Get-Date -Format yyyyMMdd_HHmmss)"
New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null

# Optional verification: set to $true to attempt remote checks on machines listed below
$EnableVerification = $true
# Provide machine names or FQDNs to verify (empty => no verification)
$VerificationTargets = @("XXX.com","PDC.DC.com")  # <-- change these as needed

# NTLM canonical detection patterns (case-insensitive)
$NTLMPatterns = @(
    'Restrict NTLM',
    'Audit NTLM',
    'NTLMv2',
    'LmCompatibilityLevel',
    'Allow LM',
    'Network security: Allow',
    'RestrictIncomingNTLMTraffic',
    'RestrictOutgoingNTLM',
    'NTLM'
)

# Common registry locations to query for verification (best-effort)
$VerifyRegistryPaths = @(
    'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider',
    'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\NTLM'
)
#endregion

#region Helpers
$global:FullReportBuilder = New-Object System.Text.StringBuilder
function Add-ContentReport {
    param(
        [Parameter(Mandatory=$true)][string]$html,
        [switch]$LineBreak
    )
    [void]$global:FullReportBuilder.AppendLine($html)
    if ($LineBreak) { [void]$global:FullReportBuilder.AppendLine("<br/>") }
}

$global:ReportHeaderHTML = @"
<html>
<head>
<title>Active Directory Inventory Report - NTLM-focused</title>
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
<h2>Author: Steve McKee IGTPLC</h2>
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

#region Module Imports
Import-ModuleSafe ActiveDirectory | Out-Null
Import-ModuleSafe DnsServer | Out-Null
Import-ModuleSafe DhcpServer | Out-Null
Import-ModuleSafe GroupPolicy | Out-Null
#endregion

#region GPOs - NTLM Configuration (Strict Detection) + Verification
Write-Output "Analyzing GPOs for NTLM configuration (strict detection)..."
try {
    $ntlmGpoData = @()
    $allGpos = Get-GPO -All -ErrorAction SilentlyContinue

    foreach ($gpo in $allGpos) {
        try {
            $reportXml = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
            if (-not $reportXml) { continue }

            [xml]$xml = $reportXml
            $matchedNodes = @()

            # Scan all XML nodes for any NTLM-related keywords
            $allNodes = $xml.SelectNodes("//*")
            foreach ($n in $allNodes) {
                try {
                    $inner = if ($n.InnerText) { $n.InnerText } else { '' }
                    $attrs = @()
                    if ($n.Attributes) {
                        foreach ($a in $n.Attributes) { $attrs += $a.Value }
                    }
                    $combined = ($inner + ' ' + ($attrs -join ' ')).ToLower()
                    foreach ($pat in $NTLMPatterns) {
                        if ($combined -match [regex]::Escape($pat.ToLower())) {
                            if (-not ($matchedNodes | Where-Object { $_ -eq $n })) { $matchedNodes += $n }
                            break
                        }
                    }
                } catch { }
            }

            if ($matchedNodes.Count -gt 0) {
                $links = @()
                try {
                    $guidPlain = ([string]$gpo.Id).Trim('{}')
                    $escapedGuid = [regex]::Escape($guidPlain)
                    if (Get-Command Get-ADOrganizationalUnit -ErrorAction SilentlyContinue) {
                        $ous = Get-ADOrganizationalUnit -Filter * -Properties gPLink -ErrorAction SilentlyContinue
                        foreach ($ou in $ous) {
                            if ($ou.gPLink -and ($ou.gPLink -match $escapedGuid)) {
                                $links += $ou.DistinguishedName
                            }
                        }
                    }
                } catch { }

                if ($links.Count -eq 0) {
                    try {
                        $linkNodes = $xml.SelectNodes("//LinksTo/Link")
                        foreach ($lnk in $linkNodes) {
                            $pathNode = $lnk.SelectSingleNode("Properties/Path")
                            if ($pathNode) { $links += $pathNode.InnerText }
                        }
                    } catch {}
                }

                foreach ($node in $matchedNodes) {
                    $settingName = ''
                    if ($node.Attributes -and $node.Attributes['name']) { $settingName = $node.Attributes['name'].Value }
                    elseif ($node.Attributes -and $node.Attributes['displayname']) { $settingName = $node.Attributes['displayname'].Value }
                    elseif ($node.Name) { $settingName = $node.Name }

                    $value = ''
                    if ($node.Attributes) {
                        foreach ($a in $node.Attributes) {
                            if ($a.Name -match '(?i)value|state|configured|setting') { $value = $a.Value; break }
                        }
                    }
                    if (-not $value) {
                        $candidate = $node.SelectSingleNode(".//State")
                        if ($candidate) { $value = $candidate.InnerText }
                    }
                    if (-not $value) {
                        $value = ($node.InnerText -replace '\s{2,}', ' ').Trim()
                    }

                    $parent = if ($node.ParentNode) { $node.ParentNode.Name } else { '' }
                    $grandParent = if ($node.ParentNode -and $node.ParentNode.ParentNode) { $node.ParentNode.ParentNode.Name } else { '' }
                    $context = ($grandParent + '/' + $parent) -replace '^/', ''

                    $ntlmGpoData += [PSCustomObject]@{
                        GPOName     = $gpo.DisplayName
                        GPOId       = [string]$gpo.Id
                        GPOEnabled  = $gpo.GpoStatus
                        SettingName = $settingName
                        Context     = $context
                        Value       = $value
                        LinkedOUs   = ($links -join '; ')
                    }
                }
            }
        } catch {
            Write-Warning "Failed NTLM analysis for $($gpo.DisplayName): $_"
        }
    }

    if ($ntlmGpoData.Count -eq 0) {
        $ntlmGpoData += [PSCustomObject]@{
            GPOName = '<No NTLM-related settings found>'
            GPOId = ''
            GPOEnabled = ''
            SettingName = ''
            Context = ''
            Value = ''
            LinkedOUs = ''
        }
    }

    $ntlmCsv = Join-Path $OutputPath 'gpos_ntlm.csv'
    $ntlmGpoData | Export-Csv $ntlmCsv -NoTypeInformation -Force

    Add-ContentReport "<details open><summary>NTLM Configuration GPOs (strict detection)</summary>"
    Add-ContentReport "<p>GPOs containing canonical NTLM-related settings (patterns: $($NTLMPatterns -join ', ')).</p>"
    Add-ContentReport (($ntlmGpoData | Select-Object GPOName,GPOId,GPOEnabled,SettingName,Context,Value,LinkedOUs | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
catch {
    Write-Warning "NTLM GPO analysis failed overall: $_"
}
#endregion

#region Optional: Remote Verification of NTLM-related Registry Settings
if ($EnableVerification -and $VerificationTargets.Count -gt 0) {
    Write-Output "Performing optional verification on target machines..."
    $verResults = @()

    foreach ($target in $VerificationTargets) {
        try {
            $script = {
                param($paths)
                $found = @()
                foreach ($p in $paths) {
                    try {
                        if (Test-Path $p) {
                            $props = Get-ItemProperty -Path $p -ErrorAction Stop
                            foreach ($pn in $props.PSObject.Properties) {
                                if ($pn.Name -in @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')) { continue }
                                $found += [PSCustomObject]@{
                                    Path = $p
                                    Name = $pn.Name
                                    Value = ($pn.Value -join ', ')
                                }
                            }
                        } else {
                            $found += [PSCustomObject]@{ Path = $p; Name = '<NotPresent>'; Value = '' }
                        }
                    } catch {
                        $found += [PSCustomObject]@{ Path = $p; Name = '<Error>'; Value = $_.Exception.Message }
                    }
                }
                return $found
            }

            $res = $null
            try {
                $res = Invoke-Command -ComputerName $target -ScriptBlock $script -ArgumentList ($VerifyRegistryPaths) -ErrorAction Stop -ThrottleLimit 4
            } catch {
                $res = @([PSCustomObject]@{ Path = '<ConnectionFailed>'; Name = '<Error>'; Value = $_.Exception.Message })
            }

            foreach ($r in $res) {
                $verResults += [PSCustomObject]@{
                    Target = $target
                    Path = $r.Path
                    Name = $r.Name
                    Value = $r.Value
                }
            }
        } catch {
            Write-Warning "Verification for $target failed: $_"
            $verResults += [PSCustomObject]@{
                Target = $target
                Path = '<VerificationFailed>'
                Name = ''
                Value = $_.Exception.Message
            }
        }
    }

    if ($verResults.Count -eq 0) {
        $verResults += [PSCustomObject]@{
            Target = '<No results>'
            Path = ''
            Name = ''
            Value = 'No verification results produced.'
        }
    }

    $verResults | Export-Csv (Join-Path $OutputPath 'ntlm_verification.csv') -NoTypeInformation -Force
    Add-ContentReport "<details><summary>NTLM Verification (remote registry checks)</summary>"
    Add-ContentReport "<p>Attempted to read common NTLM-related registry keys on specified machines.</p>"
    Add-ContentReport (($verResults | ConvertTo-Html -Fragment) -join "`r`n")
    Add-ContentReport "</details>"
}
else {
    Add-ContentReport "<details><summary>NTLM Verification</summary><p>Verification disabled or no targets specified.</p></details>"
}
#endregion

#region Finalize Report
Write-Output "Finalizing report..."
$reportFile = Join-Path $OutputPath 'FullReport.html'
$final = $global:ReportHeaderHTML + $global:FullReportBuilder.ToString() + $global:ReportFooterHTML
$final | Out-File -FilePath $reportFile -Encoding UTF8 -Force
Write-Output "Report saved to: $reportFile"
Write-Output "CSVs saved to: $OutputPath"
#endregion
