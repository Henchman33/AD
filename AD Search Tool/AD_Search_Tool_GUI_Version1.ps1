<### 
BEGIN FULL SCRIPT: ADSearchTool_Enterprise.ps1
Created by Steve McKee - Server Admininistrtor II
Run using PowerShell ISE elevated "as Administrator" from a Domain Controller.
Enjoy!
###>

param(
    [switch]$ScheduledMode,
    [string]$Presets = "",
    [string]$ExportFolderArg = "",
    [string]$Formats = ""
)

# -----------------------------
# Assemblies & prerequisites
# -----------------------------
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# -----------------------------
# Configuration & Helpers
# -----------------------------
$Global:AppFolder    = "C:\Temp\ADSearchTool"
$Global:ExportFolder = Join-Path $Global:AppFolder "Export"
$Global:ConfigFile   = Join-Path $Global:AppFolder "config.json"

If (!(Test-Path $Global:AppFolder)) { New-Item -Path $Global:AppFolder -ItemType Directory -Force | Out-Null }
If (!(Test-Path $Global:ExportFolder)) { New-Item -Path $Global:ExportFolder -ItemType Directory -Force | Out-Null }

function Ensure-ModuleLoaded {
    param([string]$Name)
    if (Get-Module -ListAvailable -Name $Name) {
        Import-Module $Name -ErrorAction SilentlyContinue
        return $true
    } else { return $false }
}

$HasAD  = Ensure-ModuleLoaded -Name ActiveDirectory
$HasGPO = Ensure-ModuleLoaded -Name GroupPolicy

function SafeFileName { param([string]$n) return ($n -replace '[^\w\-\._ ]','_').Trim() }

# DPAPI wrappers for credential save/load
function Protect-Credential {
    param([PSCredential]$Credential)
    if (-not $Credential) { return $null }
    $plain = ($Credential.UserName + "`n" + ($Credential.GetNetworkCredential().Password))
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($plain)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    return [System.Convert]::ToBase64String($protected)
}
function Unprotect-Credential {
    param([string]$ProtectedString)
    if (-not $ProtectedString) { return $null }
    try {
        $bytes = [System.Convert]::FromBase64String($ProtectedString)
        $un = [System.Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $plain = [System.Text.Encoding]::UTF8.GetString($un)
        $parts = $plain -split "`n",2
        $username = $parts[0]
        $password = if ($parts.Count -ge 2) { $parts[1] } else { "" }
        return New-Object System.Management.Automation.PSCredential ($username,(ConvertTo-SecureString $password -AsPlainText -Force))
    } catch { return $null }
}

# Config save/load
function Save-Config { param($cfg) $cfg | ConvertTo-Json -Depth 6 | Out-File -FilePath $Global:ConfigFile -Encoding UTF8 }
function Load-Config { if (Test-Path $Global:ConfigFile) { try { Get-Content $Global:ConfigFile -Raw | ConvertFrom-Json } catch { $null } } else { $null } }

# -----------------------------
# Export function (header + format handling)
# -----------------------------
function Export-Results {
    param(
        [Parameter(Mandatory=$true)][array]$Results,
        [Parameter(Mandatory=$true)][string]$Category,
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][string]$ExportPath,
        [Parameter(Mandatory=$true)][string[]]$Formats
    )

    if (!(Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $base = SafeFileName($Category) + "_" + $timestamp
    $total = $Results.Count

    $header = @"
Active Directory Search Export
------------------------------
Category : $Category
Filter   : $Filter
Exported : $(Get-Date)
Total    : $total
------------------------------
"@.Trim()

    if ($total -eq 0) {
        foreach ($fmt in $Formats) {
            switch ($fmt.ToLower()) {
                "csv" { $target = Join-Path $ExportPath ($base + ".csv"); "# $header" | Out-File -FilePath $target -Encoding UTF8 }
                "html" { 
                    $target = Join-Path $ExportPath ($base + ".html")
                    Add-Type -AssemblyName System.Web
                    $htmlHeader = [System.Web.HttpUtility]::HtmlEncode($header)
                    "<html><body><pre>$htmlHeader</pre><h3>No results</h3></body></html>" | Out-File -FilePath $target -Encoding UTF8
                }
                "xml" { 
                    $target = Join-Path $ExportPath ($base + ".xml")
                    $xmlHeader = [System.Security.SecurityElement]::Escape($header)
                    "<results><summary>$xmlHeader</summary></results>" | Out-File -FilePath $target -Encoding UTF8
                }
                "txt" { $target = Join-Path $ExportPath ($base + ".txt"); $header | Out-File -FilePath $target -Encoding UTF8 }
                "excel" { $target = Join-Path $ExportPath ($base + ".csv"); "# $header" | Out-File -FilePath $target -Encoding UTF8 }
                default { $target = Join-Path $ExportPath ($base + ".txt"); $header | Out-File -FilePath $target -Encoding UTF8 }
            }
        }
        return
    }

    foreach ($fmt in $Formats) {
        switch ($fmt.ToLower()) {
            "csv" {
                $file = Join-Path $ExportPath ($base + ".csv")
                $headerLines = $header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }
                $csvLines = $Results | ConvertTo-Csv -NoTypeInformation
                $headerLines + $csvLines | Out-File -FilePath $file -Encoding UTF8
            }
            "xml" {
                $file = Join-Path $ExportPath ($base + ".xml")
                $Results | Export-Clixml -Path $file
            }
            "html" {
                $file = Join-Path $ExportPath ($base + ".html")
                Add-Type -AssemblyName System.Web
                $htmlHeader = [System.Web.HttpUtility]::HtmlEncode($header)
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$htmlHeader</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $file -Encoding UTF8
            }
            "txt" {
                $file = Join-Path $ExportPath ($base + ".txt")
                $header | Out-File -FilePath $file -Encoding UTF8
                $Results | Out-String | Out-File -FilePath $file -Append -Encoding UTF8
            }
            "excel" {
                $file = Join-Path $ExportPath ($base + ".xlsx")
                if (Get-Module -ListAvailable -Name ImportExcel) {
                    try {
                        Import-Module ImportExcel -ErrorAction Stop
                        $Results | Export-Excel -Path $file -WorksheetName "Results" -AutoSize -Title ("AD Search Results - " + $Category)
                    } catch {
                        $csvfile = Join-Path $ExportPath ($base + ".csv")
                        ($header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }) + ($Results | ConvertTo-Csv -NoTypeInformation) | Out-File -FilePath $csvfile -Encoding UTF8
                    }
                } else {
                    $csvfile = Join-Path $ExportPath ($base + ".csv")
                    ($header -split "(`r`n|`n|`r)" | ForEach-Object { "# $_" }) + ($Results | ConvertTo-Csv -NoTypeInformation) | Out-File -FilePath $csvfile -Encoding UTF8
                }
            }
            "pdf" {
                $htmlFile = Join-Path $ExportPath ($base + ".html")
                $pdfFile  = Join-Path $ExportPath ($base + ".pdf")
                Add-Type -AssemblyName System.Web
                $htmlHeader = [System.Web.HttpUtility]::HtmlEncode($header)
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$htmlHeader</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $htmlFile -Encoding UTF8
                $wk = (Get-Command wkhtmltopdf -ErrorAction SilentlyContinue).Path
                if ($wk) { & $wk $htmlFile $pdfFile }
            }
            "docx" {
                $htmlFile = Join-Path $ExportPath ($base + ".html")
                $docxFile = Join-Path $ExportPath ($base + ".docx")
                Add-Type -AssemblyName System.Web
                $htmlHeader = [System.Web.HttpUtility]::HtmlEncode($header)
                $html = $Results | ConvertTo-Html -PreContent ("<pre>$htmlHeader</pre>") -Title ("AD Search Results - " + $Category)
                $html | Out-File -FilePath $htmlFile -Encoding UTF8
                try {
                    $word = New-Object -ComObject Word.Application -ErrorAction Stop
                    $doc = $word.Documents.Open($htmlFile)
                    $doc.SaveAs([ref] $docxFile, [ref] 16)
                    $doc.Close()
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                } catch { }
            }
            default {
                $file = Join-Path $ExportPath ($base + ".txt")
                $header | Out-File -FilePath $file -Encoding UTF8
                $Results | Out-String | Out-File -FilePath $file -Append -Encoding UTF8
            }
        }
    }
}

# -----------------------------
# Core Search functions
# -----------------------------
if ($HasAD) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

function Search-Users {
    param([string]$filter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $props = @("Name","sAMAccountName","distinguishedName","Enabled","LockedOut","LastLogonDate","whenCreated","memberOf","userPrincipalName")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADUser -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADUser -Filter { Name -like $f -or sAMAccountName -like $f -or mail -like $f -or userPrincipalName -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    return $res | Select-Object @{n='Type';e={'User'}}, Name,sAMAccountName,distinguishedName,Enabled,LockedOut,LastLogonDate,whenCreated,userPrincipalName,@{n='MemberOf';e={$_.memberOf -join '; '}}
}

function Search-Computers {
    param([string]$filter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $props = @("Name","OperatingSystem","OperatingSystemVersion","distinguishedName","whenCreated","lastLogonDate")
    if ($filter -match '^\(|\=|\&|\|') {
        $res = Get-ADComputer -LDAPFilter $filter -Properties $props -ErrorAction SilentlyContinue
    } else {
        $f = $filter
        $res = Get-ADComputer -Filter { Name -like $f -or OperatingSystem -like $f } -Properties $props -ErrorAction SilentlyContinue
    }
    return $res | Select-Object @{n='Type';e={'Computer'}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}}
}

function Search-OUs {
    param([string]$filter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $res = Get-ADOrganizationalUnit -Filter { Name -like $filter } -Properties distinguishedName,whenCreated -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'OU'}}, Name,distinguishedName,whenCreated
}

function Search-GPOs {
    param([string]$filter)
    if (-not $HasGPO) { throw "GroupPolicy module not available." }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All | Where-Object { $_.DisplayName -like $filter }
    $out = foreach ($g in $gpos) {
        try {
            $links = ""
            $gpoLinks = Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction SilentlyContinue
            if ($gpoLinks) {
                [xml]$xmlReport = $gpoLinks
                $linkNodes = $xmlReport.SelectNodes("//LinksTo")
                if ($linkNodes) {
                    $links = ($linkNodes | ForEach-Object { $_.SOMPath }) -join "; "
                }
            }
            [pscustomobject]@{
                Type = "GPO"
                Name = $g.DisplayName
                Id = $g.Id
                Owner = $g.Owner
                CreationTime = $g.CreationTime
                ModificationTime = $g.ModificationTime
                Links = $links
            }
        } catch {
            [pscustomobject]@{
                Type = "GPO"
                Name = $g.DisplayName
                Id = $g.Id
                Owner = $g.Owner
                CreationTime = $g.CreationTime
                ModificationTime = $g.ModificationTime
                Links = "Unable to retrieve links"
            }
        }
    }
    return $out
}

function Search-Groups {
    param([string]$filter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $res = Get-ADGroup -Filter { Name -like $filter } -Properties member,GroupScope,GroupCategory -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'Group'}}, Name,GroupScope,GroupCategory,distinguishedName,@{n='Members';e={$_.member -join '; '}}
}

function Search-ServiceAccounts {
    param([string]$filter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $res = Get-ADUser -Filter { servicePrincipalName -like $filter -or sAMAccountName -like $filter } -Properties servicePrincipalName,description,distinguishedName -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={'ServiceAccount'}}, Name,sAMAccountName,servicePrincipalName,distinguishedName,description
}

function Search-ServersOrWorkstations {
    param([string]$filter,[switch]$Servers)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $osFilter = if ($Servers) { "*Server*" } else { "*Windows*" }
    $res = Get-ADComputer -Filter { OperatingSystem -like $osFilter -and Name -like $filter } -Properties OperatingSystem,OperatingSystemVersion,distinguishedName,lastLogonDate -ErrorAction SilentlyContinue
    return $res | Select-Object @{n='Type';e={if ($Servers) {'Server'}else{'Workstation'}}}, Name,OperatingSystem,OperatingSystemVersion,distinguishedName,@{n='LastLogon';e={$_.lastLogonDate}}
}

function Search-Subnets {
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    try {
        $configNaming = (Get-ADRootDSE).configurationNamingContext
        $base = "CN=Subnets,CN=Sites,$configNaming"
        $subnets = Get-ADObject -SearchBase $base -Filter * -Properties name,location,siteObject -ErrorAction SilentlyContinue
        return $subnets | Select-Object @{n='Type';e={'ADSubnet'}}, name,@{n='Location';e={$_.location}},@{n='DistinguishedName';e={$_.DistinguishedName}}
    } catch { return @() }
}

function Search-GPOFirewallRules {
    param([string]$filter)
    if (-not $HasGPO) { throw "GroupPolicy module unavailable." }
    Import-Module GroupPolicy -ErrorAction Stop
    $gpos = Get-GPO -All
    $matches = @()
    foreach ($g in $gpos) {
        try {
            $xml = Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction SilentlyContinue
            if (-not $xml) { continue }
            [xml]$gxml = $xml
            $policies = @()
            # Look for firewall policies in GPO XML
            $firewallNodes = $gxml.SelectNodes("//q1:FirewallSettings")
            if ($firewallNodes) {
                foreach ($node in $firewallNodes) {
                    $policyName = $node.Name
                    $policySetting = $node.InnerText
                    if ($policyName -like $filter -or $policySetting -like $filter) {
                        $links = ""
                        $linkNodes = $gxml.SelectNodes("//LinksTo")
                        if ($linkNodes) {
                            $links = ($linkNodes | ForEach-Object { $_.SOMPath }) -join "; "
                        }
                        $matches += [pscustomobject]@{
                            Type = "GPOFirewall"
                            GPOName = $g.DisplayName
                            Policy = $policyName
                            Setting = $policySetting
                            Links = $links
                        }
                    }
                }
            }
        } catch { }
    }
    return $matches
}

# -----------------------------
# Lockout search
# -----------------------------
function Search-LockedOutUsers {
    param([switch]$ResolveOrigin)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $locked = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue | Get-ADUser -Properties LockedOut,LastLogonDate,whenCreated,sAMAccountName,distinguishedName -ErrorAction SilentlyContinue
    $out = @()
    foreach ($u in $locked) {
        $record = [pscustomobject]@{
            Type = "LockedUser"
            Name = $u.Name
            sAMAccountName = $u.sAMAccountName
            DistinguishedName = $u.DistinguishedName
            LockedOut = $u.LockedOut
            LastLogon = $u.LastLogonDate
        }
        if ($ResolveOrigin) {
            $dcs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
            $origin = $null
            foreach ($dc in $dcs) {
                try {
                    $query = @"
<QueryList>
  <Query Id='0' Path='Security'>
    <Select Path='Security'>
      *[System[(EventID=4740)]] and *[EventData[Data and (Data='$($u.sAMAccountName)')]]
    </Select>
  </Query>
</QueryList>
"@
                    $events = Get-WinEvent -ComputerName $dc.HostName -FilterXml $query -MaxEvents 1 -ErrorAction SilentlyContinue
                    if ($events -and $events.Count -gt 0) {
                        $ev = $events[0]
                        $data = [xml]$ev.ToXml()
                        $td = $data.Event.EventData.Data
                        $caller = ($td | Where-Object { $_.Name -eq "TargetUserName" }).'#text'
                        $origin = @{ DomainController = $dc.HostName; CallerComputer = $caller; EventTime = $ev.TimeCreated }
                        break
                    }
                } catch { }
            }
            if ($origin) { 
                $originValue = "DC: $($origin.DomainController), Caller: $($origin.CallerComputer), Time: $($origin.EventTime)"
            } else { 
                $originValue = "Origin not found or insufficient permissions"
            }
            $record | Add-Member -MemberType NoteProperty -Name "LockoutOrigin" -Value $originValue -Force
        }
        $out += $record
    }
    return $out
}

# -----------------------------
# Advanced Presets
# -----------------------------

function Preset-GroupMembership {
    param([string]$GroupFilter)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    if (-not $GroupFilter) { throw "Group name or -like filter required (e.g. 'Domain Admins' or '*Admins*')." }
    $groups = Get-ADGroup -Filter { Name -like $GroupFilter } -ErrorAction SilentlyContinue
    $out = @()
    foreach ($g in $groups) {
        try {
            $members = Get-ADGroupMember -Identity $g.SamAccountName -Recursive -ErrorAction SilentlyContinue
            foreach ($m in $members) {
                $out += [pscustomobject]@{ Group = $g.Name; MemberName = $m.Name; MemberType = $m.ObjectClass; DistinguishedName = $m.DistinguishedName }
            }
        } catch { }
    }
    return $out
}

function Preset-NTLMAudit {
    param([int]$Days = 7)
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $dcs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
    $cut = (Get-Date).AddDays(-$Days)
    $out = @()
    foreach ($dc in $dcs) {
        try {
            $xml = @"
<QueryList>
  <Query Id='0' Path='Security'>
    <Select Path='Security'>*[System[(EventID=4624) and TimeCreated[timediff(@SystemTime) &lt;= $(($Days)*24*60*60*1000)]]]</Select>
  </Query>
</QueryList>
"@
            $events = Get-WinEvent -ComputerName $dc.HostName -FilterXml $xml -MaxEvents 1000 -ErrorAction SilentlyContinue
            foreach ($ev in $events) {
                try {
                    $x = [xml]$ev.ToXml()
                    $data = $x.Event.EventData.Data
                    $authPkg = ($data | Where-Object { $_.Name -eq "AuthenticationPackageName" }).'#text'
                    if ($authPkg -and $authPkg -match "NTLM") {
                        $acct = ($data | Where-Object { $_.Name -eq "TargetUserName" }).'#text'
                        $work = ($data | Where-Object { $_.Name -eq "WorkstationName" }).'#text'
                        $out += [pscustomobject]@{ DC = $dc.HostName; Time = $ev.TimeCreated; Account = $acct; Workstation = $work; EventID = $ev.Id }
                    }
                } catch { }
            }
            $evs2 = Get-WinEvent -ComputerName $dc.HostName -FilterHashtable @{ LogName='Security'; Id=4776; StartTime=$cut } -MaxEvents 500 -ErrorAction SilentlyContinue
            foreach ($ev in $evs2) {
                try {
                    $x = [xml]$ev.ToXml()
                    $data = $x.Event.EventData.Data
                    $acct = ($data | Where-Object { $_.Name -eq "TargetUserName" }).'#text'
                    $work = ($data | Where-Object { $_.Name -eq "Workstation" }).'#text'
                    $out += [pscustomobject]@{ DC = $dc.HostName; Time = $ev.TimeCreated; Account = $acct; Workstation = $work; EventID = $ev.Id }
                } catch { }
            }
        } catch { }
    }
    return $out | Sort-Object Time -Descending
}

function Preset-OUDelegation {
    param([string]$OUFilter = "*")
    if (-not $HasAD) { throw "ActiveDirectory module not available." }
    $ous = Get-ADOrganizationalUnit -Filter { Name -like $OUFilter } -Properties DistinguishedName -ErrorAction SilentlyContinue
    $report = @()
    foreach ($ou in $ous) {
        try {
            $dn = $ou.DistinguishedName
            $acl = Get-Acl -Path ("AD:" + $dn)
            foreach ($ace in $acl.Access) {
                $report += [pscustomobject]@{
                    OU = $ou.Name
                    OU_DN = $dn
                    IdentityReference = $ace.IdentityReference.ToString()
                    AccessControlType = $ace.AccessControlType
                    ActiveDirectoryRights = $ace.ActiveDirectoryRights
                    IsInherited = $ace.IsInherited
                    InheritanceType = $ace.InheritanceType
                }
            }
        } catch { }
    }
    return $report
}

# -----------------------------
# Presets registry
# -----------------------------
$Global:PresetsArray = @(
    [pscustomobject]@{ Name="Disabled Accounts"; Category="User"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { Enabled -eq $false } -Properties Enabled,whenCreated -ErrorAction SilentlyContinue | Select-Object Name,sAMAccountName,distinguishedName,Enabled,whenCreated } },
    [pscustomobject]@{ Name="Locked-Out Users"; Category="Locked-out Users (basic)"; Filter="*"; ScriptBlock={ Search-LockedOutUsers } },
    [pscustomobject]@{ Name="Service Accounts (SPN)"; Category="Service Accounts"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { servicePrincipalName -like '*' } -Properties servicePrincipalName -ErrorAction SilentlyContinue | Select-Object Name,sAMAccountName,servicePrincipalName,distinguishedName } },
    [pscustomobject]@{ Name="Password Never Expires"; Category="User"; Filter="*"; ScriptBlock={ Get-ADUser -Filter { PasswordNeverExpires -eq $true } -Properties PasswordNeverExpires -ErrorAction SilentlyContinue | Select-Object Name,sAMAccountName,distinguishedName,PasswordNeverExpires } },
    [pscustomobject]@{ Name="Domain Admins Members"; Category="Security Group"; Filter="Domain Admins"; ScriptBlock={ Get-ADGroupMember -Identity "Domain Admins" -Recursive -ErrorAction SilentlyContinue | Select-Object Name,sAMAccountName,distinguishedName,@{n='Group';e={'Domain Admins'}} } },
    [pscustomobject]@{ Name="Recently Created Accounts (30d)"; Category="User"; Filter="*"; ScriptBlock={ $since=(Get-Date).AddDays(-30); Get-ADUser -Filter { whenCreated -ge $since } -Properties whenCreated -ErrorAction SilentlyContinue | Select-Object Name,sAMAccountName,distinguishedName,whenCreated } },
    [pscustomobject]@{ Name="Inactive Computers (90d)"; Category="Computer"; Filter="*"; ScriptBlock={ $cut=(Get-Date).AddDays(-90); Get-ADComputer -Filter * -Properties LastLogonDate -ErrorAction SilentlyContinue | Where-Object { $_.LastLogonDate -lt $cut -or -not $_.LastLogonDate } | Select-Object Name,OperatingSystem,distinguishedName,@{n='LastLogon';e={$_.LastLogonDate}} } },
    [pscustomobject]@{ Name="Domain Controllers"; Category="Servers (by OS or group)"; Filter="*Server*"; ScriptBlock={ Get-ADDomainController -Filter * -ErrorAction SilentlyContinue | Select-Object HostName,Site,OperatingSystem } },
    [pscustomobject]@{ Name="GPOs with Log on as a service"; Category="GPO"; Filter="*"; ScriptBlock={ param($f); if (-not (Get-Module -ListAvailable -Name GroupPolicy)) { throw 'GroupPolicy missing' } ; Get-GPO -All -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $f } | ForEach-Object { $links=""; try { $report=Get-GPOReport -Guid $_.Id -ReportType Xml -ErrorAction SilentlyContinue; if($report){[xml]$x=$report; $ln=$x.SelectNodes("//LinksTo"); if($ln){$links=($ln|%{$_.SOMPath})-join'; '}} } catch {}; [pscustomobject]@{Name=$_.DisplayName;Id=$_.Id;Links=$links} } } },
    [pscustomobject]@{ Name="Group Membership by Group"; Category="Security Group"; Filter="Domain Admins"; ScriptBlock={ param($f) ; Preset-GroupMembership -GroupFilter $f } },
    [pscustomobject]@{ Name="NTLM Audit (last 7 days)"; Category="NTLM Audit"; Filter="7"; ScriptBlock={ param($days) ; Preset-NTLMAudit -Days ([int]$days) } },
    [pscustomobject]@{ Name="OU Delegation Report"; Category="OU Delegation"; Filter="*"; ScriptBlock={ param($f) ; Preset-OUDelegation -OUFilter $f } }
)

# -----------------------------
# Scheduler install/uninstall functions
# -----------------------------
function Install-ScheduledReport {
    param(
        [string]$TaskName = "ADSearchTool_Nightly",
        [string]$ScriptPath = $PSCommandPath,
        [string]$ExportPath = $Global:ExportFolder,
        [string[]]$PresetNames = @("Disabled Accounts","Inactive Computers (90d)"),
        [string[]]$Formats = @("csv","html","excel"),
        [string]$RunTime = "02:00"
    )
    if (-not (Test-Path $ScriptPath)) { throw "Script not found: $ScriptPath" }
    $presetArg = $PresetNames -join ";"
    $fmtArg = $Formats -join ";"
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -ScheduledMode -Presets `"$presetArg`" -ExportFolderArg `"$ExportPath`" -Formats `"$fmtArg`""
    $trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($RunTime))
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable)
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force
    Write-Host "Scheduled task '$TaskName' installed to run daily at $RunTime as SYSTEM."
}

function Uninstall-ScheduledReport {
    param([string]$TaskName = "ADSearchTool_Nightly")
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Scheduled task '$TaskName' removed."
    } else { Write-Host "Scheduled task '$TaskName' not found." }
}

# -----------------------------
# Scheduled-mode runner (Task Scheduler calls the script with -ScheduledMode)
# -----------------------------
if ($ScheduledMode) {
    try {
        $presetNames = if ($Presets) { $Presets -split ";" } else { @() }
        $exportPath = if ($ExportFolderArg) { $ExportFolderArg } else { $Global:ExportFolder }
        $formats = if ($Formats) { $Formats -split ";" } else { @("csv","html","excel") }
        foreach ($pname in $presetNames) {
            $preset = $Global:PresetsArray | Where-Object { $_.Name -eq $pname } | Select-Object -First 1
            if (-not $preset) { continue }
            $sb = $preset.ScriptBlock
            $res = if ($sb.Parameters.Count -gt 0) { & $sb $preset.Filter } else { & $sb }
            $arr = @(); foreach ($r in $res) { $arr += $r }
            Export-Results -Results $arr -Category $preset.Name -Filter $preset.Filter -ExportPath $exportPath -Formats $formats
        }
    } catch { }
    return
}

# -----------------------------
# WPF UI XAML (Search/Presets/Settings/Scheduler)
# -----------------------------
$XAML = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        xmlns:wf='clr-namespace:System.Windows.Forms;assembly=System.Windows.Forms'
        Title='AD Search Tool - Enterprise' Height='760' Width='1200' WindowStartupLocation='CenterScreen'>
  <Grid Margin='8'>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <StackPanel Orientation='Horizontal' Grid.Row='0' Margin='0,0,0,6'>
      <Label Content='Category:' VerticalAlignment='Center'/>
      <ComboBox x:Name='cmbCategory' Width='300' Margin='8,0,0,0'/>
      <Label Content='Filter:' VerticalAlignment='Center' Margin='12,0,0,0'/>
      <TextBox x:Name='txtFilter' Width='420' Margin='8,0,0,0'/>
      <Button x:Name='btnRun' Content='Run Search' Width='120' Margin='12,0,0,0'/>
      <Button x:Name='btnClear' Content='Clear' Width='80' Margin='6,0,0,0'/>
    </StackPanel>

    <StackPanel Orientation='Horizontal' Grid.Row='1' Margin='0,0,0,6'>
      <CheckBox x:Name='chkResolveSIDs' Content='Resolve SIDs' IsChecked='True' Margin='6'/>
      <CheckBox x:Name='chkEventLookup' Content='Lockout origin lookup (slow)' IsChecked='False' Margin='6'/>
      <Label Content='Export folder:' VerticalAlignment='Center' Margin='6'/>
      <TextBox x:Name='txtExportFolder' Width='360' Margin='6'/>
      <Button x:Name='btnOpenExport' Content='Open' Width='70' Margin='6'/>
      <Button x:Name='btnExport' Content='Export' Width='90' Margin='12,0,0,0'/>
    </StackPanel>

    <TabControl Grid.Row='2'>
      <TabItem Header='Results'>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height='*'/>
            <RowDefinition Height='200'/>
          </Grid.RowDefinitions>
          <DataGrid x:Name='dgResults' Grid.Row='0' AutoGenerateColumns='True' IsReadOnly='True'/>
          <WindowsFormsHost Grid.Row='1' x:Name='hostChart'/>
        </Grid>
      </TabItem>

      <TabItem Header='Presets'>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height='Auto'/>
            <RowDefinition Height='*'/>
          </Grid.RowDefinitions>
          <StackPanel Orientation='Horizontal' Margin='6'>
            <Button x:Name='btnRunPreset' Content='Run Selected Preset' Width='150' Margin='0,0,8,0'/>
            <Button x:Name='btnRefreshPresets' Content='Refresh Presets' Width='120'/>
          </StackPanel>
          <ListBox x:Name='lstPresets' Grid.Row='1' Margin='6'/>
        </Grid>
      </TabItem>

      <TabItem Header='Scheduler'>
        <StackPanel Margin='8'>
          <Label Content='Scheduled Report Settings (Nightly)'/>
          <StackPanel Orientation='Horizontal' Margin='0,6,0,6'>
            <Label Content='Preset(s) (semicolon-separated):' VerticalAlignment='Center'/>
            <TextBox x:Name='txtSchedulePresets' Width='420' Margin='6'/>
          </StackPanel>
          <StackPanel Orientation='Horizontal' Margin='0,6,0,6'>
            <Label Content='Export Formats (csv;html;excel):' VerticalAlignment='Center'/>
            <TextBox x:Name='txtScheduleFormats' Width='240' Margin='6'/>
            <Label Content='Time (HH:mm):' VerticalAlignment='Center' Margin='12,0,0,0'/>
            <TextBox x:Name='txtScheduleTime' Width='80' Margin='6' Text='02:00'/>
          </StackPanel>
          <StackPanel Orientation='Horizontal' Margin='0,6,0,6'>
            <Button x:Name='btnInstallSchedule' Content='Install Schedule' Width='140' Margin='6'/>
            <Button x:Name='btnUninstallSchedule' Content='Uninstall Schedule' Width='140' Margin='6'/>
          </StackPanel>
        </StackPanel>
      </TabItem>

      <TabItem Header='Settings'>
        <StackPanel Margin='8'>
          <Label Content='Save / Load Settings' />
          <WrapPanel>
            <Button x:Name='btnSaveSettings' Content='Save Settings' Width='120' Margin='4'/>
            <Button x:Name='btnLoadSettings' Content='Load Settings' Width='120' Margin='4'/>
            <Button x:Name='btnClearSettings' Content='Clear Settings' Width='120' Margin='4'/>
          </WrapPanel>
          <Separator Margin='6'/>
          <Label Content='Optional: Store alternate credentials (protected to current user)'/>
          <WrapPanel>
            <Label Content='Username:' VerticalAlignment='Center'/>
            <TextBox x:Name='txtAltUser' Width='240' Margin='6'/>
            <Label Content='Password:' VerticalAlignment='Center'/>
            <PasswordBox x:Name='txtAltPass' Width='240' Margin='6'/>
            <Button x:Name='btnSaveCred' Content='Save Creds' Width='120' Margin='6'/>
            <Button x:Name='btnClearCred' Content='Clear Creds' Width='100' Margin='6'/>
          </WrapPanel>
        </StackPanel>
      </TabItem>

    </TabControl>

    <StatusBar Grid.Row='3' Height='26'>
      <StatusBarItem><TextBlock x:Name='txtStatus' Text='Ready.'/></StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
"@

# Load XAML
[xml]$xamlObj = $XAML
$reader = (New-Object System.Xml.XmlNodeReader $xamlObj)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Helper to find elements
function Get-Element([string]$name) { return $window.FindName($name) }

# Map UI elements
$cmbCategory = Get-Element "cmbCategory"
$txtFilter = Get-Element "txtFilter"
$btnRun = Get-Element "btnRun"
$btnClear = Get-Element "btnClear"
$chkResolveSIDs = Get-Element "chkResolveSIDs"
$chkEventLookup = Get-Element "chkEventLookup"
$txtExportFolder = Get-Element "txtExportFolder"
$btnOpenExport = Get-Element "btnOpenExport"
$btnExport = Get-Element "btnExport"
$dgResults = Get-Element "dgResults"
$hostChart = Get-Element "hostChart"
$lstPresets = Get-Element "lstPresets"
$btnRunPreset = Get-Element "btnRunPreset"
$btnRefreshPresets = Get-Element "btnRefreshPresets"
$txtSchedulePresets = Get-Element "txtSchedulePresets"
$txtScheduleFormats = Get-Element "txtScheduleFormats"
$txtScheduleTime = Get-Element "txtScheduleTime"
$btnInstallSchedule = Get-Element "btnInstallSchedule"
$btnUninstallSchedule = Get-Element "btnUninstallSchedule"
$btnSaveSettings = Get-Element "btnSaveSettings"
$btnLoadSettings = Get-Element "btnLoadSettings"
$btnClearSettings = Get-Element "btnClearSettings"
$txtAltUser = Get-Element "txtAltUser"
$txtAltPass = Get-Element "txtAltPass"
$btnSaveCred = Get-Element "btnSaveCred"
$btnClearCred = Get-Element "btnClearCred"
$txtStatus = Get-Element "txtStatus"

# Set defaults
$cmbCategory.ItemsSource = @("User","Computer","OU","GPO","Security Group","Service Accounts","Servers (by OS or group)","Workstations (by OS or group)","Locked-out Users (basic)","Locked-out Users (with origin/event lookup)","Subnets (AD Sites & Services)","Firewall (GPO firewall rules)","All: Users+Computers")
$cmbCategory.SelectedIndex = 0
$txtFilter.Text = "*"
$txtExportFolder.Text = $Global:ExportFolder

# Create Windows Forms Chart and attach
$chartHost = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chartHost.Width = 1000
$chartHost.Height = 200
$area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea "Main"
$chartHost.ChartAreas.Add($area)
$series = New-Object System.Windows.Forms.DataVisualization.Charting.Series "Series1"
$series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
$chartHost.Series.Add($series)
$hostChart.Child = $chartHost

# Populate presets list
function Refresh-PresetsUI {
    $lstPresets.Items.Clear()
    foreach ($p in $Global:PresetsArray) { $lstPresets.Items.Add($p.Name) | Out-Null }
}
Refresh-PresetsUI

# UI helpers
function Show-Results {
    param([array]$results)
    if (-not $results -or $results.Count -eq 0) {
        $dgResults.ItemsSource = $null
        $txtStatus.Text = "No results."
        $chartHost.Series["Series1"].Points.Clear()
        return
    }
    $dgResults.ItemsSource = $results
    $txtStatus.Text = "Found $($results.Count) item(s)."
    # Simple chart update for computers OS distribution
    try {
        $chartHost.Series["Series1"].Points.Clear()
        if ($results -and ($results[0].PSObject.Properties.Name -contains "OperatingSystem")) {
            $groups = $results | Group-Object -Property OperatingSystem | Sort-Object Count -Descending | Select-Object -First 10
            foreach ($g in $groups) {
                $pt = $chartHost.Series["Series1"].Points.Add($g.Count)
                $pt.AxisLabel = $g.Name
            }
        }
    } catch { }
}

# Main search dispatcher (used by GUI)
function Run-Search {
    param([string]$Category, [string]$Filter, [switch]$ResolveOrigin)
    $txtStatus.Text = "Running search..."
    $window.Dispatcher.Invoke([action]{},[System.Windows.Threading.DispatcherPriority]::Background)
    $results = @()
    try {
        switch ($Category) {
            "User" { $results = Search-Users -filter $Filter }
            "Computer" { $results = Search-Computers -filter $Filter }
            "OU" { $results = Search-OUs -filter $Filter }
            "GPO" {
                if (-not $HasGPO) { [System.Windows.MessageBox]::Show("GroupPolicy module not available.","Missing Module") ; $results = @() } else { $results = Search-GPOs -filter $Filter }
            }
            "Security Group" { $results = Search-Groups -filter $Filter }
            "Service Accounts" { $results = Search-ServiceAccounts -filter $Filter }
            "Servers (by OS or group)" { $results = Search-ServersOrWorkstations -filter $Filter -Servers }
            "Workstations (by OS or group)" { $results = Search-ServersOrWorkstations -filter $Filter }
            "Locked-out Users (basic)" { $results = Search-LockedOutUsers -ResolveOrigin:$false }
            "Locked-out Users (with origin/event lookup)" { $results = Search-LockedOutUsers -ResolveOrigin:$ResolveOrigin }
            "Subnets (AD Sites & Services)" { $results = Search-Subnets }
            "Firewall (GPO firewall rules)" { if (-not $HasGPO) { [System.Windows.MessageBox]::Show("GroupPolicy module not available.","Missing Module"); $results = @() } else { $results = Search-GPOFirewallRules -filter $Filter } }
            "All: Users+Computers" { $results = @(Search-Users -filter $Filter) + @(Search-Computers -filter $Filter) }
            default { $results = @() }
        }
    } catch {
        [System.Windows.MessageBox]::Show("Search error: $($_.Exception.Message)","Error")
        $txtStatus.Text = "Error during search."
        return @()
    }
    return $results
}

# -----------------------------
# UI Event Handlers
# -----------------------------
$btnRun.Add_Click({
    $category = $cmbCategory.SelectedItem
    $filter = $txtFilter.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filter)) { $filter = "*" }
    $resolve = $chkEventLookup.IsChecked
    $res = Run-Search -Category $category -Filter $filter -ResolveOrigin:$resolve
    Show-Results -results $res
})

$btnClear.Add_Click({
    $txtFilter.Text = "*"
    $dgResults.ItemsSource = $null
    $txtStatus.Text = "Cleared."
    $chartHost.Series["Series1"].Points.Clear()
})

$btnRunPreset.Add_Click({
    $sel = $lstPresets.SelectedItem
    if (-not $sel) { [System.Windows.MessageBox]::Show("Select a preset first.","Presets") ; return }
    $preset = $Global:PresetsArray | Where-Object { $_.Name -eq $sel } | Select-Object -First 1
    if (-not $preset) { [System.Windows.MessageBox]::Show("Preset not found.","Presets") ; return }
    try {
        $sb = $preset.ScriptBlock
        $res = if ($sb.Parameters.Count -gt 0) { & $sb $preset.Filter } else { & $sb }
        $arr = @(); foreach ($r in $res) { $arr += $r }
        Show-Results -results $arr
    } catch {
        [System.Windows.MessageBox]::Show("Preset execution error: $($_.Exception.Message)","Presets")
    }
})

$btnRefreshPresets.Add_Click({ Refresh-PresetsUI })

$btnOpenExport.Add_Click({
    $p = $txtExportFolder.Text
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
    Start-Process -FilePath $p
})

$btnExport.Add_Click({
    $items = $dgResults.ItemsSource
    if (-not $items) { [System.Windows.MessageBox]::Show("No results to export.","Export") ; return }
    $arr = @(); foreach ($i in $items) { $arr += $i }
    $formats = @("csv","html","excel")
    Export-Results -Results $arr -Category $cmbCategory.SelectedItem -Filter $txtFilter.Text -ExportPath $txtExportFolder.Text -Formats $formats
    [System.Windows.MessageBox]::Show("Export completed to: $($txtExportFolder.Text)","Export")
})

$btnInstallSchedule.Add_Click({
    try {
        $scriptPath = $PSCommandPath
        if (-not $scriptPath -or -not (Test-Path $scriptPath)) {
            [System.Windows.MessageBox]::Show("Unable to determine script path. Please save the script first.","Scheduler")
            return
        }
        $presetNames = ($txtSchedulePresets.Text -split ";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($presetNames.Count -eq 0) { $presetNames = @("Disabled Accounts","Inactive Computers (90d)","Domain Controllers") }
        $formats = ($txtScheduleFormats.Text -split ";") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($formats.Count -eq 0) { $formats = @("csv","html","excel") }
        $time = $txtScheduleTime.Text
        Install-ScheduledReport -TaskName "ADSearchTool_Nightly" -ScriptPath $scriptPath -ExportPath $txtExportFolder.Text -PresetNames $presetNames -Formats $formats -RunTime $time
        [System.Windows.MessageBox]::Show("Scheduled task installed (daily at $time).","Scheduler")
    } catch { [System.Windows.MessageBox]::Show("Install failed: $($_.Exception.Message)","Scheduler") }
})

$btnUninstallSchedule.Add_Click({
    try { 
        Uninstall-ScheduledReport -TaskName "ADSearchTool_Nightly"
        [System.Windows.MessageBox]::Show("Scheduled task removed.","Scheduler") 
    } catch { 
        [System.Windows.MessageBox]::Show("Uninstall failed: $($_.Exception.Message)","Scheduler") 
    }
})

$btnSaveSettings.Add_Click({
    $cfg = [ordered]@{}
    $cfg.LastCategory = $cmbCategory.SelectedItem
    $cfg.LastFilter = $txtFilter.Text
    $cfg.ExportFolder = $txtExportFolder.Text
    $cfg.SchedulePresets = $txtSchedulePresets.Text
    $cfg.ScheduleFormats = $txtScheduleFormats.Text
    $cfg.ScheduleTime = $txtScheduleTime.Text
    if ($txtAltUser.Text -and $txtAltPass.Password) {
        $cred = New-Object System.Management.Automation.PSCredential ($txtAltUser.Text, (ConvertTo-SecureString $txtAltPass.Password -AsPlainText -Force))
        $cfg.CredProtected = Protect-Credential -Credential $cred
    }
    Save-Config -cfg $cfg
    [System.Windows.MessageBox]::Show("Settings saved to $Global:ConfigFile","Settings")
})

$btnLoadSettings.Add_Click({
    $cfg = Load-Config
    if (-not $cfg) { [System.Windows.MessageBox]::Show("No saved settings.","Settings"); return }
    if ($cfg.LastCategory) { $cmbCategory.SelectedItem = $cfg.LastCategory }
    if ($cfg.LastFilter) { $txtFilter.Text = $cfg.LastFilter }
    if ($cfg.ExportFolder) { $txtExportFolder.Text = $cfg.ExportFolder }
    if ($cfg.SchedulePresets) { $txtSchedulePresets.Text = $cfg.SchedulePresets }
    if ($cfg.ScheduleFormats) { $txtScheduleFormats.Text = $cfg.ScheduleFormats }
    if ($cfg.ScheduleTime) { $txtScheduleTime.Text = $cfg.ScheduleTime }
    if ($cfg.CredProtected) {
        $cred = Unprotect-Credential -ProtectedString $cfg.CredProtected
        if ($cred) { $txtAltUser.Text = $cred.UserName }
    }
    [System.Windows.MessageBox]::Show("Settings loaded.","Settings")
})

$btnClearSettings.Add_Click({
    if (Test-Path $Global:ConfigFile) { Remove-Item $Global:ConfigFile -Force }
    [System.Windows.MessageBox]::Show("Settings file removed.","Settings")
})

$btnSaveCred.Add_Click({
    if (-not $txtAltUser.Text -or -not $txtAltPass.Password) { 
        [System.Windows.MessageBox]::Show("Enter username and password to save.","Credentials")
        return 
    }
    $cred = New-Object System.Management.Automation.PSCredential ($txtAltUser.Text, (ConvertTo-SecureString $txtAltPass.Password -AsPlainText -Force))
    $cfg = Load-Config
    if (-not $cfg) { $cfg = [ordered]@{} }
    $cfg.CredProtected = Protect-Credential -Credential $cred
    Save-Config -cfg $cfg
    [System.Windows.MessageBox]::Show("Credentials saved (protected).","Credentials")
})

$btnClearCred.Add_Click({
    $cfg = Load-Config
    if ($cfg -and $cfg.CredProtected) { 
        $cfg.PSObject.Properties.Remove('CredProtected')
        Save-Config -cfg $cfg
        $txtAltUser.Text = ""
        $txtAltPass.Password = ""
        [System.Windows.MessageBox]::Show("Stored credentials cleared.","Credentials") 
    } else { 
        [System.Windows.MessageBox]::Show("No stored credentials found.","Credentials") 
    }
})

# Keyboard shortcut F5 for Run
$window.Add_KeyDown({
    param($s,$e)
    if ($e.Key.ToString() -eq 'F5') { $btnRun.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) }
})

# Load settings into UI on start if present
$config = Load-Config
if ($config) {
    if ($config.LastFilter) { $txtFilter.Text = $config.LastFilter }
    if ($config.LastCategory) { $cmbCategory.SelectedItem = $config.LastCategory }
    if ($config.ExportFolder) { $txtExportFolder.Text = $config.ExportFolder }
    if ($config.SchedulePresets) { $txtSchedulePresets.Text = $config.SchedulePresets }
    if ($config.ScheduleFormats) { $txtScheduleFormats.Text = $config.ScheduleFormats }
    if ($config.ScheduleTime) { $txtScheduleTime.Text = $config.ScheduleTime }
    if ($config.CredProtected) {
        $cred = Unprotect-Credential -ProtectedString $config.CredProtected
        if ($cred) { $txtAltUser.Text = $cred.UserName }
    }
}

# Show window
$window.ShowDialog() | Out-Null

### END FULL SCRIPT
