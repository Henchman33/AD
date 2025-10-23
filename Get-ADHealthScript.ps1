<#
    .SYNOPSIS
    Get-ADHealth.ps1 - Domain Controller Health Check Script.

    .DESCRIPTION
    This script performs a list of common health checks to a specific domain, or the entire forest. The results are then compiled into a colour coded HTML report.

    .OUTPUTS
    The results are currently only output to HTML for email or as an HTML report file, or sent as an SMTP message with an HTML body.

    .PARAMETER DomainName
    Perform a health check on a specific Active Directory domain.

    .PARAMETER ReportFile
    Output the report details to a file in the current directory.

    .PARAMETER SendEmail
    Send the report via email. You have to configure the correct SMTP settings.

    .EXAMPLE
    .\Get-ADHealthScript.ps1 -ReportFile
    Checks all domains and all domain controllers in your current forest and creates a report.

    .EXAMPLE
    .\Get-ADHealth.ps1 -DomainName XXX.com -ReportFile
    Checks all the domain controllers in the specified domain "XXX.com" and creates a report.

    .EXAMPLE
    .\Get-ADHealth.ps1 -DomainName XXX.com -SendEmail
    Checks all the domain controllers in the specified domain "XXX.com" and sends the resulting report as an email message.

    .NOTES
    Written by: Steve McKee
    Website:    https://github.com/Henchman33
#>

[CmdletBinding()]
Param(
    [Parameter( Mandatory = $false)]
    [string]$DomainName,

    [Parameter( Mandatory = $false)]
    [switch]$ReportFile,

    [Parameter( Mandatory = $false)]
    [switch]$SendEmail
)

#...................................
# Global Variables
#...................................

$allTestedDomainControllers = [System.Collections.Generic.List[Object]]::new()
$allDomainControllers = [System.Collections.Generic.List[Object]]::new()
$now = Get-Date
$date = $now.ToShortDateString()
$reportTime = $now
$reportFileNameTime = $now.ToString("yyyyMMdd_HHmmss")
$reportemailsubject = "Domain Controller Health Report"

$smtpsettings = @{
    To         = 'email@domain.com'
    From       = 'adhealth@yourdomain.com'
    Subject    = "$reportemailsubject - $date"
    SmtpServer = "mail.domain.com"
    Port       = "25"
    #Credential = (Get-Credential)
    #UseSsl     = $true
}

#...................................
# Functions
#...................................

# This function gets all the domains in the forest.
Function Get-AllDomains() {
    Write-Verbose "Running function Get-AllDomains"
    $allDomains = (Get-ADForest).Domains
    return $allDomains
}

# This function gets all the domain controllers in a specified domain.
Function Get-AllDomainControllers ($ComputerName) {
    Write-Verbose "Running function Get-AllDomainControllers"
    $allDomainControllers = Get-ADDomainController -Filter * -Server $ComputerName | Sort-Object HostName
    return $allDomainControllers
}

# This function tests the domain controller against DNS.
Function Get-DomainControllerNSLookup($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerNSLookup"
    try {
        $domainControllerNSLookupResult = Resolve-DnsName $ComputerName -Type A | Select-Object -ExpandProperty IPAddress
        $domainControllerNSLookupResult = 'Success'
    }
    catch {
        $domainControllerNSLookupResult = 'Fail'
    }
    return $domainControllerNSLookupResult
}

# This function tests the connectivity to the domain controller.
Function Get-DomainControllerPingStatus($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerPingStatus"
    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        $domainControllerPingStatus = "Success"
    }
    else {
        $domainControllerPingStatus = 'Fail'
    }
    return $domainControllerPingStatus
}

# This function tests the domain controller uptime.
Function Get-DomainControllerUpTime($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerUpTime"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $W32OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction SilentlyContinue
            $timespan = (Get-Date) - $W32OS.LastBootUpTime
            [int]$uptime = "{0:00}" -f $timespan.TotalHours
        }
        catch {
            $uptime = 'CIM Failure'
        }
    }
    else {
        $uptime = 'Fail'
    }
    return $uptime
}

# This function checks the time synchronization offset.
function Get-TimeDifference($ComputerName) {
    Write-Verbose "Running function Get-TimeDifference"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $currentTime, $timeDifference = (& w32tm /stripchart /computer:$ComputerName /samples:1 /dataonly)[-1].Trim("s") -split ',\s*'
            $diff = [double]$timeDifference
            $diffRounded = [Math]::Round($diff, 1, [MidPointRounding]::AwayFromZero)
        }
        catch {
            $diffRounded = 'Fail'
        }
    }
    else {
        $diffRounded = 'Fail'
    }
    return $diffRounded
}

# This function checks the DNS, NTDS and Netlogon services.
Function Get-DomainControllerServices($ComputerName) {
    Write-Verbose "Running function DomainControllerServices"
    $thisDomainControllerServicesTestResult = [PSCustomObject]@{
        DNSService      = $null
        NTDSService     = $null
        NETLOGONService = $null
    }

    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        if ((Get-Service -ComputerName $ComputerName -Name DNS -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.DNSService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.DNSService = 'Fail'
        }
        if ((Get-Service -ComputerName $ComputerName -Name NTDS -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.NTDSService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.NTDSService = 'Fail'
        }
        if ((Get-Service -ComputerName $ComputerName -Name netlogon -ErrorAction SilentlyContinue).Status -eq 'Running') {
            $thisDomainControllerServicesTestResult.NETLOGONService = 'Success'
        }
        else {
            $thisDomainControllerServicesTestResult.NETLOGONService = 'Fail'
        }
    }
    else {
        $thisDomainControllerServicesTestResult.DNSService = 'Fail'
        $thisDomainControllerServicesTestResult.NTDSService = 'Fail'
        $thisDomainControllerServicesTestResult.NETLOGONService = 'Fail'
    }
    return $thisDomainControllerServicesTestResult
}

# This function runs the DCDiag tests and saves them in a variable for later processing.
Function Get-DomainControllerDCDiagTestResults($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerDCDiagTestResults"

    # Initialize the object with all properties set to null
    $DCDiagTestResults = [PSCustomObject]@{
        ServerName         = $ComputerName
        Connectivity       = $null
        Advertising        = $null
        FrsEvent           = $null
        DFSREvent          = $null
        SysVolCheck        = $null
        KccEvent           = $null
        KnowsOfRoleHolders = $null
        MachineAccount     = $null
        NCSecDesc          = $null
        NetLogons          = $null
        ObjectsReplicated  = $null
        Replications       = $null
        RidManager         = $null
        Services           = $null
        SystemLog          = $null
        VerifyReferences   = $null
        CheckSDRefDom      = $null
        CrossRefValidation = $null
        LocatorCheck       = $null
        Intersite          = $null
        FSMOCheck          = $null
    }

    if ((Test-Connection $ComputerName -Count 1 -quiet) -eq $True) {
        # Define an array of parameters for Dcdiag.exe
        $params = @(
            "/s:$ComputerName",
            "/test:Connectivity",
            "/test:Advertising",
            "/test:FrsEvent",
            "/test:DFSREvent",
            "/test:SysVolCheck",
            "/test:KccEvent",
            "/test:KnowsOfRoleHolders",
            "/test:MachineAccount",
            "/test:NCSecDesc",
            "/test:NetLogons",
            "/test:ObjectsReplicated",
            "/test:Replications",
            "/test:RidManager",
            "/test:Services",
            "/test:SystemLog",
            "/test:VerifyReferences",
            "/test:CheckSDRefDom",
            "/test:CrossRefValidation",
            "/test:LocatorCheck",
            "/test:Intersite",
            "/test:FSMOCheck"
        )

        $DCDiagTest = (Dcdiag.exe @params) -split ('[\r\n]')

        $TestName = $null
        $TestStatus = $null

        $DCDiagTest | ForEach-Object {
            switch -Regex ($_) {
                "Starting test:" {
                    $TestName = ($_ -replace ".*Starting test:").Trim()
                }
                "passed test|failed test" {
                    $TestStatus = if ($_ -match "passed test") { "Passed" } else { "Failed" }
                }
            }
            if ($TestName -and $TestStatus) {
                # Set the property value directly
                $DCDiagTestResults.$TestName = $TestStatus
                $TestName = $null
                $TestStatus = $null
            }
        }
    }
    else {
        # If the domain controller is not reachable, set all tests to 'Failed'
        foreach ($property in $DCDiagTestResults.PSObject.Properties.Name) {
            if ($property -ne "ServerName") {
                $DCDiagTestResults.$property = "Failed"
            }
        }
    }
    return $DCDiagTestResults
}

# This function checks the free space in percentage on the OS drive
Function Get-DomainControllerOSDriveFreeSpace ($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerOSDriveFreeSpace"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $thisOSDriveLetter = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop).SystemDrive
            $thisOSDiskDrive = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$thisOSDriveLetter'" -ErrorAction Stop
            $thisOSPercentFree = [math]::Round($thisOSDiskDrive.FreeSpace / $thisOSDiskDrive.Size * 100)
        }
        catch {
            $thisOSPercentFree = 'CIM Failure'
        }
    }
    else {
        $thisOSPercentFree = "Fail"
    }
    return $thisOSPercentFree
}

# This function checks the free disk space on the OS drive in GB
Function Get-DomainControllerOSDriveFreeSpaceGB ($ComputerName) {
    Write-Verbose "Running function Get-DomainControllerOSDriveFreeSpaceGB"
    if ((Test-Connection $ComputerName -Count 1 -Quiet) -eq $True) {
        try {
            $thisOSDriveLetter = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop).SystemDrive
            $thisOSDiskDrive = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$thisOSDriveLetter'" -ErrorAction Stop
            # Convert bytes to GB, rounding to 2 decimal places
            $freeSpaceGB = [math]::Round($thisOSDiskDrive.FreeSpace / 1GB, 2)
        }
        catch {
            $freeSpaceGB = 'CIM Failure'
        }
    }
    else {
        $freeSpaceGB = 'Fail'
    }
    return $freeSpaceGB
}

# This function generates HTML code from the results of the above functions.
Function New-ServerHealthHTMLTableCell() {
    param( $lineitem )
    $htmltablecell = $null

    switch ($($reportline."$lineitem")) {
        "Success" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Passed" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Pass" { $htmltablecell = "<td class=""pass"">$($reportline."$lineitem")</td>" }
        "Warn" { $htmltablecell = "<td class=""warn"">$($reportline."$lineitem")</td>" }
        "Fail" { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        "Failed" { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        "Could not test server uptime." { $htmltablecell = "<td class=""fail"">$($reportline."$lineitem")</td>" }
        default { $htmltablecell = "<td>$($reportline."$lineitem")</td>" }
    }
    return $htmltablecell
}

if (!($DomainName)) {
    Write-Host "No domain specified, using all domains in forest" -ForegroundColor Yellow
    $allDomains = Get-AllDomains
    $reportFileName = 'forest_health_report_' + (Get-ADForest).name + '_' + $reportFileNameTime + '.html'
}
else {
    Write-Host "Domain name specified on cmdline" -ForegroundColor Cyan
    $allDomains = $DomainName
    $reportFileName = 'dc_health_report_' + $DomainName + '_' + $reportFileNameTime + '.html'
}

foreach ($domain in $allDomains) {
    Write-Host "Testing domain" $domain -ForegroundColor Green
    $allDomainControllers = Get-AllDomainControllers $domain

    # Force array first, then get count
    $allDomainControllers = @($allDomainControllers)
    $totalDCs = $allDomainControllers.Count

    # Initialize counter for display
    $currentDCNumber = 0

    foreach ($domainController in $allDomainControllers) {
        $currentDCNumber++
        $stopWatch = [system.diagnostics.stopwatch]::StartNew()
        Write-Host "Testing domain controller ($currentDCNumber of $totalDCs) $($domainController.HostName)" -ForegroundColor Cyan
        $DCDiagTestResults = Get-DomainControllerDCDiagTestResults $domainController.HostName

        $thisDomainController = [PSCustomObject]@{
            Server                            = ($domainController.HostName).ToLower()
            Site                              = $domainController.Site
            "OS Version"                      = $domainController.OperatingSystem
            "IPv4 Address"                    = $domainController.IPv4Address
            "Operation Master Roles"          = $domainController.OperationMasterRoles
            "DNS"                             = Get-DomainControllerNSLookup $domainController.HostName
            "Ping"                            = Get-DomainControllerPingStatus $domainController.HostName
            "Uptime (hours)"                  = Get-DomainControllerUpTime $domainController.HostName
            "OS Free Space (%)"               = Get-DomainControllerOSDriveFreeSpace $domainController.HostName
            "OS Free Space (GB)"              = Get-DomainControllerOSDriveFreeSpaceGB $domainController.HostName
            "Time offset (seconds)"           = Get-TimeDifference $domainController.HostName
            "DNS Service"                     = (Get-DomainControllerServices $domainController.HostName).DNSService
            "NTDS Service"                    = (Get-DomainControllerServices $domainController.HostName).NTDSService
            "NetLogon Service"                = (Get-DomainControllerServices $domainController.HostName).NETLOGONService
            "DCDIAG: Connectivity"            = $DCDiagTestResults.Connectivity
            "DCDIAG: Advertising"             = $DCDiagTestResults.Advertising
            "DCDIAG: FrsEvent"                = $DCDiagTestResults.FrsEvent
            "DCDIAG: DFSREvent"               = $DCDiagTestResults.DFSREvent
            "DCDIAG: SysVolCheck"             = $DCDiagTestResults.SysVolCheck
            "DCDIAG: KccEvent"                = $DCDiagTestResults.KccEvent
            "DCDIAG: FSMO KnowsOfRoleHolders" = $DCDiagTestResults.KnowsOfRoleHolders
            "DCDIAG: MachineAccount"          = $DCDiagTestResults.MachineAccount
            "DCDIAG: NCSecDesc"               = $DCDiagTestResults.NCSecDesc
            "DCDIAG: NetLogons"               = $DCDiagTestResults.NetLogons
            "DCDIAG: ObjectsReplicated"       = $DCDiagTestResults.ObjectsReplicated
            "DCDIAG: Replications"            = $DCDiagTestResults.Replications
            "DCDIAG: RidManager"              = $DCDiagTestResults.RidManager
            "DCDIAG: Services"                = $DCDiagTestResults.Services
            "DCDIAG: SystemLog"               = $DCDiagTestResults.SystemLog
            "DCDIAG: VerifyReferences"        = $DCDiagTestResults.VerifyReferences
            "DCDIAG: CheckSDRefDom"           = $DCDiagTestResults.CheckSDRefDom
            "DCDIAG: CrossRefValidation"      = $DCDiagTestResults.CrossRefValidation
            "DCDIAG: LocatorCheck"            = $DCDiagTestResults.LocatorCheck
            "DCDIAG: Intersite"               = $DCDiagTestResults.Intersite
            "DCDIAG: FSMO Check"              = $DCDiagTestResults.FSMOCheck
            "Processing Time (seconds)"       = $stopWatch.Elapsed.Seconds
        }

        $allTestedDomainControllers.Add($thisDomainController)
        $totalDCtoProcessCounter--
    }
}

# Common HTML head and styles
$htmlhead = "<html>
        <style>
        BODY { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 10pt; }
        H1 { font-size: 20px; }
        H2 { font-size: 16px; }
        H3 { font-size: 14px; }
        TABLE { border: 1px solid #ccc; border-collapse: collapse; font-size: 10pt;}
            TH { border: 1px solid #ccc; background: #f2f2f2; padding: 10px; color: #000000;}
                TD { border: 1px solid #ccc; padding: 10px; }
                    td.pass { background: #6BBF59;}
                        td.warn { background: #FFD966;}
                            td.fail { background: #D9534F; color: #ffffff;}
                                td.info { background: #5BC0DE;}
                                    </style>
                                    <body>
                                    <h1 align=""left"">Domain Controller Health Check Report</h1>
                                    <h3 align=""left"">Generated: $reportTime </h3>"

# Domain Controller Health Report Table Header
$htmltableheader = "<h3>Domain Controller Health Summary</h3>
                                    <h3>Forest: $((Get-ADForest).Name)</h3>
                                    <p>
                                    <table style=""width: 100%; border-collapse: separate; "">
                                    <tr>
                                    <th style=""text-align: center"">Server</th>
                                    <th style=""text-align: center; "">Site</th>
                                    <th style=""text-align: center; "">OS Version</th>
                                    <th style=""text-align: center; "">IPv4 Address</th>
                                    <th style=""text-align: center; "">Operation Master Roles</th>
                                    <th style=""text-align: center; "">DNS</th>
                                    <th style=""text-align: center; "">Ping</th>
                                    <th style=""text-align: center; "">Uptime (hours)</th>
                                    <th style=""text-align: center; "">OS Free Space (%)</th>
                                    <th style=""text-align: center; "">OS Free Space (GB)</th>
                                    <th style=""text-align: center; "">Time offset (seconds)</th>
                                    <th style=""text-align: center; "">DNS Service</th>
                                    <th style=""text-align: center; "">NTDS Service</th>
                                    <th style=""text-align: center; "">NetLogon Service</th>
                                    <th style=""text-align: center; "">DCDIAG: Connectivity</th>
                                    <th style=""text-align: center; "">DCDIAG: Advertising</th>
                                    <th style=""text-align: center; "">DCDIAG: FrsEvent</th>
                                    <th style=""text-align: center; "">DCDIAG: DFSREvent</th>
                                    <th style=""text-align: center; "">DCDIAG: SysVolCheck</th>
                                    <th style=""text-align: center; "">DCDIAG: KccEvent</th>
                                    <th style=""text-align: center; "">DCDIAG: FSMO KnowsOfRoleHolders</th>
                                    <th style=""text-align: center; "">DCDIAG: MachineAccount</th>
                                    <th style=""text-align: center; "">DCDIAG: NCSecDesc</th>
                                    <th style=""text-align: center; "">DCDIAG: NetLogons</th>
                                    <th style=""text-align: center; "">DCDIAG: ObjectsReplicated</th>
                                    <th style=""text-align: center; "">DCDIAG: Replications</th>
                                    <th style=""text-align: center; "">DCDIAG: RidManager</th>
                                    <th style=""text-align: center; "">DCDIAG: Services</th>
                                    <th style=""text-align: center; "">DCDIAG: SystemLog</th>
                                    <th style=""text-align: center; "">DCDIAG: VerifyReferences</th>
                                    <th style=""text-align: center; "">DCDIAG: CheckSDRefDom</th>
                                    <th style=""text-align: center; "">DCDIAG: CrossRefValidation</th>
                                    <th style=""text-align: center; "">DCDIAG: LocatorCheck</th>
                                    <th style=""text-align: center; "">DCDIAG: Intersite</th>
                                    <th style=""text-align: center; "">DCDIAG: FSMO Check</th>
                                    <th style=""text-align: center; "">Processing Time (seconds)</th>
                                    </tr>"

# Domain Controller Health Report Table
$serverhealthhtmltable = $serverhealthhtmltable + $htmltableheader

# This section will process through the $allTestedDomainControllers array object and create and colour the HTML table based on certain conditions.
foreach ($reportline in $allTestedDomainControllers) {

    if (Test-Path variable:fsmoRoleHTML) {
        Remove-Variable fsmoRoleHTML
    }

    if (($reportline."Operation Master Roles").Count -gt 0) {
        $fsmoRoleHTML = ($reportline."Operation Master Roles" | ForEach-Object { "$_`r`n" }) -join '<br>'
    }
    else {
        $fsmoRoleHTML = 'None<br>'
    }

    $htmltablerow = "<tr>"
    $htmltablerow += "<td>$($reportline.Server)</td>"
    $htmltablerow += "<td>$($reportline.Site)</td>"
    $htmltablerow += "<td>$($reportline."OS Version")</td>"
    $htmltablerow += "<td>$($reportline."IPv4 Address")</td>"
    $htmltablerow += "<td>$fsmoRoleHTML</td>"
    $htmltablerow += (New-ServerHealthHTMLTableCell "DNS" )
    $htmltablerow += (New-ServerHealthHTMLTableCell "Ping")

    if ($($reportline."Uptime (hours)") -eq "CIM Failure") {
        $htmltablerow += "<td class=""warn"">Could not test server uptime.</td>"
    }
    elseif ($($reportline."Uptime (hours)") -eq "Fail") {
        $htmltablerow += "<td class=""fail"">Fail</td>"
    }
    else {
        $hours = [int]$($reportline."Uptime (hours)")
        if ($hours -le 24) {
            $htmltablerow += "<td class=""warn"">$hours</td>"
        }
        else {
            $htmltablerow += "<td class=""pass"">$hours</td>"
        }
    }

    $osSpace = $reportline."OS Free Space (%)"
    if ($osSpace -eq "CIM Failure") {
        $htmltablerow += "<td class=""warn"">Could not test server free space.</td>"
    }
    elseif ($osSpace -eq "Fail") {
        $htmltablerow += "<td class=""fail"">$osSpace</td>"
    }
    elseif ($osSpace -le 5) {
        $htmltablerow += "<td class=""fail"">$osSpace</td>"
    }
    elseif ($osSpace -le 30) {
        $htmltablerow += "<td class=""warn"">$osSpace</td>"
    }
    else {
        $htmltablerow += "<td class=""pass"">$osSpace</td>"
    }

    $osSpaceGB = $reportline."OS Free Space (GB)"
    if ($osSpaceGB -eq "CIM Failure") {
        $htmltablerow += "<td class=""warn"">Could not test server free space.</td>"
    }
    elseif ($osSpaceGB -eq "Fail") {
        $htmltablerow += "<td class=""fail"">$osSpaceGB</td>"
    }
    elseif ($osSpaceGB -lt 5) {
        $htmltablerow += "<td class=""fail"">$osSpaceGB</td>"
    }
    elseif ($osSpaceGB -lt 10) {
        $htmltablerow += "<td class=""warn"">$osSpaceGB</td>"
    }
    else {
        $htmltablerow += "<td class=""pass"">$osSpaceGB</td>"
    }

    $time = $reportline."Time offset (seconds)"
    if ($time -ge 1) {
        $htmltablerow += "<td class=""fail"">$time</td>"
    }
    else {
        $htmltablerow += "<td class=""pass"">$time</td>"
    }

    $htmltablerow += (New-ServerHealthHTMLTableCell "DNS Service")
    $htmltablerow += (New-ServerHealthHTMLTableCell "NTDS Service")
    $htmltablerow += (New-ServerHealthHTMLTableCell "NetLogon Service")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: Connectivity")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: Advertising")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: FrsEvent")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: DFSREvent")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: SysVolCheck")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: KccEvent")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: FSMO KnowsOfRoleHolders")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: MachineAccount")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: NCSecDesc")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: NetLogons")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: ObjectsReplicated")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: Replications")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: RidManager")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: Services")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: SystemLog")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: VerifyReferences")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: CheckSDRefDom")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: CrossRefValidation")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: LocatorCheck")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: Intersite")
    $htmltablerow += (New-ServerHealthHTMLTableCell "DCDIAG: FSMO Check")

    $processingTime = $reportline."Processing Time (seconds)"
    $htmltablerow += "<td>$processingTime</td>"

    [array]$serverhealthhtmltable += $htmltablerow
}

$serverhealthhtmltable += "</table></p>"
$htmltail = "* DNS test is performed using Resolve-DnsName. This cmdlet is only available from Windows 2012 onwards.
                                    </body>
                                    </html>"

$htmlreport = $htmlhead + $serverhealthhtmltable + $htmltail

if ($ReportFile) {
    $htmlreport | Out-File $reportFileName -Encoding UTF8
}

if ($SendEmail) {
    try {
        # Send email with both inline HTML and attachment
        $htmlreport | Out-File $reportFileName -Encoding UTF8
        Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Attachments $reportFileName -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction Stop
        Write-Host "Email sent successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to send email. Error: $_" -ForegroundColor Red
    }
}
