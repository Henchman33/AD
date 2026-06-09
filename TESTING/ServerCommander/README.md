  ## Server Commander All-In-One - Enterprise Server Management GUI for System Administrators

  A professional WPF-based all-in-one tool for daily server administration tasks:
    - Computer/Server Info (SystemInfo, DriverQuery, WMI Explorer integration)
    - Remote PowerShell Code Runner (single or multi-host, import from file)
    - Services, Processes, Event Logs, Disk management
    - Registry, Shares, Scheduled Tasks, WSUS/Windows Update status
    - Network diagnostics (ping, traceroute, open ports, DNS, netstat)
    - RDP Launcher and PSExec/PAExec integration
    - AD Computer lookup and SCCM/MECM quick queries
    - External Tools launcher (AdExplorer, WMIExplorer, PSExec, PAExec, sydi-server)
    - CMTrace-compatible logging
    - Dark/Light theme (Catppuccin Mocha-inspired dark + clean light)
    - Credential management per host/domain
    - Export results (CSV/JSON/TXT)
	 Download these tools to your workstation to the following paths
   PSExec       = "C:\Tools\PSExec.exe"
    PAExec       = "C:\Tools\PAExec.exe"
    AdExplorer   = "C:\Tools\AdExplorer.exe"
    WMIExplorer  = "C:\Tools\WmiExplorer.ps1"
    SydiServer   = "C:\Tools\sydi-server.vbs"
    SystemInfo   = "systeminfo.exe"
    DriverQuery  = "driverquery.exe"
    CMTrace      = "C:\Windows\CCM\CMTrace.exe"
.NOTES
  Author : Steve McKee
  Version: 1.0
  Requires: PowerShell 5.1, RSAT (optional), WinRM for remoting
  External : PSExec, PAExec, AdExplorer, WMIExplorer.ps1
  PSVersion: 5.1 ONLY
