# Define the target Domain Controller's hostname or IP address
$TargetDC = "10.209.22.20" #new dc name or ip address

# Define the essential AD ports based on Microsoft documentation
$ports = @(
    # Core AD and Authentication Services
    88,    # Kerberos Authentication [citation:2][citation:8]
    389,   # LDAP [citation:2][citation:8]
    636,   # LDAP over SSL [citation:2][citation:8]
    3268,  # Global Catalog (GC) [citation:2][citation:8]
    3269,  # Global Catalog over SSL [citation:2][citation:8]
    464,   # Kerberos Password Change [citation:2][citation:8]
    
    # RPC and Core Windows Services
    135,   # RPC Endpoint Mapper [citation:2][citation:8]
    445,   # SMB for replication, SYSVOL, and Group Policy [citation:2][citation:8]
    
    # Time Synchronization
    123,   # W32Time (NTP) [citation:2][citation:8]
    
    # Web Services for AD Management (Windows Server 2012 and later)
    9389,  # Active Directory Web Services (ADWS) [citation:2][citation:8]
    
    # DNS (Critical for domain name resolution)
    53     # DNS [citation:5][citation:2][citation:8]
)

Write-Host "Testing critical AD ports against $TargetDC..." -ForegroundColor Cyan

# Loop through each port and test the connection
foreach ($port in $ports) {
    # -InformationLevel Quiet makes the command return $true or $false
    $result = Test-NetConnection -ComputerName $TargetDC -Port $port -InformationLevel Quiet -WarningAction SilentlyContinue

    # Display the result
    if ($result) {
        Write-Host "Port $port`: SUCCESS" -ForegroundColor Green
    } else {
        Write-Host "Port $port`: FAILED" -ForegroundColor Red
    }
}

Write-Host "`nNote: Testing the ephemeral RPC port range (49152-65535) for services like LSA, SAM, NetLogon, and DFSR is not practical with a single test." -ForegroundColor Yellow
Write-Host "These ports must be allowed through the firewall but are assigned dynamically by the system." -ForegroundColor Yellow
