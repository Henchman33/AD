# A quick and dirty PowerShell script to search for a particular MAC address on Windows DHCP servers across all scopes.

$dhcpServers = @("dhcp-server-01","dhcp-server-02") # Define your servers here
$results = @()
$clientId = Read-Host -Prompt "Input the MAC address you want to search for"

foreach ($dhcpServer in $dhcpServers) {
    Write-Host "Searching $dhcpServer..."
    Get-DhcpServerv4Scope -ComputerName $dhcpServer | ForEach-Object {
        $results += Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $_.ScopeId -ClientId $clientId -ErrorAction SilentlyContinue
    }
    $results | ForEach-Object {$_ | Add-Member -membertype noteproperty -Name Server -Value $dhcpServer -ErrorAction SilentlyContinue}
}

Write-Output $results | select Server, IPAddress, ScopeId, ClientId, Name, Description | ft
