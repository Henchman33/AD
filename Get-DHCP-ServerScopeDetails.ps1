<#
.SYNOPSIS
    DHCP Inventory / Audit Script
.DESCRIPTION
    Collects DHCP Server, Scope, and Option data for auditing and inventory.
    Compatible with Windows PowerShell 5.0+ and ISE.
    Author: Stephen McKee - IGTPLC
#>

# -------------------------------
# CONFIGURATION
# -------------------------------
$DhcpServers = Get-DhcpServerInDC | Select-Object -ExpandProperty DnsName  # auto-detect DHCP servers registered in AD
# Or manually specify:
# $DhcpServers = "DHCP01","DHCP02","DHCP03"

$OutputFile  = "C:\DHCP_Inventory.csv"

# -------------------------------
# FUNCTION
# -------------------------------
function Get-DhcpInventory {
    param(
        [string[]]$Servers
    )

    $results = @()

    foreach ($serverName in $Servers) {
        Write-Host "Querying DHCP server: $serverName" -ForegroundColor Cyan

        try {
            # Server-level options
            $serverOptions = Get-DhcpServerv4OptionValue -ComputerName $serverName -ErrorAction Stop

            $server_dns    = ($serverOptions | Where-Object {$_.OptionId -eq 6}).Value
            $server_domain = ($serverOptions | Where-Object {$_.OptionId -eq 15}).Value
            $server_ntp    = ($serverOptions | Where-Object {$_.OptionId -eq 42}).Value

        } catch {
            Write-Warning "Unable to retrieve server-level options for ${serverName}: $_"
            $server_dns = $server_domain = $server_ntp = $null
        }

        try {
            $scopes = Get-DhcpServerv4Scope -ComputerName $serverName -ErrorAction Stop
        } catch {
            Write-Warning "Could not enumerate scopes on ${serverName}: $_"
            continue
        }

        foreach ($scope in $scopes) {
            $scopeId = $scope.ScopeId

            try {
                $scopeOptions = Get-DhcpServerv4OptionValue -ComputerName $serverName -ScopeId $scopeId -ErrorAction Stop

                $scope_dns    = ($scopeOptions | Where-Object {$_.OptionId -eq 6}).Value
                $scope_domain = ($scopeOptions | Where-Object {$_.OptionId -eq 15}).Value
                $scope_ntp    = ($scopeOptions | Where-Object {$_.OptionId -eq 42}).Value
                $scope_router = ($scopeOptions | Where-Object {$_.OptionId -eq 3}).Value

            } catch {
                Write-Warning "Unable to get options for ${serverName} scope ${scopeId}: $_"
                $scope_dns = $scope_domain = $scope_ntp = $scope_router = $null
            }

            # Prepare safe strings for CSV output
            $scopeDNSStr    = if ($scope_dns)    { ($scope_dns -join ",") }    else { "" }
            $scopeDomainStr = if ($scope_domain) { ($scope_domain -join ",") } else { "" }
            $scopeNTPStr    = if ($scope_ntp)    { ($scope_ntp -join ",") }    else { "" }
            $scopeRouterStr = if ($scope_router) { ($scope_router -join ",") } else { "" }

            $serverDNSStr   = if ($server_dns)   { ($server_dns -join ",") }   else { "" }
            $serverDomainStr= if ($server_domain){ ($server_domain -join ",") }else { "" }
            $serverNTPStr   = if ($server_ntp)   { ($server_ntp -join ",") }   else { "" }

            # Build object
            $results += [PSCustomObject]@{
                ServerName     = $serverName
                ScopeName      = $scope.Name
                ScopeId        = $scope.ScopeId
                StartRange     = $scope.StartRange
                EndRange       = $scope.EndRange
                SubnetMask     = $scope.SubnetMask
                LeaseDuration  = $scope.LeaseDuration
                Scope_DNS      = $scopeDNSStr
                Scope_Domain   = $scopeDomainStr
                Scope_Router   = $scopeRouterStr
                Scope_NTP      = $scopeNTPStr
                Server_DNS     = $serverDNSStr
                Server_Domain  = $serverDomainStr
                Server_NTP     = $serverNTPStr
            }
        }
    }

    return $results
}

# -------------------------------
# RUN INVENTORY
# -------------------------------
$inventory = Get-DhcpInventory -Servers $DhcpServers

# Export to CSV
$inventory | Export-Csv -NoTypeInformation -Path $OutputFile -Encoding UTF8

Write-Host "DHCP inventory complete. Results exported to $OutputFile" -ForegroundColor Green
