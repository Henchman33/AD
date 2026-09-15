# Ensure the Active Directory module is loaded
Import-Module ActiveDirectory

# Query AD for unique print server names
$PrintServers = Get-ADObject -Filter 'objectCategory -eq "printQueue"' -Properties ServerName |
                Select-Object -ExpandProperty ServerName -Unique | 
                Sort-Object

# Output the list of servers
$PrintServers
