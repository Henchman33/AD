# Use LDAP filter to find published printer queues - may produce duplicate names
$Searcher = [adsisearcher]'(&(objectCategory=printQueue)(uNCName=*))'
$Searcher.PageSize = 10000

# Extract and isolate unique server names
$PrintServers = $Searcher.FindAll() | ForEach-Object {
    $_.Properties.servername
} | Select-Object -Unique | Sort-Object

# Output the list of servers
$PrintServers
