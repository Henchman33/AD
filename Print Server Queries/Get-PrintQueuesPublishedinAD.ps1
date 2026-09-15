#import-module ActiveDirectory # From Remote Server Admin Tools (RSAT) Windows Desktop/server OS feature

# get print queues published in Active Directory
$printqueues=Get-AdObject -filter "objectCategory -eq 'printqueue'" -Properties *

# list the print queue server names and filter out duplicates where one server has multiple print queues
$PrintServers=$printqueues | select servername | Sort-Object -property servername -Unique

# Export to CSV file reports (calculated property printshares as semicolon delimited list) https://ss64.com/ps/select-object.html 
$PrintServers | Export-Csv -NoTypeInformation -Encoding UTF8 -path $env:temp\printservers.csv
$resultTable=$printqueues | Select-object serverName,Name, @{Name='ShareNames';Expression={$_.printShareName -join ';'}} 
$resultTable | Export-Csv -NoTypeInformation -Encoding UTF8 -path $env:temp\printqueues.csv

# Open reports
invoke-item $env:temp\printservers.csv
invoke-item $env:temp\printqueues.csv
