Get-ADObject -Filter 'objectCategory -eq "printQueue"' -Property Name, Location, ServerName, Description | Select-Object Name, Location, ServerName, Description
