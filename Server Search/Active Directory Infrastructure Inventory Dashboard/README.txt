Windows Server 2025 – Yes, it will be detected.
The script filters computer objects using -Filter { OperatingSystem -like "*Server*" }. Microsoft’s AD population for Windows Server 2025 will almost certainly set the OperatingSystem attribute to something like "Windows Server 2025" or "Windows Server 2025 Datacenter", which contains "Server". So it will be included automatically.

Linux (or other non‑Windows) servers – Only if their AD OperatingSystem attribute contains the substring "Server".
In practice, many Linux distributions that are domain‑joined (via SSSD or similar) set OperatingSystem to "Linux", "Red Hat Enterprise Linux", "Ubuntu", etc. – these will be missed because they don’t include "Server".
How to include all server‑like computers (including Linux)

If you want to capture all computer objects and then categorise them by OS, you can either:

    Remove the OS filter and fetch all computer objects, then filter within your code.

    Use an expanded filter that matches common server OS patterns.

Option 1: Fetch all computers (simplest)

Replace the server query block (around line 63) with:
powershell

$servers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonTimestamp, Description, Enabled, whenCreated, ServicePrincipalNames, DistinguishedName

Then, in the loop, you can still compute $status and optionally add a flag like IsServer = $true if $s.OperatingSystem -match "Server|Linux|CentOS|Red Hat|Ubuntu|Debian" to differentiate.
Option 2: Expand the filter to include Linux patterns
powershell

$servers = Get-ADComputer -Filter { (OperatingSystem -like "*Server*") -or (OperatingSystem -like "*Linux*") -or (OperatingSystem -like "*Ubuntu*") -or (OperatingSystem -like "*CentOS*") -or (OperatingSystem -like "*Red Hat*") } -Properties ...

This will catch most domain‑joined Linux servers that have a recognizable OS name.
What about other OSes (e.g., FreeBSD, macOS Server)?

You can similarly extend the filter with additional -or clauses. The script is flexible – just adjust the -Filter parameter to suit your environment.

If you want to be exhaustive, fetching all computers and then applying a regex in PowerShell is more maintainable and won’t miss any.
Quick modification for your script

Look for this line (around line 63 in the provided code):
powershell

$servers = Get-ADComputer -Filter { OperatingSystem -like "*Server*" } -Properties ...

Replace it with:
powershell

$servers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, LastLogonTimestamp, Description, Enabled, whenCreated, ServicePrincipalNames, DistinguishedName

Then, inside the foreach ($s in $servers) loop, you can optionally add an IsServer property based on a broader match – but the existing $status logic will still work for all computers, marking them as Active/Stale/Disabled regardless of OS.
Summary

    Windows Server 2025: ✅ detected.

    Linux servers: ❌ not detected unless they include “Server” in their OS name – modify the filter as shown above to include them.

Feel free to adjust the filter to match your exact naming conventions. The rest of the dashboard (stats, charts, exports) will work seamlessly with the expanded data.
