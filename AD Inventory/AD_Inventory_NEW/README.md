Required Modules:

#1. ActiveDirectory (REQUIRED)
Provides all AD cmdlets.
Installed via:
RSAT on Windows 10/11
AD DS tools on Windows Server

Check:

Get-Module ActiveDirectory -ListAvailable

Install RSAT on Windows 11:

Get-WindowsCapability -Name RSAT.ActiveDirectory* -Online | Add-WindowsCapability -Online
Recommended Export / Reporting Modules


#2. 
ImportExcel (HIGHLY RECOMMENDED)
Used for:
Native .xlsx export
Multiple worksheets
Auto-sizing
Tables
Formatting
Pivot tables/charts later if desired

Install:
Install-Module ImportExcel -Scope CurrentUser -Force

Check:
Get-Module ImportExcel -ListAvailable



#3. PSWriteHTML (OPTIONAL BUT EXCELLENT)
Greatly improves:
Interactive HTML reports
Searchable tables
Collapsible sections
Charts/graphs
Professional dashboards

Install:
Install-Module PSWriteHTML -Scope CurrentUser -Force
This is one of the best PowerShell reporting modules available.



#4. PSHTML (OPTIONAL)
Alternative HTML generation framework.
Install:
Install-Module PSHTML -Scope CurrentUser -Force
Mapping / Visio / Diagram Modules



#5. Visio Automation (OPTIONAL)
If Microsoft Visio is installed locally, we can later enhance the script to:
Automatically generate .vsdx
Create OU hierarchy diagrams
DC topology diagrams
Trust relationship maps
Group nesting diagrams

Requirements:
Microsoft Visio installed
COM automation enabled
No PowerShell module required.



#6. Graphviz (HIGHLY RECOMMENDED)
Used for:
Professional PNG/SVG/PDF diagrams
Forest topology maps
OU hierarchy maps
Relationship graphs

Download:
https://graphviz.org/download/
After install:
dot -V
Your script already exports .dot files compatible with Graphviz.

To generate PNG automatically:
dot -Tpng AD_Map.dot -o AD_Map.png
GPO Reporting



#7. GroupPolicy (Recommended)
Needed for:
GPO inventory
GPO settings export
GPO link mapping
Usually installed automatically with RSAT.



Check:
Get-Module GroupPolicy -ListAvailable
Best Enterprise Setup
I recommend this exact stack:
Install-Module ImportExcel -Force
Install-Module PSWriteHTML -Force
Install-Module PSHTML -Force



Plus:
RSAT Active Directory tools
GroupPolicy module
Graphviz installed locally
Recommended Future Enhancements



*** Your script can later be upgraded to include ***
Security Analysis
Stale accounts
Privileged account audit
Kerberoast exposure
AS-REP roastable users
Unconstrained delegation
SIDHistory
AdminCount
Nested privileged groups
AD Health
Replication health
DFSR status
FSMO health
SYSVOL consistency
Lingering object detection
Tiering Analysis
Tier 0 exposure
Admin workstation mapping
Privileged path analysis
BloodHound Export

Can export directly into:
BloodHound
Neo4j
Purple Knight style reports
Azure / Hybrid
Azure AD Connect
Entra ID sync analysis
Hybrid identity mapping
Automatic Visio Generation



Possible using:
Visio COM
Draw.io export
Mermaid.js
SVG topology rendering
Best Practice Run Method



Run PowerShell ISE or PowerShell as:
Run as Administrator
Then:
Set-ExecutionPolicy RemoteSigned
And launch:
.\Improved_AD_Inventory.ps1
The script exports automatically to:
Desktop\AD_Inventory_YYYYMMDD_HHMMSS
