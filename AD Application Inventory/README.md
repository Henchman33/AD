What This Script Provides

    Comprehensive Discovery – Finds applications through SCPs, explicit application objects, servers with SPNs, dedicated service accounts, and application‑specific OUs.

    Stale Detection – Flags stale application servers and service accounts based on LastLogonTimestamp.

    Rich Dashboard – Interactive HTML with stats cards, SPN service‑type chart, category pie chart, and a filterable table.

    Export Options – One‑click exports to Excel, CSV (from dashboard), and Word; also saves master CSV and per‑category CSVs to the output folder.

    Offline Testing – Uses sample data if AD is unavailable.

    Organised Output – All files are stored in a timestamped folder on your desktop.

How to Use

    Save the script as AD_Application_Inventory.ps1.

    Run it in PowerShell ISE (as Administrator) on a domain‑joined machine.

    Check your desktop for the Application_Search_<timestamp> folder.

    Open AD_App_Inventory.html in any browser to explore the dashboard.

    Use the Excel, CSV, and Word buttons inside the dashboard for dynamic exports.

This script gives you a complete picture of applications registered in Active Directory – essential for any AD assessment.
