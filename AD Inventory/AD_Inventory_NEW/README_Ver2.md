# What's new vs your original v2.0:

## OU/Container Inventory (the core new feature)

### All OUs and built-in CN containers (CN=Users, CN=Computers, CN=Builtin, etc.) are enumerated. 

### Each is automatically classified by purpose — Tier 0/1/2, Domain Controllers, Servers, Workstations, Service Accounts, PAW, gMSA, Staging, Test/Lab, Disabled/Archived, and more — based on name pattern matching. 

### Each card shows exact counts of Users, Servers, Workstations, DCs, Groups, and GPO links directly inside that container.

### Desktop output — exports go to %USERPROFILE%\Desktop\ADInventory_DOMAIN_timestamp\ and Explorer opens automatically when done.

### Actual Excel workbook via Excel.Application COM — 14 worksheets with blue headers, alternating rows, auto-fit columns, frozen header rows, and red/amber severity highlighting. 

### Gracefully skipped if Excel isn't installed.

### 15 CSV files — one per data category, named and numbered for easy sorting.

### SVG Forest Map — a dark-themed hierarchical tree diagram showing Forest → Domains → OUs → Sub-OUs with color-coded nodes per type, stats in each node, and a full legend. 

### Always generated, no dependencies.

### Visio diagram via Visio.Application COM — full organizational chart with colored shapes and connectors. 

### Gracefully skipped with a log message if Visio isn't installed.

### Bug fixes from v2.0 — the broken ?.Format(0) PKI expression, the mismatched-quote security alert string, and the GPO status inline-if side-effect are all corrected.
