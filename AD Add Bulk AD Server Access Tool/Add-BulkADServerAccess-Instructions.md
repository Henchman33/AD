# Add Bulk AD Server Access Tool — Instructions
## Version 2.0 | PowerShell ISE / Console GUI

---

## OVERVIEW

A GUI-based PowerShell tool for adding Active Directory users and Managed Service Accounts
(MSAs/gMSAs) to local groups on multiple servers simultaneously. Designed for System
Administrators managing multi-domain environments.

**Key Features**
- Add multiple AD users to multiple servers in one operation
- Add Managed Service Accounts (MSA / gMSA) to servers
- Credential profiles for multiple domains (held in memory, not written to disk)
- Server and user list paste-in or CSV import/export
- Dropdown group selection (Administrators, RDP Users, Remote Management Users, and more)
- Real-time progress bar and live log viewer
- Error Grid with per-operation status (SUCCESS / SKIPPED / WARNING / ERROR)
- Timestamped log files saved automatically to your Desktop

---

## REQUIREMENTS

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or higher |
| Execution Policy | Must allow script execution (see Step 1 below) |
| RSAT / AD Module | Required if you extend the script with AD queries |
| WinRM | Must be enabled and configured on target servers |
| Network | Must have WinRM access (TCP 5985/5986) to target servers |
| Permissions | Your account (or the credential profile used) must have local admin rights on target servers |

---

## STEP 1 — SET EXECUTION POLICY (One-Time)

Open PowerShell **as Administrator** and run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Or for ISE specifically, you can bypass for a single session:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

---

## STEP 2 — ENABLE WINRM ON TARGET SERVERS (If Not Already Done)

On each target server, run **as Administrator**:

```powershell
Enable-PSRemoting -Force
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
```

> For domain-joined servers in the same domain, WinRM is typically already enabled via GPO.
> For workgroup or cross-domain servers, you may need to add specific IPs to TrustedHosts.

---

## STEP 3 — RUN THE TOOL

**In PowerShell ISE:**
1. Open `Add-BulkADServerAccess.ps1` in ISE
2. Press **F5** to run the script
3. The GUI window will launch

**In standard PowerShell console:**
```powershell
.\Add-BulkADServerAccess.ps1
```

The tool automatically re-launches itself in **STA mode** if needed (required for WinForms).

---

## STEP 4 — LOG FILE LOCATION

Logs are automatically written to:

```
%USERPROFILE%\Desktop\Add Bulk AD Server Access Tool\Logs\log_YYYYMMDD_HHmmss.log
```

Example:
```
C:\Users\Steve\Desktop\Add Bulk AD Server Access Tool\Logs\log_20250615_091532.log
```

The log path is shown at the top of the tool window and in the Live Log tab.

---

## TAB-BY-TAB GUIDE

---

### TAB: CREDENTIAL PROFILES

**Use this tab FIRST if your target servers are in a different domain than your logon.**

**To create a credential profile:**
1. Enter a short **Profile Name** — e.g. `CORP`, `DMZ`, `LAB`, `PRODADMIN`
2. Enter the **Domain\Username** — e.g. `CORP\svc-admin` or `DMZ\Administrator`
3. Enter the **Password**
4. Optionally check **Show password** to verify
5. Click **💾 Save Profile**

The profile name now appears in the credential dropdown on the Main Operation and MSA tabs.

**Notes:**
- Credentials are held in memory only — NOT written to any file
- You can save multiple profiles for different domains
- Use "Current Windows Session" if your logged-on account already has rights
- To remove a profile: select it in the list and click **🗑 Remove Selected**

---

### TAB: MAIN OPERATION

This is where you run bulk user-to-server additions.

**Target Servers (left box)**
- Paste one server name per line: `SERVER01`, `SERVER02`, `DC01.corp.local`
- Or click **📂 Import CSV** — the CSV should have server names in the first column
- Click **💾 Export CSV** to save the current list for reuse

**Users to Add (center box)**
- Paste one `DOMAIN\username` per line
- Example: `CORP\jsmith`, `LAB\a.jones`, `DMZ\svc-account`
- Or click **📂 Import CSV** — accounts should be in the first column

**Options (right panel)**

| Option | Description |
|---|---|
| Target Local Group | Dropdown to select: Administrators, Remote Desktop Users, Remote Management Users, Backup Operators, etc. You can also type a custom group name directly. |
| Credential Profile | Select which saved credential to use when connecting to the servers. |
| Test connectivity (ping) | Recommended ON. Skips unreachable servers instead of timing out. |
| Verify membership after add | Confirms the account is in the group after adding. |
| Skip if already a member | Marks already-present accounts as SKIPPED instead of ERROR. |

**The counter** at the bottom-right shows: `Servers: X | Users: Y | Operations: X×Y`

**To run:**
1. Fill in servers and users
2. Select group and credential profile
3. Click **▶ RUN**
4. The tool switches to the Live Log tab — watch progress in real time
5. The progress bar and status label track each operation
6. When complete, a summary dialog shows success/skip/error counts

---

### TAB: MANAGED SERVICE ACCOUNTS

Use this tab specifically for adding **MSA** or **gMSA** accounts to local groups on servers.

**Important prerequisites:**
- The gMSA must already exist in AD (`New-ADServiceAccount`)
- The gMSA must already be installed on the target server (`Install-ADServiceAccount`)
- gMSA names typically end with `$` — e.g. `CORP\svc-sqlengine$`

**Steps:**
1. Paste target servers in the left box (or Import CSV)
2. Paste MSA/gMSA account names in the center box — format: `DOMAIN\account$`
3. Select the **Target Local Group** (usually `Administrators` or `Remote Desktop Users`)
4. Select a **Credential Profile** if needed
5. Click **▶ ADD MSAs**

Results appear in the Live Log and Error Grid tabs.

---

### TAB: LIVE LOG

Real-time color-coded log of every operation:

| Color | Meaning |
|---|---|
| 🟢 Green | SUCCESS — account added successfully |
| 🟡 Yellow | WARNING or SKIPPED — already a member, or verification issue |
| 🔴 Red | ERROR — failed to add, unreachable, permission denied |
| Grey | INFO — general status messages |

**Buttons:**
- **🗑 Clear View** — clears the on-screen log (does NOT delete the log file)
- **💾 Save Log** — saves the current view to a separate file
- **📂 Open Log Folder** — opens the Desktop log folder in Explorer

---

### TAB: ERROR GRID

A table showing every operation result with columns:

| Column | Description |
|---|---|
| Timestamp | When the operation ran |
| Server | Target server name |
| Account | DOMAIN\username being added |
| Group | Local group targeted |
| Status | SUCCESS / SKIPPED / WARNING / ERROR |
| Message | Detailed result or error message |

**Buttons:**
- **💾 Export Errors CSV** — exports the full grid to a CSV file
- **🗑 Clear Grid** — clears the grid (does not affect the log file)

You can also reach this tab by clicking **📋 Error Grid** on the Main Operation tab.

---

## CSV FORMAT

**Servers CSV** — one column, header optional:
```
Server
SERVER01
SERVER02
DC01.corp.local
```

**Users CSV** — one column, header optional:
```
Account
CORP\jsmith
CORP\a.jones
DMZ\svc-deploy
```

**MSA CSV** — same format as users, gMSA names with `$`:
```
Account
CORP\svc-sqlengine$
CORP\svc-iisapppool$
```

---

## COMMON SCENARIOS

**Scenario 1: Add helpdesk users to Remote Desktop Users on lab servers**
1. Credential Profile tab → save `LAB\labadmin` as profile "LAB"
2. Main Operation → paste servers, paste users (LAB\user1, LAB\user2)
3. Group = `Remote Desktop Users`
4. Credential Profile = `LAB`
5. ▶ RUN

**Scenario 2: Add a service account to Administrators on production servers**
1. Credential Profile tab → save `PROD\domain-admin` as "PROD"
2. Main Operation → paste prod servers, paste `PROD\svc-monitoring`
3. Group = `Administrators`
4. Enable "Verify membership after add"
5. ▶ RUN

**Scenario 3: Add a gMSA to multiple servers**
1. MSA tab → paste servers
2. Paste `CORP\svc-sqlengine$` in the accounts box
3. Group = `Administrators`
4. ▶ ADD MSAs
5. Check Error Grid for results

**Scenario 4: Multi-domain environment (CORP + DMZ servers in same run)**
- Option A: Run twice — once with CORP credential profile for CORP servers, once with DMZ profile for DMZ servers
- Option B: Split the lists and run separate sessions

---

## TROUBLESHOOTING

| Problem | Fix |
|---|---|
| "Access Denied" errors | Verify your credential profile has local admin rights on the target server |
| "Server unreachable" errors | Check WinRM is enabled; check firewall TCP 5985; verify hostname resolves |
| "Group not found" | The group name must match exactly — check spelling, capitalization |
| Tool opens then closes | Script may be running in MTA mode — the STA auto-relaunch handles this, but check PS version |
| Already a member = ERROR | Enable "Skip if already a member" checkbox |
| Verify fails after add | Network latency — wait a moment and recheck; or disable verify for large batches |
| gMSA add fails | Ensure `Install-ADServiceAccount` has been run on the target server first |

---

## SECURITY NOTES

- Credentials entered in the tool are stored **in memory only** for the session duration
- No passwords are written to the log file
- The log file records operations and outcomes only (server, account, group, status, message)
- Run the tool from a **PAW (Privileged Access Workstation)** in production environments
- Close the tool when finished to clear credential profiles from memory

---

## EXTENDING THE TOOL

The script is designed to be readable and extensible. Some ideas:

- **Add more groups to the dropdown:** Find the `@('Administrators','Remote Desktop Users'...)` array and add entries
- **Pre-populate servers from AD:** Use `Get-ADComputer` and pipe to the textbox
- **Email notification on completion:** Add `Send-MailMessage` after the summary block
- **Scheduled/unattended mode:** The core `Add-AccountToLocalGroup` function can be called directly from a script without the GUI
