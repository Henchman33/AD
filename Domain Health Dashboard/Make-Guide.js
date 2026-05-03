# Good, now I have everything I need. This is a large, complex report — let me build it iteratively.

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
        VerticalAlign, Header, Footer, LevelFormat, PageBreak } = require('docx');
const fs = require('fs');

const BLUE       = '1B4F9E';
const LIGHT_BLUE = 'EBF3FF';
const ALT_ROW    = 'F5F8FF';
const GREY       = '555555';
const brd = { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD' };
const borders = { top: brd, bottom: brd, left: brd, right: brd };

function h(text, level, color = BLUE) {
  const sizes = { 1: 40, 2: 32, 3: 26 };
  return new Paragraph({
    spacing: { before: level === 1 ? 360 : 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: sizes[level] || 24, color, font: 'Arial' })]
  });
}
function p(text, bold = false, color = '222222', size = 20) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text, bold, color, size, font: 'Arial' })]
  });
}
function note(text, color = '1B4F9E', bg = 'EBF3FF') {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color, space: 8 } },
    children: [new TextRun({ text, size: 18, color, font: 'Arial' })]
  });
}
function code(text) {
  return new Paragraph({
    spacing: { after: 60 },
    indent: { left: 360 },
    children: [new TextRun({ text, size: 18, font: 'Courier New', color: '1A3A5C' })]
  });
}
function bullet(text, indent = 360) {
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: indent + 360, hanging: 360 },
    children: [
      new TextRun({ text: '• ', bold: true, size: 20, color: BLUE, font: 'Arial' }),
      new TextRun({ text, size: 20, color: '222222', font: 'Arial' })
    ]
  });
}
function tbl(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const hdrRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders, width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: 'FFFFFF', size: 18, font: 'Arial' })] })]
    }))
  });
  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, ci) => new TableCell({
        borders, width: { size: colWidths[ci], type: WidthType.DXA },
        shading: { fill: ri % 2 === 0 ? 'FFFFFF' : ALT_ROW, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: String(cell), size: 18, font: 'Arial' })] })]
      }))
    })
  );
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: colWidths, rows: [hdrRow, ...dataRows] });
}
function spacer(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({ children: [new TextRun('')] }));
}

// ── BUILD DOC ──
const doc = new Document({
  numbering: {
    config: [{
      reference: 'steps',
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { bold: true, color: BLUE, font: 'Arial', size: 20 } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 2 } },
          children: [
            new TextRun({ text: 'AD Health Dashboard — PowerShell Script', bold: true, size: 22, color: BLUE, font: 'Arial' }),
            new TextRun({ text: '    Setup & Usage Guide', size: 20, color: GREY, font: 'Arial' }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 2 } },
          children: [new TextRun({ text: 'Stephen McKee — Server Administrator 2    |    AD Health Dashboard v2.0    |    Confidential — Internal Use Only', size: 16, color: '999999', font: 'Arial' })]
        })]
      })
    },
    children: [
      // ─ TITLE BLOCK ─
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: 'AD Health Dashboard', bold: true, size: 56, color: BLUE, font: 'Arial' })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: 'PowerShell Data Collection Script  —  Setup & Usage Guide', size: 26, color: GREY, font: 'Arial' })]
      }),
      tbl(
        ['Attribute', 'Value'],
        [
          ['Script Name',  'Get-ADHealthDashboard.ps1'],
          ['Version',      '2.0'],
          ['Author',       'Stephen McKee — Server Administrator 2'],
          ['PowerShell',   '5.1+ (Windows PowerShell; PS 7 also supported)'],
          ['Audience',     'Domain Administrators / Server Administrators'],
          ['Output',       'Self-contained HTML report with CSV/XLSX/TXT/DOCX export'],
        ],
        [2880, 6480]
      ),
      ...spacer(1),

      // ─ SECTION 1 ─
      h('1. Overview', 1),
      p('The Get-ADHealthDashboard.ps1 script connects to your Active Directory domain, collects health data across all key service areas, and generates a single self-contained HTML report. The report is designed for C-level executive review and can be emailed, shared via SharePoint, or printed directly from the browser.'),
      ...spacer(1),
      p('Sections collected by the script:', true),
      bullet('Domain Controller health — uptime, FSMO roles, GC status, IP'),
      bullet('AD Replication — partner status, consecutive failures, last success time'),
      bullet('User Account Health — stale, locked, password never expires, expiring passwords'),
      bullet('Fine-Grained Password Policies (PSO) — all policies and their applied subjects'),
      bullet('Group Policy — linked/unlinked GPOs, WMI filters, enforced policies'),
      bullet('PKI / Certificate Services — expiry tracking across DC and local machine cert stores'),
      bullet('DNS Zones — AD-integrated zones, scavenging status, record counts'),
      bullet('DHCP Scope Utilization — utilization %, failover mode, high-use warnings'),
      bullet('Privileged Group Membership — Domain Admins, Enterprise Admins, Schema Admins, and more'),
      bullet('Security Alerts — auto-generated from collected data; severity-ranked findings with remediation steps'),
      ...spacer(1),

      // ─ SECTION 2 ─
      h('2. Prerequisites', 1),
      h('2.1  Run Environment', 2),
      p('The script should be run from one of the following:'),
      bullet('A Domain Controller (recommended — all data available locally)'),
      bullet('A member server with RSAT tools installed (see Section 2.2)'),
      bullet('An admin workstation with RSAT and domain connectivity'),
      ...spacer(1),
      note('IMPORTANT: The script must be run as a Domain Administrator, or an account with equivalent read rights across all AD objects, DNS, DHCP, and GPMC. Run PowerShell as Administrator.'),
      ...spacer(1),

      h('2.2  Required & Optional Modules', 2),
      tbl(
        ['Module', 'Required?', 'Install / Source', 'Sections Affected'],
        [
          ['ActiveDirectory',  'REQUIRED',  'Add-WindowsFeature RSAT-AD-PowerShell\nor automatic on DCs',  'All sections'],
          ['GroupPolicy',      'Recommended','Add-WindowsFeature GPMC\nor Install via RSAT in Server Manager', 'GPO Health section'],
          ['DnsServer',        'Recommended','Add-WindowsFeature RSAT-DNS-Server', 'DNS Zone Health section'],
          ['DHCPServer',       'Recommended','Add-WindowsFeature RSAT-DHCP',       'DHCP Utilization section'],
        ],
        [2000, 1400, 3200, 2760]
      ),
      ...spacer(1),
      note('If an optional module is absent, the script will log a warning and that section of the report will display a placeholder message. All other sections will still populate normally.'),
      ...spacer(1),

      h('2.3  Execution Policy', 2),
      p('If your environment restricts script execution, temporarily allow it for the session:'),
      code('Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass'),
      p('Or permanently for the local machine (requires elevation):'),
      code('Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine'),
      ...spacer(1),

      // ─ SECTION 3 ─
      h('3. Installation', 1),
      h('3.1  Save the Script', 2),
      ...['Save Get-ADHealthDashboard.ps1 to a folder on your DC or admin workstation.',
          'Recommended path: C:\\Scripts\\ADHealth\\Get-ADHealthDashboard.ps1',
          'Ensure the output directory exists or let the script create it (default: C:\\Reports\\ADHealth\\).'].map((t, i) =>
        new Paragraph({
          numbering: { reference: 'steps', level: 0 },
          spacing: { after: 80 },
          children: [new TextRun({ text: t, size: 20, font: 'Arial', color: '222222' })]
        })
      ),
      ...spacer(1),

      h('3.2  Install RSAT (If Not on a DC)', 2),
      p('On Windows Server 2019 / 2022:'),
      code('Add-WindowsFeature RSAT-AD-PowerShell, GPMC, RSAT-DNS-Server, RSAT-DHCP'),
      p('On Windows 10 / 11 admin workstation:'),
      code('Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'),
      code('Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0'),
      code('Add-WindowsCapability -Online -Name Rsat.Dns.Tools~~~~0.0.1.0'),
      code('Add-WindowsCapability -Online -Name Rsat.DHCP.Tools~~~~0.0.1.0'),
      ...spacer(1),

      // ─ SECTION 4 ─
      h('4. Running the Script', 1),
      h('4.1  Basic Usage (Current Domain, All Defaults)', 2),
      code('.\\Get-ADHealthDashboard.ps1'),
      p('This collects data from the current machine\'s domain and saves the HTML report to C:\\Reports\\ADHealth\\'),
      ...spacer(1),

      h('4.2  Specify a Target Domain', 2),
      code('.\\Get-ADHealthDashboard.ps1 -DomainFQDN "corp.contoso.com"'),
      ...spacer(1),

      h('4.3  Custom Output Path', 2),
      code('.\\Get-ADHealthDashboard.ps1 -OutputPath "D:\\Reports\\Quarterly"'),
      ...spacer(1),

      h('4.4  Open Report Automatically After Generation', 2),
      code('.\\Get-ADHealthDashboard.ps1 -OpenOnComplete'),
      ...spacer(1),

      h('4.5  Full Example with All Parameters', 2),
      code('.\\Get-ADHealthDashboard.ps1 \\'),
      code('    -DomainFQDN   "corp.contoso.com" \\'),
      code('    -OutputPath   "C:\\Reports\\ADHealth" \\'),
      code('    -Author       "Stephen McKee - Server Administrator 2" \\'),
      code('    -OpenOnComplete'),
      ...spacer(1),

      h('4.6  Run as Scheduled Task (Recommended for Weekly Reports)', 2),
      p('To schedule the report to run every Monday at 06:00 AM:'),
      code('$action  = New-ScheduledTaskAction -Execute "PowerShell.exe" `'),
      code('            -Argument "-NonInteractive -ExecutionPolicy Bypass `'),
      code('             -File C:\\Scripts\\ADHealth\\Get-ADHealthDashboard.ps1 -OpenOnComplete:$false"'),
      code('$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6am'),
      code('$settings= New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable'),
      code('Register-ScheduledTask -TaskName "AD Health Dashboard" `'),
      code('  -Action $action -Trigger $trigger -Settings $settings `'),
      code('  -RunLevel Highest -User "DOMAIN\\svc-adreport" -Password "P@ssword"'),
      note('Use a dedicated service account (gMSA preferred) with read-only AD rights for scheduled execution. Avoid using Domain Admin credentials in scheduled tasks.'),
      ...spacer(1),

      // ─ SECTION 5 ─
      h('5. Script Parameters', 1),
      tbl(
        ['Parameter', 'Type', 'Required', 'Default', 'Description'],
        [
          ['-DomainFQDN',     'String',  'No',  '$env:USERDNSDOMAIN',         'Target domain FQDN. Defaults to the current user\'s domain.'],
          ['-OutputPath',     'String',  'No',  'C:\\Reports\\ADHealth',       'Folder where HTML report and log are saved.'],
          ['-Author',         'String',  'No',  'Stephen McKee - Server Administrator 2', 'Name shown in report header and all exports.'],
          ['-OpenOnComplete', 'Switch',  'No',  'False',                       'If set, opens the HTML report in the default browser on completion.'],
        ],
        [1600, 1000, 1000, 2000, 3760]
      ),
      ...spacer(1),

      // ─ SECTION 6 ─
      h('6. Report Features', 1),
      tbl(
        ['Feature', 'Description'],
        [
          ['Live Domain Editing',  'The domain name in the report header is editable — useful for presenting reports across multiple domains without re-running.'],
          ['Collapsible Sections', 'Each data section can be expanded/collapsed individually. Click the section header to toggle.'],
          ['Search',               'The search bar filters all sections simultaneously. Matching rows are highlighted; non-matching sections are hidden.'],
          ['Per-Section Export',   'Each section has individual export buttons: CSV, XLSX, TXT, and DOCX. Exports use the current table data only.'],
          ['KPI Cards',            'Top-level summary cards provide at-a-glance health indicators with colour-coded severity (green/amber/red).'],
          ['Security Alerts',      'Automatically generated from collected data. Ranked by severity with recommended remediation actions.'],
          ['Print / PDF',          'Use Ctrl+P or browser print. All export buttons and search bars are hidden in print view.'],
        ],
        [2200, 7160]
      ),
      ...spacer(1),

      // ─ SECTION 7 ─
      h('7. Output Files', 1),
      tbl(
        ['File', 'Description'],
        [
          ['ADHealthDashboard_<domain>_<timestamp>.html', 'Self-contained HTML report — all CSS, JS, and data embedded. No external dependencies required to open.'],
          ['ADHealthDashboard_<timestamp>.log',           'Detailed execution log with timestamps. Review if any section returns no data.'],
        ],
        [3600, 5760]
      ),
      ...spacer(1),

      // ─ SECTION 8 ─
      h('8. Troubleshooting', 1),
      tbl(
        ['Symptom', 'Cause', 'Resolution'],
        [
          ['Section shows "Install RSAT to collect data"',     'Optional module not installed',       'Install the relevant RSAT feature (see Section 2.2) and re-run.'],
          ['"Access Denied" in the log',                       'Insufficient AD read permissions',    'Run as Domain Admin or grant read rights to the service account.'],
          ['No replication data collected',                    'Run from a non-DC workstation',       'Run from a Domain Controller for full replication metadata access.'],
          ['Script blocked by Execution Policy',               'Restricted execution policy',         'Run: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass'],
          ['Certificate section shows no data',                'No certs in LocalMachine\\My store',  'Run on a DC with enrolled certificates, or check cert store manually.'],
          ['DHCP shows 0 scopes',                             'DHCP servers not authorized in AD',    'Run Get-DhcpServerInDC to verify authorized DHCP servers are reachable.'],
          ['Report file is very large (>10MB)',               'Large GPO XML reports in environment', 'Reduce GPO count or comment out the GPO section in the script.'],
        ],
        [2400, 2200, 4760]
      ),
      ...spacer(1),

      // ─ SECTION 9 ─
      h('9. Security Considerations', 1),
      bullet('The script is read-only — it makes no changes to Active Directory, DNS, DHCP, or Group Policy.'),
      bullet('The HTML report contains live AD data including account names and group membership. Treat it as a confidential internal document.'),
      bullet('Do not email the report externally without redacting sensitive information.'),
      bullet('Store reports in a secured network location with access restricted to IT management and above.'),
      bullet('For scheduled tasks, use a dedicated gMSA (Group Managed Service Account) with minimum required read rights rather than a Domain Admin account.'),
      bullet('Log files contain no passwords or secrets but do include account names and hostnames — apply the same handling as the report.'),
      ...spacer(1),

      // ─ SECTION 10 ─
      h('10. Version History', 1),
      tbl(
        ['Version', 'Date', 'Author', 'Changes'],
        [
          ['2.0', '2025-05-03', 'Stephen McKee', 'Full rewrite with live data collection, 10 sections, auto security alerts, DOCX/XLSX/CSV/TXT export per section.'],
          ['1.0', '2024-01-01', 'Stephen McKee', 'Initial release — static HTML template only.'],
        ],
        [800, 1200, 2400, 5000]
      ),
      ...spacer(2),

      new Paragraph({
        spacing: { before: 240 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
        children: [new TextRun({ text: 'AD Health Dashboard v2.0  —  Stephen McKee, Server Administrator 2  —  Confidential / Internal Use Only', size: 16, color: '999999', font: 'Arial' })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/AD-HealthDashboard-Setup-Guide.docx', buf);
  console.log('DOCX written OK');
});
