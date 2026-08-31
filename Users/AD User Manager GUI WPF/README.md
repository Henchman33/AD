PowerShell AD User Management GUI

**How to use alternate credentials**

1. Click **Credentials…**
2. Enter the account that has rights in the target domain (e.g. CHILD\Admin or admin@child.contoso.com).
3. Select / type the domain FQDN and click **Load / Connect**.

The tool will use those credentials for every subsequent operation (create, search, bulk actions, etc.).

If you need anything else (light/dark theme toggle, more attributes on create, or async bulk processing), just let me know.

| Request              | Implementation                                                                                                                     |     |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------- | --- |
| **Credential login** | “Credentials…” button → Get-Credential. Status line shows which identity is in use. All AD cmdlets now accept and use -Credential. |     |
| **White background** | Window + content areas are pure white / very light gray (#F7F9FC panels).                                                          |     |
| **Tabs**             | Darker professional blue (#1B4F72) with **white lettering**. Selected / hover state uses a slightly lighter blue (#2874A6).        |     |
| **Overall look**     | Clean light theme, high contrast, professional corporate feel.                                                                     |     |
