AD Search Tool is a powerful feature that allows you to search Active Directory using different types of filters. Here's a comprehensive guide on how to use it:
Basic Filter Syntax
1. Simple Wildcard Searches

    Use * as a wildcard to match any characters

    Use ? to match a single character

Examples:

    * - Find everything in the selected category

    john* - Find items starting with "john" (e.g., john.smith, john.doe)

    *admin* - Find items containing "admin" anywhere

    *server* - Find items containing "server"

    SRV?? - Find items starting with "SRV" followed by exactly 2 characters

2. Category-Specific Searching

The filter behaves differently depending on which category you select:

For Users:

    Searches: Name, sAMAccountName, Email, UserPrincipalName

    Example: jsmith finds users with name or username containing "jsmith"

For Computers:

    Searches: Computer Name, Operating System

    Example: *server* finds computers with "server" in name or OS

For Groups:

    Searches: Group Name only

    Example: *admin* finds groups with "admin" in the name

For OUs:

    Searches: Organizational Unit Name

    Example: *sales* finds OUs with "sales" in the name

Advanced LDAP Filter Syntax

When you start your filter with certain special characters, the tool switches to LDAP filter mode, which is much more powerful.
LDAP Filter Indicators

Start your filter with any of these to use LDAP syntax:

    ( - Opening parenthesis

    = - Equals sign

    & - AND operator

    | - OR operator

    ! - NOT operator

Common LDAP Filters

1. Search by specific attribute:
text

(sAMAccountName=jsmith)

Finds user with exact username "jsmith"

2. Combine conditions with AND (&):
text

(&(objectCategory=person)(objectClass=user)(department=IT))

Finds users in the IT department

3. Combine conditions with OR (|):
text

(|(department=Sales)(department=Marketing))

Finds users in Sales OR Marketing departments

4. Complex combinations:
text

(&(objectCategory=person)(objectClass=user)(|(department=IT)(title=*Manager*)))

Finds users who are in IT department OR have "Manager" in their title

5. Search for enabled/disabled accounts:
text

(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))

Finds disabled accounts
text

(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))

Finds enabled accounts

6. Search by date:
text

(whenCreated>=20240101000000.0Z)

Finds accounts created on or after Jan 1, 2024
Practical Examples
User Search Examples:

    john* - Simple: Users with "john" in name/username

    (sn=Smith) - LDAP: Users with last name "Smith"

    (&(objectClass=user)(mail=*@company.com)) - LDAP: Users with company email

    (&(objectClass=user)(title=*Manager*)) - LDAP: Users with "Manager" in title

    (&(objectClass=user)(memberOf=CN=Domain Admins,CN=Users,DC=domain,DC=com)) - LDAP: Domain Admins members

Computer Search Examples:

    *ws* - Simple: Computers with "ws" in name

    (operatingSystem=*Server*) - LDAP: Servers

    (&(objectCategory=computer)(operatingSystem=*10*)) - LDAP: Windows 10 computers

    (&(objectCategory=computer)(lastLogonTimestamp<=131928768000000000)) - LDAP: Inactive computers (timestamp format)

Group Search Examples:

    *security* - Simple: Groups with "security" in name

    (groupType:1.2.840.113556.1.4.803:=2147483648) - LDAP: Security groups

    (groupType:1.2.840.113556.1.4.803:=2) - LDAP: Distribution groups

Special Features
Presets

The tool comes with built-in presets for common searches:

    "Disabled Accounts" - Finds all disabled user accounts

    "Locked-Out Users" - Finds currently locked out users

    "Service Accounts (SPN)" - Finds accounts with Service Principal Names

    "Password Never Expires" - Finds accounts with non-expiring passwords

    "Domain Admins Members" - Lists all Domain Admins members

    "Recently Created Accounts (30d)" - Accounts created in last 30 days

    "Inactive Computers (90d)" - Computers not logged in for 90+ days

Lockout Origin Lookup

When searching for locked-out users, check "Lockout origin lookup" to find which computer caused the lockout (requires admin rights to read event logs).
Export Options

After getting results, use the Export button to save in multiple formats (CSV, HTML, Excel, etc.)
Tips & Best Practices

    Start Simple: Use wildcards first (*admin*) before complex LDAP filters

    Be Specific: The more specific your filter, the faster the search

    Test in Production: Test new filters on small subsets first

    Use Presets: Check the built-in presets before creating custom filters

    Save Settings: Use the Settings tab to save frequently used filters

    Monitor Performance: Complex LDAP filters on large domains can be slow

Troubleshooting

If searches return no results:

    Check your spelling

    Try a broader filter (* first)

    Ensure you have the right category selected

    Verify you have proper AD permissions

If LDAP filters fail:

    Check parentheses - they must be balanced

    Ensure attribute names are correct (case-sensitive in some cases)

    Try the filter in a simpler form first

For domain-specific attributes:
You may need to use your domain's specific attribute names. Use Get-ADUser -Filter * -Properties * | Get-Member to see all available attributes.
