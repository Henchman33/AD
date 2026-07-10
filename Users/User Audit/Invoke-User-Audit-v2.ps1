#Requires -Modules ActiveDirectory - SENS outputGPT

Import-Module ActiveDirectory

$Domains = @(
    "MYIGT.COM",
    "AD.IGT.COM",
    "IGTSAP.AD.IGT.COM",
    "IS.AD.IGT.COM",
    "EC.AD.IGT.COM"
)

$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$RootFolder = Join-Path `
    ([Environment]::GetFolderPath("Desktop")) `
    "DUMPSEC\$TimeStamp"

New-Item -Path $RootFolder -ItemType Directory -Force | Out-Null

$LogFile = Join-Path $RootFolder "Execution.log"

Start-Transcript -Path $LogFile -Force

Write-Host "Starting Active Directory Security Collection..." -ForegroundColor Cyan

$AllUsers = @()
$Tier0Accounts = @()
$Tier1Accounts = @()
$ServiceAccounts = @()
$LockedAccounts = @()
$DelegationAccounts = @()
$StaleAccounts = @()

$Tier0Groups = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators",
    "Account Operators",
    "Backup Operators",
    "Server Operators",
    "Print Operators",
    "Protected Users",
    "Group Policy Creator Owners"
)

foreach ($Domain in $Domains)
{
    try
    {
        Write-Host "Processing $Domain"

        $Users = Get-ADUser `
            -Server $Domain `
            -Filter * `
            -Properties * `
            -ResultPageSize 2000 `
            -ResultSetSize $null

        foreach ($User in $Users)
        {
            $LastLogon = $null

            if ($User.LastLogonDate)
            {
                $LastLogon = $User.LastLogonDate
            }

            $PasswordExpiryDate = $null

            try
            {
                $PasswordExpiryDate =
                    [datetime]::FromFileTime(
                        $User.'msDS-UserPasswordExpiryTimeComputed'
                    )
            }
            catch
            {
            }

            $DaysSinceLogon = $null

            if ($LastLogon)
            {
                $DaysSinceLogon =
                    (New-TimeSpan -Start $LastLogon -End (Get-Date)).Days
            }

            $DaysSincePasswordSet = $null

            if ($User.PasswordLastSet)
            {
                $DaysSincePasswordSet =
                    (New-TimeSpan `
                        -Start $User.PasswordLastSet `
                        -End (Get-Date)).Days
            }

            $AccountStatus =
                if ($User.Enabled)
                {
                    "Enabled"
                }
                else
                {
                    "Disabled"
                }

            $Record = [PSCustomObject]@{

                DomainName                = $Domain
                DisplayName               = $User.DisplayName
                CommonName                = $User.CN
                SamAccountName            = $User.SamAccountName
                FirstName                 = $User.GivenName
                LastName                  = $User.Surname
                FullName                  = $User.Name
                Email                     = $User.Mail
                Alias                     = $User.MailNickname

                Ext8                      = $User.extensionAttribute8
                Ext9                      = $User.extensionAttribute9
                Ext10                     = $User.extensionAttribute10
                Ext13                     = $User.extensionAttribute13
                Ext14                     = $User.extensionAttribute14
                Ext15                     = $User.extensionAttribute15

                Description               = $User.Description
                DistinguishedName         = $User.DistinguishedName
                OU                        = ($User.DistinguishedName -replace '^CN=.*?,','')

                Enabled                   = $User.Enabled
                AccountStatus             = $AccountStatus

                AccountLocked             = $User.LockedOut

                AccountExpirationDate     = $User.AccountExpirationDate

                PasswordLastSet           = $User.PasswordLastSet
                DaysSincePasswordSet      = $DaysSincePasswordSet

                PasswordNeverExpires      = $User.PasswordNeverExpires
                CannotChangePassword      = $User.CannotChangePassword

                PasswordExpiryDate        = $PasswordExpiryDate

                LastLogonDate             = $User.LastLogonDate
                LastLogonTimestamp        = $User.LastLogonTimeStamp

                DaysSinceLastLogon        = $DaysSinceLogon

                SmartCardRequired         = $User.SmartcardLogonRequired

                TrustedForDelegation      = $User.TrustedForDelegation

                SIDHistory                = ($User.SIDHistory -join ";")

                ServicePrincipalNames     = ($User.ServicePrincipalName -join ";")

                AdminCount                = $User.AdminCount

                WhenCreated               = $User.WhenCreated
                WhenChanged               = $User.WhenChanged
            }

            $AllUsers += $Record

            if ($User.LockedOut)
            {
                $LockedAccounts += $Record
            }

            if ($User.TrustedForDelegation)
            {
                $DelegationAccounts += $Record
            }

            if ($User.ServicePrincipalName)
            {
                $ServiceAccounts += $Record
            }

            if ($DaysSinceLogon -gt 90)
            {
                $StaleAccounts += $Record
            }
        }

        foreach ($Group in $Tier0Groups)
        {
            try
            {
                $Members =
                    Get-ADGroupMember `
                        -Server $Domain `
                        -Identity $Group `
                        -Recursive `
                        -ErrorAction Stop

                foreach ($Member in $Members)
                {
                    $Tier0Accounts += [PSCustomObject]@{
                        Domain = $Domain
                        Group  = $Group
                        Name   = $Member.Name
                        Type   = $Member.ObjectClass
                    }
                }
            }
            catch
            {
            }
        }

        try
        {
            $Tier1OU =
                Get-ADUser `
                    -Server $Domain `
                    -LDAPFilter "(adminCount=1)" `
                    -Properties *

            foreach ($Admin in $Tier1OU)
            {
                $Tier1Accounts += [PSCustomObject]@{
                    Domain            = $Domain
                    Name              = $Admin.Name
                    SamAccountName    = $Admin.SamAccountName
                    DistinguishedName = $Admin.DistinguishedName
                }
            }
        }
        catch
        {
        }
    }
    catch
    {
        Write-Warning "$Domain failed: $_"
    }
}

$CSVUsers = Join-Path $RootFolder "AD_Users.csv"
$CSVTier0 = Join-Path $RootFolder "Tier0_Accounts.csv"
$CSVTier1 = Join-Path $RootFolder "Tier1_Accounts.csv"
$CSVService = Join-Path $RootFolder "ServiceAccounts.csv"
$CSVLocked = Join-Path $RootFolder "LockedAccounts.csv"
$CSVDelegation = Join-Path $RootFolder "DelegationAccounts.csv"
$CSVStale = Join-Path $RootFolder "StaleAccounts.csv"

$AllUsers | Export-Csv $CSVUsers -NoTypeInformation -Encoding UTF8
$Tier0Accounts | Export-Csv $CSVTier0 -NoTypeInformation -Encoding UTF8
$Tier1Accounts | Export-Csv $CSVTier1 -NoTypeInformation -Encoding UTF8
$ServiceAccounts | Export-Csv $CSVService -NoTypeInformation -Encoding UTF8
$LockedAccounts | Export-Csv $CSVLocked -NoTypeInformation -Encoding UTF8
$DelegationAccounts | Export-Csv $CSVDelegation -NoTypeInformation -Encoding UTF8
$StaleAccounts | Export-Csv $CSVStale -NoTypeInformation -Encoding UTF8

if (Get-Module -ListAvailable ImportExcel)
{
    $XLSX = Join-Path $RootFolder "AD_Security_Report.xlsx"

    $AllUsers | Export-Excel $XLSX -WorksheetName Users -AutoSize
    $Tier0Accounts | Export-Excel $XLSX -WorksheetName Tier0 -AutoSize
    $Tier1Accounts | Export-Excel $XLSX -WorksheetName Tier1 -AutoSize
    $ServiceAccounts | Export-Excel $XLSX -WorksheetName ServiceAccounts -AutoSize
    $LockedAccounts | Export-Excel $XLSX -WorksheetName LockedAccounts -AutoSize
    $DelegationAccounts | Export-Excel $XLSX -WorksheetName Delegation -AutoSize
    $StaleAccounts | Export-Excel $XLSX -WorksheetName StaleAccounts -AutoSize
}

$HTML = Join-Path $RootFolder "AD_Security_Report.html"

$Summary = @"
<h1>Active Directory Security Assessment</h1>

<p>Total Users: $($AllUsers.Count)</p>
<p>Tier0 Accounts: $($Tier0Accounts.Count)</p>
<p>Tier1 Accounts: $($Tier1Accounts.Count)</p>
<p>Service Accounts: $($ServiceAccounts.Count)</p>
<p>Locked Accounts: $($LockedAccounts.Count)</p>
<p>Delegation Accounts: $($DelegationAccounts.Count)</p>
<p>Stale Accounts: $($StaleAccounts.Count)</p>
"@

$AllUsers |
    ConvertTo-Html `
        -Title "AD Security Report" `
        -PreContent $Summary |
    Out-File $HTML

Stop-Transcript

Write-Host ""
Write-Host "Reports written to:" -ForegroundColor Green
Write-Host $RootFolder -ForegroundColor Yellow
