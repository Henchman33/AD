#requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Finds AD users in a target OU where extensionAttribute8 is NOT set and exports them to a CSV.

.PARAMETER SearchBase
    Distinguished name of the OU to search. Subtree scope (the OU and everything under it).

.PARAMETER Server
    Optional. Target DC or domain (e.g. myigt.com). Defaults to PDC emulator discovery.

.PARAMETER OutputPath
    Optional. CSV output path. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\Get-MissingExtAttr8.ps1 -SearchBase "OU=Users,OU=Corp,DC=myigt,DC=com" -Server myigt.com
#>

[CmdletBinding()]
param(
    [string]$SearchBase = "OU=Users,OU=SYNCTEST,DC=XX,DC=XXX,DC=COM",

    [string]$Server = (Get-ADDomainController -Discover -Service PrimaryDC).HostName[0],

    [string]$OutputPath = ".\ExtAttr8-Missing_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

Import-Module ActiveDirectory -ErrorAction Stop

# LDAP filter pulls only enabled users missing extensionAttribute8, all server-side.
# (!(attr=*)) is the "not set" case; in AD an attribute is either present with a value or absent.
# (!(userAccountControl:1.2.840.113556.1.4.803:=2)) excludes disabled accounts (ACCOUNTDISABLE bit).
$ldapFilter = '(&(objectCategory=person)(objectClass=user)(!(extensionAttribute8=*))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'

$results = Get-ADUser -LDAPFilter $ldapFilter -SearchBase $SearchBase -Server $Server `
    -Properties extensionAttribute8, DisplayName, Enabled, Created |
    Select-Object SamAccountName, DisplayName, UserPrincipalName, Enabled, extensionAttribute8, Created, DistinguishedName

if (-not $results) {
    Write-Host "No users missing extensionAttribute8 under $SearchBase." -ForegroundColor Green
    return
}

# Classify each user into one bucket.
# Precedence: a Consultants OU in the DN (structural) wins over the (CONTRACTOR) display-name tag.
$consultants = @($results | Where-Object { $_.DistinguishedName -match 'OU=Consultants' })
$contractors = @($results | Where-Object { $_.DistinguishedName -notmatch 'OU=Consultants' -and $_.DisplayName -like '*(CONTRACTOR)*' })
$standard    = @($results | Where-Object { $_.DistinguishedName -notmatch 'OU=Consultants' -and $_.DisplayName -notlike '*(CONTRACTOR)*' })

# Derive paired output paths from $OutputPath so the timestamps stay in sync.
$dir  = [IO.Path]::GetDirectoryName($OutputPath); if ([string]::IsNullOrEmpty($dir)) { $dir = '.' }
$base = [IO.Path]::GetFileNameWithoutExtension($OutputPath)
$ext  = [IO.Path]::GetExtension($OutputPath)
$ContractorPath = Join-Path $dir ("{0}-Contractors{1}" -f $base, $ext)
$ConsultantPath = Join-Path $dir ("{0}-Consultants{1}" -f $base, $ext)

if ($standard.Count) {
    $standard | Sort-Object SamAccountName | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("{0} standard user(s) exported to {1}" -f $standard.Count, $OutputPath) -ForegroundColor Yellow
}
else {
    Write-Host "No standard users missing extensionAttribute8." -ForegroundColor Green
}

if ($contractors.Count) {
    $contractors | Sort-Object SamAccountName | Export-Csv -Path $ContractorPath -NoTypeInformation -Encoding UTF8
    Write-Host ("{0} contractor(s) exported to {1}" -f $contractors.Count, $ContractorPath) -ForegroundColor Yellow
}
else {
    Write-Host "No contractors found in the result set." -ForegroundColor Green
}

if ($consultants.Count) {
    $consultants | Sort-Object SamAccountName | Export-Csv -Path $ConsultantPath -NoTypeInformation -Encoding UTF8
    Write-Host ("{0} consultant(s) exported to {1}" -f $consultants.Count, $ConsultantPath) -ForegroundColor Yellow
}
else {
    Write-Host "No consultants found in the result set." -ForegroundColor Green
}
