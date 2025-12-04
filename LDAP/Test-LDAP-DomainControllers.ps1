# Install ADEssentials module if not already installed
# 1st - Install-Module ADEssentials -AllowClobber -Force
# 2nd - Import-Module ADEssentials -AllowClobber -Force
# Fixed Test-LDAP code on line 12 "had Test-LDAP

# Get a list of Domain Controllers
$DCs = Get-ADDomainController -Filter *

# Test LDAP connection to each DC
foreach ($DC in $DCs) {
    Write-Host "Testing LDAP connection to $($DC.Name)..."
    Test-LDAP -ComputerName $DC.Name | Format-Table
}
