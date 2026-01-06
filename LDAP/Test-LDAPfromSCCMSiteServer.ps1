# Define variables
# Some environments may need the FQDN prefix before the service account name "blah.ad.blah.com\serviceaccountname" 
$Username = " Service Account Used for System Discovery" # Use a user account with read rights in the subdomain
$Password = "Password for above Service Account"
$SubdomainDC = "Blah.Blah.com" # Hostname or IP from Step 1
$LdapPort = 389 # or 636

# Create a secure password object
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)

# The LDAP path should point to an existing object in the target subdomain, e.g., the Domain DN
# Example: "LDAP://dc01.sub.domain.com"
$LDAPPath = "LDAP://blah.blah.com/DC=blah,DC=blah,DC=blah,DC=com"

try {
    # Attempt to create a DirectoryEntry object, which performs the bind operation
    $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry($LDAPPath, $Credential.UserName, $Credential.GetNetworkCredential().Password)
    
    # Attempt to refresh the cache to force the connection and authentication
    $directoryEntry.RefreshCache()
    Write-Host "Success: LDAP connection and bind to $SubdomainDC successful." -ForegroundColor Green
}
catch {
    Write-Host "Error: Failed to make LDAP connection or bind." -ForegroundColor Red
    Write-Error $_.Exception.Message
}
