# PowerShell Script: Add a user to the local Administrators group

# Replace with the domain and username you want to add
 
$domain = "XYZ.com" # <-- change to your domain name
 
$username = "DaveChappelle"  # Change to the actual username
 
$fullUser = "$domain\$username"

# Get the local Administrators group
 
try {
 
    $adminGroup = [ADSI]"WinNT://./Administrators,group"
 
} catch {
 
    Write-Error "Failed to access the local Administrators group. $_"
 
    exit 1
 
}

# Check if the user is already a member
 
$alreadyMember = $false
 
foreach ($member in @($adminGroup.Members())) {
 
    $memberName = $member.GetType().InvokeMember("Name", 'GetProperty', $null, $member, $null)
 
    if ($memberName -ieq $username) {
 
        $alreadyMember = $true
 
        break
 
    }
 
}

# Add the user if not already a member
 
if (-not $alreadyMember) {
 
    try {
 
        $adminGroup.Add("WinNT://$fullUser")
 
        Write-Host "✅ User '$fullUser' added to the local Administrators group."
 
    } catch {
 
        Write-Error "❌ Failed to add user '$fullUser'. $_"
 
    }
 
} else {
 
    Write-Host "ℹ️ User '$fullUser' is already a member of the Administrators group."
 
}
 
