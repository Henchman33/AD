# Remove-DisabledUserProfiles.ps1
# Article: https://sccmnotes.wordpress.com/2024/06/
param([Parameter(Mandatory=$true)][string] $ComputerName)

Write-Host "`nTesting network connection to $ComputerName ..." -NoNewline
$PingTest = Test-NetConnection -ComputerName $ComputerName -WarningAction SilentlyContinue -InformationLevel Quiet

if ($PingTest)
{
    Write-Host " test successful."
    Write-Host "Creating remote PowerShell session to $ComputerName ..."
    Try
    {
        $RemotePSSession = New-PSSession $ComputerName -ErrorAction Stop
    }
    Catch
    {
        Write-Host "`nError establishing remote PowerShell session."
        Write-Host $_
        exit
    }

    # Take note of free disk space
    $InitialDisk = Invoke-Command -Session $RemotePSSession {Get-PSDrive C} | Select-Object Free

    Write-Host "Getting user profiles on $ComputerName ..." -NoNewline
    $Profiles = Get-WmiObject -Computer $ComputerName -Class Win32_UserProfile
    $ProfilesDeleted = 0
    Write-Host " done."

    foreach ($Profile in $Profiles)
    {
        $SID = ""
        $FullUserName = ""
        $NetBIOSDomainName = ""

        $objSID = New-Object System.Security.Principal.SecurityIdentifier($Profile.SID)
        $SID = $objSID.Value

        # Check for a NULL local path
        if ($Profile.LocalPath -eq $NULL)
        {
            Write-Host "LocalPath of SID $SID is null" -ForegroundColor "Yellow"
        }
        # Check to see if the specified local path exists
        elseif (!(Invoke-Command -Session $RemotePSSession {param($Profile) Test-Path $Profile } -argumentList $Profile.LocalPath))
        {
            Write-Host "LocalPath" $Profile.LocalPath "of SID $SID does not exist" -ForegroundColor "Yellow"
        }
        elseif (!$Profile.Loaded -and !$Profile.Special)
        {
            try
            {
                # See if the SID of the user matches to a domain account
                $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
                $FullUserName = $objUser.Value
                $NetBIOSDomainName = $FullUserName.SubString(0, $FullUserName.IndexOf("\"))

                Write-Host "Checking status of $FullUserName..." -NoNewline
                try
                {
                    $ADUserEnabled = Get-ADUser -LDAPFilter "(|(objectSid=$SID)(sIDHistory=$SID))" -Server $NetBIOSDomainName | Select-Object Enabled

                    # Check to see if the account was found
                    if (!$ADUserEnabled)
                    {
                        Write-Host " account not found."
                    }
                    else
                    {
                        Write-Host " account enabled is" $ADUserEnabled.Enabled
                        if (!$ADUserEnabled.Enabled)
                        {
                            Write-Host "Deleting profile for disabled user account $FullUserName ..." -ForegroundColor "Green"
                            try
                            {
                                Get-CimInstance -ComputerName $ComputerName -Class Win32_UserProfile | Where-Object { $_.SID -eq $SID } | Remove-CimInstance
                                $ProfilesDeleted = $ProfilesDeleted + 1
                            }
                            catch
                            {
                                Write-Host "Error while deleting profile $FullUserName"
                                Write-Host $_
                            }
                        }
                    }
                }
                catch
                {
                    Write-Host " error retrieving status of $FullUserName"
                    Write-Host $_
                }
            }
            catch
            {
                Write-Host "Error translating $SID to NT Account."
                Write-Host $_
            }
        }
        else
        {
            Write-Host "Skipping" $Profile.LocalPath "because loaded =" $Profile.Loaded "and special =" $Profile.Special
        }
    }
    if ($ProfilesDeleted -eq 0)
    {
        Write-Host "`nNo disabled user profiles found on $ComputerName" -ForegroundColor "Yellow"
    }
    else
    {
        Write-Host "`n$ProfilesDeleted disabled user profile" -NoNewline -ForegroundColor "Green"
        if ($ProfilesDeleted -gt 1) {Write-Host "s" -NoNewline -ForegroundColor "Green"}
        Write-Host " deleted on $ComputerName" -ForegroundColor "Green"
    }
    # Get free disk space after
    $CurrentDisk = Invoke-Command -Session $RemotePSSession {Get-PSDrive C} | Select-Object Used, Free
    Write-Host "`n"
    Write-Host "Initial disk space: "$InitialDisk.Free.ToString('N0')
    Write-Host "Current disk space: "$CurrentDisk.Free.ToString('N0')

    Remove-PSSession $RemotePSSession
}
else
{
    # Test-NetConnection failed
    Write-Host " test failed."
}
