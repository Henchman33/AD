<#
Get-OU-WindowsServerActivation
================================================================================================================
This PowerShell Script will search the OU you want to find if the devices have activated it's version of Windows
Modify the $SearchBase - "OU=X, OU=Server, DC=XYZPDQ, DC=PDQ,DC=com" to fit your needs.
Run in PowerShell ISE as a Administrator
===============================================================================================================
#>

Import-Module ActiveDirectory

# ===== CONFIGURATION =====
$SearchBase = "OU=4 Prod,OU=Servers,DC=igtsap,DC=ad,DC=igt,DC=com"
$OutputFile = "C:\Temp\ServerActivationStatus.csv"
# =========================

Write-Host "Querying servers in OU: $SearchBase" -ForegroundColor Cyan

# Get all enabled server computer accounts from the OU
$Servers = Get-ADComputer -SearchBase $SearchBase `
                          -Filter { Enabled -eq $true } `
                          -Properties OperatingSystem |
          Where-Object { $_.OperatingSystem -like "*Server*" }

$Results = foreach ($Server in $Servers) {
    Write-Host "Checking $($Server.Name)..." -ForegroundColor Yellow

    try {
        $Activation = Invoke-Command -ComputerName $Server.Name -ScriptBlock {
            $lic = Get-CimInstance SoftwareLicensingProduct |
                   Where-Object {
                       $_.ApplicationID -eq '55c92734-d682-4d71-983e-d6ec3f16059f' -and
                       $_.PartialProductKey
                   }

            if ($lic.LicenseStatus -eq 1) {
                "Activated"
            } else {
                "Not Activated"
            }
        }

        [PSCustomObject]@{
            ServerName      = $Server.Name
            OperatingSystem = $Server.OperatingSystem
            ActivationState = $Activation
        }
    }
    catch {
        [PSCustomObject]@{
            ServerName      = $Server.Name
            OperatingSystem = $Server.OperatingSystem
            ActivationState = "Unreachable / Error"
        }
    }
}

# Display in ISE
$Results | Out-GridView -Title "Windows Activation Status"

# Export to CSV
$Results | Export-Csv -Path $OutputFile -NoTypeInformation

Write-Host "Report saved to $OutputFile" -ForegroundColor Green
