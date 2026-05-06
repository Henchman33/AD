```powershell
<#
===========================================================================
Script Name: Enable-SAPServiceAccount-AES.ps1
Author: AD-GPT
Purpose:
    Enables Kerberos AES128 and AES256 encryption support for SAP service
    accounts in the specified OU.

Behavior:
    - Searches for accounts beginning with "SAPService"
    - Checks current msDS-SupportedEncryptionTypes value
    - Enables AES128 + AES256 if not already enabled
    - Skips accounts already configured correctly
    - Logs ALL actions to:
        %USERPROFILE%\Desktop\SAP Service Accounts Kerberos
    - Creates:
        - Transcript log
        - CSV report
        - Human-readable summary log

IMPORTANT:
    This script DOES NOT rotate passwords.
    Password rotation is still required later to generate AES keys.

Requirements:
    - RSAT ActiveDirectory module
    - Run as Domain Admin / delegated admin
    - PowerShell 5.1+
===========================================================================#
#>

#region INITIALIZATION

Import-Module ActiveDirectory -ErrorAction Stop

# Timestamp
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Desktop logging folder
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$LogFolder = Join-Path $DesktopPath "SAP Service Accounts Kerberos"

# Create folder if missing
if (!(Test-Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# Log files
$TranscriptLog = Join-Path $LogFolder "Transcript_$TimeStamp.txt"
$CsvReport     = Join-Path $LogFolder "AES_Results_$TimeStamp.csv"
$SummaryLog    = Join-Path $LogFolder "Summary_$TimeStamp.txt"

# Start transcript
Start-Transcript -Path $TranscriptLog

#endregion

#region CONFIGURATION

# OU Path
$SearchBase = "OU=SAP Users,DC=igtsap,DC=ad,DC=igt,DC=com"

# Account naming convention
$AccountPrefix = "SAPService*"

# AES128 + AES256
# 0x18 = 24 decimal
$AESValue = 24

#endregion

#region FUNCTIONS

function Write-Log {
    param (
        [string]$Message
    )

    $Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Date - $Message"

    Write-Host $Entry
    Add-Content -Path $SummaryLog -Value $Entry
}

#endregion

#region PROCESSING

Write-Log "===================================================="
Write-Log "Starting SAP Service Account AES Enablement Script"
Write-Log "OU: $SearchBase"
Write-Log "===================================================="

try {

    # Get SAP service accounts
    $Accounts = Get-ADUser `
        -Filter "SamAccountName -like '$AccountPrefix'" `
        -SearchBase $SearchBase `
        -Properties msDS-SupportedEncryptionTypes,
                    PasswordLastSet,
                    Enabled,
                    ServicePrincipalName

    if (!$Accounts) {
        Write-Log "No SAP service accounts found."
        Stop-Transcript
        return
    }

    $Results = @()

    foreach ($Account in $Accounts) {

        $CurrentValue = $Account.'msDS-SupportedEncryptionTypes'

        # Convert null to 0 for easier comparison
        if ($null -eq $CurrentValue) {
            $CurrentValue = 0
        }

        # Determine AES status
        $AES128Enabled = (($CurrentValue -band 0x08) -ne 0)
        $AES256Enabled = (($CurrentValue -band 0x10) -ne 0)

        Write-Log "----------------------------------------------------"
        Write-Log "Processing Account: $($Account.SamAccountName)"

        if ($AES128Enabled -and $AES256Enabled) {

            Write-Log "AES128 and AES256 already enabled. Skipping."

            $Status = "Already Configured"

        }
        else {

            try {

                Write-Log "Enabling AES128 + AES256..."

                Set-ADUser `
                    -Identity $Account `
                    -Replace @{
                        'msDS-SupportedEncryptionTypes' = $AESValue
                    }

                Write-Log "SUCCESS - AES encryption enabled."

                $Status = "Updated"

            }
            catch {

                Write-Log "ERROR - Failed to update account."
                Write-Log $_.Exception.Message

                $Status = "Failed"

            }

        }

        # Refresh account after modification
        $UpdatedAccount = Get-ADUser `
            -Identity $Account `
            -Properties msDS-SupportedEncryptionTypes

        # Build result object
        $Result = [PSCustomObject]@{

            SamAccountName               = $Account.SamAccountName
            DistinguishedName            = $Account.DistinguishedName
            Enabled                      = $Account.Enabled
            PasswordLastSet              = $Account.PasswordLastSet
            OriginalEncryptionValue      = $CurrentValue
            UpdatedEncryptionValue       = $UpdatedAccount.'msDS-SupportedEncryptionTypes'
            AES128Enabled                = (($UpdatedAccount.'msDS-SupportedEncryptionTypes' -band 0x08) -ne 0)
            AES256Enabled                = (($UpdatedAccount.'msDS-SupportedEncryptionTypes' -band 0x10) -ne 0)
            ServicePrincipalNames        = ($Account.ServicePrincipalName -join "; ")
            Status                       = $Status

        }

        $Results += $Result

    }

    # Export CSV
    $Results | Export-Csv -Path $CsvReport -NoTypeInformation -Encoding UTF8

    Write-Log "===================================================="
    Write-Log "Processing complete."
    Write-Log "CSV Report: $CsvReport"
    Write-Log "Transcript: $TranscriptLog"
    Write-Log "Summary Log: $SummaryLog"
    Write-Log "===================================================="

}
catch {

    Write-Log "FATAL ERROR"
    Write-Log $_.Exception.Message

}

#endregion

#region CLEANUP

Stop-Transcript

#endregion
```
