<#
.SYNOPSIS
Bulk Creation and Licensing of Teams Resource Accounts (Auto Attendants & Call Queues)

.DESCRIPTION
This PowerShell script streamlines the creation of Microsoft Teams Resource Accounts for Auto Attendants and Call Queues by processing data from separate CSV files. It also offers the option to automatically assign the 'Microsoft Teams Phone Resource Account' license and set the Usage Location to 'AU' for these accounts.

.NOTES
Author: Your Name/Organization
Version: 1.0
Date: 2025-04-09

Requires:
 - Microsoft Teams PowerShell Module (Install-Module MicrosoftTeams)
 - MSOnline PowerShell Module (Install-Module MSOnline)
 - CSV files for Auto Attendants and Call Queues with 'UPN' and 'DisplayName' columns.
 - Sufficient administrative permissions in your Microsoft 365 tenant.

.PARAMETER AACsvPath
Path to the CSV file containing Auto Attendant Resource Account details (UPN, DisplayName).

.PARAMETER CQCsvPath
Path to the CSV file containing Call Queue Resource Account details (UPN, DisplayName).

.PARAMETER License
Switch parameter. Include this to attempt to license the created Resource Accounts.

.EXAMPLE
.\ResourceAccountImport.ps1 -AACsvPath "C:\temp\aa_accounts.csv" -CQCsvPath "C:\temp\cq_accounts.csv" -License
This command will import Auto Attendant and Call Queue data from the specified CSV files, create the Resource Accounts, and attempt to assign the necessary license.

.EXAMPLE
.\ResourceAccountImport.ps1 -AACsvPath "C:\temp\aa_accounts.csv" -CQCsvPath "C:\temp\cq_accounts.csv"
This command will import Auto Attendant and Call Queue data from the specified CSV files and create the Resource Accounts without attempting to assign licenses.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$AACsvPath,

    [Parameter(Mandatory=$false)]
    [string]$CQCsvPath,

    [Parameter(Mandatory=$false)]
    [switch]$License
)

# --- Helper Function to Select a CSV File ---
function Get-FileName($PromptTitle = "Select the CSV file")
{
    $initialDirectory = $PSScriptRoot

    # Ensure System.Windows.Forms assembly is loaded
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Warning "Couldn't load System.Windows.Forms assembly. File dialog might not work."
        return $null # Indicate failure
    }

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.Title = $PromptTitle # Use the provided title
    if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $OpenFileDialog.FileName
    } else {
        Write-Warning "No file selected."
        return $null # Indicate cancellation or failure
    }
}

# --- Core Function to Create Resource Accounts ---
function Create-ResourceAccount
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [array]$CsvData,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$AppId,

        [Parameter(Mandatory=$true)]
        [ValidateSet("AutoAttendant","CallQueue")]
        [string]$RaType
    )

    Write-Output "`nStarting creation of $RaType Resource Accounts..."
    $createdAccounts = [System.Collections.Generic.List[string]]::new() # Track successful UPNs

    foreach($Entry in $CsvData){
        # Validate required CSV columns
        if (-not $Entry.PSObject.Properties.Name -contains 'UPN' -or
            -not $Entry.PSObject.Properties.Name -contains 'DisplayName') {
             Write-Warning "Skipping entry due to missing 'UPN' or 'DisplayName' column: $($Entry | Out-String)"
             continue # Skip to the next entry
        }

        $UPN = $Entry.UPN
        $DisplayName = $Entry.DisplayName

        # Basic UPN format validation
        if ($UPN -notlike '*@*') {
            Write-Warning "Skipping entry with potentially invalid UPN format: $UPN"
            continue
        }

        Write-Host "Attempting to create RA: UPN='$UPN', DisplayName='$DisplayName', Type='$RaType'"
        try {
            # Create the Application Instance
            New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId $AppId -DisplayName $DisplayName -ErrorAction Stop | Out-Null
            Write-Output "$UPN ... created successfully as $RaType type." -ForegroundColor Green
            $createdAccounts.Add($UPN) # Add to the success list
        } catch {
            Write-Error "Failed to create Resource Account $UPN. Error: $($_.Exception.Message)"
            # Not adding to success list if creation failed
        }
    }
    Write-Output "$RaType Resource Account creation process finished."
    return $createdAccounts # Return list of created UPNs
}

# --- Core Function to Set License and Usage Location ---
function Set-ResourceAccountLicense
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [array]$CsvData, # Needed to iterate and get UPNs

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LicenseSkuId,

        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[string]]$SuccessfullyCreatedUpns # Only process created accounts
    )

    # Fixed Usage Location
    $FixedUsageLocation = 'AU'

    Write-Output "`nStarting to assign Usage Location ('$FixedUsageLocation') and License..."
    $sleepSeconds = 30 # Short delay before first attempt
    Write-Host "Waiting $sleepSeconds seconds for Azure AD replication before starting licensing..." -ForegroundColor Yellow
    Start-Sleep -Seconds $sleepSeconds

    foreach($Entry in $CsvData) {
        $UPN = $Entry.UPN
        # Process only if the UPN was successfully created
        if ($SuccessfullyCreatedUpns -contains $UPN) {
            Write-Host "Attempting to set Usage Location and License for: $UPN"
            try {
                # 1. Set Usage Location
                Write-Host " - Setting Usage Location to '$FixedUsageLocation' for $UPN..."
                Set-MsolUser -UserPrincipalName $UPN -UsageLocation $FixedUsageLocation -ErrorAction Stop

                # Optional short delay
                Start-Sleep -Seconds 3

                # 2. Assign License
                Write-Host " - Assigning License '$LicenseSkuId' to $UPN..."
                Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $LicenseSkuId -ErrorAction Stop

                Write-Output "$UPN ... Usage Location set to '$FixedUsageLocation' and License '$LicenseSkuId' assigned successfully." -ForegroundColor Green
            } catch {
                Write-Error "Failed to set Usage Location or License for $UPN. Error: $($_.Exception.Message)"
                # Might need manual intervention
            }
            # Small delay to avoid throttling
            Start-Sleep -Seconds 2
        }
        # If not in SuccessfullyCreatedUpns, skip silently
    }
    Write-Output "`nUsage Location ('$FixedUsageLocation') and Licensing process finished."
}


# --- Main Script Logic ---
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "This script will:"
Write-Host " 1. Bulk create Resource Accounts for Auto Attendants and Call Queues from separate CSV files."
Write-Host " 2. Optionally assign the 'Microsoft Teams Phone Resource Account' license and set Usage Location to 'AU'."
Write-Host ""
Write-Host "Prerequisites:"
Write-Host " - Microsoft Teams PowerShell Module installed."
Write-Host " - MSOnline PowerShell Module installed."
Write-Host " - Separate CSV files for AAs and CQs (with 'UPN' and 'DisplayName' columns)."
Write-Host " - Sufficient administrative permissions."
Write-Host "================================================================================================================================================="
Write-Host ""

# --- Connect to Services ---
# Connect to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..."
try {
    Connect-MicrosoftTeams -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Check module and credentials. Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit."
    exit 1
}
Start-Sleep -Seconds 1

# Connect to MSOnline for Licensing
$msOnlineConnected = $false
$LicenseID = $null
if ($License) {
    Write-Host "Connecting to MSOnline (for licensing)..."
    try {
        if (-not (Get-Module -ListAvailable -Name MSOnline)) {
            Write-Warning "MSOnline PowerShell module not found. Please install it: Install-Module MSOnline. Licensing will be skipped."
        } else {
            Connect-MsolService -ErrorAction Stop
            Write-Host "Successfully connected to MSOnline." -ForegroundColor Green
            $msOnlineConnected = $true

            # --- Get the Resource Account License SKU ---
            Write-Host "Finding 'Microsoft Teams Phone Resource Account' license SKU..."
            try {
                $LicenseSku = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -like "*:PHONESYSTEM_VIRTUALUSER"} -ErrorAction Stop
                if ($null -eq $LicenseSku) {
                    Write-Warning "License SKU 'PHONESYSTEM_VIRTUALUSER' not found. Licensing will be skipped."
                } elseif ($LicenseSku.Count -gt 1) {
                    Write-Warning "Multiple matching SKUs found. Using the first one: $($LicenseSku[0].AccountSkuId)"
                    $LicenseID = $LicenseSku[0].AccountSkuId
                    Write-Host "Found License SKU: $LicenseID" -ForegroundColor Green
                } else {
                    $LicenseID = $LicenseSku.AccountSkuId
                    Write-Host "Found License SKU: $LicenseID" -ForegroundColor Green
                }
            } catch {
                Write-Warning "Failed to query MSOL Account SKUs. Licensing will be skipped. Error: $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Warning "Failed to connect to MSOnline. Check module and credentials. Licensing skipped. Error: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds 1
} else {
    Write-Host "Skipping MSOnline connection and licensing."
}

# Variables to hold created UPNs and CSV data
$createdAAupns = $null
$createdCQupns = $null
$AA_Csv_Data = $null
$CQ_Csv_Data = $null

# --- Process Auto Attendant CSV ---
Write-Host "`n--- Processing Auto Attendants ---" -ForegroundColor Cyan
$AA_OnlineApplicationID = "ce933385-9390-45d1-9512-c8d228074e07"
$AA_ConfigurationType = "AutoAttendant"

if (-not $AACsvPath) {
    Write-Host "Please select the CSV file for Auto Attendants."
    Write-Host "Ensure it has 'UPN' and 'DisplayName' columns."
    $AA_Path = Get-FileName -PromptTitle "Select the AUTO ATTENDANT CSV file"
} else {
    $AA_Path = $AACsvPath
    Write-Host "Using provided Auto Attendant CSV path: $AA_Path"
}

if ($AA_Path) {
    Write-Host "Importing Auto Attendant CSV from: $AA_Path"
    try {
        $AA_Csv_Data = Import-Csv -Path $AA_Path -ErrorAction Stop
        if ($null -eq $AA_Csv_Data -or $AA_Csv_Data.Count -eq 0) {
            Write-Error "Auto Attendant CSV file '$AA_Path' is empty or could not be read."
        } else {
            Write-Host "Successfully imported $($AA_Csv_Data.Count) entries from AA CSV." -ForegroundColor Green
            # Create Auto Attendant Resource Accounts
            $createdAAupns = Create-ResourceAccount -CsvData $AA_Csv_Data -AppId $AA_OnlineApplicationID -RaType $AA_ConfigurationType

            # --- License Auto Attendants if requested and possible ---
            if ($License -and $msOnlineConnected -and $LicenseID -ne $null -and $createdAAupns -ne $null -and $createdAAupns.Count -gt 0) {
                Write-Host "`n--- Licensing Auto Attendants (UsageLocation: AU) ---" -ForegroundColor Cyan
                Set-ResourceAccountLicense -CsvData $AA_Csv_Data -LicenseSkuId $LicenseID -SuccessfullyCreatedUpns $createdAAupns
            } elseif ($License -and -not $msOnlineConnected) {
                Write-Warning "Skipping Auto Attendant licensing due to failed MSOnline connection."
            } elseif ($License -and $LicenseID -eq $null) {
                Write-Warning "Skipping Auto Attendant licensing as the license SKU was not found."
            } elseif ($createdAAupns -eq $null -or $createdAAupns.Count -eq 0) {
                Write-Warning "No Auto Attendant Resource Accounts were successfully created, skipping licensing."
            }
        }
    } catch {
        Write-Error "Failed to import Auto Attendant CSV file '$AA_Path'. Error: $($_.Exception.Message)"
        Write-Warning "Skipping Auto Attendant creation."
    }
} else {
    Write-Warning "No Auto Attendant CSV file selected. Skipping AA processing."
}

# --- Process Call Queue CSV ---
Write-Host "`n--- Processing Call Queues ---" -ForegroundColor Cyan
$CQ_OnlineApplicationID = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
$CQ_ConfigurationType = "CallQueue"

if (-not $CQCsvPath) {
    Write-Host "Please select the CSV file for Call Queues."
    Write-Host "Ensure it has 'UPN' and 'DisplayName' columns."
    $CQ_Path = Get-FileName -PromptTitle "Select the CALL QUEUE CSV file"
} else {
    $CQ_Path = $CQCsvPath
    Write-Host "Using provided Call Queue CSV path: $CQ_Path"
}

if ($CQ_Path) {
    Write-Host "Importing Call Queue CSV from: $CQ_Path"
    try {
        $CQ_Csv_Data = Import-Csv -Path $CQ_Path -ErrorAction Stop
        if ($null -eq $CQ_Csv_Data -or $CQ_Csv_Data.Count -eq 0) {
            Write-Error "Call Queue CSV file '$CQ_Path' is empty or could not be read."
        } else {
            Write-Host "Successfully imported $($CQ_Csv_Data.Count) entries from CQ CSV." -ForegroundColor Green
            # Create Call Queue Resource Accounts
            $createdCQupns = Create-ResourceAccount -CsvData $CQ_Csv_Data -AppId $CQ_OnlineApplicationID -RaType $CQ_ConfigurationType

            # --- License Call Queues if requested and possible ---
            if ($License -and $msOnlineConnected -and $LicenseID -ne $null -and $createdCQupns -ne $null -and $createdCQupns.Count -gt 0) {
                Write-Host "`n--- Licensing Call Queues (UsageLocation: AU) ---" -ForegroundColor Cyan
                Set-ResourceAccountLicense -CsvData $CQ_Csv_Data -LicenseSkuId $LicenseID -SuccessfullyCreatedUpns $createdCQupns
            } elseif ($License -and -not $msOnlineConnected) {
                Write-Warning "Skipping Call Queue licensing due to failed MSOnline connection."
            } elseif ($License -and $LicenseID -eq $null) {
                Write-Warning "Skipping Call Queue licensing as the license SKU was not found."
            } elseif ($createdCQupns -eq $null -or $createdCQupns.Count -eq 0) {
                Write-Warning "No Call Queue Resource Accounts were successfully created, skipping licensing."
            }
        }
    } catch {
        Write-Error "Failed to import Call Queue CSV file '$CQ_Path'. Error: $($_.Exception.Message)"
        Write-Warning "Skipping Call Queue creation."
    }
} else {
    Write-Warning "No Call Queue CSV file selected. Skipping CQ processing."
}

# --- Disconnect ---
Write-Host ""
Write-Host "Script finished processing Resource Account creation and optional licensing (UsageLocation set to AU for licensed accounts)."
Write-Host "Disconnecting from services..."
# Optional: Disconnect sessions
Disconnect-MicrosoftTeams
# Disconnect-MsolService # Generally not needed

Read-Host "Press Enter to exit."
