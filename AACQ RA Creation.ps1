Function Get-FileName($PromptTitle = "Select the CSV file")
{
    $initialDirectory = $PSScriptRoot

    # Ensure System.Windows.Forms assembly is loaded
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Warning "Could not load System.Windows.Forms assembly. File dialog may not work."
        Return $null # Indicate failure
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
function CreateRA
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

    Write-Output "`nStart creating $RaType Resource Accounts..."
    $createdAccounts = [System.Collections.Generic.List[string]]::new() # Keep track of successfully created UPNs

    foreach($Entry in $CsvData){
        # Validate required CSV columns exist for the current entry (UsageLocation no longer needed here)
        if (-not $Entry.PSObject.Properties.Name -contains 'UPN' -or
            -not $Entry.PSObject.Properties.Name -contains 'DisplayName') {
             Write-Warning "Skipping entry due to missing 'UPN' or 'DisplayName' column: $($Entry | Out-String)"
             continue # Skip to the next entry
        }

        $UPN = $Entry.UPN
        $DisplayName = $Entry.DisplayName

        # Basic validation for UPN format
        if ($UPN -notlike '*@*') {
            Write-Warning "Skipping entry with potentially invalid UPN format: $UPN"
            continue
        }

        # Usage Location is now set during the licensing phase
        Write-Host "Attempting to create RA: UPN='$UPN', DisplayName='$DisplayName', Type='$RaType'"
        try {
            # Create the Application Instance (which also creates the underlying user)
            New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId $AppId -DisplayName $DisplayName -ErrorAction Stop | Out-Null
            Write-Output "$UPN ... created successfully as $RaType type." -ForegroundColor Green
            $createdAccounts.Add($UPN) # Add UPN to the list for licensing
        } catch {
            Write-Error "Failed to create Resource Account $UPN. Error: $($_.Exception.Message)"
            # If creation fails, it won't be added to the list for licensing
        }
    }
    Write-Output "$RaType Resource Account creation process finished."
    return $createdAccounts # Return the list of UPNs that were created
}

# --- Core Function to Set License and Usage Location ---
function SetRALicense
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [array]$CsvData, # Still needed to iterate and get UPNs

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$LicenseSkuId,

        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[string]]$SuccessfullyCreatedUpns # Only process accounts confirmed created
    )

    # Hardcoded Usage Location
    $FixedUsageLocation = 'AU'

    Write-Output "`nStart assigning Usage Location ('$FixedUsageLocation') and License..."
    $sleepSeconds = 45 # Delay before first attempt
    Write-Host "Waiting $sleepSeconds seconds for Azure AD replication before starting licensing..." -ForegroundColor Yellow
    Start-Sleep -Seconds $sleepSeconds

    foreach($Entry in $CsvData) {
        $UPN = $Entry.UPN
        # Only process if this UPN was successfully created in the CreateRA step
        if ($SuccessfullyCreatedUpns -contains $UPN) {

            Write-Host "Attempting to set Usage Location and License for: $UPN"
            try {
                # 1. Set Usage Location (using the hardcoded value)
                Write-Host " - Setting Usage Location to '$FixedUsageLocation' for $UPN..."
                Set-MsolUser -UserPrincipalName $UPN -UsageLocation $FixedUsageLocation -ErrorAction Stop

                # Optional short delay between setting location and license
                Start-Sleep -Seconds 5

                # 2. Assign License
                Write-Host " - Assigning License '$LicenseSkuId' to $UPN..."
                Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $LicenseSkuId -ErrorAction Stop

                Write-Output "$UPN ... Usage Location set to '$FixedUsageLocation' and License '$LicenseSkuId' assigned successfully." -ForegroundColor Green
            } catch {
                Write-Error "Failed to set Usage Location or License for $UPN. Error: $($_.Exception.Message)"
                # If it failed, the user might need manual intervention later
            }
            # Add a small delay before processing the next user to avoid throttling
            Start-Sleep -Seconds 2
        }
        # If UPN is not in SuccessfullyCreatedUpns, skip silently as it wasn't created or failed earlier.
    }
    Write-Output "`nUsage Location ('$FixedUsageLocation') and Licensing process finished."
}


# --- Main Script Logic ---
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "This script will:"
Write-Host " 1. Bulk create Resource Accounts based on separate CSV files for AAs and CQs."
Write-Host " 2. Optionally attempt to assign the 'Microsoft Teams Phone Resource Account' license to successfully created accounts."
Write-Host " 3. Automatically set the Usage Location to 'AU' for licensed accounts."
Write-Host ""
Write-Host "Prereqs:"
Write-Host " - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)"
Write-Host " - MSOnline PowerShell Module installed (`Install-Module MSOnline`)"
Write-Host " - Separate CSV files prepared for AAs and CQs (each with 'UPN' and 'DisplayName' columns)"
Write-Host " - Sufficient permissions (e.g., Teams Admin, User Admin/License Admin)"
Write-Host "================================================================================================================================================="
Write-Host ""

# --- Connect to Services ---
# Connect to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..."
try {
    # Simple connect - assuming first time run
    Connect-MicrosoftTeams -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
     Write-Error "Failed to connect to Microsoft Teams. Please check module installation and credentials. Error: $($_.Exception.Message)"
     Read-Host "Press Enter to exit."
     exit 1
}
Start-Sleep -Seconds 2

# Connect to MSOnline for Licensing
Write-Host "Connecting to MSOnline (for licensing)..."
$msOnlineConnected = $false # Flag to track connection status
try {
    # Check if MSOnline module exists
    if (-not (Get-Module -ListAvailable -Name MSOnline)) {
        Write-Warning "MSOnline PowerShell module not found. Please install it: Install-Module MSOnline. Licensing will be skipped."
    } else {
        # Simple connect
        Connect-MsolService -ErrorAction Stop
        Write-Host "Successfully connected to MSOnline." -ForegroundColor Green
        $msOnlineConnected = $true
    }
} catch {
     Write-Warning "Failed to connect to MSOnline. Please check module installation and credentials. Licensing will be skipped. Error: $($_.Exception.Message)"
}
Start-Sleep -Seconds 2

# --- Get the Resource Account License SKU ---
$LicenseID = $null
if ($msOnlineConnected) {
    Write-Host "Finding the 'Microsoft Teams Phone Resource Account' license SKU..."
    try {
        # Find the SKU. The AccountSkuId often contains "PHONESYSTEM_VIRTUALUSER"
        $LicenseSku = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -like "*:PHONESYSTEM_VIRTUALUSER"} -ErrorAction Stop
        if ($null -eq $LicenseSku) {
            Write-Warning "Could not find the 'Microsoft Teams Phone Resource Account' (PHONESYSTEM_VIRTUALUSER) license SKU in your tenant. Licensing will be skipped."
        } elseif ($LicenseSku.Count -gt 1) {
            Write-Warning "Multiple SKUs found matching '*:PHONESYSTEM_VIRTUALUSER'. Using the first one: $($LicenseSku[0].AccountSkuId)"
            $LicenseSku = $LicenseSku[0]
            $LicenseID = $LicenseSku.AccountSkuId
            Write-Host "Found License SKU: $LicenseID" -ForegroundColor Green
        } else {
            $LicenseID = $LicenseSku.AccountSkuId
            Write-Host "Found License SKU: $LicenseID" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Failed to query MSOL Account SKUs. Licensing will be skipped. Error: $($_.Exception.Message)"
    }
} else {
    Write-Warning "Skipping license SKU retrieval as MSOnline connection failed."
}

# Variables to hold results
$createdAAupns = $null
$createdCQupns = $null
$AA_Csv_Data = $null
$CQ_Csv_Data = $null

# --- Process Auto Attendant CSV ---
Do{
    Write-Output ""
    $CreateAA_Choice = Read-Host -Prompt 'Would you like to create Auto Attendant Resource Accounts? (Y/N)'
}until(($CreateAA_Choice -eq "Y") -or ($CreateAA_Choice -eq "N"))

if ($CreateAA_Choice -eq "Y") {
    Write-Host "`n--- Processing Auto Attendants ---" -ForegroundColor Cyan
    $AA_OnlineApplicationID = "ce933385-9390-45d1-9512-c8d228074e07"
    $AA_ConfigurationType = "AutoAttendant"

    Write-Host "Select the CSV file containing the Auto Attendant Resource Accounts."
    Write-Host "Please ensure it has 'UPN' and 'DisplayName' columns."
    $AA_Path = Get-FileName -PromptTitle "Select the AUTO ATTENDANT CSV file"
    if (-not $AA_Path) {
        Write-Warning "No CSV file selected for Auto Attendants. Skipping AA processing."
    } else {
        Write-Host "Attempting to import AA CSV from: $AA_Path"
        try {
            $AA_Csv_Data = Import-Csv -Path $AA_Path -ErrorAction Stop # Store data
            if ($null -eq $AA_Csv_Data -or $AA_Csv_Data.Count -eq 0) {
                Write-Error "AA CSV file '$AA_Path' is empty or could not be read properly."
            } else {
                Write-Host "Successfully imported $($AA_Csv_Data.Count) entries from AA CSV." -ForegroundColor Green
                # Create the Auto Attendant Resource Accounts and get list of successfully created UPNs
                $createdAAupns = CreateRA -CsvData $AA_Csv_Data -AppId $AA_OnlineApplicationID -RaType $AA_ConfigurationType

                # --- License Auto Attendants? ---
                if ($msOnlineConnected -and $LicenseID -ne $null -and $createdAAupns -ne $null -and $createdAAupns.Count -gt 0) {
                    Do {
                        Write-Output ""
                        $LicenseAA_Choice = Read-Host -Prompt "Would you like to license the created Auto Attendant Resource Accounts now? (Y/N)"
                    } until (($LicenseAA_Choice -eq "Y") -or ($LicenseAA_Choice -eq "N"))

                    if ($LicenseAA_Choice -eq "Y") {
                        Write-Host "`n--- Licensing Auto Attendants (UsageLocation: AU) ---" -ForegroundColor Cyan
                        SetRALicense -CsvData $AA_Csv_Data -LicenseSkuId $LicenseID -SuccessfullyCreatedUpns $createdAAupns
                    } else {
                        Write-Host "`nSkipping licensing for Auto Attendant Resource Accounts."
                    }
                } elseif (-not $msOnlineConnected) {
                    Write-Warning "Skipping licensing of Auto Attendants because connection to MSOnline failed."
                } elseif ($LicenseID -eq $null) {
                    Write-Warning "Skipping licensing of Auto Attendants because the license SKU was not found."
                }
            }
        } catch {
            Write-Error "Failed to import AA CSV file '$AA_Path'. Error: $($_.Exception.Message)"
            Write-Warning "Skipping Auto Attendant creation due to CSV import error."
        }
    }
} else {
    Write-Host "`nSkipping Auto Attendant Resource Account creation."
}


# --- Process Call Queue CSV ---
Do{
    Write-Output ""
    $CreateCQ_Choice = Read-Host -Prompt 'Would you like to create Call Queue Resource Accounts? (Y/N)'
}until(($CreateCQ_Choice -eq "Y") -or ($CreateCQ_Choice -eq "N"))

if ($CreateCQ_Choice -eq "Y") {
    Write-Host "`n--- Processing Call Queues ---" -ForegroundColor Cyan
    $CQ_OnlineApplicationID = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
    $CQ_ConfigurationType = "CallQueue"

    Write-Host "Select the CSV file containing the Call Queue Resource Accounts."
    Write-Host "Please ensure it has 'UPN' and 'DisplayName' columns."
    $CQ_Path = Get-FileName -PromptTitle "Select the CALL QUEUE CSV file"
    if (-not $CQ_Path) {
        Write-Warning "No CSV file selected for Call Queues. Skipping CQ processing."
    } else {
        Write-Host "Attempting to import CQ CSV from: $CQ_Path"
        try {
            $CQ_Csv_Data = Import-Csv -Path $CQ_Path -ErrorAction Stop # Store data
            if ($null -eq $CQ_Csv_Data -or $CQ_Csv_Data.Count -eq 0) {
                Write-Error "CQ CSV file '$CQ_Path' is empty or could not be read properly."
            } else {
                Write-Host "Successfully imported $($CQ_Csv_Data.Count) entries from CQ CSV." -ForegroundColor Green
                # Create the Call Queue Resource Accounts and get list of successfully created UPNs
                $createdCQupns = CreateRA -CsvData $CQ_Csv_Data -AppId $CQ_OnlineApplicationID -RaType $CQ_ConfigurationType

                # --- License Call Queues? ---
                if ($msOnlineConnected -and $LicenseID -ne $null -and $createdCQupns -ne $null -and $createdCQupns.Count -gt 0) {
                    Do {
                        Write-Output ""
                        $LicenseCQ_Choice = Read-Host -Prompt "Would you like to license the created Call Queue Resource Accounts now? (Y/N)"
                    } until (($LicenseCQ_Choice -eq "Y") -or ($LicenseCQ_Choice -eq "N"))

                    if ($LicenseCQ_Choice -eq "Y") {
                        Write-Host "`n--- Licensing Call Queues (UsageLocation: AU) ---" -ForegroundColor Cyan
                        SetRALicense -CsvData $CQ_Csv_Data -LicenseSkuId $LicenseID -SuccessfullyCreatedUpns $createdCQupns
                    } else {
                        Write-Host "`nSkipping licensing for Call Queue Resource Accounts."
                    }
                } elseif (-not $msOnlineConnected) {
                    Write-Warning "Skipping licensing of Call Queues because connection to MSOnline failed."
                } elseif ($LicenseID -eq $null) {
                    Write-Warning "Skipping licensing of Call Queues because the license SKU was not found."
                }
            }
        } catch {
            Write-Error "Failed to import CQ CSV file '$CQ_Path'. Error: $($_.Exception.Message)"
            Write-Warning "Skipping Call Queue creation due to CSV import error."
        }
    }
} else {
    Write-Host "`nSkipping Call Queue Resource Account creation."
}

# --- Disconnect ---
Write-Host ""
Write-Host "Script finished processing Resource Account creation and optional licensing (UsageLocation set to AU for licensed accounts)."
Write-Host "Disconnecting from services..."
# Optional: Disconnect sessions
Disconnect-MicrosoftTeams
# Disconnect-MsolService # Generally not required, session expires

Read-Host "Press Enter to exit."