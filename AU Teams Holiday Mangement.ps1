<#
.SYNOPSIS
    Imports new Public Holiday schedules into Microsoft Teams Voice.

.DESCRIPTION
    This script retrieves Public Holiday data from data.gov.au for specified
    Australian states and creates new holiday schedules in Microsoft Teams Voice.

.NOTES
    Prerequisites:
     - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)
     - Connectivity to data.gov.au
     - Sufficient permissions to manage Teams Voice call flow schedules.
#>
#region Initial Setup
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "Starting New Public Holiday Import Script"
Write-Host "This script will:"
Write-Host " 1. Retrieve Public Holiday data from data.gov.au for specified Australian states."
Write-Host " 2. Create new Public Holiday schedules in Microsoft Teams Voice."
Write-Host ""
Write-Host "Prerequisites:"
Write-Host " - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)"
Write-Host " - Connectivity to data.gov.au"
Write-Host " - Sufficient permissions to manage Teams Voice call flow schedules."
Write-Host ""
Write-Host "Notes"
Write-Host " - To Skip Specfic holidays, modify the filter in the Get-PublicHolidays function."
Write-Host ""
Write-Host "================================================================================================================================================="
#endregion

#region Connect to Teams
try {
    Connect-MicrosoftTeams | Out-Null
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Ensure the module is installed and you have the correct permissions."
    Exit
}
#endregion

#region Functions
function Get-PublicHolidays {
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("sa", "qld", "act", "nsw", "wa", "nt", "vic", "tas")]
        [string[]] $Jurisdictions
    )

    if (-not $Jurisdictions) {
        $JurisdictionInput = Read-Host "Enter the Australian state codes (comma-separated, e.g., sa,qld,act):"
        $Jurisdictions = $JurisdictionInput.Split(',') | ForEach-Object { $_.Trim() }
        foreach ($Jurisdiction in $Jurisdictions){
             if ($Jurisdiction -notin ("sa", "qld", "act", "nsw", "wa", "nt", "vic", "tas")) {
                 Write-Warning "Invalid state code: $Jurisdiction.  It will be ignored. Please use sa, qld, act, nsw, wa, nt, vic, or tas."
             }
        }
        $Jurisdictions = $Jurisdictions | Where-Object {$_ -in ("sa", "qld", "act", "nsw", "wa", "nt", "vic", "tas")}
        if($Jurisdictions.Count -eq 0){
             Write-Error "No valid state codes provided."
             return
        }
    }

    $Holidays = @()
    $CurrentYear = (Get-Date).Year
    $ResourceID = '33673aca-0857-42e5-b8f0-9981b4755686'
    $URI = "https://data.gov.au/data/api/3/action/datastore_search_sql?sql=SELECT * from `"$ResourceID`""
    $allJurisdictionCodes = @()

    foreach ($Jurisdiction in $Jurisdictions) {
        try {
            $Results = (Invoke-RestMethod -Uri $URI -Method Get -ErrorAction Stop).Result.records |
                 Where-Object {
                     $_.Jurisdiction -eq $Jurisdiction -and
                     ([datetime]::ParseExact($_.date, 'yyyyMMdd', $null)).Year -eq $CurrentYear
                 }
        }
        catch {
            Write-Error "Failed to retrieve public holiday data from data.gov.au for ${Jurisdiction}: $($_.Exception.Message)"
            continue
        }

        foreach ($Holiday in $Results) {
            if ($Holiday.'Holiday Name' -eq "Bank Holiday") { continue }

            $HolidayDate = ([datetime]::ParseExact($Holiday.date, 'yyyyMMdd', $null)).ToShortDateString()
            $HolidayEndDate = ([datetime]::ParseExact($Holiday.date, 'yyyyMMdd', $null)).AddDays(1).ToShortDateString()

            $Holidays += [PSCustomObject]@{
                StartDate = $HolidayDate
                EndDate   = $HolidayEndDate
                Name      = $Holiday.'Holiday Name'
                State     = $Jurisdiction.ToUpper()
            }
        }
        $allJurisdictionCodes += $Jurisdiction.ToUpper()
    }
    return $Holidays, ($allJurisdictionCodes -join ", ")
}

function Import-NewHolidays {
    $TeamsVoiceHolidays = Get-CsOnlineSchedule | Where-Object { $_.FixedSchedule }
    $sameAction = Read-Host "Will the same action apply to all imported holidays? (Y/N)"
     while ($sameAction -notin ("Y", "N")) {
         Write-Warning "Invalid input. Please enter 'Y' or 'N'."
         $sameAction = Read-Host "Will the same action apply to all imported holidays? (Y/N)"
    }
    $startTime = "00:00"
    $endTime   = "23:45"

    Write-Host "`nFetching Public Holidays from data.gov.au..." -ForegroundColor Yellow
    $results = Get-PublicHolidays
    $PublicHolidays = $results[0]
    $JurisdictionCode = $results[1]

    if ($PublicHolidays) {
        if ($sameAction -eq "N") {
            foreach ($Holiday in $PublicHolidays) {
                $scheduleName = "$($Holiday.State) $($Holiday.Name)"
                $startDateTime = Get-Date "$($Holiday.StartDate) $startTime"
                $endDateTime = Get-Date "$($Holiday.EndDate) $endTime"
                $startStr = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                $endStr = $endDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr

                $existing = $TeamsVoiceHolidays | Where-Object { $_.Name -eq $scheduleName }
                if ($existing) {
                    Write-Warning "Holiday schedule '$($scheduleName)' already exists. Skipping creation."
                } else {
                    New-CsOnlineSchedule -Name $scheduleName -FixedSchedule -DateTimeRanges @($dtr)
                }
            }
        } else {
            $holidaysByState = @{}
            foreach ($Holiday in $PublicHolidays) {
                if (-not $holidaysByState.ContainsKey($Holiday.State)) {
                    $holidaysByState[$Holiday.State] = @()
                }
                $holidaysByState[$Holiday.State] += $Holiday
            }

            foreach ($state in $holidaysByState.Keys) {
                $stateHolidays = $holidaysByState[$state]
                $dateTimeRanges = @()
                foreach ($Holiday in $stateHolidays) {
                    Write-Host " - $($Holiday.Name) - Start: $($Holiday.StartDate), End: $($Holiday.EndDate) - State: $($Holiday.State)"

                    $startDateTime = Get-Date "$($Holiday.StartDate) $startTime"
                    $endDateTime = Get-Date "$($Holiday.EndDate) $endTime"
                    $startStr = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                    $endStr = $endDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                    $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr
                    $dateTimeRanges += $dtr
                }
                $scheduleName = "$state Public Holiday"
                $existing = $TeamsVoiceHolidays | Where-Object { $_.Name -eq $scheduleName }
                if ($existing) {
                    Write-Warning "Holiday schedule '$($scheduleName)' already exists. Skipping creation."
                } else {
                    New-CsOnlineSchedule -Name $scheduleName -FixedSchedule -DateTimeRanges $dateTimeRanges
                }
            }
        }
    } else {
        Write-Host "No public holidays retrieved for the specified state."
    }
}
#endregion

#region Main Execution
Import-NewHolidays
#endregion

#region Finalization
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "New Public Holiday Import Script finished."
Write-Host "================================================================================================================================================="
Write-Host ""
#endregion
