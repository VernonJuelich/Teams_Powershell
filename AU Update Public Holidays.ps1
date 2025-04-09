<#
.SYNOPSIS
    Updates existing Public Holiday schedules in Microsoft Teams Voice.

.DESCRIPTION
    This script retrieves current Public Holiday data from data.gov.au for
    specified Australian states and updates existing holiday schedules in
    Microsoft Teams Voice for holidays occurring in the current or next year.

.NOTES
    Prerequisites:
     - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)
     - Connectivity to data.gov.au
     - Sufficient permissions to manage Teams Voice call flow schedules.
     - Existing Public Holiday schedules in Teams Voice (assumes specific naming conventions).

    Assumptions:
     - This script assumes existing holiday schedules follow a naming convention
       that includes the state code (e.g., "NSW Boxing Day" or "NSW Public Holiday").

    Notes:
     - To skip specific holidays, modify the filter in the Get-PublicHolidays function.
#>
#region Initial Setup
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "Starting Public Holiday Update Script"
Write-Host "This script will:"
Write-Host " 1. Retrieve Public Holiday data from data.gov.au for specified Australian states."
Write-Host " 2. Update existing Public Holiday schedules in Microsoft Teams Voice"
Write-Host "    for holidays with dates that are in the current or next year."
Write-Host ""
Write-Host "Prerequisites:"
Write-Host " - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)"
Write-Host " - Connectivity to data.gov.au"
Write-Host " - Sufficient permissions to manage Teams Voice call flow schedules."
Write-Host ""
Write-Host "Assumptions"
Write-Host " - This script assumes existing holiday schedules follow a naming convention"
Write-Host "   that includes the state code (e.g., 'NSW Boxing Day' or 'NSW Public Holiday')."
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
                     ([datetime]::ParseExact($_.date, 'yyyyMMdd', $null)).Year -ge $CurrentYear -and
                     ([datetime]::ParseExact($_.date, 'yyyyMMdd', $null)).Year -le ($CurrentYear + 1)
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

function Update-ExistingHolidays {
    $TeamsVoiceHolidays = Get-CsOnlineSchedule | Where-Object { $_.FixedSchedule }
    $results = Get-PublicHolidays
    $PublicHolidays = $results[0]
    $JurisdictionCode = $results[1]

    Write-Host "`nCurrent Fixed Holiday Schedules in Microsoft Teams Voice:" -ForegroundColor Yellow
    if ($TeamsVoiceHolidays.Count -gt 0) {
        foreach ($Holiday in $TeamsVoiceHolidays) {
            Write-Host " - $($Holiday.Name)"
            foreach ($DateRange in $Holiday.FixedSchedule.DateTimeRanges) {
                $StartDate = [datetime]$DateRange.Start
                $EndDate   = [datetime]$DateRange.End
                Write-Host "    - Start: $($StartDate.ToString("yyyy-MM-dd HH:mm")), End: $($EndDate.ToString("yyyy-MM-dd HH:mm"))"
            }
        }
    } else {
        Write-Host "No fixed holiday schedules found in Microsoft Teams Voice."
        return
    }

    if ($PublicHolidays) {
        $updatedCount = 0
        Write-Host "`nAttempting to update existing schedules based on Public Holidays for states: $($JurisdictionCode)..." -ForegroundColor Green

        foreach ($holiday in $PublicHolidays) {
            $holidayDate = Get-Date $holiday.StartDate
            $stateCode = $holiday.State
            $holidayName = $holiday.Name
            $combinedName = "$stateCode $holidayName"
            $groupName = "$stateCode Public Holiday"

            Write-Host "`nChecking Holiday: '$combinedName' with date: $($holidayDate.ToString("yyyy-MM-dd"))"
            $matchedSchedule = $TeamsVoiceHolidays | Where-Object {$_.Name -eq $combinedName}

            if ($matchedSchedule) {
                $start = Get-Date "$($holiday.StartDate) 00:00"
                $end = Get-Date "$($holiday.EndDate) 23:45"
                $startStr = $start.ToString("yyyy-MM-ddTHH:mm:ss")
                $endStr = $end.ToString("yyyy-MM-ddTHH:mm:ss")
                $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr

                Write-Host "    ✅ Found exact match: '$($matchedSchedule.Name)'.  Updating..."
                Write-Host "    📅  Old Start Date: $($matchedSchedule.FixedSchedule.DateTimeRanges[0].Start)"
                Write-Host "    📅  New Start Date: $($start.ToString("yyyy-MM-dd HH:mm:ss"))"
                Write-Host "    📅  Old End Date: $($matchedSchedule.FixedSchedule.DateTimeRanges[0].End)"
                Write-Host "    📅  New End Date: $($end.ToString("yyyy-MM-dd HH:mm:ss"))"
                #Set-CsOnlineSchedule -Identity $matchedSchedule.Identity -FixedSchedule -DateTimeRanges @($dtr)
                $updatedCount++
            } else {
                $matchedSchedules = $TeamsVoiceHolidays | Where-Object {$_.Name -eq $groupName}
                if ($matchedSchedules) {
                    foreach ($schedule in $matchedSchedules) {
                         $start = Get-Date "$($holiday.StartDate) 00:00"
                        $end = Get-Date "$($holiday.EndDate) 23:45"
                        $startStr = $start.ToString("yyyy-MM-ddTHH:mm:ss")
                        $endStr = $end.ToString("yyyy-MM-ddTHH:mm:ss")
                        $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr
                        Write-Host "    ✅ Found group match: '$($schedule.Name)'.  Updating..."
                        Write-Host "    📅  Old Start Date: $($schedule.FixedSchedule.DateTimeRanges[0].Start)"
                        Write-Host "    📅  New Start Date: $($start.ToString("yyyy-MM-dd HH:mm:ss"))"
                        Write-Host "    📅  Old End Date: $($schedule.FixedSchedule.DateTimeRanges[0].End)"
                        Write-Host "    📅  New End Date: $($end.ToString("yyyy-MM-dd HH:mm:ss"))"
                        #Set-CsOnlineSchedule -Identity $schedule.Identity -FixedSchedule -DateTimeRanges @($dtr)
                        $updatedCount++
                    }
                }
                else{
                    Write-Host "    ❌ No match found for '$combinedName' or '$groupName'."
                }
            }
        }
        Write-Host "`nUpdate process complete.  $updatedCount schedules were considered for updating."
        if ($updatedCount -eq 0) {
            Write-Host "`n❌ No holidays were updated."
        }
    } else {
        Write-Warning "Could not retrieve Public Holiday Data from the source"
    }
}
#endregion

#region Main Execution
Update-ExistingHolidays
#endregion

#region Finalization
Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "Public Holiday Update Script finished."
Write-Host "Remember to remove the 'TEST MODE' comments to apply the changes to your Teams environment."
Write-Host "================================================================================================================================================="
Write-Host ""
#endregion