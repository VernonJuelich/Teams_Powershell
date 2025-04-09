Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "This script will:"
Write-Host " 1. Manage Public Holiday schedules in Microsoft Teams Voice."
Write-Host " 2. Retrieve Public Holiday data from data.gov.au for specified Australian states."
Write-Host " 3. Allow you to create new or update existing holiday schedules in Teams."
Write-Host "    - Updates will only apply to holidays with dates that are in the current or next year."
Write-Host ""
Write-Host "Prerequisites:"
Write-Host " - Microsoft Teams PowerShell Module installed (`Install-Module MicrosoftTeams`)"
Write-Host " - Connectivity to data.gov.au"
Write-Host " - Sufficient permissions to manage Teams Voice call flow schedules."
Write-Host ""
Write-Host "Assumptions"
Write-Host " - This scipt has been used to create the holiday. As the update section looks for Specific naming of Holidays"
Write-Host ""
Write-Host "Notes"
Write-Host " - To Skip Specfic holidays got to line 83 and copy that command filtering on the specfic holiday name"
Write-Host ""
Write-Host "================================================================================================================================================="


# Connect to Microsoft Teams
try {
    Connect-MicrosoftTeams | Out-Null
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Ensure the module is installed and you have the correct permissions."
    Exit  # Terminate the script if connection fails
}

# Prompt for import type
$importType = Read-Host "Do you want to perform a 'new' import of holidays or 'update' existing schedules? (Enter 'new' or 'update')"
while ($importType -notin ("new", "update")) {
    Write-Warning "Invalid input. Please enter 'Y' or 'N'."
    $importType = Read-Host "Do you want to perform a 'new' import of holidays or 'update' existing schedules? (Enter 'new' or 'update')"
}

# ------------------ FUNCTIONS ------------------

function Get-PublicHolidays {
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("sa", "qld", "act", "nsw", "wa", "nt", "vic", "tas")]
        [string[]] $Jurisdictions  # Changed to string array
    )

    if (-not $Jurisdictions) {
        $JurisdictionInput = Read-Host "Enter the Australian state codes (comma-separated, e.g., sa,qld,act):"
        $Jurisdictions = $JurisdictionInput.Split(',') | ForEach-Object { $_.Trim() } # Split and trim
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
    $allJurisdictionCodes = @() # Array to hold all jurisdiction codes.

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
            # Continue to the next jurisdiction, even if one fails
            continue
        }

        foreach ($Holiday in $Results) {
            if ($Holiday.'Holiday Name' -eq "Bank Holiday") { continue } #Filter Specfic Holidays
            #if ($Holiday.'Holiday Name' -eq "Bank Holiday") { continue }


            $HolidayDate = ([datetime]::ParseExact($Holiday.date, 'yyyyMMdd', $null)).ToShortDateString()
            $HolidayEndDate = ([datetime]::ParseExact($Holiday.date, 'yyyyMMdd', $null)).AddDays(1).ToShortDateString()

            $Holidays += [PSCustomObject]@{
                StartDate = $HolidayDate
                EndDate   = $HolidayEndDate
                Name      = $Holiday.'Holiday Name'
                State     = $Jurisdiction.ToUpper() # Add State information in Upper case
            }
        }
        $allJurisdictionCodes += $Jurisdiction.ToUpper() # Add to the array
    }
    return $Holidays, ($allJurisdictionCodes -join ", ") # Return all codes as comma-separated string
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
    $JurisdictionCode = $results[1] # This is now a comma-separated string

    if ($PublicHolidays) {
        if ($sameAction -eq "N") {
            # Create individual schedules
            foreach ($Holiday in $PublicHolidays) {
                $scheduleName = "$($Holiday.State) $($Holiday.Name)"
                $startDateTime = Get-Date "$($Holiday.StartDate) $startTime"
                $endDateTime = Get-Date "$($Holiday.EndDate) $endTime"
                $startStr = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                $endStr = $endDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
                $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr

                $existing = $TeamsVoiceHolidays | Where-Object { $_.Name -eq $scheduleName }
                if ($existing) {
                    $update = Read-Host "Holiday schedule '$($scheduleName)' already exists. Update it? (Y/N)"
                    while ($update -notin ("Y", "N")) {
                        Write-Warning "Invalid input. Please enter 'Y' or 'N'."
                        $update = Read-Host "Holiday schedule '$($scheduleName)' already exists. Update it? (Y/N)"
                    }
                    if ($update -eq "Y") {
                        Write-Host "TEST MODE: Would update schedule '$($scheduleName)' with Start: $startStr, End: $endStr"
                        # Set-CsOnlineSchedule -Identity $existing.Identity -FixedSchedule -DateTimeRanges @($dtr)
                    } else {
                        Write-Host "Skipping update for '$($scheduleName)'."
                    }
                } else {
                    Write-Host "TEST MODE: Would create new schedule '$($scheduleName)' with Start: $startStr, End: $endStr"
                    # New-CsOnlineSchedule -Name $scheduleName -FixedSchedule -DateTimeRanges @($dtr)
                }
            }
        } else {
            # Create a dictionary to group holidays by state
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
                    $update = Read-Host "Holiday schedule '$($scheduleName)' already exists. Update it? (Y/N)"
                    while ($update -notin ("Y", "N")) {
                        Write-Warning "Invalid input. Please enter 'Y' or 'N'."
                        $update = Read-Host "Holiday schedule '$($scheduleName)' already exists. Update it? (Y/N)"
                    }
                    if ($update -eq "Y") {
                        Write-Host "TEST MODE: Would update schedule '$($scheduleName)' with  date ranges."
                        # Set-CsOnlineSchedule -Identity $existing.Identity -FixedSchedule -DateTimeRanges $dateTimeRanges
                    } else {
                        Write-Host "Skipping update for '$($scheduleName)'."
                    }
                } else {
                    Write-Host "TEST MODE: Would create new schedule '$($scheduleName)' with  date ranges."
                        # New-CsOnlineSchedule -Name $scheduleName -FixedSchedule -DateTimeRanges $dateTimeRanges
                }
            }
        }
    } else {
        Write-Host "No public holidays retrieved for the specified state."
    }
}

# ------------------ UPDATE MODE ------------------
if ($importType -ceq "update") {
    $TeamsVoiceHolidays = Get-CsOnlineSchedule | Where-Object { $_.FixedSchedule }

    Write-Host "`nCurrent Fixed Holiday Schedules in Microsoft Teams Voice:" -ForegroundColor Yellow
    if ($TeamsVoiceHolidays.Count -gt 0) {
        foreach ($Holiday in $TeamsVoiceHolidays) {
            Write-Host " - $($Holiday.Name)"
            foreach ($DateRange in $Holiday.FixedSchedule.DateTimeRanges) {
                $StartDate = [datetime]$DateRange.Start
                $EndDate   = [datetime]$DateRange.End
                Write-Host "   - Start: $($StartDate.ToString("yyyy-MM-dd HH:mm")), End: $($EndDate.ToString("yyyy-MM-dd HH:mm"))"
            }
        }
    } else {
        Write-Host "No fixed holiday schedules found in Microsoft Teams Voice."
    }

    $results = Get-PublicHolidays
    $PublicHolidays = $results[0]
    $JurisdictionCode = $results[1] # Comma-separated string

    if ($PublicHolidays) {
        $updatedCount = 0
        $currentDate = Get-Date
        $currentYear = $currentDate.Year
        $nextYear = $currentYear + 1  # Calculate the next year

        # Write-Host "`nCurrent Year: $currentYear"  # Remove
        # Write-Host "`nNext Year: $nextYear"      # Remove
        Write-Host "`nAttempting to update existing schedules based on Public Holidays for states: $($JurisdictionCode)..." -ForegroundColor Green

        foreach ($holiday in $PublicHolidays) {
            $holidayDate = Get-Date $holiday.StartDate
            $stateCode = $holiday.State
            $holidayName = $holiday.Name
            $combinedName = "$stateCode $holidayName"
            $groupName = "$stateCode Public Holiday"

            # Check if the holiday is in the current or next year.  CHANGED CONDITION
            if ($holidayDate.Year -eq $currentYear -or $holidayDate.Year -eq $nextYear) { # Changed the condition here
                Write-Host "`nChecking Holiday: '$combinedName' with date: $($holidayDate.ToString("yyyy-MM-dd"))"
                # 1. Try to match the combined name (e.g., "NSW Boxing Day")
                $matchedSchedule = $TeamsVoiceHolidays | Where-Object {$_.Name -eq $combinedName}

                if ($matchedSchedule) {
                    $start = Get-Date "$($holiday.StartDate) 00:00"
                    $end = Get-Date "$($holiday.EndDate) 23:45"
                    $startStr = $start.ToString("yyyy-MM-ddTHH:mm:ss") # Ensure correct format
                    $endStr = $end.ToString("yyyy-MM-ddTHH:mm:ss")   # Ensure correct format
                    $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr

                    Write-Host "   ✅ Found exact match: '$($matchedSchedule.Name)'.  Updating..."
                    Write-Host "   📅  Old Start Date: $($matchedSchedule.FixedSchedule.DateTimeRanges[0].Start)"
                    Write-Host "   📅  New Start Date: $($start.ToString("yyyy-MM-dd HH:mm:ss"))"
                    Write-Host "   📅  Old End Date: $($matchedSchedule.FixedSchedule.DateTimeRanges[0].End)"
                    Write-Host "   📅  New End Date: $($end.ToString("yyyy-MM-dd HH:mm:ss"))"
                    Write-Host "TEST MODE: Would update '$($matchedSchedule.Name)' to Start: $($dtr.Start), End: $($dtr.End)"
                    #Set-CsOnlineSchedule -Identity $matchedSchedule.Identity -FixedSchedule -DateTimeRanges @($dtr)
                    $updatedCount++
                } else {
                    # 2. If no exact match, try to match the group name (e.g., "NSW Public Holiday")
                    $matchedSchedules = $TeamsVoiceHolidays | Where-Object {$_.Name -eq $groupName} # Change to find ALL matching
                    if ($matchedSchedules) {
                        # Iterate through all schedules with the group name
                        foreach ($schedule in $matchedSchedules) {
                            $start = Get-Date "$($holiday.StartDate) 00:00"
                            $end = Get-Date "$($holiday.EndDate) 23:45"
                            $startStr = $start.ToString("yyyy-MM-ddTHH:mm:ss") # Ensure correct format
                            $endStr = $end.ToString("yyyy-MM-ddTHH:mm:ss")   # Ensure correct format
                            $dtr = New-CsOnlineDateTimeRange -Start $startStr -End $endStr
                            Write-Host "   ✅ Found group match: '$($schedule.Name)'.  Updating..."
                            Write-Host "   📅  Old Start Date: $($schedule.FixedSchedule.DateTimeRanges[0].Start)"
                            Write-Host "   📅  New Start Date: $($start.ToString("yyyy-MM-dd HH:mm:ss"))"
                            Write-Host "   📅  Old End Date: $($schedule.FixedSchedule.DateTimeRanges[0].End)"
                            Write-Host "   📅  New End Date: $($end.ToString("yyyy-MM-dd HH:mm:ss"))"
                            Write-Host "TEST MODE: Would update '$($schedule.Name)' to Start: $($dtr.Start), End: $($dtr.End)"
                            #Set-CsOnlineSchedule -Identity $schedule.Identity -FixedSchedule -DateTimeRanges @($dtr)
                            $updatedCount++
                        }
                    }
                    else{
                         Write-Host "   ❌ No match found for '$combinedName' or '$groupName'."
                    }
                }
            }
            else{
                Write-Host "   ⏩ Skipping '$combinedName' as it is not in the current or next year." # Changed message
            }
        }
        Write-Host "`nUpdate process complete.  $updatedCount schedules were considered for updating."
        if ($updatedCount -eq 0) {
            Write-Host "`n❌ No holidays were updated."
            $doImport = Read-Host "Would you like to import the future public holidays instead? (Y/N)"
            while ($doImport -notin ("Y", "N")) {
                Write-Warning "Invalid input. Please enter 'Y' or 'N'."
                $doImport = Read-Host "Would you like to import the public holidays instead? (Y/N)"
            }
            if ($doImport -eq "Y") {
                Import-NewHolidays
            } else {
                Write-Host "No new holidays were imported."
            }
        }
    } else {
        Write-Warning "Could not retrieve Public Holiday Data from the source"
    }
}

# ------------------ NEW MODE ------------------
if ($importType -ceq "new") {
    Import-NewHolidays
}

Write-Host ""
Write-Host "================================================================================================================================================="
Write-Host "Script execution finished."
Write-Host "Remember to remove the 'TEST MODE' comments to apply the changes to your Teams environment."
Write-Host "================================================================================================================================================="
Write-Host ""
