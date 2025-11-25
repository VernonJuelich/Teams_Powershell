<#
.SYNOPSIS
    WPF GUI for managing Teams holiday schedules with embedded Create/Update logic.

.DESCRIPTION
    - Connects to Microsoft Teams.
    - Displays existing holiday schedules.
    - Creates or updates state-based public holiday schedules from JSON.
    - Shows reminder to link new/updated schedules to Auto Attendants.
    - GenNet branding with clickable logo (opens https://gennet.com.au).
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName System.Windows.Forms

# Resolve script directory robustly
if ($PSScriptRoot) {
    $ScriptDir = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    # Fallback when run interactively in ISE / console
    $ScriptDir = (Get-Location).Path
}

# Global: Teams connection state
$global:IsConnected = $false

#-----------------------------#
# Logging Helper
#-----------------------------#
function Add-LogLine {
    param(
        [string]$Message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] $Message"
    if ($script:LogTextBox -and $script:LogTextBox.Dispatcher.CheckAccess()) {
        $script:LogTextBox.AppendText($logLine + [Environment]::NewLine)
        $script:LogTextBox.ScrollToEnd()
    }
    else {
        Write-Host $logLine
    }
}

#-----------------------------#
# Teams Connection Helpers
#-----------------------------#

function Connect-Teams {
    try {
        if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
            [System.Windows.MessageBox]::Show(
                "MicrosoftTeams module not found. Please install it first (Install-Module MicrosoftTeams).",
                "Module Missing",
                'OK',
                'Error'
            ) | Out-Null
            return
        }

        Add-LogLine "Connecting to Microsoft Teams..."
        Import-Module MicrosoftTeams -ErrorAction Stop

        # Interactive modern auth
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        $global:IsConnected = $true

        if ($script:lblConnectionStatus) {
            try {
                $tenantName = (Get-CsTenant | Select-Object -ExpandProperty DisplayName)
            } catch {
                $tenantName = "Unknown tenant"
            }
            $script:lblConnectionStatus.Content = "Connected to: $tenantName"
            $script:lblConnectionStatus.Foreground = "Green"
        }

        Add-LogLine "Connected to Microsoft Teams successfully."
    }
    catch {
        $global:IsConnected = $false
        if ($script:lblConnectionStatus) {
            $script:lblConnectionStatus.Content = "Not connected"
            $script:lblConnectionStatus.Foreground = "Red"
        }
        Add-LogLine "ERROR connecting to Teams: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Microsoft Teams. Please check your account/permissions.",
            "Connection Error",
            'OK',
            'Error'
        ) | Out-Null
    }
}

function Ensure-TeamsConnection {
    if (-not $global:IsConnected) {
        Add-LogLine "No active Teams session. Prompting user to connect..."
        Connect-Teams
    }
    return $global:IsConnected
}

#-----------------------------#
# Holiday Creation / Update Functions
#-----------------------------#

function Get-HolidayJsonFiles {
    param(
        [string]$Directory
    )
    if (-not (Test-Path $Directory)) {
        throw "Holiday JSON directory '$Directory' not found."
    }

    $jsonFiles = Get-ChildItem -Path $Directory -Filter "*.json" -File | Sort-Object Name
    if (-not $jsonFiles) {
        throw "No JSON files found in '$Directory'."
    }

    return $jsonFiles
}

function Build-HolidayRangesFromJson {
    param(
        [string]$Directory,
        [string]$State
    )

    $jsonFiles = Get-HolidayJsonFiles -Directory $Directory
    $allRanges = @()

    foreach ($file in $jsonFiles) {
        $data = Get-Content $file.FullName | ConvertFrom-Json
        $rawDates = $data | Where-Object { $_.states -contains $State } | Select-Object -ExpandProperty date
        Add-LogLine "Raw dates for $State in $($file.Name): $($rawDates -join ', ')"

        # Filter and convert valid YYYY-MM-DD dates
        $dates = $rawDates | Where-Object { $_ -and ($_ -match '^\d{4}-\d{2}-\d{2}$') } | ForEach-Object { [DateTime]$_ }
        if (-not $dates -or $dates.Count -eq 0) {
            Add-LogLine "No valid holiday dates found for $State in $($file.Name)"
            continue
        }

        # Build DateTimeRanges for Teams
        $dateRanges = foreach ($d in $dates) {
            @{
                Start = $d.ToString("yyyy-MM-ddT00:00:00")
                End   = $d.ToString("yyyy-MM-ddT23:59:59")
            }
        }

        $allRanges += $dateRanges
    }

    return $allRanges
}

function Create-HolidaySchedule {
    param(
        [string]$State,
        [string]$HolidayJsonDir,
        [switch]$WhatIf
    )

    Add-LogLine "Preparing to CREATE holiday schedule for state: $State"

    $ranges = Build-HolidayRangesFromJson -Directory $HolidayJsonDir -State $State
    if (-not $ranges -or $ranges.Count -eq 0) {
        Add-LogLine "No date ranges found for $State. Skipping create."
        return
    }

    # NEW NAMING: "NSW Public Holidays", "VIC Public Holidays", etc.
    $scheduleName = "$State Public Holidays"
    Add-LogLine "Calculated schedule name: $scheduleName"
    Add-LogLine "Number of date ranges to create: $($ranges.Count)"

    if ($WhatIf) {
        Add-LogLine "[WhatIf] Would create schedule '$scheduleName' with $($ranges.Count) ranges."
        return
    }

    try {
        Add-LogLine "Creating schedule '$scheduleName'..."
        New-CsOnlineSchedule -Name $scheduleName -Type Holiday -DateTimeRanges $ranges -ErrorAction Stop | Out-Null
        Add-LogLine "Successfully created schedule '$scheduleName'."
        Add-LogLine "IMPORTANT: Connect this new holiday schedule to the correct Auto Attendant in the Teams admin centre."
    }
    catch {
        Add-LogLine "ERROR creating schedule '$scheduleName' : $($_.Exception.Message)"
    }
}

function Update-HolidaySchedule {
    param(
        [string]$State,
        [string]$HolidayJsonDir,
        [switch]$WhatIf
    )

    Add-LogLine "Preparing to UPDATE holiday schedule for state: $State"

    # NEW NAMING: "NSW Public Holidays", etc.
    $scheduleName = "$State Public Holidays"
    # Try exact match first
    $schedule = Get-CsOnlineSchedule -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "$State Public Holidays" }

    # If no exact match, try startswith (handles trailing spaces, etc.)
    if (-not $schedule) {
        $schedule = Get-CsOnlineSchedule -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "$State Public Holidays*" }
    }

    # If still no match, try contains
    if (-not $schedule) {
        $schedule = Get-CsOnlineSchedule -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*$State*" -and $_.Name -like "*Public Holidays*" }
    }

    if (-not $schedule) {
        Add-LogLine "No existing schedule located for: '$State Public Holidays'. Possible naming mismatch."
        return
    }
    if (-not $schedule) {
        Add-LogLine "No existing schedule named '$scheduleName'. Cannot update, skipping."
        return
    }

    $existingRanges = @()
    if ($schedule.DateTimeRanges) {
        $existingRanges = $schedule.DateTimeRanges | ForEach-Object {
            @{
                Start = ([DateTime]$_.Start).ToString("yyyy-MM-ddT00:00:00")
                End   = ([DateTime]$_.End).ToString("yyyy-MM-ddT23:59:59")
            }
        }
    }

    $holidayRanges = Build-HolidayRangesFromJson -Directory $HolidayJsonDir -State $State

    # Filter to only future dates (from today onwards)
    $today = (Get-Date).Date
    Add-LogLine "Filtering out holidays that have already occurred. Today: $today"

    $holidayRanges = $holidayRanges | Where-Object {
        ([DateTime]$_.Start).Date -ge $today
    }

    if (-not $holidayRanges -or $holidayRanges.Count -eq 0) {
        Add-LogLine "No future holiday ranges found for $State. No update needed."
        return
    }

    $existingRangeStrings = $existingRanges | ForEach-Object {
        "$($_.Start)|$($_.End)"
    }

    $combined = @()
    $combined += $existingRanges

    foreach ($r in $holidayRanges) {
        $key = "$($r.Start)|$($r.End)"
        if ($existingRangeStrings -notcontains $key) {
            $combined += $r
        }
        else {
            Add-LogLine "Skipping duplicate holiday range: $($r.Start) to $($r.End)"
        }
    }

    # De-duplicate by date (one range per calendar day)
    $dateGroups = $combined | Group-Object {
        ([DateTime]$_.Start).Date
    }

    $finalRanges = @()
    foreach ($g in $dateGroups) {
        $d = $g.Group | Select-Object -First 1
        $finalRanges += @{
            Start = ([DateTime]$d.Start).ToString("yyyy-MM-ddT00:00:00")
            End   = ([DateTime]$d.End).ToString("yyyy-MM-ddT23:59:59")
        }
    }

    Add-LogLine "Will update schedule '$scheduleName' with $($finalRanges.Count) total ranges."

    if ($WhatIf) {
        Add-LogLine "[WhatIf] Would update schedule '$scheduleName' with $($finalRanges.Count) ranges (future only)."
        return
    }

    try {
        Add-LogLine "Updating schedule '$scheduleName'..."
        Set-CsOnlineSchedule -Identity $schedule.Id -DateTimeRanges $finalRanges -ErrorAction Stop | Out-Null
        Add-LogLine "Successfully updated schedule '$scheduleName'."
        Add-LogLine "IMPORTANT: Confirm this holiday schedule is still linked to the correct Auto Attendant in the Teams admin centre."
    }
    catch {
        Add-LogLine "ERROR updating schedule '$scheduleName' : $($_.Exception.Message)"
    }
}

#-----------------------------#
# WPF UI (XAML)
#-----------------------------#
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Teams Holiday Schedule Manager"
        WindowStartupLocation="CenterScreen"
        WindowState="Maximized"
        WindowStyle="SingleBorderWindow"
        ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Header -->
            <RowDefinition Height="2*"/>   <!-- Main area: connection/json/mode/states + schedules -->
            <RowDefinition Height="*"/>    <!-- Log -->
            <RowDefinition Height="Auto"/> <!-- Buttons -->
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2*"/>
            <ColumnDefinition Width="3*"/>
        </Grid.ColumnDefinitions>

        <!-- Header with clickable GenNet logo -->
        <Border Grid.Row="0" Grid.ColumnSpan="2"
                Background="#050505"
                Padding="12">
            <DockPanel LastChildFill="False">
                <!-- Logo button -->
                <Button x:Name="btnLogo"
                        DockPanel.Dock="Left"
                        Margin="0,0,20,0"
                        Background="Transparent"
                        BorderThickness="0"
                        Cursor="Hand"
                        ToolTip="Open GenNet website">
                    <Image x:Name="imgLogo"
                           Height="60"
                           Stretch="Uniform"/>
                </Button>

                <!-- Title + Subtitle -->
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="GenNet Teams Holiday Schedule Manager"
                               FontSize="26"
                               FontWeight="Bold"
                               Foreground="White"/>
                    <TextBlock Text="Accelerating innovation in Teams Voice &amp; Contact Centre"
                               Margin="0,4,0,0"
                               FontSize="14"
                               Foreground="#EDF0FF"/>
                    <TextBlock x:Name="txtScriptDir"
                               FontSize="12"
                               Margin="0,6,0,0"
                               TextWrapping="Wrap"
                               Foreground="#EDF0FF"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- LEFT: Connection → Holiday JSON → Mode → States → WhatIf -->
        <ScrollViewer Grid.Row="1" Grid.Column="0" Margin="0,0,10,0" VerticalScrollBarVisibility="Auto">
            <StackPanel>

                <!-- Connection first (status only) -->
                <GroupBox Header="Connection" FontSize="14" Margin="0,0,0,10">
                    <StackPanel Margin="8">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Text="Teams connection status:" VerticalAlignment="Center"/>
                            <Label x:Name="lblConnectionStatus"
                                   Content="Not connected"
                                   Margin="5,0,0,0"
                                   VerticalAlignment="Center"
                                   Foreground="Red"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>

                <!-- Holiday JSON directory -->
                <GroupBox Header="Holiday JSON Directory" FontSize="14" Margin="0,0,0,10">
                    <StackPanel Margin="8">
                        <StackPanel Orientation="Horizontal">
                            <TextBox x:Name="txtHolidayDir" Width="250" Margin="0,0,5,0"/>
                            <Button x:Name="btnBrowseDir" Content="Browse..." Width="80"/>
                        </StackPanel>
                        <TextBlock Text="Directory containing holiday JSON files (e.g. holiday.json files by year)."
                                   Margin="0,5,0,0" TextWrapping="Wrap" Foreground="Gray"/>
                    </StackPanel>
                </GroupBox>

                <!-- Mode -->
                <GroupBox Header="Mode" Margin="0,0,0,10" FontSize="14">
                    <StackPanel Margin="8">
                        <RadioButton x:Name="rbCreate" Content="Create NEW state public holiday schedules"
                                     IsChecked="True" Margin="0,0,0,8" FontSize="13"/>
                        <RadioButton x:Name="rbUpdate" Content="UPDATE existing state public holiday schedules"
                                     FontSize="13"/>
                        <TextBlock Text="Note: After creating new holiday schedules, you must connect them to the appropriate Auto Attendant in the Teams admin centre."
                                   Margin="0,10,0,0"
                                   TextWrapping="Wrap"
                                   Foreground="#0063A5"
                                   FontStyle="Italic"
                                   FontSize="12"/>
                    </StackPanel>
                </GroupBox>

                <!-- States: row with checkboxes + select/clear below -->
                <GroupBox Header="States" Margin="0,0,0,10" FontSize="14">
                    <StackPanel Margin="8">
                        <WrapPanel Margin="0,0,0,5">
                            <CheckBox x:Name="cbNSW" Content="NSW" Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbVIC" Content="VIC" Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbQLD" Content="QLD" Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbSA"  Content="SA"  Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbWA"  Content="WA"  Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbTAS" Content="TAS" Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbACT" Content="ACT" Margin="0,0,10,0"/>
                            <CheckBox x:Name="cbNT"  Content="NT"  Margin="0,0,10,0"/>
                        </WrapPanel>
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnSelectAll" Content="Select All" Width="80" Margin="0,5,5,0"/>
                            <Button x:Name="btnClearAll"  Content="Clear All"  Width="80" Margin="0,5,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>

                <!-- WhatIf -->
                <GroupBox Header="WhatIf / Dry Run" FontSize="14" Margin="0,0,0,10">
                    <StackPanel Margin="8">
                        <CheckBox x:Name="cbWhatIf" Content="WhatIf (simulate only, no changes in Teams)"
                                  Margin="0,0,0,8" FontSize="13"/>
                        <TextBlock Text="When enabled, Run will only log what it WOULD do without changing any Teams schedules."
                                   TextWrapping="Wrap"
                                   FontSize="11"
                                   Foreground="Gray"/>
                    </StackPanel>
                </GroupBox>

            </StackPanel>
        </ScrollViewer>

        <!-- RIGHT: Existing Holiday Schedules -->
        <GroupBox Header="Existing Holiday Schedules" Grid.Row="1" Grid.Column="1" FontSize="14">
            <ListView x:Name="lvSchedules" Margin="5">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="State" DisplayMemberBinding="{Binding State}" Width="80"/>
                        <GridViewColumn Header="Schedule Name" DisplayMemberBinding="{Binding Name}" Width="220"/>
                        <GridViewColumn Header="Ranges (count)" DisplayMemberBinding="{Binding RangeCount}" Width="120"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </GroupBox>

        <!-- Row 2: Log output spanning both columns -->
        <GroupBox Header="Log Output" Grid.Row="2" Grid.ColumnSpan="2" FontSize="14" Margin="0,10,0,0">
            <TextBox x:Name="txtLog"
                     Margin="5"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     TextWrapping="Wrap"
                     IsReadOnly="True"/>
        </GroupBox>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="3" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnRun" Content="Run" Width="150" Height="40" Margin="0,0,15,0"
                    Background="#0078D7" Foreground="White" FontWeight="Bold" FontSize="15"/>
            <Button x:Name="btnClose" Content="Close" Width="120" Height="40" FontSize="15"/>
        </StackPanel>
    </Grid>
</Window>
"@

#-----------------------------#
# Build Window from XAML
#-----------------------------#
$reader  = New-Object System.Xml.XmlNodeReader $xaml
$window  = [Windows.Markup.XamlReader]::Load($reader)

# Resolve controls into script: scope
$script:txtScriptDir        = $window.FindName("txtScriptDir")
$script:lblConnectionStatus = $window.FindName("lblConnectionStatus")
$script:txtHolidayDir       = $window.FindName("txtHolidayDir")
$script:btnBrowseDir        = $window.FindName("btnBrowseDir")
$script:rbCreate            = $window.FindName("rbCreate")
$script:rbUpdate            = $window.FindName("rbUpdate")
$script:cbNSW               = $window.FindName("cbNSW")
$script:cbVIC               = $window.FindName("cbVIC")
$script:cbQLD               = $window.FindName("cbQLD")
$script:cbSA                = $window.FindName("cbSA")
$script:cbWA                = $window.FindName("cbWA")
$script:cbTAS               = $window.FindName("cbTAS")
$script:cbACT               = $window.FindName("cbACT")
$script:cbNT                = $window.FindName("cbNT")
$script:btnSelectAll        = $window.FindName("btnSelectAll")
$script:btnClearAll         = $window.FindName("btnClearAll")
$script:cbWhatIf            = $window.FindName("cbWhatIf")
$script:lvSchedules         = $window.FindName("lvSchedules")
$script:LogTextBox          = $window.FindName("txtLog")
$script:btnRun              = $window.FindName("btnRun")
$script:btnClose            = $window.FindName("btnClose")
$script:btnLogo             = $window.FindName("btnLogo")
$script:imgLogo             = $window.FindName("imgLogo")

if ($script:txtScriptDir) {
    $script:txtScriptDir.Text = "Script directory: $ScriptDir"
}

#-----------------------------#
# Load logo from web + click action (no log entry)
#-----------------------------#
if ($script:imgLogo) {
    try {
        $logoUrl = [Uri]"https://gennet.com.au/wp-content/uploads/2019/08/main-logo.png"

        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.UriSource = $logoUrl
        # No Freeze() to avoid Freezable error when loading from web
        $bitmap.EndInit()

        $script:imgLogo.Source = $bitmap
    }
    catch {
        Add-LogLine "Failed to load GenNet logo from the web: $($_.Exception.Message)"
    }
}

if ($script:btnLogo) {
    $script:btnLogo.Add_Click({
        try {
            Start-Process "https://gennet.com.au"
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Unable to open https://gennet.com.au",
                "Browser Error",
                'OK',
                'Warning'
            ) | Out-Null
        }
    })
}

#-----------------------------#
# Populate Existing Holiday Schedules (using "*Public Holidays*")
#-----------------------------#
function Refresh-HolidaySchedulesView {
    if (-not (Ensure-TeamsConnection)) { return }

    try {
        # Use name pattern like "*Public Holidays*" and handle FixedSchedule / DateTimeRanges
        $holidaySchedules = Get-CsOnlineSchedule -ErrorAction Stop |
            Where-Object { $_.Name -like "*Public Holidays*" }

        $holidayData = foreach ($schedule in $holidaySchedules) {

            # Choose correct range source
            $ranges = if ($schedule.FixedSchedule -and $schedule.FixedSchedule.DateTimeRanges) {
                $schedule.FixedSchedule.DateTimeRanges
            }
            else {
                $schedule.DateTimeRanges
            }

            $dates = @()
            if ($ranges) {
                $dates = $ranges | ForEach-Object {
                    ([DateTime]$_.Start).ToShortDateString()
                } | Sort-Object -Unique
            }

            [PSCustomObject]@{
                Name       = $schedule.Name
                State      = ($schedule.Name -split ' ')[0]
                RangeCount = if ($ranges) { $ranges.Count } else { 0 }
                Dates      = ($dates -join ', ')
            }
        }

        $script:lvSchedules.ItemsSource = $holidayData
        Add-LogLine "Loaded $($holidayData.Count) holiday schedules into the list."
    }
    catch {
        Add-LogLine "Failed to retrieve holiday schedules: $($_.Exception.Message)"
    }
}

#-----------------------------#
# Event Handlers
#-----------------------------#

$script:btnBrowseDir.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $ScriptDir
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:txtHolidayDir.Text = $dialog.SelectedPath
    }
})

$script:btnSelectAll.Add_Click({
    $script:cbNSW.IsChecked = $true
    $script:cbVIC.IsChecked = $true
    $script:cbQLD.IsChecked = $true
    $script:cbSA.IsChecked  = $true
    $script:cbWA.IsChecked  = $true
    $script:cbTAS.IsChecked = $true
    $script:cbACT.IsChecked = $true
    $script:cbNT.IsChecked  = $true
})

$script:btnClearAll.Add_Click({
    $script:cbNSW.IsChecked = $false
    $script:cbVIC.IsChecked = $false
    $script:cbQLD.IsChecked = $false
    $script:cbSA.IsChecked  = $false
    $script:cbWA.IsChecked  = $false
    $script:cbTAS.IsChecked = $false
    $script:cbACT.IsChecked = $false
    $script:cbNT.IsChecked  = $false
})

$script:btnRun.Add_Click({
    if (-not (Ensure-TeamsConnection)) { return }

    $holidayDir = $script:txtHolidayDir.Text.Trim()
    if (-not $holidayDir) {
        [System.Windows.MessageBox]::Show(
            "Please select the Holiday JSON directory before running.",
            "Missing Directory",
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    if (-not (Test-Path $holidayDir)) {
        [System.Windows.MessageBox]::Show(
            "The specified directory does not exist: $holidayDir",
            "Invalid Directory",
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    $states = @()
    if ($script:cbNSW.IsChecked) { $states += "NSW" }
    if ($script:cbVIC.IsChecked) { $states += "VIC" }
    if ($script:cbQLD.IsChecked) { $states += "QLD" }
    if ($script:cbSA.IsChecked)  { $states += "SA"  }
    if ($script:cbWA.IsChecked)  { $states += "WA"  }
    if ($script:cbTAS.IsChecked) { $states += "TAS" }
    if ($script:cbACT.IsChecked) { $states += "ACT" }
    if ($script:cbNT.IsChecked)  { $states += "NT"  }

    if (-not $states -or $states.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one state.",
            "No States Selected",
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    $mode   = if ($script:rbCreate.IsChecked) { "Create" } else { "Update" }
    $whatIf = [bool]$script:cbWhatIf.IsChecked

    Add-LogLine "-----------------------------"
    Add-LogLine "Run initiated. Mode: $mode; States: $($states -join ', '); WhatIf: $whatIf"
    Add-LogLine "Holiday JSON directory: $holidayDir"

    foreach ($state in $states) {
        if ($mode -eq "Create") {
            Create-HolidaySchedule -State $state -HolidayJsonDir $holidayDir -WhatIf:$whatIf
        }
        else {
            Update-HolidaySchedule -State $state -HolidayJsonDir $holidayDir -WhatIf:$whatIf
        }
    }

    if ($whatIf) {
        Add-LogLine "WhatIf mode: No changes were made in Microsoft Teams. Review the log to see what would have happened."
    }
    else {
        Add-LogLine "Run completed. Schedules were updated in Microsoft Teams."
    }

    Refresh-HolidaySchedulesView
})

$script:btnClose.Add_Click({
    $window.Close()
})

# On load: auto-connect and load existing holidays
$window.Add_SourceInitialized({
    Add-LogLine "Window initialised. Script directory: $ScriptDir"
    if ($script:txtHolidayDir -and -not $script:txtHolidayDir.Text) {
        $script:txtHolidayDir.Text = $ScriptDir
    }

    # Force connect to Teams when the app opens
    Connect-Teams

    # If connected, load existing holiday schedules immediately
    if ($global:IsConnected) {
        Refresh-HolidaySchedulesView
    }
})

# Show the window
[void]$window.ShowDialog()
