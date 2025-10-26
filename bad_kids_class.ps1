using namespace System.Windows.Forms
using namespace System.Drawing
using namespace System.Speech.Synthesis
using namespace System.Globalization

<#
SchedulerScreenSaver class

Purpose
This program helps parents manage their children's daily routines by providing
automated, fair, and trustworthy scheduling of tasks and chores. Key features:

- Automated Task Management
  • Provides clear, spoken reminders for daily activities like bedtime prep,
    homework time, and chores
  • Speaks announcements so kids hear exactly who needs to do what and when
  • Shows visual timers and countdowns for time-limited activities
  • Adapts schedules for school days vs weekends automatically

- Fair Task Rotation
  • Manages shared resources like bathroom/shower time fairly by rotating tasks
  • Remembers whose turn it is across program restarts
  • Provides enough time between rotated tasks (e.g., 20 min shower windows)
  • Eliminates arguments about "who goes first" by tracking rotation state

- Kid-Friendly Interface
  • Always-visible display shows current time and upcoming tasks
  • Each child has their own task panel showing their specific schedule
  • Large, clear alerts appear when it's time for an activity
  • Spoken reminders ensure kids don't miss notifications
  • Moves around screen to prevent burn-in and maintain visibility

- Smart Scheduling
  • Handles different schedules for school days vs weekends
  • Supports both individual tasks and shared/rotating responsibilities
  • Prevents task overlap by spacing activities appropriately
  • Logs task completion to track adherence to schedules

The program acts as an impartial "referee" for scheduling, eliminating common
sources of argument like "it's not fair" or "it's not my turn" by maintaining
consistent rotation state and providing clear, timely notifications that everyone
can trust.

High-level behavior
- Loads schedule data from `schedule.csv` and optional persistent rotation state
  from `task_state.csv` (TaskState). Today's tasks are computed at load/reset.
- A Timer (1s) updates the clock, checks scheduled tasks, shows alerts, and
  moves the main panel around the screen.
- When a scheduled task falls within a short trigger window it is logged,
  optionally spoken via the SpeechSynthesizer, and an on-screen alert is shown.

Key fields
- $Schedule (array): Raw tasks loaded from CSV. Each task has Time, Name,
  DaysOfWeek, and Action.
- $TodaysTasks (hashtable): Computed tasks for the current day. Each entry holds
  the original task object, its scheduled DateTime, and flags (Called, Completed).
- $TaskState (hashtable): Persistent rotation state for multi-person tasks.
  Keys are rotation identifiers (Time|Days|Action) and values are the last
  person who was assigned; used to determine the next assignee.
- $PersonTaskPanels (hashtable): UI Panel controls keyed by person name where
  each person's tasks for today are displayed.
- $MainPanel, labels, and other UI controls for displaying information.

Important methods (brief)
- InitializeComponents(): Builds the Form and main UI, sets up labels and
  the main panel used to display data.
- InitializeTimer(): Starts a 1-second timer that calls OnTimerTick.
- LoadSchedule(): Reads `schedule.csv` into $Schedule and calls LoadTodaysTasks.
- LoadTodaysTasks(): Scans $Schedule and builds $TodaysTasks for the current day
  (parses times and sets DateTime values used for comparisons).
- LoadTaskState()/SaveTaskState(): Read and persist rotation state to
  `task_state.csv` so multi-person rotations continue across runs.
- CheckScheduledTasks(): Iterates $TodaysTasks to find tasks within a trigger
  window (±60s), applies rotation logic to pick the assigned person, logs and
  triggers alerts/speech, marks tasks as Called/Completed.
- UpdateNextTaskDisplay(): Computes and updates the Next task label. It now
  deterministically selects the next upcoming task or shows "No more tasks
  scheduled for today" when appropriate.
- UpdatePersonTaskDisplays(): Builds/updates per-person panels inside
  $MainPanel, showing only the tasks assigned to that person (rotation-aware).
- MoveMainPanel(): Randomly moves $MainPanel, ensuring it stays fully on screen.
- CleanupAndExit(): Gracefully stops timers, disposes speech and form resources.

Data shapes and conventions
- Schedule CSV columns: Time (HH:mm), Name (single or colon-separated list),
  DaysOfWeek (single day, comma list, or range such as Monday-Friday), Action (text).
- Rotation key used for TaskState: "$($task.Time)|$($task.DaysOfWeek)|$($task.Action)"
- Task date-times are stored as DateTime objects in $TodaysTasks to allow
  reliable comparisons against [DateTime]::Now.

Notes and tips
- Name splitting trims whitespace before comparisons to avoid accidental
  duplicate matching (e.g., "Kevin: Chris" vs "Kevin:Chris").
- Person panels are added as children of $MainPanel so they move together when
  $MainPanel is repositioned.
- The trigger window in CheckScheduledTasks is intentionally small (±60s)
  to avoid duplicate firings; tasks are marked Called/Completed after firing.
- For debugging, set the constructor flag to $true to run in a windowed,
  non-topmost mode and see console output.

#>
class SchedulerScreenSaver {
    [Form]$Form
    [Timer]$Timer
    [SpeechSynthesizer]$SpeechSynth
    [Panel]$MainPanel
    [Label]$DateLabel
    [Label]$TimeLabel
    [Label]$LastTaskLabel
    [Label]$NextTaskLabel
    [Label]$ExitLabel
    [Label]$AlertLabel
    [bool]$DebugMode = $false  # Debug mode flag
    [array]$Schedule
    [hashtable]$TaskState
    [hashtable]$PersonTaskPanels
    [hashtable]$TodaysTasks  # Track tasks for today and their completion status
    [DateTime]$LastScheduleLoad  # Track when we last loaded the schedule
    [string]$ScheduleFile = "schedule.csv"
    [string]$StateFile = "task_state.csv"
    [string]$LogFile = "TaskLog.csv"
    [Point]$MainPanelPosition
    [Size]$MainPanelSize
    [int]$MoveDirectionX = 1
    [int]$MoveDirectionY = 1
    [bool]$IsShowingAlert = $false
    [DateTime]$AlertEndTime

    SchedulerScreenSaver([bool]$debug = $false) {
        Write-Host "Creating form (Debug Mode: $debug)"
        $this.DebugMode = $false

        $this.InitializeComponents()
        $this.LoadSchedule()
        $this.LoadTaskState()
        $this.InitializeTimer()
        Write-Host "Showing the Form"
        $this.Form.ShowDialog()
    }

    [void]InitializeComponents() {
        # Form setup
        Write-Host "Initializing form (Debug Mode: $($this.DebugMode))"
        $this.Form = [Form]::new()
        $this.Form.Text = "Scheduler Screen Saver"
        
        if ($this.DebugMode) {
            # Debug mode - normal window
            $this.Form.WindowState = [FormWindowState]::Normal
            $this.Form.FormBorderStyle = [FormBorderStyle]::Sizable
            $this.Form.TopMost = $false
            $this.Form.Size = [System.Drawing.Size]::new(1024, 768)
            $this.Form.StartPosition = [FormStartPosition]::CenterScreen
        }
        else {
            # Full screen mode for non-debug
            $this.Form.WindowState = [FormWindowState]::Maximized
            $this.Form.FormBorderStyle = [FormBorderStyle]::None
            $this.Form.TopMost = $true
        }
        $this.Form.BackColor = [Color]::Black
        $this.Form.KeyPreview = $true
        $this.Form.tag = $this
        $this.Form.Add_KeyDown({ 
                param($s, $e) 
                if ($e.KeyCode -eq [Keys]::Escape) { 
                    $s.tag.CleanupAndExit()
                    $s.Close()
                } })
        
        # Add mouse click handler
        $this.Form.Add_MouseClick({
                param($s, $e)
                $s.tag.CleanupAndExit()
                $s.Close()
            })

        # Speech synthesizer
        $this.SpeechSynth = [SpeechSynthesizer]::new()

        # Main panel for static info
        $this.MainPanel = [Panel]::new()
        $this.MainPanel.BackColor = [Color]::Transparent
        $this.MainPanel.AutoSize = $true
        $this.MainPanel.AutoSizeMode = [AutoSizeMode]::GrowAndShrink
        $this.MainPanelSize = [Size]::new(800, 200)  # Initial size, will grow as needed
        $this.MainPanel.MinimumSize = $this.MainPanelSize
        $this.MainPanelPosition = [Point]::new(100, 50)
        $this.MainPanel.Location = $this.MainPanelPosition

        # Labels
        $labelFont = [Font]::new("Arial", 24, [FontStyle]::Bold)
        $smallFont = [Font]::new("Arial", 18, [FontStyle]::Regular)
        $alertFont = [Font]::new("Arial", 48, [FontStyle]::Bold)

        $this.DateLabel = [Label]::new()
        $this.DateLabel.Font = $labelFont
        $this.DateLabel.ForeColor = [Color]::White
        $this.DateLabel.AutoSize = $true
        $this.DateLabel.Location = [Point]::new(0, 0)

        $this.TimeLabel = [Label]::new()
        $this.TimeLabel.Font = $labelFont
        $this.TimeLabel.ForeColor = [Color]::White
        $this.TimeLabel.AutoSize = $true
        $this.TimeLabel.Location = [Point]::new(0, 50)

        $this.LastTaskLabel = [Label]::new()
        $this.LastTaskLabel.Font = $smallFont
        $this.LastTaskLabel.ForeColor = [Color]::LightGreen
        $this.LastTaskLabel.AutoSize = $true
        $this.LastTaskLabel.Location = [Point]::new(0, 110)
        $this.LastTaskLabel.Text = "Last: No Previous Tasks"

        $this.NextTaskLabel = [Label]::new()
        $this.NextTaskLabel.Font = $smallFont
        $this.NextTaskLabel.ForeColor = [Color]::LightBlue
        $this.NextTaskLabel.AutoSize = $true
        $this.NextTaskLabel.Location = [Point]::new(0, 150)

        $this.ExitLabel = [Label]::new()
        $this.ExitLabel.Text = "Press Esc to exit"
        $this.ExitLabel.Font = $smallFont
        $this.ExitLabel.ForeColor = [Color]::Gray
        $this.ExitLabel.AutoSize = $true
        $this.ExitLabel.Location = [Point]::new(0, 200)

        $this.AlertLabel = [Label]::new()
        $this.AlertLabel.Font = $alertFont
        $this.AlertLabel.ForeColor = [Color]::Yellow
        $this.AlertLabel.AutoSize = $true
        $this.AlertLabel.Visible = $false
        $this.AlertLabel.Location = [Point]::new(100, 300)

        # Add static controls
        $this.MainPanel.Controls.Add($this.DateLabel)
        $this.MainPanel.Controls.Add($this.TimeLabel)
        $this.MainPanel.Controls.Add($this.LastTaskLabel)
        $this.MainPanel.Controls.Add($this.NextTaskLabel)
        $this.MainPanel.Controls.Add($this.ExitLabel)
        $this.Form.Controls.Add($this.MainPanel)
        $this.Form.Controls.Add($this.AlertLabel)

        # Initialize person task panels hashtable
        $this.PersonTaskPanels = @{}
    }

    [void]InitializeTimer() {
        Write-Host "Initializing timer"
        $this.Timer = [Timer]::new()
        $this.Timer.Interval = 5000  # Check every second
        $this.Timer.tag = $this

        $this.Timer.Add_Tick({ 
                param($sender, $e)
                $sender.tag.OnTimerTick($sender.tag) 
            })
        $this.Timer.Start()
    }

    [void]OnTimerTick([SchedulerScreenSaver] $inthis) {
        Write-Host "Timer tick"
        try {
            $now = [DateTime]::Now
            
            # Check if we need to reload schedule (at midnight or if not loaded today)
            if ($this.LastScheduleLoad.Date -ne $now.Date) {
                Write-Host "New day detected - reloading schedule"
                $this.LoadSchedule()
            }

            # Update time display
            $this.DateLabel.Text = $now.ToString("dddd, MMMM dd, yyyy")
            $this.TimeLabel.Text = $now.ToString("h:mm:ss tt")  # 'h' for 12-hour with AM/PM

            # Handle alert timeout
            if ($this.IsShowingAlert -and $now -ge $this.AlertEndTime) {
                $this.AlertLabel.Visible = $false
                $this.IsShowingAlert = $false
                $this.MainPanel.Visible = $true
                # Show person panels again
                foreach ($panel in $this.PersonTaskPanels.Values) {
                    $panel.Visible = $true
                }
            }

            # Check for scheduled tasks
            $this.CheckScheduledTasks()

            # Update person task displays
            $this.UpdatePersonTaskDisplays()

            # Move main panel randomly every 10 seconds
            $this.MoveMainPanel()
        }
        catch {
            Write-Host "Error in timer tick: $($_.Exception.Message)"
        }
    }

    [void]MoveMainPanel() {
        $screen = [Screen]::PrimaryScreen.Bounds
        $actualPanelHeight = $this.MainPanel.Height
        $actualPanelWidth = $this.MainPanel.Width

        # Use smaller movement increments for larger panels
        $xMove = [Math]::Max(30, [Math]::Min(150, 800 / ($actualPanelHeight / 200)))
        $yMove = [Math]::Max(20, [Math]::Min(100, 600 / ($actualPanelHeight / 200)))

        $newX = $this.MainPanelPosition.X + ($this.MoveDirectionX * (Get-Random -Minimum ($xMove / 2) -Maximum $xMove))
        $newY = $this.MainPanelPosition.Y + ($this.MoveDirectionY * (Get-Random -Minimum ($yMove / 2) -Maximum $yMove))

        # Bounce off edges, using actual panel size
        if ($newX -le 0 -or $newX + $actualPanelWidth -ge $screen.Width) {
            $this.MoveDirectionX *= -1
            $newX = [Math]::Max(0, [Math]::Min($screen.Width - $actualPanelWidth, $newX))
        }
        if ($newY -le 0 -or $newY + $actualPanelHeight -ge $screen.Height) {
            $this.MoveDirectionY *= -1
            $newY = [Math]::Max(0, [Math]::Min($screen.Height - $actualPanelHeight, $newY))
        }

        # Ensure panel stays fully visible
        $newX = [Math]::Max(0, [Math]::Min($screen.Width - $actualPanelWidth, $newX))
        $newY = [Math]::Max(0, [Math]::Min($screen.Height - $actualPanelHeight, $newY))

        $this.MainPanelPosition = [Point]::new($newX, $newY)
        $this.MainPanel.Location = $this.MainPanelPosition
    }

    [void]LoadSchedule() {
        Write-Host "Loading schedule from $($this.ScheduleFile)"

        if (Test-Path $this.ScheduleFile) {
            $this.Schedule = Import-Csv -Path $this.ScheduleFile
            $this.LastScheduleLoad = [DateTime]::Now
            $this.LoadTodaysTasks()
        }
        else {
            $this.Schedule = @()
            Write-Host "Warning: $this.ScheduleFile not found. Creating sample file."
            $sample = @(
                [PSCustomObject]@{Time = "22:00"; Name = "Frank:Lisa:Tom"; DaysOfWeek = "Sunday-Thursday"; Action = "go to bed" }
                [PSCustomObject]@{Time = "08:00"; Name = "Alice"; DaysOfWeek = "Monday-Friday"; Action = "take vitamins" }
                [PSCustomObject]@{Time = "19:00"; Name = "Frank:Alice"; DaysOfWeek = "Wednesday"; Action = "family dinner" }
            )
            $sample | Export-Csv -Path $this.ScheduleFile -NoTypeInformation
            $this.LoadTodaysTasks()
        }
    }

    [void]LoadTodaysTasks() {
        Write-Host "Loading today's tasks"
        $now = [DateTime]::Now
        $currentDay = $now.DayOfWeek.ToString()
        $this.TodaysTasks = @{}

        foreach ($task in $this.Schedule) {
            if ($this.IsDayInRange($currentDay, $task.DaysOfWeek)) {
                try {
                    $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                    $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, 
                        $taskTime.Hour, $taskTime.Minute, 0)
                    
                    # Create a unique key for the task
                    $taskKey = "$($task.Time)|$($task.Name)|$($task.Action)"
                    
                    # Add to today's tasks with completion status
                    $this.TodaysTasks[$taskKey] = @{
                        Task      = $task
                        DateTime  = $taskDateTime
                        Completed = $false
                        Called    = $false
                    }
                }
                catch {
                    Write-Host "Error parsing task time: $($_.Exception.Message)"
                }
            }
        }
    }
    

    [void]LoadTaskState() {
        Write-Host "Loading task state from $($this.StateFile)"

        $this.TaskState = @{}
        if (Test-Path $this.StateFile) {
            $stateData = Import-Csv -Path $this.StateFile
            foreach ($item in $stateData) {
                $this.TaskState[$item.TaskKey] = $item.NextPerson
            }
        }
    }

    [void]SaveTaskState() {
        Write-Host "Saving task state to $($this.StateFile)"
        $stateArray = @()
        foreach ($key in $this.TaskState.Keys) {
            $stateArray += [PSCustomObject]@{
                TaskKey    = $key
                NextPerson = $this.TaskState[$key]
            }
        }
        if ($stateArray.Count -gt 0) {
            $stateArray | Export-Csv -Path $this.StateFile -NoTypeInformation
        }
    }

    [void]CheckScheduledTasks() {
        Write-Host "Checking scheduled tasks"
        $now = [DateTime]::Now
        if (-not $this.TodaysTasks) {
            Write-Host "TodaysTasks not loaded - loading now"
            $this.LoadTodaysTasks()
        }

        foreach ($taskKey in $this.TodaysTasks.Keys) {
            $taskInfo = $this.TodaysTasks[$taskKey]
            # Skip tasks already completed or called
            if ($taskInfo.Completed -or $taskInfo.Called) {
                continue
            }

            $task = $taskInfo.Task
            $taskDateTime = $taskInfo.DateTime

            try {
                $timeDiffSeconds = ($now - $taskDateTime).TotalSeconds
                Write-Host "Checking task '$($task.Action)' at $($task.Time), diff: $timeDiffSeconds seconds"

                # If the task is more than 60 seconds in the past, mark it completed and stop checking it
                if ($timeDiffSeconds -gt 60) {
                    Write-Host "Task expired, marking completed: $taskKey"
                    $this.TodaysTasks[$taskKey].Completed = $true
                    continue
                }

                # If within the 60-second window (past or future within 60s), trigger it
                if ([Math]::Abs($timeDiffSeconds) -le 60) {
                    # Determine person for rotation
                    $names = $task.Name -split ':'
                    $person = $names[0]
                    $rotationKey = "$($task.Time)|$($task.DaysOfWeek)|$($task.Action)"

                    if ($names.Count -gt 1) {
                        if ($this.TaskState.ContainsKey($rotationKey)) {
                            $lastPerson = $this.TaskState[$rotationKey]
                            $currentIndex = [Array]::IndexOf($names, $lastPerson)
                            if ($currentIndex -eq -1 -or $currentIndex -eq $names.Count - 1) {
                                $person = $names[0]
                            }
                            else {
                                $person = $names[$currentIndex + 1]
                            }
                        }
                        $this.TaskState[$rotationKey] = $person
                        $this.SaveTaskState()
                    }

                    # Log task
                    $logEntry = [PSCustomObject]@{
                        Date       = $now.ToString("yyyy-MM-dd")
                        Time       = $now.ToString("HH:mm:ss")  # Keep 24-hour in logs
                        TaskTime   = $task.Time
                        DaysOfWeek = $task.DaysOfWeek
                        Person     = $person
                        Action     = $task.Action
                    }
                    $logEntry | Export-Csv -Path $this.LogFile -Append -NoTypeInformation

                    # Update UI
                    $this.LastTaskLabel.Text = "Last: $person - $($task.Action) at $($task.Time)"
                    $this.UpdateNextTaskDisplay()

                    # Show alert
                    $alertMessage = "$person, $($task.Action)!"
                    $this.AlertLabel.Text = $alertMessage
                    $this.AlertLabel.Visible = $true
                    $this.IsShowingAlert = $true
                    $this.AlertEndTime = $now.AddSeconds(8)
                    $this.MainPanel.Visible = $false
                    # Hide person panels during alert
                    #foreach ($panel in $this.PersonTaskPanels.Values) {
                    #    $panel.Visible = $false
                    #}

                    # Speak alert
                    $speechText = "It is $($now.ToString("h:mm tt")). $alertMessage"  # 12-hour for speech
                    $this.SpeechSynth.SpeakAsync($speechText)

                    # Mark as called/completed to avoid re-triggering
                    $this.TodaysTasks[$taskKey].Called = $true
                    $this.TodaysTasks[$taskKey].Completed = $true

                    # Only trigger one task per timer tick
                    break
                }
            }
            catch {
                Write-Host "Error processing task '$($task.Action)': $($_.Exception.Message)"
            }
        }

        # Update next task if not showing alert
        if (-not $this.IsShowingAlert) {
            $this.UpdateNextTaskDisplay()
        }
    }

    [bool]IsDayInRange([string]$currentDay, [string]$range) {
        # Handle empty or null cases
        if ([string]::IsNullOrWhiteSpace($range)) {
            Write-Host "Warning: Empty day range specified"
            return $false
        }

        # Special cases for all days
        if ($range -eq "Sunday-Saturday" -or $range -eq "All" -or $range -eq "*") {
            return $true
        }

        # Split multiple days if present (e.g., "Monday,Wednesday")
        if ($range.Contains(",")) {
            $days = $range.Split(",").Trim()
            return $days -contains $currentDay
        }

        # Handle day ranges (e.g., "Monday-Friday" or "Sunday-Thursday")
        if ($range -match "^(.*?)-(.*)$") {
            $startDay = $matches[1].Trim()
            $endDay = $matches[2].Trim()
            
            $daysOfWeek = @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
            $startIndex = $daysOfWeek.IndexOf($startDay)
            $endIndex = $daysOfWeek.IndexOf($endDay)
            $currentIndex = $daysOfWeek.IndexOf($currentDay)

            # Check for invalid day names
            if ($startIndex -eq -1 -or $endIndex -eq -1 -or $currentIndex -eq -1) {
                Write-Host "Warning: Invalid day name in range $range"
                return $false
            }

            # Handle week wraparound
            if ($startIndex -gt $endIndex) {
                $endIndex += 7
            }
            if ($currentIndex -lt $startIndex) {
                $currentIndex += 7
            }

            return $currentIndex -ge $startIndex -and $currentIndex -le $endIndex
        }

        # Single day case
        return $currentDay -eq $range
    }

    [void]UpdateNextTaskDisplay() {
        Write-Host "Updating next task display"
        $now = [DateTime]::Now

        # Collect today's tasks with parsed DateTime objects
        $todayTasks = @()
        foreach ($task in $this.Schedule | Where-Object { $this.IsDayInRange($now.DayOfWeek.ToString(), $_.DaysOfWeek) }) {
            try {
                $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, $taskTime.Hour, $taskTime.Minute, 0)
                $todayTasks += [PSCustomObject]@{ Task = $task; DateTime = $taskDateTime }
            }
            catch {
                # Ignore tasks with unparsable times
                continue
            }
        }

        # Sort by scheduled time
        $todayTasks = $todayTasks | Sort-Object -Property DateTime

        # Find the next upcoming task for today
        $upcoming = $todayTasks | Where-Object { $_.DateTime -gt $now }
        if ($upcoming.Count -gt 0) {
            $next = $upcoming[0]
            $displayTime = $next.DateTime.ToString("h:mm tt")
            $this.NextTaskLabel.Text = "Next: $($next.Task.Action) at $displayTime (Today)"
            Write-Host "Next task (today): $($next.Task.Action) at $($next.Task.Time)"
            return
        }

        # If there were tasks today but none are upcoming, say so
        if ($todayTasks.Count -gt 0) {
            $this.NextTaskLabel.Text = "No more tasks scheduled for today"
            Write-Host "All today's tasks have passed"
            return
        }

        # Otherwise, look for the first task tomorrow (or later)
        $tomorrow = $now.AddDays(1)
        $tomorrowDay = $tomorrow.DayOfWeek.ToString()
        $tomorrowTasks = @()
        foreach ($task in $this.Schedule | Where-Object { $this.IsDayInRange($tomorrowDay, $_.DaysOfWeek) }) {
            try {
                $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                $taskDateTime = [DateTime]::new($tomorrow.Year, $tomorrow.Month, $tomorrow.Day, $taskTime.Hour, $taskTime.Minute, 0)
                $tomorrowTasks += [PSCustomObject]@{ Task = $task; DateTime = $taskDateTime }
            }
            catch { continue }
        }
        $tomorrowTasks = $tomorrowTasks | Sort-Object -Property DateTime
        if ($tomorrowTasks.Count -gt 0) {
            $next = $tomorrowTasks[0]
            $displayTime = $next.DateTime.ToString("h:mm tt")
            $this.NextTaskLabel.Text = "Next: $($next.Task.Action) at $displayTime (Tomorrow)"
            Write-Host "Next task (tomorrow): $($next.Task.Action) at $($next.Task.Time)"
            return
        }

        # No tasks at all
        $this.NextTaskLabel.Text = "Next: No upcoming tasks"
        Write-Host "No upcoming tasks found"
    }

    [void]UpdatePersonTaskDisplays() {
        Write-Host "Updating person task displays"

        $now = [DateTime]::Now
        $currentDay = $now.DayOfWeek.ToString()

        # Get all unique persons from schedule (trim whitespace)
        $allPersons = @()
        foreach ($task in $this.Schedule) {
            $names = ($task.Name -split ':') | ForEach-Object { $_.Trim() }
            $allPersons += $names
        }
        $uniquePersons = $allPersons | Sort-Object -Unique

        # Create or update panels for each person
        $yOffset = $this.ExitLabel.Bottom + 20  # Start after the exit label
        foreach ($person in $uniquePersons) {
            # Get today's tasks for this person
            $personTasks = @()
            foreach ($task in $this.Schedule) {
                $names = ($task.Name -split ':') | ForEach-Object { $_.Trim() }
                if (-not $this.IsDayInRange($currentDay, $task.DaysOfWeek)) { continue }

                # If task is assigned to multiple people (rotation), only show it for the
                # person who is currently assigned according to TaskState (same logic as CheckScheduledTasks)
                if ($names.Count -gt 1) {
                    $rotationKey = "$($task.Time)|$($task.DaysOfWeek)|$($task.Action)"
                    $assigned = $names[0]
                    if ($this.TaskState.ContainsKey($rotationKey)) {
                        $lastPerson = $this.TaskState[$rotationKey]
                        $currentIndex = [Array]::IndexOf($names, $lastPerson)
                        if ($currentIndex -eq -1 -or $currentIndex -ge $names.Count - 1) {
                            $assigned = $names[0]
                        }
                        else {
                            $assigned = $names[$currentIndex + 1]
                        }
                    }

                    if ($assigned -eq $person) {
                        $personTasks += [PSCustomObject]@{
                            Time   = $task.Time
                            Action = $task.Action
                        }
                    }
                }
                else {
                    if ($names -contains $person) {
                        $personTasks += [PSCustomObject]@{
                            Time   = $task.Time
                            Action = $task.Action
                        }
                    }
                }
            }

            if ($personTasks.Count -eq 0) {
                continue
            }

            # Sort tasks by time
            $personTasks = $personTasks | Sort-Object { [DateTime]::ParseExact($_.Time, "HH:mm", $null) }

            # Create or update panel
            if (-not $this.PersonTaskPanels.ContainsKey($person)) {
                $panel = [Panel]::new()
                $panel.BackColor = [Color]::FromArgb(50, 50, 50, 50)  # Semi-transparent
                $panel.AutoSize = $false
                $panel.Size = [Size]::new(1000, 100 + ($personTasks.Count * 30))
                $panel.Location = [Point]::new(10, $yOffset)  # Position relative to MainPanel
                $this.MainPanel.Controls.Add($panel)
                $this.PersonTaskPanels[$person] = $panel
            }
            else {
                $panel = $this.PersonTaskPanels[$person]
                # Clear existing controls
                $panel.Controls.Clear()
                # Keep size and location consistent with creation
                $panel.Size = [Size]::new(1000, 100 + ($personTasks.Count * 30))
                $panel.Location = [Point]::new(10, $yOffset)
            }

            # Add person name header
            $nameLabel = [Label]::new()
            $nameLabel.Text = "$person's Tasks:"
            $nameLabel.Font = [Font]::new("Arial", 20, [FontStyle]::Bold)
            $nameLabel.ForeColor = [Color]::White
            $nameLabel.AutoSize = $true
            $nameLabel.Location = [Point]::new(10, 10)
            $panel.Controls.Add($nameLabel)

            # Add task items
            $yPos = 50
            $now = [DateTime]::Now
            foreach ($task in $personTasks) {
                $taskLabel = [Label]::new()
                # Parse time for both display and comparison
                $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                $displayTime = $taskTime.ToString("h:mm tt")  # Convert to AM/PM for display
                $taskLabel.Text = "$displayTime - $($task.Action)"
                $taskLabel.Font = [Font]::new("Arial", 16, [FontStyle]::Regular)
                $taskLabel.AutoSize = $true
                $taskLabel.Location = [Point]::new(30, $yPos)
                $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, $taskTime.Hour, $taskTime.Minute, 0)
                
                # Generate unique key for this task to check completion status
                $taskKey = "$($task.Time)|$person|$($task.Action)"
                
                if ($taskDateTime -lt $now -or 
                    ($this.TodaysTasks.ContainsKey($taskKey) -and $this.TodaysTasks[$taskKey].Completed)) {
                    $taskLabel.ForeColor = [Color]::Gray
                }
                else {
                    $taskLabel.ForeColor = [Color]::LightCyan
                }
                
                $panel.Controls.Add($taskLabel)
                $yPos += 30
            }

            # Recompute height after adding controls to ensure correct stacking
            $panel.Height = 100 + ($personTasks.Count * 30)
            $yOffset += $panel.Height + 20
        }

        # Hide panels for persons not in today's schedule
        $currentPersons = $uniquePersons
        $panelsToRemove = @()
        foreach ($key in $this.PersonTaskPanels.Keys) {
            if ($currentPersons -notcontains $key) {
                $panelsToRemove += $key
            }
        }
        foreach ($key in $panelsToRemove) {
            $this.MainPanel.Controls.Remove($this.PersonTaskPanels[$key])
            $this.PersonTaskPanels.Remove($key)
        }
    }

    [void]CleanupAndExit() {
        Write-Host "Cleaning up and exiting"
        try {
            # First, try to exit fullscreen and show window normally
            if ($this.Form) {
                $this.Form.WindowState = [FormWindowState]::Normal
                $this.Form.FormBorderStyle = [FormBorderStyle]::Sizable
                $this.Form.TopMost = $false
            }

            # Stop and dispose timer
            try {
                if ($this.Timer) {
                    $this.Timer.Stop()
                    $this.Timer.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing timer: $($_.Exception.Message)"
            }

            # Dispose speech synthesizer
            try {
                if ($this.SpeechSynth) {
                    $this.SpeechSynth.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing speech synthesizer: $($_.Exception.Message)"
            }

            # Dispose form
            try {
                if ($this.Form) {
                    $this.Form.Close()
                    $this.Form.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing form: $($_.Exception.Message)"
            }

            # Force exit as last resort
            #$currentPid = $PID
            #Start-Process powershell -ArgumentList "-NoProfile -Command Stop-Process -Id $currentPid -Force" -WindowStyle Hidden
            
            #[Environment]::Exit(0)
        }
        catch {
            Write-Host "Error during cleanup: $($_.Exception.Message)"
            # Force kill the process if all else fails
            #Stop-Process -Id $PID -Force
        }
    }
}
