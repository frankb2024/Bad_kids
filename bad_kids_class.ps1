

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
- Story, Quote & Joke Features (NEW)
  • Reads non-repeating stories, quotes, or jokes from CSV files
  • Tracks used entries via tracker files
  • Resets rotation when all items are used
  • Alerts now wrap text and display cleanly
  • Person panels arranged in three columns (max 9 total)
- Enhanced Stability
  • Comprehensive try/catch blocks with detailed error context
  • Loop-level error isolation to prevent single-task failures from crashing the app
The program acts as an impartial "referee" for scheduling, eliminating common
sources of argument like "it's not fair" or "it's not my turn" by maintaining
consistent rotation state and providing clear, timely notifications that everyone
can trust.
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
    [TextBox]$AlertTextBox  # Replaces AlertLabel for word wrapping
    [bool]$DebugMode = $false  # Debug mode flag
    [bool]$AutoMuteDuringSpeech = $true
    [bool]$IsSpeaking = $false
    [array]$Schedule
    [hashtable]$TaskState
    [hashtable]$PersonTaskPanels
    [hashtable]$TodaysTasks  # Track tasks for today and their completion status
    [DateTime]$LastScheduleLoad  # Track when we last loaded the schedule
    [string]$ScheduleFile = "schedule.csv"
    [string]$StateFile = "task_state.csv"
    [string]$LogFile = "TaskLog.csv"
    [string]$StoriesFile = "stories.csv"
    [string]$QuotesFile = "Daily Wisdom for Future Success.csv"
    [string]$JokesFile = "jokes.csv"
    [string]$StoryTrackerFile = "story_tracker.txt"
    [string]$QuoteTrackerFile = "quote_tracker.txt"
    [string]$JokeTrackerFile = "joke_tracker.txt"
    [array]$Stories = @()
    [array]$Quotes = @()
    [array]$Jokes = @()
    [Point]$MainPanelPosition
    [Size]$MainPanelSize
    [int]$MoveDirectionX = 1
    [double]$SpeechVolume = 0.35
    [int]$MoveDirectionY = 1
    [bool]$IsShowingAlert = $false
    [DateTime]$AlertEndTime

    SchedulerScreenSaver([bool]$debug = $false) {
        try {
            Write-Host "Creating form (Debug Mode: $debug)"
            $this.DebugMode = $debug
            # Ensure volume starts at 0
            try {
                [Audio]::SetVolume(0.0)
            }
            catch {
                Write-Host "Error setting initial volume: $($_.Exception.Message)"
            }
            $this.InitializeComponents()
            $this.LoadSchedule()
            $this.LoadTaskState()
            $this.InitializeTimer()
            Write-Host "Showing the Form"
            $this.Form.ShowDialog()
        }
        catch {
            Write-Host "FATAL ERROR in SchedulerScreenSaver constructor: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]InitializeComponents() {
        try {
            # Form setup
            Write-Host "Initializing form (Debug Mode: $($this.DebugMode))"
            $this.Form = [Form]::new()
            $this.Form.Text = "Scheduler Screen Saver"
            $this.StoriesFile = Join-Path -Path $PSScriptRoot -ChildPath $this.StoriesFile
            $this.QuotesFile = Join-Path -Path $PSScriptRoot -ChildPath $this.QuotesFile
            $this.JokesFile = Join-Path -Path $PSScriptRoot -ChildPath $this.JokesFile
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
                    try {
                        if ($e.KeyCode -eq [Keys]::Escape) { 
                            $s.tag.CleanupAndExit()
                            $s.Close()
                        }
                        elseif ($e.KeyCode -eq [Keys]::I) {
                            # Inject a test scheduled task 30 seconds from now (frank:john)
                            $s.tag.InjectScheduledTaskIn30Seconds()
                        }
                        elseif ($e.KeyCode -eq [Keys]::T) {
                            # Inject a single-person task for Frank in 30 seconds
                            $s.tag.InjectPersonTaskIn30Seconds('Frank')
                        }
                        elseif ($e.KeyCode -eq [Keys]::J) {
                            # Inject a joke-run in 30 seconds
                            $s.tag.InjectContentTaskIn30Seconds('jokes')
                        }
                        elseif ($e.KeyCode -eq [Keys]::W) {
                            # Inject a wisdom quote in 30 seconds
                            $s.tag.InjectContentTaskIn30Seconds('quotes')
                        }
                        elseif ($e.KeyCode -eq [Keys]::Q) {
                            # Inject a story in 30 seconds
                            $s.tag.InjectContentTaskIn30Seconds('story')
                        }
                    }
                    catch {
                        Write-Host "KeyDown handler error: $($_.Exception.Message)"
                    }
                })
            # Add mouse click handler
            $this.Form.Add_MouseClick({
                    param($s, $e)
                    $s.tag.CleanupAndExit()
                    $s.Close()
                })
            # Speech synthesizer
            $this.SpeechSynth = [SpeechSynthesizer]::new()
            # Wire automatic mute/unmute around speech if enabled
            $owner = $this
            try {
                $this.SpeechSynth.add_SpeakStarted({ param($s, $e)
                        try {
                            if ($owner.AutoMuteDuringSpeech) {
                                [Audio]::SetVolume($this.SpeechVolume)
                                $owner.IsSpeaking = $true
                            }
                        }
                        catch {
                            Write-Host "Warning: failed to unmute for speech: $($_.Exception.Message)"
                        }
                    })
                $this.SpeechSynth.add_SpeakCompleted({ param($s, $e)
                        try {
                            if ($owner.AutoMuteDuringSpeech) {
                                [Audio]::SetVolume(0.0)
                                $owner.IsSpeaking = $false
                            }
                        }
                        catch {
                            Write-Host "Warning: failed to mute after speech: $($_.Exception.Message)"
                        }
                    })
            }
            catch {
                Write-Host "Speech event hookup failed: $($_.Exception.Message)"
            }
            # Start muted when running normally (not in debug) so unexpected system sounds don't wake people
            if ($this.AutoMuteDuringSpeech ) {
                try {
                    if ($this.DebugMode -eq $false) {
                        [Audio]::SetVolume(0.0)
                    }
                }
                catch { 
                    Write-Host "Initial mute failed: $($_.Exception.Message)"
                }
            }
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
            $this.LastTaskLabel.Text = "Please wait..."
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
            # Use a TextBox for word wrapping (styled as label)
            $this.AlertTextBox = [TextBox]::new()
            $this.AlertTextBox.Font = [Font]::new("Arial", 24, [FontStyle]::Bold)
            $this.AlertTextBox.ForeColor = [Color]::Yellow
            $this.AlertTextBox.BackColor = [Color]::Black
            $this.AlertTextBox.ReadOnly = $true
            $this.AlertTextBox.Multiline = $true
            $this.AlertTextBox.ScrollBars = [ScrollBars]::Vertical
            $this.AlertTextBox.Visible = $false
            $this.AlertTextBox.Location = [Point]::new(100, 300)
            $this.AlertTextBox.Size = [Size]::new(800, 400)
            $this.AlertTextBox.WordWrap = $true
            # Add static controls
            $this.MainPanel.Controls.Add($this.DateLabel)
            $this.MainPanel.Controls.Add($this.TimeLabel)
            $this.MainPanel.Controls.Add($this.LastTaskLabel)
            $this.MainPanel.Controls.Add($this.NextTaskLabel)
            $this.MainPanel.Controls.Add($this.ExitLabel)
            $this.Form.Controls.Add($this.MainPanel)
            $this.Form.Controls.Add($this.AlertTextBox)
            # Ensure we attempt to mute after the form is shown (avoids early startup race
            # where audio subsystems may not yet be ready). Retry once if it fails.
            $this.Form.Add_Shown({ param($s, $e)
                    $owner = $s.tag
                    if ($owner.AutoMuteDuringSpeech) {
                        try {
                            if ($this.DebugMode -eq $false) {
                                [Audio]::SetVolume(0.0)
                            } 
                        }
                        catch {
                            Write-Host "Initial mute failed on Shown; retrying: $($_.Exception.Message)"
                            Start-Sleep -Milliseconds 200
                            try {
                                [Audio]::SetVolume(0.0) 
                            }
                            catch { 
                                Write-Host "Retry mute failed: $($_.Exception.Message)" 
                            }
                        }
                    }
                })
            # Initialize person task panels hashtable
            $this.PersonTaskPanels = @{}
        }
        catch {
            Write-Host "ERROR in InitializeComponents: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]InitializeTimer() {
        try {
            Write-Host "Initializing timer"
            $this.Timer = [Timer]::new()
            $this.Timer.Interval = 5000  # Check every second
            $this.Timer.tag = $this
            $this.Timer.Add_Tick({ 
                    param($sender, $e)
                    try {
                        $sender.tag.OnTimerTick($sender.tag) 
                    }
                    catch {
                        Write-Host "UNHANDLED ERROR in Timer Tick: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
                        # Do not rethrow — keep timer alive
                    }
                })
            $this.Timer.Start()
        }
        catch {
            Write-Host "ERROR in InitializeTimer: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]OnTimerTick([SchedulerScreenSaver] $inthis) {
        try {
            $now = [DateTime]::Now
            # Check if we need to reload schedule (at midnight or if not loaded today)
            if ($this.LastScheduleLoad.Date -ne $now.Date) {
                Write-Host "New day detected - reloading schedule"
                $this.LoadSchedule()
            }
            if ($this.LastTaskLabel.Text.Equals("Please wait...", [System.StringComparison]::InvariantCultureIgnoreCase)) {
                $this.LastTaskLabel.Text = "Last: No Previous Tasks"
            }
            # Update time display
            $this.DateLabel.Text = $now.ToString("dddd, MMMM dd, yyyy")
            $this.TimeLabel.Text = $now.ToString("h:mm:ss tt")  # 'h' for 12-hour with AM/PM
            # Handle alert timeout
            if ($this.IsShowingAlert -and $now -ge $this.AlertEndTime) {
                $this.AlertTextBox.Visible = $false
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
            Write-Host "ERROR in OnTimerTick: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            # Do not rethrow — keep timer running
        }
    }

    [void]MoveMainPanel() {
        try {
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
        catch {
            Write-Host "ERROR in MoveMainPanel: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]LoadSchedule() {
        try {
            Write-Host "Loading schedule from $($this.ScheduleFile)"
            if (Test-Path $this.ScheduleFile) {
                $this.Schedule = Import-Csv -Path $this.ScheduleFile
                $this.LastScheduleLoad = [DateTime]::Now
                $this.LoadStoriesAndQuotes()
                $this.LoadJokes()
                $this.LoadTodaysTasks()
            }
            else {
                $this.Schedule = @()
                Write-Host "Warning: $this.ScheduleFile not found. Creating sample file."
                $sample = @(
                    [PSCustomObject]@{Time = "22:00"; Name = "Frank:Lisa:Tom"; DaysOfWeek = "Sunday-Thursday"; Action = "go to bed"; short_title = "bedtime" }
                    [PSCustomObject]@{Time = "08:00"; Name = "Alice"; DaysOfWeek = "Monday-Friday"; Action = "take vitamins"; short_title = "vitamins" }
                    [PSCustomObject]@{Time = "19:00"; Name = "Frank:Alice"; DaysOfWeek = "Wednesday"; Action = "family dinner"; short_title = "dinner" }
                )
                $sample | Export-Csv -Path $this.ScheduleFile -NoTypeInformation
                $this.LoadStoriesAndQuotes()
                $this.LoadJokes()
                $this.LoadTodaysTasks()
            }
        }
        catch {
            Write-Host "ERROR in LoadSchedule: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]LoadStoriesAndQuotes() {
        try {
            # Load stories
            if (Test-Path $this.StoriesFile) {
                $this.Stories = @(Import-Csv -Path $this.StoriesFile)
                Write-Host "Loaded $($this.Stories.Count) stories"
            }
            else {
                Write-Host "Warning: $($this.StoriesFile) not found"
                $this.Stories = @()
            }
            # Load quotes
            if (Test-Path $this.QuotesFile) {
                $this.Quotes = @(Import-Csv -Path $this.QuotesFile)
                Write-Host "Loaded $($this.Quotes.Count) quotes"
            }
            else {
                Write-Host "Warning: $($this.QuotesFile) not found"
                $this.Quotes = @()
            }
        }
        catch {
            Write-Host "ERROR in LoadStoriesAndQuotes: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]LoadJokes() {
        try {
            if (Test-Path $this.JokesFile) {
                $this.Jokes = @(Import-Csv -Path $this.JokesFile)
                Write-Host "Loaded $($this.Jokes.Count) jokes"
            }
            else {
                Write-Host "Warning: $($this.JokesFile) not found"
                $this.Jokes = @()
            }
        }
        catch {
            Write-Host "ERROR in LoadJokes: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    

    [array]GetUnusedIndices([array]$items, [string]$trackerFile) {
        try {
            if ($items.Count -eq 0) { return @() }
            $used = @()
            if (Test-Path $trackerFile) {
                $used = Get-Content $trackerFile | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
            }
            $all = 0..($items.Count - 1)
            $unused = $all | Where-Object { $used -notcontains $_ }
            if ($unused.Count -eq 0) {
                if (Test-Path $trackerFile) { Remove-Item $trackerFile }
                return $all
            }
            return $unused
        }
        catch {
            Write-Host "ERROR in GetUnusedIndices: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            return @()
        }
    }

    [array]GetUnusedJokeIndices() {
        try {
            if ($this.Jokes.Count -eq 0) { return @() }
            $used = @()
            if (Test-Path $this.JokeTrackerFile) {
                $used = Get-Content $this.JokeTrackerFile | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
            }
            $all = 0..($this.Jokes.Count - 1)
            $unused = $all | Where-Object { $used -notcontains $_ }
            if ($unused.Count -eq 0) {
                if (Test-Path $this.JokeTrackerFile) { Remove-Item $this.JokeTrackerFile }
                return $all
            }
            return $unused
        }
        catch {
            Write-Host "ERROR in GetUnusedJokeIndices: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            return @()
        }
    }

    [void]RecordUsedIndex([int]$index, [string]$trackerFile) {
        try {
            Add-Content -Path $trackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedIndex (index=$index, file=$trackerFile): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]RecordUsedJokeIndex([int]$index) {
        try {
            Add-Content -Path $this.JokeTrackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedJokeIndex (index=$index): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]InjectScheduledTaskIn30Seconds() {
        try {
            $now = [DateTime]::Now
            $due = $now.AddSeconds(30)
            # Use the actual due time for display, but keep a fixed RotationAnchor so successive
            # injections rotate between names consistently (RotationAnchor is used when
            # determining rotation assignment)
            $timeString = $due.ToString('HH:mm')
            $rotationAnchor = "00:00"
            $names = "frank:john"
            $days = "Sunday-Saturday"
            $action = "Injected Test"
            $short = "Injected"

            $taskObj = [PSCustomObject]@{
                Time           = $timeString
                Name           = $names
                DaysOfWeek     = $days
                Action         = $action
                short_title    = $short
                RotationAnchor = $rotationAnchor
            }

            # Create a unique dictionary key for TodaysTasks so multiple injections can coexist
            $uniqueSuffix = $due.ToString("HHmmss")
            $taskKey = "$($taskObj.Time)|$($taskObj.Name)|$($taskObj.Action)|$uniqueSuffix"

            $this.TodaysTasks[$taskKey] = @{
                Task      = $taskObj
                DateTime  = $due
                Completed = $false
                Called    = $false
            }

            Write-Host "Injected scheduled task '$action' for [$names] at $($due.ToString('HH:mm:ss')) (key=$taskKey)"
        }
        catch {
            Write-Host "ERROR in InjectScheduledTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    [void]InjectPersonTaskIn30Seconds([string]$personName) {
        try {
            $now = [DateTime]::Now
            $due = $now.AddSeconds(30)
            $timeString = $due.ToString('HH:mm')
            $names = $personName
            $days = "Sunday-Saturday"
            $action = "Injected Manual"
            $short = "Injected"

            $taskObj = [PSCustomObject]@{
                Time        = $timeString
                Name        = $names
                DaysOfWeek  = $days
                Action      = $action
                short_title = $short
            }

            $uniqueSuffix = $due.ToString("HHmmss")
            $taskKey = "$($taskObj.Time)|$($taskObj.Name)|$($taskObj.Action)|$uniqueSuffix"

            $this.TodaysTasks[$taskKey] = @{
                Task      = $taskObj
                DateTime  = $due
                Completed = $false
                Called    = $false
            }

            Write-Host "Injected single-person task '$action' for [$names] at $($due.ToString('HH:mm:ss')) (key=$taskKey)"
        }
        catch {
            Write-Host "ERROR in InjectPersonTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    [void]InjectContentTaskIn30Seconds([string]$contentAction) {
        try {
            $now = [DateTime]::Now
            $due = $now.AddSeconds(30)
            $timeString = $due.ToString('HH:mm')
            $names = ""
            $days = "Sunday-Saturday"
            # contentAction should be one of 'jokes','story','quotes'
            $action = $contentAction
            $short = "Injected $action"

            $taskObj = [PSCustomObject]@{
                Time        = $timeString
                Name        = $names
                DaysOfWeek  = $days
                Action      = $action
                short_title = $short
            }

            $uniqueSuffix = $due.ToString("HHmmss")
            $taskKey = "$($taskObj.Time)|$($taskObj.Name)|$($taskObj.Action)|$uniqueSuffix"

            $this.TodaysTasks[$taskKey] = @{
                Task      = $taskObj
                DateTime  = $due
                Completed = $false
                Called    = $false
            }

            Write-Host "Injected content task '$action' at $($due.ToString('HH:mm:ss')) (key=$taskKey)"
        }
        catch {
            Write-Host "ERROR in InjectContentTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    [void]LoadTodaysTasks() {
        try {
            Write-Host "Loading today's tasks"
            $now = [DateTime]::Now
            $currentDay = $now.DayOfWeek.ToString()
            $this.TodaysTasks = @{}
            foreach ($task in $this.Schedule) {
                try {
                    if ($this.IsDayInRange($currentDay, $task.DaysOfWeek)) {
                        $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                        $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, 
                            $taskTime.Hour, $taskTime.Minute, 0)
                        $taskKey = "$($task.Time)|$($task.Name)|$($task.Action)"
                        $this.TodaysTasks[$taskKey] = @{
                            Task      = $task
                            DateTime  = $taskDateTime
                            Completed = $false
                            Called    = $false
                        }
                    }
                }
                catch {
                    Write-Host "ERROR processing task in LoadTodaysTasks - Action: '$($task.Action)', Name: '$($task.Name)': $($_.Exception.Message)`n$($_.Exception.StackTrace)"
                    continue
                }
            }
        }
        catch {
            Write-Host "ERROR in LoadTodaysTasks: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]LoadTaskState() {
        try {
            Write-Host "Loading task state from $($this.StateFile)"
            $this.TaskState = @{}
            if (Test-Path $this.StateFile) {
                $stateData = Import-Csv -Path $this.StateFile
                foreach ($item in $stateData) {
                    try {
                        $rawKey = $item.TaskKey.ToString().Trim()
                        $nextPerson = $item.NextPerson.ToString().Trim()
                        # If rawKey already looks like Time|Days|Action (3 parts), use it as-is
                        $parts = $rawKey -split '\|'
                        if ($parts.Count -eq 3) {
                            $normalized = "$($parts[0].Trim())|$($parts[1].Trim())|$($parts[2].Trim())"
                            $this.TaskState[$normalized] = $nextPerson
                            continue
                        }

                        # Migration: older keys may omit Time (e.g., Days|Action or just Action).
                        # Map such keys to all matching schedule entries by Action and optional DaysOfWeek.
                        $mapped = $false
                        if ($this.Schedule -and $this.Schedule.Count -gt 0) {
                            foreach ($sched in $this.Schedule) {
                                try {
                                    $schedAction = $sched.Action.ToString().Trim()
                                    $schedDays = $sched.DaysOfWeek.ToString().Trim()
                                    # If rawKey contains both days and action separated by '|', try to match
                                    if ($rawKey -match '\|') {
                                        $rparts = $rawKey -split '\|'
                                        $rDays = $rparts[0].Trim()
                                        $rAction = $rparts[1].Trim()
                                        if ($schedAction.Equals($rAction, [System.StringComparison]::InvariantCultureIgnoreCase) -and $schedDays.Equals($rDays, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                            $rotationKey = "$($sched.Time.ToString().Trim())|$schedDays|$schedAction"
                                            $this.TaskState[$rotationKey] = $nextPerson
                                            $mapped = $true
                                        }
                                    }
                                    else {
                                        # rawKey likely only contains action text; match by action (case-insensitive)
                                        if ($schedAction.Equals($rawKey, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                            $rotationKey = "$($sched.Time.ToString().Trim())|$schedDays|$schedAction"
                                            $this.TaskState[$rotationKey] = $nextPerson
                                            $mapped = $true
                                        }
                                    }
                                }
                                catch {
                                    continue
                                }
                            }
                        }
                        if (-not $mapped) {
                            # As a fallback, just store the raw key so it's not lost
                            $this.TaskState[$rawKey] = $nextPerson
                        }
                    }
                    catch {
                        Write-Host "ERROR loading task state item - TaskKey: '$($item.TaskKey)': $($_.Exception.Message)"
                        continue
                    }
                }
            }
        }
        catch {
            Write-Host "ERROR in LoadTaskState: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]SaveTaskState() {
        try {
            Write-Host "Saving task state to $($this.StateFile)"
            $stateArray = @()
            # Normalize and expand keys before writing so saved keys follow the canonical
            # Time|DaysOfWeek|Action format. If an older/raw key maps to multiple schedule
            # entries (same Action across times), expand it to each matching rotation key.
            $written = @{}
            foreach ($key in $this.TaskState.Keys) {
                $value = $this.TaskState[$key]
                $rawKey = $key.ToString().Trim()
                $parts = $rawKey -split '\|'
                $mapped = $false

                if ($parts.Count -eq 3) {
                    $normalized = "$($parts[0].Trim())|$($parts[1].Trim())|$($parts[2].Trim())"
                    if (-not $written.ContainsKey($normalized)) {
                        $stateArray += [PSCustomObject]@{ TaskKey = $normalized; NextPerson = $value }
                        $written[$normalized] = $true
                    }
                    continue
                }

                # Try to map older/raw keys to schedule entries
                if ($this.Schedule -and $this.Schedule.Count -gt 0) {
                    foreach ($sched in $this.Schedule) {
                        try {
                            $schedAction = $sched.Action.ToString().Trim()
                            $schedDays = $sched.DaysOfWeek.ToString().Trim()
                            if ($rawKey -match '\|') {
                                $rparts = $rawKey -split '\|'
                                $rDays = $rparts[0].Trim()
                                $rAction = $rparts[1].Trim()
                                if ($schedAction.Equals($rAction, [System.StringComparison]::InvariantCultureIgnoreCase) -and $schedDays.Equals($rDays, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                    $rotationKey = "$($sched.Time.ToString().Trim())|$schedDays|$schedAction"
                                    if (-not $written.ContainsKey($rotationKey)) {
                                        $stateArray += [PSCustomObject]@{ TaskKey = $rotationKey; NextPerson = $value }
                                        $written[$rotationKey] = $true
                                    }
                                    $mapped = $true
                                }
                            }
                            else {
                                if ($schedAction.Equals($rawKey, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                    $rotationKey = "$($sched.Time.ToString().Trim())|$schedDays|$schedAction"
                                    if (-not $written.ContainsKey($rotationKey)) {
                                        $stateArray += [PSCustomObject]@{ TaskKey = $rotationKey; NextPerson = $value }
                                        $written[$rotationKey] = $true
                                    }
                                    $mapped = $true
                                }
                            }
                        }
                        catch {
                            continue
                        }
                    }
                }

                if (-not $mapped) {
                    # Fallback: write raw key as-is
                    if (-not $written.ContainsKey($rawKey)) {
                        $stateArray += [PSCustomObject]@{ TaskKey = $rawKey; NextPerson = $value }
                        $written[$rawKey] = $true
                    }
                }
            }
            # Write atomically to avoid partial writes.
            $temp = "$($this.StateFile).tmp"
            if ($stateArray.Count -gt 0) {
                $stateArray | Export-Csv -Path $temp -NoTypeInformation -Force
                try {
                    Move-Item -Path $temp -Destination $this.StateFile -Force
                }
                catch {
                    # If move fails, attempt to write directly as fallback
                    Write-Host "Warning: atomic move failed when saving task state: $($_.Exception.Message). Falling back to direct write."
                    $stateArray | Export-Csv -Path $this.StateFile -NoTypeInformation -Force
                }
            }
            else {
                # No state to save; remove existing file if present
                if (Test-Path $this.StateFile) { Remove-Item $this.StateFile -ErrorAction SilentlyContinue }
            }
        }
        catch {
            Write-Host "ERROR in SaveTaskState: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]CheckScheduledTasks() {
        $helloTimer = [System.Windows.Forms.Timer]::new()
                                $helloTimer.Add_Tick({ param($s, $ev)
                                        try {
                                            $s.Stop()
                                            [Audio]::SetVolume(0.0)
                                        }
                                        catch {
                                            Write-Host "HelloTimer tick error: $($_.Exception.Message)"
                                        }
                                    })
                                    
        try {
            if ($this.DebugMode) { Write-Host "Checking scheduled tasks" }
            $now = [DateTime]::Now
            if (-not $this.TodaysTasks) {
                Write-Host "TodaysTasks not loaded - loading now"
                $this.LoadTodaysTasks()
            }
            foreach ($taskKey in $this.TodaysTasks.Keys) {
                try {
                    $taskInfo = $this.TodaysTasks[$taskKey]
                    if ($taskInfo.Completed -or $taskInfo.Called) {
                        continue
                    }
                    $task = $taskInfo.Task
                    $taskDateTime = $taskInfo.DateTime
                    $timeDiffSeconds = ($now - $taskDateTime).TotalSeconds
                    if ($this.DebugMode) {
                        Write-Host "Checking task '$($task.Action)' at $($task.Time), diff: $timeDiffSeconds seconds"
                    }
                    if ($timeDiffSeconds -gt 60) {
                        Write-Host "Task expired, marking completed: $taskKey"
                        $this.TodaysTasks[$taskKey].Completed = $true
                        continue
                    }
                    if ([Math]::Abs($timeDiffSeconds) -le 60) {
                        $personForLog = "Unknown"
                        $alertMessage = ""
                        $speechText = ""
                        # Special handling for "story", "quotes", or "jokes" actions
                        if ($task.Action -eq "story") {
                            $unused = $this.GetUnusedIndices($this.Stories, $this.StoryTrackerFile)
                            if ($unused.Count -eq 0) {
                                $speechText = "Story time! No stories available."
                                $displayText = "Story time! No stories available."
                            }
                            else {
                                $selectedIndex = Get-Random -InputObject $unused
                                $story = $this.Stories[$selectedIndex]
                                $this.RecordUsedIndex($selectedIndex, $this.StoryTrackerFile)
                                $speechText = "Story time! $($story.Title). $($story.Story) Moral: $($story.Moral)"
                                $displayText = "📖 STORY TIME 📖`n`nTitle: $($story.Title)`n`n$($story.Story)`n`nMoral: $($story.Moral)"
                            }
                            $alertMessage = $displayText
                            $personForLog = "System"
                        }
                        elseif ($task.Action -eq "quotes") {
                            $unused = $this.GetUnusedIndices($this.Quotes, $this.QuoteTrackerFile)
                            if ($unused.Count -eq 0) {
                                $speechText = "Time for wisdom! No quotes available."
                                $displayText = "Time for wisdom! No quotes available."
                            }
                            else {
                                $selectedIndex = Get-Random -InputObject $unused
                                $quote = $this.Quotes[$selectedIndex]
                                $this.RecordUsedIndex($selectedIndex, $this.QuoteTrackerFile)
                                $speechText = "Time for wisdom! $($quote.Quote) $($quote.Explanation)"
                                $displayText = "💡 DAILY WISDOM 💡`n`n$($quote.Quote)`n`n$($quote.Explanation)"
                            }
                            $alertMessage = $displayText
                            $personForLog = "System"
                        }
                        elseif ($task.Action -eq "jokes") {
                            $unused = $this.GetUnusedJokeIndices()
                            if ($unused.Count -eq 0) {
                                $speechText = "Time for a joke! No jokes available."
                                $displayText = "Time for a joke! No jokes available."
                            }
                            else {
                                $selectedIndex = Get-Random -InputObject $unused
                                $joke = $this.Jokes[$selectedIndex]
                                $this.RecordUsedJokeIndex($selectedIndex)
                                $speechText = "Here's a joke! $($joke.Joke)"
                                $displayText = "😂 JOKE TIME 😂`n`n$($joke.Joke)"
                            }
                            $alertMessage = $displayText
                            $personForLog = "System"
                        }
                        else {
                            # Normalize names (trim whitespace) and build rotation key consistently
                            $names = @((($task.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                            if ($names.Count -eq 0) {
                                # No assigned person; treat as system task
                                $person = "Unassigned"
                            }
                            else {
                                $person = $names[0]
                            }

                            # Use RotationAnchor for injected tasks when present so rotation
                            # alternation works independent of the displayed due time.
                            $timePart = if ($task.PSObject.Properties.Name -contains 'RotationAnchor') { $task.RotationAnchor.ToString().Trim() } else { $task.Time.ToString().Trim() }
                            $daysPart = $task.DaysOfWeek.ToString().Trim()
                            $actionPart = $task.Action.ToString().Trim()
                            $rotationKey = "$timePart|$daysPart|$actionPart"

                            if ($names.Count -gt 1) {
                                if ($this.DebugMode) { Write-Host "RotationKey='$rotationKey'; Names=[$( $names -join ',' )]" }
                                if ($this.TaskState.ContainsKey($rotationKey)) {
                                    $lastPerson = $this.TaskState[$rotationKey]
                                    if ($null -ne $lastPerson) { $lastPerson = $lastPerson.ToString().Trim() }
                                    # Find index case-insensitively
                                    $currentIndex = -1
                                    for ($i = 0; $i -lt $names.Count; $i++) {
                                        if ($names[$i].Equals($lastPerson, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                            $currentIndex = $i
                                            break
                                        }
                                    }
                                    if ($this.DebugMode) { Write-Host "LastPerson='$lastPerson'; CurrentIndex=$currentIndex" }
                                    if ($currentIndex -eq -1 -or $currentIndex -eq ($names.Count - 1)) {
                                        $person = $names[0]
                                    }
                                    else {
                                        $person = $names[$currentIndex + 1]
                                    }
                                }
                                # Store normalized person name
                                $this.TaskState[$rotationKey] = $person.ToString().Trim()
                                $this.SaveTaskState()
                            }
                            else {
                                # Single person assigned
                                $this.TaskState[$rotationKey] = $person.ToString().Trim()
                                $this.SaveTaskState()
                            }

                            $alertMessage = "$person, $($task.Action)!"
                            $speechText = "It is $($now.ToString("h:mm tt")). $alertMessage"
                            $personForLog = $person
                        }
                        # Log task
                        $logEntry = [PSCustomObject]@{
                            Date       = $now.ToString("yyyy-MM-dd")
                            Time       = $now.ToString("HH:mm:ss")
                            TaskTime   = $task.Time
                            DaysOfWeek = $task.DaysOfWeek
                            Person     = $personForLog
                            Action     = $task.Action
                        }
                        $logEntry | Export-Csv -Path $this.LogFile -Append -NoTypeInformation
                        # Update UI
                        $this.LastTaskLabel.Text = "Last: $($task.Action) at $($task.Time)"
                        $this.UpdateNextTaskDisplay()
                        # Show alert
                        $this.AlertTextBox.Text = $alertMessage
                        $this.AlertTextBox.Visible = $true
                        $this.IsShowingAlert = $true
                        $this.AlertEndTime = $now.AddSeconds(15)
                        $this.MainPanel.Visible = $false
                        # Speak alert
                        Write-Host "Speaking: $speechText"
                        try {
                            [Audio]::SetVolume($this.SpeechVolume)
                            $this.SpeechSynth.Speak($speechText)
                            

                            # Start a one-off UI timer that waits 60 seconds and then writes "hello" to the screen
                            try {
                                $helloTimer.Interval = 60000  # 60 seconds
                                $helloTimer.Tag = $this
                                $helloTimer.Start()
                            }
                            catch {
                                Write-Host "Failed to start hello timer: $($_.Exception.Message)"
                            }
                        }
                        catch {
                            Write-Host "Error during speech/volume control: $($_.Exception.Message)"
                        }
                        # Mark as called/completed
                        $this.TodaysTasks[$taskKey].Called = $true
                        $this.TodaysTasks[$taskKey].Completed = $true
                        break
                    }
                }
                catch {
                    Write-Host "ERROR processing task key '$($taskKey)' "
                    #Write-Host "Action: '$($task.Action)"
                    #Write-Host "Person: '$($task.Name)': "
                    Write-Host "$($_.Exception.Message)"
                    Write-Host "$($_.Exception.StackTrace)"
                    continue
                }
            }
            if (-not $this.IsShowingAlert) {
                $this.UpdateNextTaskDisplay()
            }
        }
        catch {
            Write-Host "ERROR in CheckScheduledTasks: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [bool]IsDayInRange([string]$currentDay, [string]$range) {
        try {
            if ([string]::IsNullOrWhiteSpace($range)) {
                Write-Host "Warning: Empty day range specified"
                return $false
            }
            if ($range -eq "Sunday-Saturday" -or $range -eq "All" -or $range -eq "*") {
                return $true
            }
            if ($range.Contains(",")) {
                $days = $range.Split(",").Trim()
                return $days -contains $currentDay
            }
            if ($range -match "^(.*?)-(.*)$") {
                $startDay = $matches[1].Trim()
                $endDay = $matches[2].Trim()
                $daysOfWeek = @("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
                $startIndex = $daysOfWeek.IndexOf($startDay)
                $endIndex = $daysOfWeek.IndexOf($endDay)
                $currentIndex = $daysOfWeek.IndexOf($currentDay)
                if ($startIndex -eq -1 -or $endIndex -eq -1 -or $currentIndex -eq -1) {
                    Write-Host "Warning: Invalid day name in range $range"
                    return $false
                }
                if ($startIndex -gt $endIndex) {
                    $endIndex += 7
                }
                if ($currentIndex -lt $startIndex) {
                    $currentIndex += 7
                }
                return $currentIndex -ge $startIndex -and $currentIndex -le $endIndex
            }
            return $currentDay -eq $range
        }
        catch {
            Write-Host "ERROR in IsDayInRange (currentDay='$currentDay', range='$range'): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            return $false
        }
    }

    [void]UpdateNextTaskDisplay() {
        try {
            if ($this.DebugMode) { Write-Host "Updating next task display" }
            $now = [DateTime]::Now
            $todayTasks = @()
            foreach ($task in $this.Schedule | Where-Object { $this.IsDayInRange($now.DayOfWeek.ToString(), $_.DaysOfWeek) }) {
                try {
                    $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                    $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, $taskTime.Hour, $taskTime.Minute, 0)
                    $todayTasks += [PSCustomObject]@{ Task = $task; DateTime = $taskDateTime }
                }
                catch {
                    Write-Host "ERROR parsing task time in UpdateNextTaskDisplay - Action: '$($task.Action)': $($_.Exception.Message)"
                    continue
                }
            }
            $todayTasks = $todayTasks | Sort-Object -Property DateTime
            $upcoming = $todayTasks | Where-Object { $_.DateTime -gt $now }
            if ($upcoming.Count -gt 0) {
                $next = $upcoming[0]
                $displayTime = $next.DateTime.ToString("h:mm tt")
                $this.NextTaskLabel.Text = "Next: $($next.Task.Action) at $displayTime (Today)"
                return
            }
            if ($todayTasks.Count -gt 0) {
                $this.NextTaskLabel.Text = "No more tasks scheduled for today"
                return
            }
            $tomorrow = $now.AddDays(1)
            $tomorrowDay = $tomorrow.DayOfWeek.ToString()
            $tomorrowTasks = @()
            foreach ($task in $this.Schedule | Where-Object { $this.IsDayInRange($tomorrowDay, $_.DaysOfWeek) }) {
                try {
                    $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                    $taskDateTime = [DateTime]::new($tomorrow.Year, $tomorrow.Month, $tomorrow.Day, $taskTime.Hour, $taskTime.Minute, 0)
                    $tomorrowTasks += [PSCustomObject]@{ Task = $task; DateTime = $taskDateTime }
                }
                catch {
                    Write-Host "ERROR parsing tomorrow task time - Action: '$($task.Action)': $($_.Exception.Message)"
                    continue
                }
            }
            $tomorrowTasks = $tomorrowTasks | Sort-Object -Property DateTime
            if ($tomorrowTasks.Count -gt 0) {
                $next = $tomorrowTasks[0]
                $displayTime = $next.DateTime.ToString("h:mm tt")
                $this.NextTaskLabel.Text = "Next: $($next.Task.Action) at $displayTime (Tomorrow)"
                return
            }
            $this.NextTaskLabel.Text = "Next: No upcoming tasks"
        }
        catch {
            Write-Host "ERROR in UpdateNextTaskDisplay: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]UpdatePersonTaskDisplays() {
        try {
            if ($this.DebugMode) { Write-Host "Updating person task displays" }
            $now = [DateTime]::Now
            $currentDay = $now.DayOfWeek.ToString()
            $allPersons = @()
            foreach ($task in $this.Schedule) {
                $names = @((($task.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                $allPersons += $names
            }
            # Include any names from injected/today's tasks so they appear in panels
            foreach ($tinfo in $this.TodaysTasks.Values) {
                try {
                    $t = $tinfo.Task
                    if ($null -ne $t -and -not [string]::IsNullOrWhiteSpace($t.Name)) {
                        $names = @((($t.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                        $allPersons += $names
                    }
                }
                catch { continue }
            }
            $uniquePersons = $allPersons | Sort-Object -Unique
            # Clear existing panels
            foreach ($panel in $this.PersonTaskPanels.Values) {
                $this.MainPanel.Controls.Remove($panel)
            }
            $this.PersonTaskPanels = @{}
            # Setup layout constants for 3 columns
            $panelWidth = 300
            $maxPerColumn = 2
            $columnGap = 20
            $leftX = 10
            $middleX = $leftX + $panelWidth + $columnGap
            $rightX = $middleX + $panelWidth + $columnGap
            $yStart = $this.ExitLabel.Bottom + 20
            $yOffsetCol1 = $yStart
            $yOffsetCol2 = $yStart
            $yOffsetCol3 = $yStart
            $countCol1 = 0
            $countCol2 = 0
            $countCol3 = 0
            foreach ($person in $uniquePersons) {
                try {
                    $personTasks = @()
                    # Track which display keys we've already added to avoid duplicates
                    $addedKeys = @{}
                    foreach ($task in $this.Schedule) {
                        try {
                            $names = ($task.Name -split ':') | ForEach-Object { $_.Trim() }
                            if (-not $this.IsDayInRange($currentDay, $task.DaysOfWeek)) { continue }
                            if ($names.Count -gt 1) {
                                $rotationKey = "$($task.Time)|$($task.DaysOfWeek)|$($task.Action)"
                                $assigned = $names[0]
                                if ($this.TaskState.ContainsKey($rotationKey)) {
                                    $lastPerson = $this.TaskState[$rotationKey]
                                    if ($null -ne $lastPerson) { $lastPerson = $lastPerson.ToString().Trim() }
                                    $currentIndex = -1
                                    for ($i = 0; $i -lt $names.Count; $i++) {
                                        if ($names[$i].Equals($lastPerson, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                            $currentIndex = $i
                                            break
                                        }
                                    }
                                    if ($currentIndex -eq -1 -or $currentIndex -ge ($names.Count - 1)) {
                                        $assigned = $names[0]
                                    }
                                    else {
                                        $assigned = $names[$currentIndex + 1]
                                    }
                                }
                                if ($assigned -eq $person) {
                                    $displayKey = "$($task.Time)|$person|$($task.Action)"
                                    if (-not $addedKeys.ContainsKey($displayKey)) {
                                        $personTasks += [PSCustomObject]@{
                                            Time        = $task.Time
                                            Action      = $task.Action
                                            short_title = if ($task.PSObject.Properties.Name -contains 'short_title') { $task.short_title } else { $null }
                                        }
                                        $addedKeys[$displayKey] = $true
                                    }
                                }
                            }
                            else {
                                if ($names -contains $person) {
                                    $displayKey = "$($task.Time)|$person|$($task.Action)"
                                    if (-not $addedKeys.ContainsKey($displayKey)) {
                                        $personTasks += [PSCustomObject]@{
                                            Time        = $task.Time
                                            Action      = $task.Action
                                            short_title = if ($task.PSObject.Properties.Name -contains 'short_title') { $task.short_title } else { $null }
                                        }
                                        $addedKeys[$displayKey] = $true
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Host "ERROR processing task for person '$person' - Action: '$($task.Action)': $($_.Exception.Message)"
                            continue
                        }
                    }
                    # Also account for injected or today's tasks present in TodaysTasks
                    foreach ($tinfo in $this.TodaysTasks.Values) {
                        try {
                            $t = $tinfo.Task
                            if (-not $this.IsDayInRange($currentDay, $t.DaysOfWeek)) { continue }
                            $names = @((($t.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                            # Determine rotation time anchor (injected tasks may set RotationAnchor)
                            $timeForRotation = if ($t.PSObject.Properties.Name -contains 'RotationAnchor') { $t.RotationAnchor } else { $t.Time }
                            if ($names.Count -gt 1) {
                                $rotationKey = "$($timeForRotation)|$($t.DaysOfWeek)|$($t.Action)"
                                $assigned = $names[0]
                                if ($this.TaskState.ContainsKey($rotationKey)) {
                                    $lastPerson = $this.TaskState[$rotationKey]
                                    if ($null -ne $lastPerson) { $lastPerson = $lastPerson.ToString().Trim() }
                                    $currentIndex = -1
                                    for ($i = 0; $i -lt $names.Count; $i++) {
                                        if ($names[$i].Equals($lastPerson, [System.StringComparison]::InvariantCultureIgnoreCase)) {
                                            $currentIndex = $i; break
                                        }
                                    }
                                    if ($currentIndex -eq -1 -or $currentIndex -ge ($names.Count - 1)) { $assigned = $names[0] }
                                    else { $assigned = $names[$currentIndex + 1] }
                                }
                                if ($assigned -eq $person) {
                                    $displayKey = "$($t.Time)|$person|$($t.Action)"
                                    if (-not $addedKeys.ContainsKey($displayKey)) {
                                        $personTasks += [PSCustomObject]@{
                                            Time        = $t.Time
                                            Action      = $t.Action
                                            short_title = if ($t.PSObject.Properties.Name -contains 'short_title') { $t.short_title } else { $null }
                                        }
                                        $addedKeys[$displayKey] = $true
                                    }
                                }
                            }
                            else {
                                if ($names -contains $person) {
                                    $displayKey = "$($t.Time)|$person|$($t.Action)"
                                    if (-not $addedKeys.ContainsKey($displayKey)) {
                                        $personTasks += [PSCustomObject]@{
                                            Time        = $t.Time
                                            Action      = $t.Action
                                            short_title = if ($t.PSObject.Properties.Name -contains 'short_title') { $t.short_title } else { $null }
                                        }
                                        $addedKeys[$displayKey] = $true
                                    }
                                }
                            }
                        }
                        catch {
                            continue
                        }
                    }
                    if ($personTasks.Count -eq 0) { continue }
                    $personTasks = $personTasks | Sort-Object { [DateTime]::ParseExact($_.Time, "HH:mm", $null) }
                    $panel = [Panel]::new()
                    $panel.BackColor = [Color]::FromArgb(50, 50, 50, 50)
                    $panel.Size = [Size]::new($panelWidth, 100 + ($personTasks.Count * 30))
                    # Assign to column
                    if ($countCol1 -lt $maxPerColumn) {
                        $panel.Location = [Point]::new($leftX, $yOffsetCol1)
                        $yOffsetCol1 += $panel.Height + 20
                        $countCol1++
                    }
                    elseif ($countCol2 -lt $maxPerColumn) {
                        $panel.Location = [Point]::new($middleX, $yOffsetCol2)
                        $yOffsetCol2 += $panel.Height + 20
                        $countCol2++
                    }
                    else {
                        $panel.Location = [Point]::new($rightX, $yOffsetCol3)
                        $yOffsetCol3 += $panel.Height + 20
                        $countCol3++
                    }
                    # Add content
                    $nameLabel = [Label]::new()
                    if ($person -eq "") {
                        $nameLabel.Text = "Unassigned Tasks:"
                    }
                    else {
                        $nameLabel.Text = "$person's Tasks:"
                    }
                    $nameLabel.Font = [Font]::new("Arial", 16, [FontStyle]::Bold)
                    $nameLabel.ForeColor = [Color]::White
                    $nameLabel.AutoSize = $true
                    $nameLabel.Location = [Point]::new(10, 10)
                    $panel.Controls.Add($nameLabel)
                    $yPos = 40
                    foreach ($task in $personTasks) {
                        try {
                            $taskLabel = [Label]::new()
                            $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                            $displayTime = $taskTime.ToString("h:mm tt")
                            $displayText = if (![string]::IsNullOrWhiteSpace($task.short_title)) {
                                $task.short_title
                            }
                            else {
                                $truncated = $task.Action.Substring(0, [Math]::Min(18, $task.Action.Length))
                                if ($task.Action.Length -gt 18) { "$truncated..." } else { $truncated }
                            }
                            $taskLabel.Text = "$displayTime - $displayText"
                            $taskLabel.Font = [Font]::new("Arial", 14, [FontStyle]::Regular)
                            $taskLabel.AutoSize = $true
                            $taskLabel.Location = [Point]::new(20, $yPos)
                            $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, $taskTime.Hour, $taskTime.Minute, 0)
                            $taskKey = "$($task.Time)|$person|$($task.Action)"
                            if ($taskDateTime -lt $now -or ($this.TodaysTasks.ContainsKey($taskKey) -and $this.TodaysTasks[$taskKey].Completed)) {
                                $taskLabel.ForeColor = [Color]::Gray
                            }
                            else {
                                $taskLabel.ForeColor = [Color]::LightCyan
                            }
                            $panel.Controls.Add($taskLabel)
                            $yPos += 25
                        }
                        catch {
                            Write-Host "ERROR creating task label for person '$person' - Action: '$($task.Action)': $($_.Exception.Message)"
                            continue
                        }
                    }
                    $this.MainPanel.Controls.Add($panel)
                    $this.PersonTaskPanels[$person] = $panel
                }
                catch {
                    Write-Host "ERROR building panel for person '$person': $($_.Exception.Message)`n$($_.Exception.StackTrace)"
                    continue
                }
            }
        }
        catch {
            Write-Host "ERROR in UpdatePersonTaskDisplays: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]CleanupAndExit() {
        try {
            Write-Host "Cleaning up and exiting"
            if ($this.Form) {
                $this.Form.WindowState = [FormWindowState]::Normal
                $this.Form.FormBorderStyle = [FormBorderStyle]::Sizable
                $this.Form.TopMost = $false
            }
            try {
                if ($this.Timer) {
                    $this.Timer.Stop()
                    $this.Timer.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing timer: $($_.Exception.Message)"
            }
            try {
                if ($this.SpeechSynth) {
                    if ($this.AutoMuteDuringSpeech) {
                        [Audio]::SetVolume($this.SpeechVolume) 
                    } 
                    $this.SpeechSynth.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing speech synthesizer: $($_.Exception.Message)"
            }
            try {
                if ($this.Form) {
                    $this.Form.Close()
                    $this.Form.Dispose()
                }
            }
            catch {
                Write-Host "Error disposing form: $($_.Exception.Message)"
            }
        }
        catch {
            Write-Host "ERROR in CleanupAndExit: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }
}

