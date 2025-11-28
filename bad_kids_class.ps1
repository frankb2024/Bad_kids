

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
    [string]$StateFile2 = "task_state2.csv"
    [DateTime]$ScheduleFileLastWrite
    [hashtable]$TaskState2
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
    [PictureBox]$ClockBox
    [PictureBox]$ArtBox
    [int]$MoveDirectionX = 1
    [double]$SpeechVolume = 0.35
    [int]$MoveDirectionY = 1
    [bool]$IsShowingAlert = $false
    [DateTime]$AlertEndTime

    # Constructor: SchedulerScreenSaver
    # Creates the form, initializes components, loads schedule and task state,
    # starts the timer, and shows the main form.
    # Parameters:
    #  - $debug: optional boolean to enable debug mode (verbose logging, windowed UI)
    SchedulerScreenSaver([bool]$debug = $false) {
        try {
            Write-Host "Creating form (Debug Mode: $debug)"
            $this.DebugMode = $debug
            $this.DebugMode = $true
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

    # InitializeComponents
    # Sets up the Windows Forms UI: main form, labels, panels, clock and art PictureBoxes,
    # keyboard and mouse handlers, speech synthesizer, and initial control layout.
    # This method intentionally isolates UI creation and is called once at startup.
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
                            # Inject a test scheduled task 30 seconds from now (frank:maria)
                            $s.tag.InjectScheduledTaskIn30Seconds()
                        }
                        elseif ($e.KeyCode -eq [Keys]::D) {
                            # Dump upcoming assignments for rotating tasks
                            $s.tag.DumpUpcomingAssignments(10)
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
            $this.AlertTextBox.Location = [Point]::new(50, 50)
            $this.AlertTextBox.Size = [Size]::new(1200, 800)
            $this.AlertTextBox.WordWrap = $true
            # Add static controls
            $this.MainPanel.Controls.Add($this.DateLabel)
            $this.MainPanel.Controls.Add($this.TimeLabel)
            $this.MainPanel.Controls.Add($this.LastTaskLabel)
            $this.MainPanel.Controls.Add($this.NextTaskLabel)
            $this.MainPanel.Controls.Add($this.ExitLabel)
            $this.Form.Controls.Add($this.MainPanel)
            $this.Form.Controls.Add($this.AlertTextBox)
            # Add a clock in the top-right of the MainPanel
            try {
                # Slightly larger clock for better readability
                $clockSize = 160
                $this.ClockBox = [System.Windows.Forms.PictureBox]::new()
                $this.ClockBox.Size = [System.Drawing.Size]::new($clockSize, $clockSize)
                # place at form upper-right so it remains visible during alerts
                $clockX = [Math]::Max(0, $this.Form.ClientSize.Width - $clockSize - 10)
                $this.ClockBox.Location = [System.Drawing.Point]::new($clockX, 10)
                $this.ClockBox.SizeMode = [PictureBoxSizeMode]::Normal
                $this.ClockBox.BackColor = [Color]::Black
                $this.ClockBox.BorderStyle = [BorderStyle]::None
                $this.ClockBox.Anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Right
                $this.Form.Controls.Add($this.ClockBox)
                # Initial draw for clock
                $this.UpdateClock([DateTime]::Now)

                # Add an art PictureBox in the lower-left of the Form (same size as clock)
                try {
                    $this.ArtBox = [System.Windows.Forms.PictureBox]::new()
                    $this.ArtBox.Size = [System.Drawing.Size]::new($clockSize, $clockSize)
                    $artY = [Math]::Max(0, $this.Form.ClientSize.Height - $clockSize - 10)
                    $this.ArtBox.Location = [System.Drawing.Point]::new(10, $artY)
                    $this.ArtBox.SizeMode = [PictureBoxSizeMode]::Normal
                    $this.ArtBox.BackColor = [Color]::Black
                    $this.ArtBox.BorderStyle = [BorderStyle]::None
                    $this.ArtBox.Anchor = [AnchorStyles]::Bottom -bor [AnchorStyles]::Left
                    $this.Form.Controls.Add($this.ArtBox)
                    # Initial simple draw
                    $this.UpdateArt([DateTime]::Now)
                }
                catch {
                    Write-Host "Warning: failed to create art control: $($_.Exception.Message)"
                }
            }
            catch {
                Write-Host "Warning: failed to create clock control: $($_.Exception.Message)"
            }
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

    # InitializeTimer
    # Creates and configures the main timer that drives periodic updates.
    # The timer tick is wired to `OnTimerTick` and uses a 5 second interval by default.
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


                    
    # OnTimerTick
    # Main periodic update handler invoked by the UI Timer. Performs:
    #  - Clock and art updates
    #  - Runtime schedule change detection and rotation rebuilds
    #  - Reloads schedule at midnight
    #  - Checks for scheduled tasks to fire
    #  - Updates person task displays and moves the main panel
    # Parameter:
    #  - $inthis: self reference passed from the timer tag
    [void]OnTimerTick([SchedulerScreenSaver] $inthis) {

        $inthis.Timer.Stop()

        function Get-Pos([string]$pos, [int]$ctrlW, [int]$ctrlH, [int]$fw, [int]$fh) {
            switch ($pos) {
                'UpperRight' { $x = [Math]::Max(0, $fw - $ctrlW - 10); $y = 10; $anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Right }
                'LowerRight' { $x = [Math]::Max(0, $fw - $ctrlW - 10); $y = [Math]::Max(0, $fh - $ctrlH - 10); $anchor = [AnchorStyles]::Bottom -bor [AnchorStyles]::Right }
                'LowerLeft' { $x = 10; $y = [Math]::Max(0, $fh - $ctrlH - 10); $anchor = [AnchorStyles]::Bottom -bor [AnchorStyles]::Left }
                default { $x = [Math]::Max(0, $fw - $ctrlW - 10); $y = 10; $anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Right }
            }
            return @{ X = $x; Y = $y; Anchor = $anchor }
        }
    
        try {
            $now = [DateTime]::Now
            # Runtime schedule change detection: if schedule.csv write time changed, rebuild task_state2 and reload today's tasks
            try {
                $currentWrite = $null
                if (Test-Path $this.ScheduleFile) { $currentWrite = (Get-Item $this.ScheduleFile).LastWriteTimeUtc }
                if ($currentWrite -and ($this.ScheduleFileLastWrite -eq $null -or $currentWrite -ne $this.ScheduleFileLastWrite)) {
                    Write-Host "Detected change in $($this.ScheduleFile) (LastWriteTimeUtc changed). Rebuilding rotation map."
                    $this.ScheduleFileLastWrite = $currentWrite
                    try { $this.EnsureTaskState2() } catch { Write-Host "EnsureTaskState2 failed during runtime detection: $($_.Exception.Message)" }
                    try { $this.LoadTodaysTasks() } catch { Write-Host "LoadTodaysTasks failed after schedule change: $($_.Exception.Message)" }
                }
            }
            catch {
                # Non-fatal; log and continue
                Write-Host "Schedule change detection error: $($_.Exception.Message)"
            }

            # Update clock visual
            # Randomize and place Clock and Art controls so they don't overlap.
            try {
                $positions = @('UpperRight', 'LowerRight', 'LowerLeft')
                if ($this.ClockBox -and $this.ArtBox) {
                    # Choose distinct positions for clock and art
                    $clockPos = $positions[(Get-Random -Minimum 0 -Maximum $positions.Count)]
                    $artPos = $clockPos  # <-- Fix: initialize to avoid "not assigned" warning
                    do { 
                        $artPos = $positions[(Get-Random -Minimum 0 -Maximum $positions.Count)] 
                    } while ($artPos -eq $clockPos)

                    $formW = 0; $formH = 0
                    try { 
                        $formW = $this.Form.ClientSize.Width
                        $formH = $this.Form.ClientSize.Height
                    }
                    catch { }

                    # Helper to compute location and anchor
                    try {
                        $cpos = Get-Pos $clockPos $this.ClockBox.Width $this.ClockBox.Height $formW $formH
                        $this.ClockBox.Anchor = $cpos.Anchor
                        $this.ClockBox.Location = [System.Drawing.Point]::new($cpos.X, $cpos.Y)
                    }
                    catch { }

                    try {
                        $apos = Get-Pos $artPos $this.ArtBox.Width $this.ArtBox.Height $formW $formH
                        $this.ArtBox.Anchor = $apos.Anchor
                        $this.ArtBox.Location = [System.Drawing.Point]::new($apos.X, $apos.Y)
                    }
                    catch { }
                }
            }
            catch {
                Write-Host "Positioning clock/art failed: $($_.Exception.Message)"
            }
            # Update clock visual
            try { $this.UpdateClock($now) } catch { Write-Host "Clock update failed: $($_.Exception.Message)" }
            # Update simple art display (circle)
            try { $this.UpdateArt($now) } catch { Write-Host "Art update failed: $($_.Exception.Message)" }
            
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
                # Restore visibility for each person panel hidden during alerts
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
        $inthis.Timer.Start()
    }

    # MoveMainPanel
    # Randomly moves the `MainPanel` position across the screen and implements
    # simple bouncing at display edges. Uses panel size to adapt movement step.
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

    # UpdateClock
    # Renders an analog clock into the `ClockBox` PictureBox for the given time.
    # Uses high-quality smoothing and safely swaps the produced Bitmap into the UI.
    # Parameters:
    #  - $time: DateTime to render the clock for
    [void]UpdateClock([DateTime]$time) {
        try {
            if (-not $this.ClockBox) { return }
            $w = $this.ClockBox.Width
            $h = $this.ClockBox.Height
            $bmp = New-Object System.Drawing.Bitmap $w, $h
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            # Clear background to match panel
            $g.Clear([Color]::Black)

            $cx = [int]($w / 2)
            $cy = [int]($h / 2)
            $radius = [int](($w - 10) / 2)

            # Draw outer circle
            $penFace = New-Object System.Drawing.Pen([Color]::White, 2)
            $g.DrawEllipse($penFace, $cx - $radius, $cy - $radius, $radius * 2, $radius * 2)

            # Draw hour ticks
            # Loop over 12 hour positions to draw the hour marker lines around the clock face
            for ($i = 0; $i -lt 12; $i++) {
                $angle = ($i / 12.0) * 2.0 * [Math]::PI
                $outerX = [int]($cx + $radius * [Math]::Cos($angle - [Math]::PI / 2))
                $outerY = [int]($cy + $radius * [Math]::Sin($angle - [Math]::PI / 2))
                $innerX = [int]($cx + ($radius - 14) * [Math]::Cos($angle - [Math]::PI / 2))
                $innerY = [int]($cy + ($radius - 14) * [Math]::Sin($angle - [Math]::PI / 2))
                $g.DrawLine($penFace, $outerX, $outerY, $innerX, $innerY)
            }

            # Draw half-hour markers (between hour ticks)
            $penHalf = New-Object System.Drawing.Pen([Color]::Gray, 1)
            # Loop over positions between hours to draw smaller half-hour markers
            for ($i = 0; $i -lt 12; $i++) {
                $angle = (($i + 0.5) / 12.0) * 2.0 * [Math]::PI
                $outerX = [int]($cx + ($radius - 4) * [Math]::Cos($angle - [Math]::PI / 2))
                $outerY = [int]($cy + ($radius - 4) * [Math]::Sin($angle - [Math]::PI / 2))
                $innerX = [int]($cx + ($radius - 10) * [Math]::Cos($angle - [Math]::PI / 2))
                $innerY = [int]($cy + ($radius - 10) * [Math]::Sin($angle - [Math]::PI / 2))
                $g.DrawLine($penHalf, $outerX, $outerY, $innerX, $innerY)
            }

            # Draw numeric hour labels (1 - 12)
            $fontSize = [Math]::Max(8, [int]($radius * 0.18))
            $fontNum = New-Object System.Drawing.Font("Arial", $fontSize, [System.Drawing.FontStyle]::Bold)
            $stringFormat = New-Object System.Drawing.StringFormat
            $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
            $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
            # Loop from 1 to 12 to place numeric labels around the clock face
            for ($i = 1; $i -le 12; $i++) {
                $angle = ($i / 12.0) * 2.0 * [Math]::PI
                $tx = $cx + [int](($radius - 28) * [Math]::Cos($angle - [Math]::PI / 2))
                $ty = $cy + [int](($radius - 28) * [Math]::Sin($angle - [Math]::PI / 2))
                $g.DrawString($i.ToString(), $fontNum, [System.Drawing.Brushes]::White, $tx, $ty, $stringFormat)
            }

            # Get time components
            $sec = $time.Second
            $min = $time.Minute + ($sec / 60.0)
            $hour = ($time.Hour % 12) + ($min / 60.0)

            # Angles (radians), adjust so 12 o'clock is -PI/2
            $secAngle = ($sec / 60.0) * 2.0 * [Math]::PI - [Math]::PI / 2
            $minAngle = ($min / 60.0) * 2.0 * [Math]::PI - [Math]::PI / 2
            $hourAngle = ($hour / 12.0) * 2.0 * [Math]::PI - [Math]::PI / 2

            # Draw hour hand
            $hourLen = [int]($radius * 0.5)
            $hourX = [int]($cx + $hourLen * [Math]::Cos($hourAngle))
            $hourY = [int]($cy + $hourLen * [Math]::Sin($hourAngle))
            $penHour = New-Object System.Drawing.Pen([Color]::LightGray, 6)
            $penHour.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $g.DrawLine($penHour, $cx, $cy, $hourX, $hourY)

            # Draw minute hand
            $minLen = [int]($radius * 0.75)
            $minX = [int]($cx + $minLen * [Math]::Cos($minAngle))
            $minY = [int]($cy + $minLen * [Math]::Sin($minAngle))
            $penMin = New-Object System.Drawing.Pen([Color]::White, 4)
            $penMin.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $g.DrawLine($penMin, $cx, $cy, $minX, $minY)

            # Draw second hand
            $secLen = [int]($radius * 0.85)
            $secX = [int]($cx + $secLen * [Math]::Cos($secAngle))
            $secY = [int]($cy + $secLen * [Math]::Sin($secAngle))
            $penSec = New-Object System.Drawing.Pen([Color]::Red, 2)
            $penSec.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $g.DrawLine($penSec, $cx, $cy, $secX, $secY)

            # Center dot
            $brushCenter = [System.Drawing.Brushes]::White
            $g.FillEllipse($brushCenter, $cx - 3, $cy - 3, 6, 6)

            $g.Dispose()

            # Swap image safely
            try {
                if ($this.ClockBox.Image -ne $null) {
                    $old = $this.ClockBox.Image
                    $this.ClockBox.Image = $null
                    $old.Dispose()
                }
            }
            catch { }
            $this.ClockBox.Image = $bmp
        }
        catch {
            Write-Host "ERROR in UpdateClock: $($_.Exception.Message)"
        }
    }

    # UpdateArt
    # Generates a kaleidoscope-like image and assigns it to `ArtBox`.
    # The renderer creates a fresh Bitmap, draws shapes mirrored across sectors,
    # and safely swaps the image into the PictureBox using a MemoryStream fallback
    # to avoid GDI+ / ImageAnimator issues.
    # Parameters:
    #  - $time: DateTime used as a seed for randomized appearance
    [void]UpdateArt([DateTime]$time) {
        try {
            if (-not $this.ArtBox) { return }
            $w = [Math]::Max(32, $this.ArtBox.Width)
            $h = [Math]::Max(32, $this.ArtBox.Height)
            $bmp = New-Object System.Drawing.Bitmap $w, $h
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $g.Clear([Color]::Black)

            $cx = [int]($w / 2)
            $cy = [int]($h / 2)
            $radius = [Math]::Min($w, $h) / 2 - 6

            # choose random but deterministic-ish seed from time seconds
            $seed = ($time.Second * 1000) -bor ($time.Millisecond)
            $rand = New-Object System.Random([int](Get-Random -Minimum 1 -Maximum 100000))

            # palette
            $palette = @(
                [Color]::FromArgb(220, 40, 120, 200),
                [Color]::FromArgb(200, 200, 50, 120),
                [Color]::FromArgb(200, 60, 200, 90),
                [Color]::FromArgb(200, 250, 200, 30),
                [Color]::FromArgb(200, 255, 128, 0)
            )

            $sectors = 6 + ($rand.Next() % 7) # 6..12 sectors
            $shapes = 60 + ($rand.Next() % 80) # total base shapes to place

            # Loop through each base shape to place randomized artwork elements
            for ($i = 0; $i -lt $shapes; $i++) {
                $angleOffset = $rand.NextDouble() * 360.0
                $shapeRadius = $rand.Next([int]($radius * 0.05), [int]($radius * 0.7))
                $shapeW = [int]($shapeRadius * (0.6 + $rand.NextDouble()))
                $shapeH = [int]($shapeRadius * (0.3 + $rand.NextDouble()))
                $relX = ($rand.NextDouble() * ($radius * 0.8) - ($radius * 0.4))
                $relY = ($rand.NextDouble() * ($radius * 0.8) - ($radius * 0.4))
                $col = $palette[$rand.Next(0, $palette.Count)]
                $brush = New-Object System.Drawing.SolidBrush $col

                # Mirror/draw the current base shape into each sector of the kaleidoscope
                for ($s = 0; $s -lt $sectors; $s++) {
                    $angle = ($s * 360.0 / $sectors) + $angleOffset
                    $g.TranslateTransform($cx, $cy)
                    $g.RotateTransform($angle)
                    if ($s % 2 -eq 0) { $g.ScaleTransform(1.0, 1.0) } else { $g.ScaleTransform(-1.0, 1.0) }

                    $rectX = [int]($relX - ($shapeW / 2))
                    $rectY = [int]($relY - ($shapeH / 2))
                    $g.FillEllipse($brush, $rectX, $rectY, $shapeW, $shapeH)

                    $g.ResetTransform()
                }
                $brush.Dispose()
            }

            # add radial rings
            $ringPen = New-Object System.Drawing.Pen([Color]::FromArgb(60, 255, 255, 255), 1)
            # Draw several concentric rings to add structure to the kaleidoscope
            for ($r = 1; $r -le 3; $r++) {
                $g.DrawEllipse($ringPen, $cx - ($radius * $r / 3), $cy - ($radius * $r / 3), ($radius * 2 * $r / 3), ($radius * 2 * $r / 3))
            }
            $ringPen.Dispose()

            $g.Dispose()

            # Swap image safely: first try to provide a fresh Bitmap instance
            try {
                if ($this.ArtBox.Image -ne $null) { $old = $this.ArtBox.Image; $this.ArtBox.Image = $null; $old.Dispose() }
            }
            catch { }

            try {
                $newImg = New-Object System.Drawing.Bitmap($bmp)
                $this.ArtBox.Image = $newImg
            }
            catch [System.ArgumentException] {
                try {
                    Write-Host "Art swap fallback: creating image via MemoryStream due to ArgumentException"
                    $ms = New-Object System.IO.MemoryStream
                    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
                    $ms.Position = 0
                    $img = [System.Drawing.Image]::FromStream($ms)
                    $this.ArtBox.Image = $img
                    $ms.Close()
                }
                catch {
                    Write-Host "Fallback art swap failed: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Host "ERROR in UpdateArt (kaleidoscope): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    # LoadSchedule
    # Loads `schedule.csv` from disk (or creates a sample file if missing),
    # sets `LastScheduleLoad`, loads related content (stories, quotes, jokes),
    # and ensures the compact rotation map (`task_state2.csv`) is present and up-to-date.
    [void]LoadSchedule() {
        try {
            Write-Host "Loading schedule from $($this.ScheduleFile)"
            if (Test-Path $this.ScheduleFile) {
                $this.Schedule = Import-Csv -Path $this.ScheduleFile
                # Cache schedule file last write time for change detection
                try { $this.ScheduleFileLastWrite = (Get-Item $this.ScheduleFile).LastWriteTimeUtc } catch { $this.ScheduleFileLastWrite = [DateTime]::Now.ToUniversalTime() }
                $this.LastScheduleLoad = [DateTime]::Now
                $this.LoadStoriesAndQuotes()
                $this.LoadJokes()
                # Ensure the rotation map is present and up-to-date
                try { $this.EnsureTaskState2() } catch { Write-Host "EnsureTaskState2 failed: $($_.Exception.Message)" }
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
                try { $this.EnsureTaskState2() } catch { Write-Host "EnsureTaskState2 failed (sample): $($_.Exception.Message)" }
                $this.LoadTodaysTasks()
            }
        }
        catch {
            Write-Host "ERROR in LoadSchedule: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    # LoadStoriesAndQuotes
    # Loads `stories.csv` and the quotes CSV into memory arrays used by content tasks.
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

    # LoadJokes
    # Loads `jokes.csv` into the `$this.Jokes` array for joke tasks.
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

    

    # GetUnusedIndices
    # Returns an array of indices into `$items` that have not yet been recorded
    # in `$trackerFile`. If all indices have been used the tracker is reset.
    # Parameters:
    #  - $items: array of items to index into
    #  - $trackerFile: path to a file that stores used indices (one per line)
    # Returns: array of integer indices that are unused
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

    # GetUnusedJokeIndices
    # Convenience wrapper of `GetUnusedIndices` specialized for jokes using
    # the class's `$Jokes` and `$JokeTrackerFile`.
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

    # RecordUsedIndex
    # Appends an index to a tracker file used to avoid repeating content until
    # all items have been used.
    # Parameters:
    #  - $index: integer index to record
    #  - $trackerFile: path to the tracker file
    [void]RecordUsedIndex([int]$index, [string]$trackerFile) {
        try {
            Add-Content -Path $trackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedIndex (index=$index, file=$trackerFile): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    # RecordUsedJokeIndex
    # Convenience wrapper to record a used joke index in the joke tracker file.
    [void]RecordUsedJokeIndex([int]$index) {
        try {
            Add-Content -Path $this.JokeTrackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedJokeIndex (index=$index): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    # InjectScheduledTaskIn30Seconds
    # Helper for testing: creates a scheduled task 30 seconds in the future
    # that behaves like a regular schedule entry (frank:maria). Useful for
    # exercising rotation, logging, and UI display without editing CSVs.
    [void]InjectScheduledTaskIn30Seconds() {
        try {
            $now = [DateTime]::Now
            # Base due time 30 seconds from now; we'll create per-person tasks around this
            $baseDue = $now.AddSeconds(30)
            # Names list: colon-separated values; duplicate names are allowed and treated individually
            $names = "chris:kevin:frank:maria"
            $days = "Sunday-Saturday"
            $action = "Injected Test"
            $short = "Injected"

            $nameList = @($names -split ':' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($nameList.Count -eq 0) { Write-Host "No names to inject"; return }

            foreach ($person in $nameList) {
                try {
                    # First task for this person at baseDue
                    $due1 = $baseDue
                    $timeString1 = $due1.ToString('HH:mm')
                    $taskObj1 = [PSCustomObject]@{
                        Time        = $timeString1
                        Name        = $person
                        DaysOfWeek  = $days
                        Action      = $action
                        short_title = $short
                    }
                    $taskKey1 = "$($taskObj1.Time)|$($taskObj1.Name)|$($taskObj1.Action)"
                    $this.TodaysTasks[$taskKey1] = @{
                        Task           = $taskObj1
                        DateTime       = $due1
                        Completed      = $false
                        Called         = $false
                        AssignedPerson = $person
                    }

                    # Second task for this person two minutes later
                    $due2 = $baseDue.AddMinutes(2)
                    $timeString2 = $due2.ToString('HH:mm')
                    $taskObj2 = [PSCustomObject]@{
                        Time        = $timeString2
                        Name        = $person
                        DaysOfWeek  = $days
                        Action      = $action
                        short_title = $short
                    }
                    $taskKey2 = "$($taskObj2.Time)|$($taskObj2.Name)|$($taskObj2.Action)"
                    $this.TodaysTasks[$taskKey2] = @{
                        Task           = $taskObj2
                        DateTime       = $due2
                        Completed      = $false
                        Called         = $false
                        AssignedPerson = $person
                    }

                    Write-Host "Injected tasks for '$($person)' at $($due1.ToString('HH:mm:ss')) and $($due2.ToString('HH:mm:ss')) (keys: $taskKey1, $taskKey2)"
                }
                catch {
                    Write-Host "ERROR injecting tasks for $($person): $($_.Exception.Message)"
                    continue
                }
            }
            # Immediately refresh UI so injected tasks appear without waiting for the next timer tick
            try {
                $this.UpdatePersonTaskDisplays()
                $this.UpdateNextTaskDisplay()
                try { $this.MainPanel.Refresh() } catch { }
                try { $this.Form.Refresh() } catch { }
            }
            catch {
                Write-Host "Warning: failed to refresh UI after injection: $($_.Exception.Message)"
            }
        }
        catch {
            Write-Host "ERROR in InjectScheduledTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    # InjectPersonTaskIn30Seconds
    # Creates a single-person scheduled task due in 30 seconds for testing.
    # Parameters:
    #  - $personName: name to assign for this injected task
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

            $taskKey = "$($taskObj.Time)|$($taskObj.Name)|$($taskObj.Action)"

            $this.TodaysTasks[$taskKey] = @{
                Task           = $taskObj
                DateTime       = $due
                Completed      = $false
                Called         = $false
                AssignedPerson = $null
            }

            Write-Host "Injected single-person task '$action' for [$names] at $($due.ToString('HH:mm:ss')) (key=$taskKey)"
            # Immediately refresh UI so the injected single-person task appears
            try {
                $this.UpdatePersonTaskDisplays()
                $this.UpdateNextTaskDisplay()
                try { $this.MainPanel.Refresh() } catch { }
                try { $this.Form.Refresh() } catch { }
            }
            catch {
                Write-Host "Warning: failed to refresh UI after InjectPersonTaskIn30Seconds: $($_.Exception.Message)"
            }
        }
        catch {
            Write-Host "ERROR in InjectPersonTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    

    # InjectContentTaskIn30Seconds
    # Creates an injected content task (jokes/story/quotes) due in 30 seconds.
    # Parameters:
    #  - $contentAction: one of 'jokes', 'story', or 'quotes'
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

            $taskKey = "$($taskObj.Time)|$($taskObj.Name)|$($taskObj.Action)"

            $this.TodaysTasks[$taskKey] = @{
                Task           = $taskObj
                DateTime       = $due
                Completed      = $false
                Called         = $false
                AssignedPerson = $null
            }

            Write-Host "Injected content task '$action' at $($due.ToString('HH:mm:ss')) (key=$taskKey)"
            # Immediately refresh UI so the injected content task appears
            try {
                $this.UpdatePersonTaskDisplays()
                $this.UpdateNextTaskDisplay()
                try { $this.MainPanel.Refresh() } catch { }
                try { $this.Form.Refresh() } catch { }
            }
            catch {
                Write-Host "Warning: failed to refresh UI after InjectContentTaskIn30Seconds: $($_.Exception.Message)"
            }
        }
        catch {
            Write-Host "ERROR in InjectContentTaskIn30Seconds: $($_.Exception.Message)"
        }
    }

    # LoadTodaysTasks
    # Builds `$this.TodaysTasks` from `$this.Schedule` for the current date.
    # For rotating tasks (Name contains ':'), this method consults the compact
    # rotation map `$this.TaskState2` and uses `GetAssignedPersonForDate` to
    # compute the AssignedPerson for each task occurrence.
    [void]LoadTodaysTasks() {
        try {
            Write-Host "Loading today's tasks"
            $now = [DateTime]::Now
            $currentDay = $now.DayOfWeek.ToString()
            # Preserve any existing injected task instances so they are not lost
            # when reloading the schedule at runtime. Merge existing keys into
            # the new collection so injected tasks remain visible immediately.
            $existing = @{}
            if ($this.TodaysTasks) {
                foreach ($k in $this.TodaysTasks.Keys) {
                    try { $existing[$k] = $this.TodaysTasks[$k] } catch { }
                }
            }
            $this.TodaysTasks = $existing
            
            # Iterate each scheduled row to build today's tasks list
            foreach ($task in $this.Schedule) {
                try {
                    if ($this.IsDayInRange($currentDay, $task.DaysOfWeek)) {
                        $taskTime = [DateTime]::ParseExact($task.Time, "HH:mm", $null)
                        $taskDateTime = [DateTime]::new($now.Year, $now.Month, $now.Day, 
                            $taskTime.Hour, $taskTime.Minute, 0)
                        $taskKey = "$($task.Time)|$($task.Name)|$($task.Action)"
                        # Compute assigned person from compact task_state2 if present
                        $assigned = $null
                        try {
                            if ($this.TaskState2 -and $this.TaskState2.Count -gt 0) {
                                $rotationKey = "$($task.Time.ToString().Trim())|$($task.DaysOfWeek.ToString().Trim())|$($task.Action.ToString().Trim())"
                                $assigned = $this.GetAssignedPersonForDate($rotationKey, $taskDateTime)
                            }
                        }
                        catch { $assigned = $null }

                        # Only add schedule-derived entries if an injected/explicit
                        # instance with the same key does not already exist. This
                        # prevents LoadTodaysTasks from overwriting manual injections.
                        if (-not $this.TodaysTasks.ContainsKey($taskKey)) {
                            $this.TodaysTasks[$taskKey] = @{
                                Task           = $task
                                DateTime       = $taskDateTime
                                Completed      = $false
                                Called         = $false
                                AssignedPerson = $assigned
                            }
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

    # LoadTaskState
    # Loads legacy `task_state.csv` which stored next-person rotation hints.
    # This method migrates older key formats to the canonical Time|Days|Action
    # keyspace used elsewhere in the program.
    [void]LoadTaskState() {
        try {
            Write-Host "Loading task state from $($this.StateFile)"
            $this.TaskState = @{}
            if (Test-Path $this.StateFile) {
                $stateData = Import-Csv -Path $this.StateFile
                # Iterate legacy state rows to migrate/normalize old rotation hints
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
                            # Scan schedule entries to find matches for legacy state keys
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

    # SaveTaskState
    # Persists the in-memory `$this.TaskState` to disk in a normalized form.
    # Writes atomically using a temporary file and Move-Item to avoid partial
    # writes.
    [void]SaveTaskState() {
        try {
            Write-Host "Saving task state to $($this.StateFile)"
            $stateArray = @()
            # Normalize and expand keys before writing so saved keys follow the canonical
            # Time|DaysOfWeek|Action format. If an older/raw key maps to multiple schedule
            # entries (same Action across times), expand it to each matching rotation key.
            $written = @{}
            # Normalize and expand in-memory task state keys before persisting
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
                    # Map older/raw keys to current schedule entries by scanning the schedule
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

    [void]EnsureTaskState2() {
        try {
            # If schedule file changed since last cached time, or task_state2 missing, rebuild
            $needRebuild = $false
            try {
                $currentWrite = (Get-Item $this.ScheduleFile).LastWriteTimeUtc
            }
            catch {
                $currentWrite = $null
            }
            if (-not (Test-Path $this.StateFile2)) { $needRebuild = $true }
            if ($currentWrite -and $this.ScheduleFileLastWrite -and ($currentWrite -ne $this.ScheduleFileLastWrite)) { $needRebuild = $true }
            if ($needRebuild) {
                Write-Host "Rebuilding $($this.StateFile2) due to schedule change or missing file"
                $this.BuildTaskState2()
                try { $this.ScheduleFileLastWrite = $currentWrite } catch { }
            }
            else {
                # Load existing state2 into memory as grouped rotation definitions
                try {
                    $this.TaskState2 = @{}
                    $rows = Import-Csv -Path $this.StateFile2
                    $grouped = $rows | Group-Object -Property TaskKey
                    # Build in-memory rotation definitions by grouping CSV rows by TaskKey
                    foreach ($g in $grouped) {
                        $taskKey = $g.Name
                        # Rows contain AnchorDate, Position, Name
                        $anchor = $g.Group | Select-Object -First 1 | Select-Object -ExpandProperty AnchorDate
                        $ordered = $g.Group | Sort-Object -Property @{Expression = { [int]$_.Position } } | ForEach-Object { $_.Name }
                        $this.TaskState2[$taskKey] = @{ AnchorDate = $anchor; Names = $ordered }
                    }
                    Write-Host "Loaded task_state2 with $($this.TaskState2.Count) rotation definitions"
                }
                catch {
                    Write-Host "Failed loading existing $($this.StateFile2): $($_.Exception.Message). Rebuilding."
                    $this.BuildTaskState2()
                }
            }
        }
        catch {
            Write-Host "ERROR in EnsureTaskState2: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    # BuildTaskState2
    # Builds a compact rotation map (`task_state2.csv`) describing rotating tasks.
    # The CSV contains one row per participant per rotating task, recording the
    # rotation's TaskKey, an AnchorDate, and the participant Position and Name.
    # This compact representation allows computing assignments by date math for
    # arbitrarily distant dates without precomputing every calendar date.
    [void]BuildTaskState2() {
        try {
            Write-Host "Building compact rotation map into $($this.StateFile2)"
            $rows = @()
            $this.TaskState2 = @{}
            $anchorDate = [DateTime]::Today.ToString('yyyy-MM-dd')
            # Walk each schedule row to generate compact rotation rows for rotating tasks
            foreach ($sched in $this.Schedule) {
                try {
                    $namesRaw = $sched.Name.ToString()
                    if ([string]::IsNullOrWhiteSpace($namesRaw)) { continue }
                    if ($namesRaw -notmatch ':') { continue }
                    $names = @($namesRaw -split ':' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    if ($names.Count -eq 0) { continue }

                    $rotationKey = "$($sched.Time.ToString().Trim())|$($sched.DaysOfWeek.ToString().Trim())|$($sched.Action.ToString().Trim())"
                    # For compact representation write one line per name with a Position (+1..N)
                    # Iterate through each participant name to write one line per participant in the compact rotation CSV
                    for ($i = 0; $i -lt $names.Count; $i++) {
                        $pos = $i + 1
                        $row = [PSCustomObject]@{
                            TaskKey    = $rotationKey
                            AnchorDate = $anchorDate
                            Position   = $pos
                            Name       = $names[$i]
                        }
                        $rows += $row
                    }
                    # keep in-memory structure
                    $this.TaskState2[$rotationKey] = @{ AnchorDate = $anchorDate; Names = $names }
                }
                catch {
                    continue
                }
            }
            # Write CSV atomically
            $temp = "$($this.StateFile2).tmp"
            if ($rows.Count -gt 0) {
                $rows | Export-Csv -Path $temp -NoTypeInformation -Force
                try { Move-Item -Path $temp -Destination $this.StateFile2 -Force } catch { Write-Host "Warning: move failed for $($this.StateFile2): $($_.Exception.Message); writing directly"; $rows | Export-Csv -Path $this.StateFile2 -NoTypeInformation -Force }
            }
            else {
                if (Test-Path $this.StateFile2) { Remove-Item $this.StateFile2 -ErrorAction SilentlyContinue }
            }

            # Debug: write compact rotation definitions
            Write-Host "Built rotation definitions (anchor date = $anchorDate):"
            # Debug print: iterate built rotation keys to log anchor and names
            foreach ($k in $this.TaskState2.Keys) {
                $def = $this.TaskState2[$k]
                Write-Host "$k => Anchor=$($def.AnchorDate) Names=($([string]::Join(',', $def.Names)))"
            }
            Write-Host "Finished building $($this.StateFile2) with $($this.TaskState2.Count) rotation entries"
        }
        catch {
            Write-Host "ERROR in BuildTaskState2: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    # GetAssignedPersonForDate
    # Given a compact rotation definition key (Time|DaysOfWeek|Action) and a
    # target date, compute which participant is assigned on that date by:
    #  - reading the rotation's AnchorDate and ordered Names from `$this.TaskState2`
    #  - counting occurrences of the scheduled day(s) between anchor and target
    #  - applying modular arithmetic to select the name by offset
    # Parameters:
    #  - $rotationKey: canonical rotation identifier (Time|DaysOfWeek|Action)
    #  - $targetDate: DateTime for which to compute the assigned person
    # Returns: string assigned person's name, or $null if not defined
    [string]GetAssignedPersonForDate([string]$rotationKey, [DateTime]$targetDate) {
        try {
            if (-not $this.TaskState2.ContainsKey($rotationKey)) { return $null }
            $def = $this.TaskState2[$rotationKey]
            $anchor = [DateTime]::ParseExact($def.AnchorDate, 'yyyy-MM-dd', $null)
            $names = $def.Names
            if (-not $names -or $names.Count -eq 0) { return $null }

            # Extract DaysOfWeek part from rotationKey (Time|DaysOfWeek|Action)
            $parts = $rotationKey -split '\|'
            $daysPart = if ($parts.Count -ge 2) { $parts[1] } else { 'Sunday-Saturday' }

            # Count schedule occurrences between anchor and target (respecting DaysOfWeek)
            $occurrences = 0
            if ($targetDate -eq $anchor) {
                $occurrences = 0
            }
            elseif ($targetDate -gt $anchor) {
                $d = $anchor.AddDays(1)
                while ($d -le $targetDate) {
                    if ($this.IsDayInRange($d.DayOfWeek.ToString(), $daysPart)) { $occurrences++ }
                    $d = $d.AddDays(1)
                }
            }
            else {
                # targetDate < anchor: count negative occurrences backward
                $d = $targetDate.AddDays(1)
                $back = 0
                while ($d -le $anchor) {
                    if ($this.IsDayInRange($d.DayOfWeek.ToString(), $daysPart)) { $back++ }
                    $d = $d.AddDays(1)
                }
                $occurrences = - $back
            }

            # anchor position is names[0] (Position 1)
            $n = $names.Count
            $anchorIndex = 0
            $idx = ($anchorIndex + $occurrences) % $n
            if ($idx -lt 0) { $idx = (($idx % $n) + $n) % $n }
            return $names[$idx]
        }
        catch {
            Write-Host "ERROR in GetAssignedPersonForDate for $rotationKey on $($targetDate.ToString('yyyy-MM-dd')): $($_.Exception.Message)"
            return $null
        }
    }

    [void]DumpUpcomingAssignments([int]$days = 10) {
        # DumpUpcomingAssignments
        # Logs the computed assignment for the next N days for each rotating task.
        # Parameters:
        #  - $days: integer number of days to include (default 10)
        # Uses GetAssignedPersonForDate() to compute assignments and writes to console.
    
        try {
            Write-Host "Today's scheduled tasks (date/time -> person -> action):"
            # Ensure TaskState2 loaded so we can compute rotating assignments
            try { if (-not $this.TaskState2 -or $this.TaskState2.Count -eq 0) { $this.EnsureTaskState2() } } catch { Write-Host "Warning: EnsureTaskState2 failed: $($_.Exception.Message)" }

            $today = [DateTime]::Today
            # Ensure today's tasks are loaded (this will include injected tasks)
            try { if (-not $this.TodaysTasks -or $this.TodaysTasks.Count -eq 0) { $this.LoadTodaysTasks() } } catch { Write-Host "Warning: LoadTodaysTasks failed: $($_.Exception.Message)" }

            $list = @()
            # Collect each today's task instance and resolve the display person
            foreach ($taskKey in $this.TodaysTasks.Keys) {
                try {
                    $tinfo = $this.TodaysTasks[$taskKey]
                    $dt = $tinfo.DateTime
                    if ($dt.Date -ne $today) { continue }
                    $task = $tinfo.Task

                    # Prefer AssignedPerson if the task instance recorded it
                    $person = $null
                    if ($tinfo.PSObject.Properties.Name -contains 'AssignedPerson' -and -not [string]::IsNullOrWhiteSpace($tinfo.AssignedPerson)) {
                        $person = $tinfo.AssignedPerson.ToString().Trim()
                    }
                    else {
                        $names = @((($task.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                        if ($names.Count -eq 0) { $person = 'Unassigned' }
                        elseif ($names.Count -eq 1) { $person = $names[0] }
                        else {
                            $rotationKey = "$($task.Time.ToString().Trim())|$($task.DaysOfWeek.ToString().Trim())|$($task.Action.ToString().Trim())"
                            try { $person = $this.GetAssignedPersonForDate($rotationKey, $dt) } catch { $person = $names[0] }
                            if (-not $person) { $person = $names[0] }
                        }
                    }

                    $list += [PSCustomObject]@{
                        DateTime = $dt
                        Time     = $dt.ToString('h:mm tt')
                        Action   = $task.Action
                        Person   = $person
                    }
                }
                catch {
                    Write-Host "Warning: failed while collecting task for dump: $($_.Exception.Message)"
                    continue
                }
            }

            if ($list.Count -eq 0) {
                Write-Host "No tasks scheduled for today."
                return
            }

            $list = $list | Sort-Object -Property DateTime
            foreach ($entry in $list) {
                Write-Host "$($entry.DateTime.ToString('yyyy-MM-dd')) $($entry.Time) - $($entry.Person) - $($entry.Action)"
            }
            Write-Host "Finished listing today's tasks"
        }
        catch {
            Write-Host "ERROR in DumpUpcomingAssignments: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]CheckScheduledTasks() {
        # CheckScheduledTasks
        # Iterates over `$this.TodaysTasks` and determines if any task is due.
        # For content tasks (story/quotes/jokes) it selects a non-repeating entry,
        # speaks and displays alerts, logs to `$this.LogFile`, and marks tasks called/completed.
        # For rotating person tasks it computes/updates rotation state and persists it.
        # This method is the runtime executor for scheduled items.
        write-Host "Checking scheduled tasks..."

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
            
            $now = [DateTime]::Now
            if (-not $this.TodaysTasks) {
                Write-Host "TodaysTasks not loaded - loading now"
                $this.LoadTodaysTasks()
            }
            
            Write-Host "Count $($this.TodaysTasks.Count)"
            # Iterate today's tasks to check each task for due/executing state
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
                        Write-Host "Checking task '$($task.Action)' $($task.Name) at $($task.Time), diff: $timeDiffSeconds seconds"
                    }

                    if ($timeDiffSeconds -gt 20) {
                        Write-Host "Task expired, marking completed: $taskKey"
                        $this.TodaysTasks[$taskKey].Completed = $true
                        continue
                    }

                    if ([Math]::Abs($timeDiffSeconds) -le 20) {
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
                                $speechText = "Story time! $($story.Title). $($story.Story)"
                                $displayText = "📖 STORY TIME 📖`n`nTitle: $($story.Title)`n`n$($story.Story)"
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
                                    # Iterate participants to find the index of the last person and determine the next assigned person
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
                            Date           = $now.ToString("yyyy-MM-dd")
                            Time           = $now.ToString("HH:mm:ss")
                            TaskTime       = $task.Time
                            DaysOfWeek     = $task.DaysOfWeek
                            Person         = $personForLog
                            AssignedPerson = $personForLog
                            Action         = $task.Action
                        }
                        $logEntry | Export-Csv -Path $this.LogFile -Append -NoTypeInformation
                        # Record the actual assigned person for this task instance so UI and
                        # any subsequent display logic can show who was called for this run.
                        try {
                            if ($this.TodaysTasks.ContainsKey($taskKey)) {
                                $this.TodaysTasks[$taskKey].AssignedPerson = $personForLog
                            }
                        }
                        catch {
                            Write-Host "Warning: failed to set AssignedPerson for key $($taskKey): $($_.Exception.Message)"
                        }
                        # Update UI
                        $this.LastTaskLabel.Text = "Last: $($task.Action) at $($task.Time)"
                        $this.UpdateNextTaskDisplay()
                        # Show alert
                        $this.AlertTextBox.Text = $alertMessage
                        $this.AlertTextBox.Visible = $true
                        $this.IsShowingAlert = $true
                        $this.AlertEndTime = $now.AddSeconds(20)
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
        # IsDayInRange
        # Returns true if `$currentDay` (e.g., 'Monday') is included in `$range`.
        # Supported range formats:
        #  - 'Sunday-Saturday' | 'All' | '*' => all days
        #  - Comma-separated list: 'Monday,Wednesday,Friday'
        #  - Range: 'Monday-Friday'
        #  - Single day name: 'Tuesday'
        # This helper centralizes day-range parsing used throughout the scheduler.
    
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
        # UpdateNextTaskDisplay
        # Recomputes and updates the 'Last' and 'Next' labels shown in the UI.
        # - Builds a list of today's tasks, resolves display person (prefers AssignedPerson),
        # - Picks the most recent past task as Last and the next future task as Next,
        # - Updates `$this.LastTaskLabel` and `$this.NextTaskLabel` accordingly.
    
        try {
            
            $now = [DateTime]::Now

            # Build a list of today's tasks (including injected/todays entries) with DateTime and resolved person
            $todayList = @()
            # Collect today's tasks and resolve display person for each (including injected tasks)
            foreach ($tinfo in $this.TodaysTasks.Values) {
                try {
                    $task = $tinfo.Task
                    $dt = $tinfo.DateTime
                    if ($dt.Date -ne $now.Date) { continue }

                    # Determine display person: prefer AssignedPerson if recorded, else compute from rotation/state
                    $displayPerson = $null
                    if ($tinfo.PSObject.Properties.Name -contains 'AssignedPerson' -and -not [string]::IsNullOrWhiteSpace($tinfo.AssignedPerson)) {
                        $displayPerson = $tinfo.AssignedPerson.ToString().Trim()
                    }
                    else {
                        # compute from task definition
                        $names = @((($task.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                        if ($names.Count -eq 0) { $displayPerson = 'Unassigned' }
                        elseif ($names.Count -eq 1) { $displayPerson = $names[0] }
                        else {
                            $rotationKey = "$($task.Time.ToString().Trim())|$($task.DaysOfWeek.ToString().Trim())|$($task.Action.ToString().Trim())"
                            $assigned = $names[0]
                            if ($this.TaskState.ContainsKey($rotationKey)) {
                                $lastPerson = $this.TaskState[$rotationKey]
                                if ($null -ne $lastPerson) { $lastPerson = $lastPerson.ToString().Trim() }
                                $currentIndex = -1
                                # Iterate participants to find the index of the last person and determine the next assigned person
                                for ($i = 0; $i -lt $names.Count; $i++) {
                                    if ($names[$i].Equals($lastPerson, [System.StringComparison]::InvariantCultureIgnoreCase)) { $currentIndex = $i; break }
                                }
                                if ($currentIndex -eq -1 -or $currentIndex -ge ($names.Count - 1)) { $assigned = $names[0] } else { $assigned = $names[$currentIndex + 1] }
                            }
                            $displayPerson = $assigned
                        }
                    }

                    $todayList += [PSCustomObject]@{ DateTime = $dt; Task = $task; DisplayPerson = $displayPerson; Called = $tinfo.Called; Completed = $tinfo.Completed }
                }
                catch { continue }
            }

            $todayList = $todayList | Sort-Object -Property DateTime

            # Previous: last task strictly before now (use most recent past)
            $previous = $todayList | Where-Object { $_.DateTime -lt $now } | Sort-Object -Property DateTime -Descending | Select-Object -First 1
            if ($previous) {
                $prevPerson = $previous.DisplayPerson
                $prevAction = $previous.Task.Action
                $this.LastTaskLabel.Text = "Last: $($prevPerson): $($prevAction)"
            }
            else {
                $this.LastTaskLabel.Text = "Last: No Previous Tasks"
            }

            # Next: first task strictly after now
            $next = $todayList | Where-Object { $_.DateTime -gt $now } | Sort-Object -Property DateTime | Select-Object -First 1
            if ($next) {
                $nextPerson = $next.DisplayPerson
                $nextAction = $next.Task.Action
                $displayTime = $next.DateTime.ToString("h:mm tt")
                $this.NextTaskLabel.Text = "Next: $($nextPerson) $($nextAction) at $($displayTime)"
            }
            else {
                $this.NextTaskLabel.Text = "No more tasks scheduled for today"
            }
        }
        catch {
            Write-Host "ERROR in UpdateNextTaskDisplay: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]UpdatePersonTaskDisplays() {
        # UpdatePersonTaskDisplays
        # Rebuilds person panels shown under the main panel. For each unique person
        # the method collects tasks assigned to them today (including injected tasks),
        # creates a compact panel showing time and short title, and grays-out completed tasks.
        # This keeps the UI in sync with rotation state and today's tasks.
    
        try {
            # Build panels exclusively from today's task instances (TodaysTasks) so assignments match runtime state
            $now = [DateTime]::Now
            $today = $now.Date

            # Map person -> list of task entries (each entry has DateTime, Action, short_title, Completed)
            $personMap = @{}
            $allPersons = @()

            # Ensure TodaysTasks is present
            if (-not $this.TodaysTasks -or $this.TodaysTasks.Count -eq 0) {
                try { $this.LoadTodaysTasks() } catch { }
            }

            foreach ($tinfo in $this.TodaysTasks.Values) {
                try {
                    $dt = $tinfo.DateTime
                    if ($dt.Date -ne $today) { continue }
                    $task = $tinfo.Task

                    # Resolve responsible person: prefer AssignedPerson recorded on the instance
                    $person = $null
                    if ($tinfo.PSObject.Properties.Name -contains 'AssignedPerson' -and -not [string]::IsNullOrWhiteSpace($tinfo.AssignedPerson)) {
                        $person = $tinfo.AssignedPerson.ToString().Trim()
                    }
                    else {
                        $names = @((($task.Name -split ':') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }))
                        if ($names.Count -eq 0) { $person = 'Unassigned' }
                        elseif ($names.Count -eq 1) { $person = $names[0] }
                        else {
                            # Multi-person rotation: compute assignment deterministically for this DateTime
                            $rotationKey = "$($task.Time.ToString().Trim())|$($task.DaysOfWeek.ToString().Trim())|$($task.Action.ToString().Trim())"
                            try { $person = $this.GetAssignedPersonForDate($rotationKey, $dt) } catch { $person = $names[0] }
                            if (-not $person) { $person = $names[0] }
                        }
                    }

                    if (-not $personMap.ContainsKey($person)) { $personMap[$person] = @() }
                    $personMap[$person] += [PSCustomObject]@{
                        DateTime    = $dt
                        Time        = $dt.ToString('HH:mm')
                        Action      = $task.Action
                        short_title = if ($task.PSObject.Properties.Name -contains 'short_title') { $task.short_title } else { $null }
                        Completed   = $tinfo.Completed
                        Called      = $tinfo.Called
                    }
                    $allPersons += $person
                }
                catch {
                    continue
                }
            }

            $uniquePersons = $allPersons | Sort-Object -Unique

            # Remove existing person panels from the UI before rebuilding
            foreach ($panel in $this.PersonTaskPanels.Values) {
                try { $this.MainPanel.Controls.Remove($panel) } catch { }
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
                    $personTasks = $personMap[$person]
                    if (-not $personTasks -or $personTasks.Count -eq 0) { continue }

                    # Sort tasks by DateTime
                    $personTasks = $personTasks | Sort-Object -Property DateTime

                    $panel = [Panel]::new()
                    $panel.BackColor = [Color]::FromArgb(50, 50, 50, 50)
                    # Size the panel tightly to the number of tasks: name label + task lines
                    $lineHeight = 25
                    $headerHeight = 30
                    $panelHeight = [Math]::Max(50, $headerHeight + ($personTasks.Count * $lineHeight) + 10)
                    $panel.Size = [Size]::new($panelWidth, $panelHeight)

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

                    # Add header
                    $nameLabel = [Label]::new()
                    if ($person -eq "") { $nameLabel.Text = "Unassigned Tasks:" } else { $nameLabel.Text = "$person's Tasks:" }
                    $nameLabel.Font = [Font]::new("Arial", 16, [FontStyle]::Bold)
                    $nameLabel.ForeColor = [Color]::White
                    $nameLabel.AutoSize = $true
                    $nameLabel.Location = [Point]::new(10, 10)
                    $panel.Controls.Add($nameLabel)

                    $yPos = 40
                    foreach ($taskEntry in $personTasks) {
                        try {
                            $taskLabel = [Label]::new()
                            $taskDateTime = $taskEntry.DateTime
                            $displayTime = $taskDateTime.ToString("h:mm tt")
                            $displayText = if (-not [string]::IsNullOrWhiteSpace($taskEntry.short_title)) { $taskEntry.short_title } else {
                                $truncated = $taskEntry.Action.Substring(0, [Math]::Min(18, $taskEntry.Action.Length))
                                if ($taskEntry.Action.Length -gt 18) { "$truncated..." } else { $truncated }
                            }
                            $taskLabel.Text = "$displayTime - $displayText"
                            $taskLabel.Font = [Font]::new("Arial", 14, [FontStyle]::Regular)
                            $taskLabel.AutoSize = $true
                            $taskLabel.Location = [Point]::new(20, $yPos)
                            if ($taskEntry.DateTime -lt $now -or $taskEntry.Completed) {
                                $taskLabel.ForeColor = [Color]::Gray
                            }
                            else {
                                $taskLabel.ForeColor = [Color]::LightCyan
                            }
                            $panel.Controls.Add($taskLabel)
                            $yPos += $lineHeight
                        }
                        catch {
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
        # CleanupAndExit
        # Gracefully disposes UI resources, timers, and the speech synthesizer.
        # Resets volume if necessary and disposes images and controls.
        # Should be called when the app exits to avoid resource leaks.
    
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
                if ($this.ClockBox) {
                    if ($this.ClockBox.Image -ne $null) { $this.ClockBox.Image.Dispose() }
                    $this.ClockBox.Dispose()
                }
                if ($this.ArtBox) {
                    try { if ($this.ArtBox.Image -ne $null) { $this.ArtBox.Image.Dispose() } } catch { }
                    try { $this.ArtBox.Dispose() } catch { }
                }
            }
            catch {
                Write-Host "Error disposing clock control: $($_.Exception.Message)"
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

