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

# HOW TO USE THIS PROGRAM

## SETUP
1. Create a CSV file named "schedule.csv" in the same directory as the script.
   The CSV should have columns: Time, Name, DaysOfWeek, Action, short_title
   - Time: 24-hour format (HH:MM)
   - Name: Single person name or colon-separated list for rotating tasks (e.g., "Frank:Alice:Tom")
   - DaysOfWeek: Days when task occurs (e.g., "Monday-Friday" or "Monday,Wednesday,Friday")
   - Action: Description of what to do
   - short_title: Brief description for display

2. Optional: Create CSV files for content:
   - "stories.csv" with columns: Title, Story
   - "Daily Wisdom for Future Success.csv" with columns: Quote, Explanation
   - "jokes.csv" with column: Joke

3. Run the program by executing the main script (bad_kids.ps1).  Please do not try to run bad_kids_class.ps1 directly.

## OPERATION
- The program will display a full-screen interface showing:
  * Current time and date
  * Last completed task
  * Next upcoming task
  * Individual task panels for each person
  * Decorative clock and kaleidoscope art.  Its hard to tell my kids to turn the wrench clockwise when clocks are a thing of the past.
  * The volume is always kept at zero except when speaking.  This is to prevent accidental noise from the computer in the middle of the night.
  * screensaver-like movement to prevent screen burn-in
  * The schedule.csv allows you to keep separate schedules for weekdays and weekends. (for example allow them sleep later on weekends) 

- When a task time arrives:
  * An alert appears with the task details
  * The task is announced via text-to-speech
  * The task is logged to "TaskLog.csv"

## KEYBOARD CONTROLS
- Esc: Exit the program (or click on the screen)
- J: Inject a joke in 30 seconds
- W: Inject a wisdom quote in 30 seconds
- Q: Inject a story in 30 seconds
- A: Advance all rotating tasks to change kids turns.

## FILE MANAGEMENT
- schedule.csv: Main task schedule
- task_state2.csv: Tracks rotation state for shared tasks. Delete if schedule changes.
- TaskLog.csv: Records completed tasks
- story_tracker.txt, quote_tracker.txt, joke_tracker.txt: Track used content
- You will also want to update the jokes.csv, stories.csv, and Daily Wisdom for Future Success.csv files with your own content.

## TROUBLESHOOTING
- If tasks aren't appearing, check the schedule.csv format
- If speech isn't working, verify Windows text-to-speech is enabled
- If rotation seems wrong, delete task_state2.csv to reset

## DEBUG MODE
- In debug mode, the window is resizable and shows additional controls

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
    [hashtable]$PersonTaskPanels
    [hashtable]$TodaysTasks  # Track tasks for today and their completion status
    [DateTime]$LastScheduleLoad  # Track when we last loaded the schedule
    [string]$ScheduleFile = "schedule.csv"
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
    [DateTime]$CurrentDate  # New: Current date for scheduling (can be modified for debugging)
    [bool]$ShouldTimerTickFire = $true

    SchedulerScreenSaver([bool]$debug = $false) {
        try {
            Write-Host "Creating form (Debug Mode: $debug)"
            $this.DebugMode = $debug
            #$this.DebugMode = $true
            $this.CurrentDate = [DateTime]::Now.Date  # Initialize with current date
           
            # Ensure volume starts at 0
            try {
                [Audio]::SetVolume(0.0)
            }
            catch {
                Write-Host "Error setting initial volume: $($_.Exception.Message)"
            }
            $this.InitializeComponents()
            $this.LoadSchedule()
            
            Write-Host "Initializing form (Debug Mode: $($this.DebugMode))"
            $this.Form = [Form]::new()
            $this.Form.Text = "Scheduler Screen Saver"
        
            if ($this.DebugMode) {
                $this.Form.WindowState = [FormWindowState]::Normal
                $this.Form.FormBorderStyle = [FormBorderStyle]::Sizable
                $this.Form.TopMost = $false
                $this.Form.Size = [System.Drawing.Size]::new(1024, 768)
                $this.Form.StartPosition = [FormStartPosition]::CenterScreen
            }
            else {
                $this.Form.WindowState = [FormWindowState]::Maximized
                $this.Form.FormBorderStyle = [FormBorderStyle]::None
                $this.Form.TopMost = $true
            }
        
            $this.Form.BackColor = [Color]::Black
            $this.Form.KeyPreview = $true
            $this.Form.tag = $this 

            $this.Form.Add_MouseClick({
                    param($s, $e)
                    $s.tag.CleanupAndExit()
                    $s.Close()
                })

            # Add Form Shown event handler to prepare the screen without starting the timer
            $this.Form.Add_Shown({
                    param($sender, $e)
                    Write-Host "Form shown, preparing screen without timer"
                    
                    # Add the key event handler now that the form is fully shown
                    $sender.Add_KeyDown({
                        param($s, $e)
                        $scheduler = $s.tag
                        
                        write-host "$($scheduler.DebugMode)"
                        
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
                        elseif ($e.KeyCode -eq [Keys]::A) {
                            # Advance all rotating tasks by one position
                            $s.tag.AdvanceAllRotatingTasks()
                        }
                        elseif ($scheduler.DebugMode -and $e.KeyCode -eq [Keys]::N) {
                            $scheduler.CurrentDate = $scheduler.CurrentDate.AddDays(1)
                            Write-Host "Debug: Advanced date to $($scheduler.CurrentDate.ToString('yyyy-MM-dd'))"
                            $scheduler.LoadSchedule()
                            $scheduler.UpdatePersonTaskDisplays()
                            $scheduler.UpdateNextTaskDisplay()
                        }
                        elseif ($e.KeyCode -eq [Keys]::Q) {
                            # Inject a wisdom quote in 30 seconds
                            $s.tag.InjectContentTaskIn30Seconds('quotes')
                        }
                        elseif ($e.KeyCode -eq [Keys]::S) {
                            # Inject a story in 30 seconds
                            $s.tag.InjectContentTaskIn30Seconds('story')
                        }

                        }
                        catch {
                            Write-Host "ERROR in key handler: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
                        }
                    })
                    
                    $sender.tag.PrepareScreenBeforeTimer()
                    Write-Host "Screen preparation complete - timer not started"
                    $sender.tag.InitializeTimer()
                })
 
                
            # Add controls to form in order: clock, art, then main panel
            $this.Form.Controls.AddRange(@(
                    $this.ClockBox,
                    $this.ArtBox,
                    $this.MainPanel,
                    $this.AlertTextBox
                ))
                
            $this.Form.ShowDialog()
        }
        catch {
            Write-Host "FATAL ERROR in SchedulerScreenSaver constructor: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }
    



    [void]InitializeControls() {
        write-host "Initializing controls"
        try {
            $this.InitializeMainPanel()
            $this.InitializeLabels()
            $this.InitializeAlertBox()
            $this.InitializeClockAndArt()
        
            # Add labels to MainPanel
            $this.MainPanel.Controls.AddRange(@(
                    $this.DateLabel,
                    $this.TimeLabel,
                    $this.LastTaskLabel,
                    $this.NextTaskLabel,
                    $this.ExitLabel
                ))
        }
        catch {
            Write-Host "ERROR in InitializeControls: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]InitializeAlertBox() {
        write-host "Initializing AlertBox"
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
    }

    [void]InitializeSpeech() {
        write-host "Initializing Speech Synthesizer"
        try {
            $this.SpeechSynth = [SpeechSynthesizer]::new()
            $owner = $this
        
            $this.SpeechSynth.add_SpeakStarted({ 
                    param($s, $e)
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
        
            $this.SpeechSynth.add_SpeakCompleted({ 
                    param($s, $e)
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
            Write-Host "ERROR in InitializeSpeech: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]InitializeMainPanel() {
        write-host "Initializing MainPanel"

        $this.MainPanel = [Panel]::new()
        $this.MainPanel.BackColor = [Color]::Transparent
        $this.MainPanel.AutoSize = $true
        $this.MainPanel.AutoSizeMode = [AutoSizeMode]::GrowAndShrink
        $this.MainPanelSize = [Size]::new(800, 200)
        $this.MainPanel.MinimumSize = $this.MainPanelSize
        $this.MainPanelPosition = [Point]::new(100, 200)  # Moved down to make room for clock/art
        $this.MainPanel.Location = $this.MainPanelPosition
    }

    [void]InitializeClockAndArt() {
        write-host "Initializing Clock and Art controls"
        try {
            $clockSize = 160
            $this.ClockBox = [System.Windows.Forms.PictureBox]::new()
            $this.ClockBox.Size = [System.Drawing.Size]::new($clockSize, $clockSize)
            $this.ClockBox.Location = [System.Drawing.Point]::new(10, 10)  # Changed from right side to top-left
            $this.ClockBox.SizeMode = [PictureBoxSizeMode]::Normal
            $this.ClockBox.BackColor = [Color]::Black
            $this.ClockBox.BorderStyle = [BorderStyle]::None
            $this.ClockBox.Anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Left  # Changed anchor

            $this.ArtBox = [System.Windows.Forms.PictureBox]::new()
            $this.ArtBox.Size = [System.Drawing.Size]::new($clockSize, $clockSize)
            $this.ArtBox.Location = [System.Drawing.Point]::new($clockSize + 20, 10)  # Position next to clock
            $this.ArtBox.SizeMode = [PictureBoxSizeMode]::Normal
            $this.ArtBox.BackColor = [Color]::Black
            $this.ArtBox.BorderStyle = [BorderStyle]::None
            $this.ArtBox.Anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Left  # Changed anchor
        }
        catch {
            Write-Host "Warning: failed to create clock/art controls: $($_.Exception.Message)"
        }
    }

    [void]InitializeComponents() {
        write-host "Initializing components"
        try {
            $this.InitializeSpeech()
            $this.InitializeControls()

            # Initialize person task panels
            $this.PersonTaskPanels = @{}
        }
        catch {
            Write-Host "ERROR in InitializeComponents: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]InitializeLabels() {
        write-host "Initializing Labels"

        $labelFont = [Font]::new("Arial", 24, [FontStyle]::Bold)
        $smallFont = [Font]::new("Arial", 18, [FontStyle]::Regular)

        # Date and Time labels at top
        $this.DateLabel = [Label]::new()
        $this.DateLabel.Font = $labelFont
        $this.DateLabel.ForeColor = [Color]::White
        $this.DateLabel.AutoSize = $true
        $this.DateLabel.Location = [Point]::new(10, 10)

        $this.TimeLabel = [Label]::new()
        $this.TimeLabel.Font = $labelFont
        $this.TimeLabel.ForeColor = [Color]::White
        $this.TimeLabel.AutoSize = $true
        $this.TimeLabel.Location = [Point]::new(10, 50)

        # Task labels below time
        $this.LastTaskLabel = [Label]::new()
        $this.LastTaskLabel.Font = $smallFont
        $this.LastTaskLabel.ForeColor = [Color]::Orange
        $this.LastTaskLabel.Text = "Last: Please wait..."
        $this.LastTaskLabel.AutoSize = $true
        $this.LastTaskLabel.Location = [Point]::new(10, 90)

        $this.NextTaskLabel = [Label]::new()
        $this.NextTaskLabel.Font = $smallFont
        $this.NextTaskLabel.ForeColor = [Color]::Lime
        $this.NextTaskLabel.Text = "Next: Please wait..."
        $this.NextTaskLabel.AutoSize = $true
        $this.NextTaskLabel.Location = [Point]::new(10, 120)

        # Exit label at bottom
        $this.ExitLabel = [Label]::new()
        $this.ExitLabel.Font = $smallFont
        $this.ExitLabel.ForeColor = [Color]::Gray
        $this.ExitLabel.Text = "Press Esc to Exit. Press A to re-asign rotating tasks. Press S for a story. Press Q for a quote. Press J for a joke."
        if ($this.DebugMode) {
            $this.ExitLabel.Text += " Press N to advance debug date."
        }
        $this.ExitLabel.AutoSize = $true
        $this.ExitLabel.Location = [Point]::new(10, 150)

        # Add labels to MainPanel
        $this.MainPanel.Controls.AddRange(@(
                $this.DateLabel,
                $this.TimeLabel,
                $this.LastTaskLabel,
                $this.NextTaskLabel,
                $this.ExitLabel
            ))
    }

    [void]InitializeTimer() {
        try {
            Write-Host "Initializing timer"
            $this.ShouldTimerTickFire = $false
            $this.Timer = [Timer]::new()
            $this.Timer.Interval = 5000
            $this.Timer.tag = $this
            $this.Timer.Add_Tick({ 
                    param($sender, $e)
                    write-host "Timer tick real event"
                    try {
                        if ($sender.tag.ShouldTimerTickFire) {
                            # Check our explicit state
                            $sender.tag.OnTimerTick() 
                        }
                    }
                    catch {
                        Write-Host "UNHANDLED ERROR in Timer Tick: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
                    }
                })
            # Timer will be started after the form is loaded
            $this.Timer.Start()
            $this.ShouldTimerTickFire = $true
        }
        catch {
            Write-Host "ERROR in InitializeTimer: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }
    
    [hashtable]GetPos([string]$pos, [int]$ctrlW, [int]$ctrlH, [int]$fw, [int]$fh) {
        # GetPos is a helper function to calculate control positions based on desired corner
        # Inputs: desired corner (UpperRight, LowerRight, LowerLeft), control width, control height, form width, form height
        $x = [Math]::Max(0, $fw - $ctrlW - 10)
        $y = 10
        $anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Right 

        switch ($pos) {
            'UpperRight' { 
                $x = [Math]::Max(0, $fw - $ctrlW - 10)
                $y = 10
                $anchor = [AnchorStyles]::Top -bor [AnchorStyles]::Right 
            }
            'LowerRight' { 
                $x = [Math]::Max(0, $fw - $ctrlW - 10)
                $y = [Math]::Max(0, $fh - $ctrlH - 10)
                $anchor = [AnchorStyles]::Bottom -bor [AnchorStyles]::Right 
            }
            'LowerLeft' { 
                $x = 10
                $y = [Math]::Max(0, $fh - $ctrlH - 10)
                $anchor = [AnchorStyles]::Bottom -bor [AnchorStyles]::Left 
            }
            default { 
            }
        }
        return @{ X = $x; Y = $y; Anchor = $anchor }
    }

    [void]MoveMainPanel() {
        # MoveMainPanel
        # This method creates the "screensaver" effect by moving the main panel around the screen.
        # The movement is designed to prevent screen burn-in on displays that show
        # the interface for extended periods.
        #
        # The movement algorithm works by:
        #
        # 1. Calculating random movement distances:
        #    - X and Y movements are randomized within ranges
        #    - Larger panels move more slowly (smaller increments)
        #    - This ensures visibility regardless of panel size
        #
        # 2. Applying movement in current direction:
        #    - Adds random distance to current position
        #    - Maintains current direction (horizontal/vertical)
        #
        # 3. Detecting and handling screen edge collisions:
        #    - If panel would go off-screen, direction is reversed
        #    - Position is clamped to ensure panel stays fully visible
        #    - This creates a "bouncing" effect at screen edges
        #
        # 4. Updating the panel position:
        #    - New position is applied to the MainPanel control
        #    - Position is stored for next movement calculation
        #
        # The movement is subtle enough to not distract from task information,
        # but sufficient to prevent static image retention on displays.
        #
        # This method is called every 5 seconds by the main timer.
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

    [void]OnTimerTick() {
        # OnTimerTick is the main periodic update method called by the timer every 5 seconds.
        # It performs the following key functions:  
        # 1. Checks for schedule changes and reloads if needed
        # 2. Updates clock and kaleidoscope art positions
        # 3. Updates time display labels
        # 4. Handles alert timeouts
        # 5. Checks for scheduled tasks to trigger
        # 6. Updates individual person task displays
        # 7. Moves the main panel to create a screensaver effect
        # The method is designed to be robust and handle errors gracefully.
        # It also includes a check to ensure the timer is running before proceeding.
       
        Write-Host "Timer tick"

        if (-not $this.ShouldTimerTickFire) {
            # Add this check
            return
        }
    

        $this.ShouldTimerTickFire = $false

        try {
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
            $this.CheckScheduleChanges()
            $this.UpdateClockAndArt($now)
            $this.UpdateTimeDisplay($now)
            $this.HandleAlertTimeout($now)
            $this.CheckScheduledTasks()
            $this.UpdatePersonTaskDisplays()
            $this.MoveMainPanel()
        }
        catch {
            Write-Host "ERROR in OnTimerTick: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
        $this.ShouldTimerTickFire = $true

    }

    
    [void]PrepareScreenBeforeTimer() {
        try {
            Write-Host "Preparing screen before timer starts"

            # Get current time
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }

            # Run all the important tasks from OnTimerTick
            $this.CheckScheduleChanges()
            $this.UpdateClockAndArt($now)
            $this.UpdateTimeDisplay($now)
            $this.HandleAlertTimeout($now)
            $this.CheckScheduledTasks()
            $this.UpdatePersonTaskDisplays()
            $this.MoveMainPanel()

            Write-Host "Screen preparation completed - all data loaded and displayed"
        }
        catch {
            Write-Host "ERROR in PrepareScreenBeforeTimer: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }


    [void]CheckScheduleChanges() {
        # CheckScheduleChanges is a method that checks for changes in the schedule file and reloads tasks if needed.
        # It performs the following steps:
        # 1. Retrieves the current write time of the schedule file.
        # 2. Compares the current write time with the last known write time.
        # 3. If a change is detected, it updates the last known write time and ensures the task state.
        # 4. It then loads the tasks for the current day.
        # The method is designed to be robust and handle errors gracefully.

        # CheckScheduleChanges is a method that checks for changes in the schedule file and reloads tasks if needed.
        # It performs the following steps:
        # 1. Retrieves the current write time of the schedule file.
        # 2. Compares the current write time with the last known write time.
        # 3. If a change is detected, it updates the last known write time and ensures the task state.
        # 4. It then loads the tasks for the current day.
        # The method is designed to be robust and handle errors gracefully.

        write-host "Checking for schedule changes"
        try {
            $currentWrite = $null
            if (Test-Path $this.ScheduleFile) { 
                $currentWrite = (Get-Item $this.ScheduleFile).LastWriteTimeUtc 
            }
        
            if ($currentWrite -and ($this.ScheduleFileLastWrite -eq $null -or $currentWrite -ne $this.ScheduleFileLastWrite)) {
                Write-Host "Detected change in $($this.ScheduleFile)"
                $this.ScheduleFileLastWrite = $currentWrite
                try { 
                    $this.EnsureTaskState2() 
                }
                catch { 
                    Write-Host "EnsureTaskState2 failed: $($_.Exception.Message)" 
                }
                try { 
                    $this.LoadTodaysTasks() 
                }
                catch { 
                    Write-Host "LoadTodaysTasks failed: $($_.Exception.Message)" 
                }
            }
        }
        catch {
            Write-Host "Schedule change detection error: $($_.Exception.Message)"
        }
    }

    [void]UpdateClockAndArt([DateTime]$now) {

        # UpdateClockAndArt updates the positions of the clock and kaleidoscope art on the form.
        # It randomly selects positions for each control from predefined corners of the screen,
        # ensuring they do not overlap. It then updates the controls' locations and anchors
        # accordingly. Finally, it calls methods to update the clock and art displays based on the
        # current time.

        write-host "Updating Clock and Art positions"
        try {
            $positions = @('UpperRight', 'LowerRight', 'LowerLeft')
            if ($this.ClockBox -and $this.ArtBox) {
                $clockPos = $positions[(Get-Random -Minimum 0 -Maximum $positions.Count)]
                $artPos = $clockPos
            
                do { 
                    $artPos = $positions[(Get-Random -Minimum 0 -Maximum $positions.Count)] 
                } while ($artPos -eq $clockPos)
            
                $formW = $this.Form.ClientSize.Width
                $formH = $this.Form.ClientSize.Height

                if ($this.ClockBox) {
                    $cpos = $this.GetPos($clockPos, $this.ClockBox.Width, $this.ClockBox.Height, $formW, $formH)
                    $this.ClockBox.Anchor = $cpos.Anchor
                    $this.ClockBox.Location = [System.Drawing.Point]::new($cpos.X, $cpos.Y)
                }

                if ($this.ArtBox) {
                    $apos = $this.GetPos($artPos, $this.ArtBox.Width, $this.ArtBox.Height, $formW, $formH)
                    $this.ArtBox.Anchor = $apos.Anchor
                    $this.ArtBox.Location = [System.Drawing.Point]::new($apos.X, $apos.Y)
                }
            }
        
            $this.UpdateClock($now)
            $this.UpdateArt($now)
        }
        catch {
            Write-Host "Clock/Art update failed: $($_.Exception.Message)"
        }
    }

    [void]SaveTaskState2() {
        # SaveTaskState2 saves the current task rotation state to a CSV file.
        # It writes the task key, anchor date, position, name, and whether the task is rotating.
        # The CSV file is saved to a temporary file, then moved to the final location to ensure atomicity.
        # If an error occurs during the process, it writes an error message and rethrows the exception.
        # The method is designed to be robust and handle errors gracefully.

        write-host "Saving TaskState2 to $($this.StateFile2)"
        
        try {
            $rows = @()
            foreach ($rotationKey in $this.TaskState2.Keys) {
                $def = $this.TaskState2[$rotationKey]
                for ($i = 0; $i -lt $def.Names.Count; $i++) {
                    $row = [PSCustomObject]@{
                        TaskKey    = $rotationKey
                        AnchorDate = $def.AnchorDate
                        Position   = $i + 1
                        Name       = $def.Names[$i]
                        IsRotating = $def.IsRotating
                    }
                    $rows += $row
                }
            }
        
            $temp = "$($this.StateFile2).tmp"
            $rows | Export-Csv -Path $temp -NoTypeInformation -Force
            Move-Item -Path $temp -Destination $this.StateFile2 -Force
        }
        catch {
            Write-Host "ERROR in SaveTaskState2: $($_.Exception.Message)"
            throw
        }
    }

    [void]AdvanceAllRotatingTasks() {
        # AdvanceAllRotatingTasks advances all rotating tasks by one position.
        # It updates the anchor date for each rotating task to the previous valid day

        try {
            Write-Host "Advancing all rotating tasks by one position"
            #if computer says that today is John's turn, but you need it to be Mary's turn instead,
            #this function will move it back one occurrence so that Mary is now up next.
            $newState = @{ }
            foreach ($rotationKey in $this.TaskState2.Keys) {
                $def = $this.TaskState2[$rotationKey]
                if ($def.IsRotating) {
                    # Move anchor date back by one occurrence
                    $parts = $rotationKey -split '\|'
                    $daysPart = if ($parts.Count -ge 2) { $parts[1] } else { 'Sunday-Saturday' }
                
                    # Find previous valid day based on DaysOfWeek
                    $anchor = [DateTime]::ParseExact($def.AnchorDate, 'yyyy-MM-dd', $null)
                    $prevDay = $anchor.AddDays(-1)
                    while (-not $this.IsDayInRange($prevDay.DayOfWeek.ToString(), $daysPart)) {
                        $prevDay = $prevDay.AddDays(-1)
                    }
                
                    # Update with new anchor date
                    $newState[$rotationKey] = @{
                        AnchorDate = $prevDay.ToString('yyyy-MM-dd')
                        Names      = $def.Names
                        IsRotating = $true
                    }
                }
                else {
                    $newState[$rotationKey] = $def
                }
            }
        
            # Update state and save
            $this.TaskState2 = $newState
            $this.SaveTaskState2()
        
            # Reload today's tasks to pick up changes
            $this.LoadTodaysTasks()
            $this.UpdatePersonTaskDisplays()
            $this.UpdateNextTaskDisplay()
        
            Write-Host "Successfully advanced all rotating tasks"
        }
        catch {
            Write-Host "ERROR in AdvanceAllRotatingTasks: $($_.Exception.Message)"
        }
    }

    [void]UpdateTimeDisplay([DateTime]$now) {
        # UpdateTimeDisplay updates the date and time labels on the form.
        # It checks if the day has changed since the last schedule load and reloads the schedule if needed.
        # It also updates the date and time labels with the current date and time.
        # If an error occurs during the process, it writes an error message and continues.

        write-host "Updating time display"
        try {
            if ($this.LastScheduleLoad.Date -ne $now.Date) {
                Write-Host "New day detected - reloading schedule"
                $this.CurrentDate = $now.Date
                Write-Host "Debug: Advanced date to $($this.CurrentDate.ToString('yyyy-MM-dd'))"
                $this.LoadSchedule()
                $this.UpdatePersonTaskDisplays()
                $this.UpdateNextTaskDisplay()
            }
        
            #if ($this.LastTaskLabel.Text.Equals("Please wait...", [System.StringComparison]::InvariantCultureIgnoreCase)) {
            #    $this.LastTaskLabel.Text = "Last: No Previous Tasks"
            #}

            $this.DateLabel.Text = $now.ToString("dddd, MMMM dd, yyyy")
            $this.TimeLabel.Text = $now.ToString("h:mm:ss tt")
        }
        catch {
            Write-Host "Time display update failed: $($_.Exception.Message)"
        }
    }

    [void]HandleAlertTimeout([DateTime]$now) {
        # HandleAlertTimeout checks if an alert is currently being shown and if its timeout has been reached.
        # If the alert timeout has been reached, it hides the alert and restores the main panel
        # and task panels to their original visibility states.
        # If an error occurs during the process, it writes an error message and continues.

        write-host "Checking for alert timeout"
        try {
            if ($this.IsShowingAlert -and $now -ge $this.AlertEndTime) {
                $this.AlertTextBox.Visible = $false
                $this.IsShowingAlert = $false
                $this.MainPanel.Visible = $true
            
                foreach ($panel in $this.PersonTaskPanels.Values) {
                    $panel.Visible = $true
                }
            }
        }
        catch {
            Write-Host "Alert timeout handling failed: $($_.Exception.Message)"
        }
    }

    [void]UpdateClock([DateTime]$time) {
        # UpdateClock
        # This method creates a custom analog clock display that updates every timer tick.
        # The clock is drawn programmatically using System.Drawing graphics rather than
        # using a pre-made image, allowing for dynamic updates and custom styling.
        #
        # The clock features:
        # - Hour, minute, and second hands
        # - Hour markers (12 positions) and half-hour markers
        # - Numeric hour labels (1-12)
        # - Center pivot point
        # - Color-coded hands (gray for hours, white for minutes, red for seconds)
        #
        # Drawing process:
        # 1. Create a bitmap with the clock dimensions
        # 2. Draw outer circle and hour markers
        # 3. Draw half-hour markers between hour markers
        # 4. Draw numeric labels (1-12) around the clock face
        # 5. Calculate hand angles based on current time
        # 6. Draw hour, minute, and second hands
        # 7. Add center pivot point
        # 8. Safely replace the previous clock image
        #
        # The clock positioning is handled by UpdateClockAndArt(), which moves
        # the clock to different corners of the screen to prevent burn-in.
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

    [void]UpdateArt([DateTime]$time) {
        # UpdateArt
        # This method creates a dynamic kaleidoscope pattern that changes with time.
        # The art is generated programmatically using mathematical patterns and randomness,
        # creating an ever-changing visual display that prevents screen burn-in.
        #
        # The kaleidoscope algorithm works by:
        #
        # 1. Creating a circular canvas with a black background
        # 2. Defining a color palette with semi-transparent colors
        # 3. Randomly determining the number of sectors (6-12)
        # 4. Generating random shapes (ellipses) in one sector:
        #    - Random position, size, and color from palette
        #    - Each shape is drawn multiple times across sectors
        # 5. Mirroring shapes across all sectors:
        #    - Alternating sectors are mirrored horizontally
        #    - Creates the classic kaleidoscope symmetry
        # 6. Adding concentric rings for structure
        # 7. Safely replacing the previous art image
        #
        # The art positioning is handled by UpdateClockAndArt(), which moves
        # the art to different corners of the screen to prevent burn-in.
        #
        # The randomness is seeded with time components to ensure:
        # - The pattern changes noticeably over time
        # - Different patterns appear on different days
        # - The same time always produces the same pattern (deterministic)
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

    [void]LoadSchedule() {
        # LoadSchedule loads the schedule from the CSV file.
        # It performs the following steps:
        # 1. Clears existing tasks
        # 2. Checks if the schedule file exists
        # 3. If it exists, imports the schedule and caches the last write time
        # 4. Loads stories, quotes, and jokes
        # 5. Ensures the task state is present and up-to-date
        # 6. Loads today's tasks
        # 7. If the schedule file does not exist, creates a sample schedule file


        try {
            Write-Host "Loading schedule from $($this.ScheduleFile)"

            # At the start of LoadSchedule(), before loading new tasks:
            $this.TodaysTasks = @{}  # Clear existing tasks

            if (Test-Path $this.ScheduleFile) {
                $this.Schedule = Import-Csv -Path $this.ScheduleFile
                # Cache schedule file last write time for change detection
                try { 
                    $this.ScheduleFileLastWrite = (Get-Item $this.ScheduleFile).LastWriteTimeUtc 
                }
                catch { 
                    $this.ScheduleFileLastWrite = [DateTime]::Now.ToUniversalTime()
                }
                $this.LastScheduleLoad = $this.CurrentDate  # Use CurrentDate instead of Now
                $this.LoadStoriesAndQuotes()
                $this.LoadJokes()
                # Ensure the rotation map is present and up-to-date
                try { 
                    $this.EnsureTaskState2() 
                    if ($this.CurrentDate.TimeOfDay -lt [TimeSpan]::FromMinutes(5)) {
                        $this.BuildTaskState2()
                    } 
                }
                catch { 
                    Write-Host "EnsureTaskState2 failed: $($_.Exception.Message)" 
                }
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
                $this.LastScheduleLoad = $this.CurrentDate  # Use CurrentDate instead of Now
                $this.LoadStoriesAndQuotes()
                $this.LoadJokes()
                # Ensure the rotation map is present and up-to-date
                try { 
                    $this.EnsureTaskState2() 
                    if ($this.CurrentDate.TimeOfDay -lt [TimeSpan]::FromMinutes(5)) {
                        $this.BuildTaskState2()
                    } 
                }
                catch { 
                    Write-Host "EnsureTaskState2 failed: $($_.Exception.Message)" 
                }
                $this.LoadTodaysTasks()
            }
        }
        catch {
            Write-Host "ERROR in LoadSchedule: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
            throw
        }
    }

    [void]LoadStoriesAndQuotes() {
        # LoadStoriesAndQuotes loads stories and quotes from their respective CSV files.
        # It performs the following steps:
        # 1. Checks if the stories and quotes files exist
        # 2. If they exist, imports the data into the Stories and Quotes properties
        # 3. If they do not exist, initializes the properties as empty arrays

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
        # LoadJokes loads jokes from the jokes CSV file.
        # It performs the following steps:
        # 1. Checks if the jokes file exists
        # 2. If it exists, imports the data into the Jokes property
        # 3. If it does not exist, initializes the property as an empty array

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
        # GetUnusedIndices returns an array of unused indices from the provided items array.
        # It reads the used indices from the specified tracker file and determines which indices are not used.

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
        # GetUnusedJokeIndices returns an array of unused joke indices.
        # It reads the used joke indices from the JokeTrackerFile and determines which indices are not used.

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
        # RecordUsedIndex records a used index by appending it to the specified tracker file.
        # It handles any errors that may occur during the file operation.
        # The function takes two parameters:
        #   $index: the index to record
        #   $trackerFile: the path to the tracker file

        try {
            Add-Content -Path $trackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedIndex (index=$index, file=$trackerFile): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]RecordUsedJokeIndex([int]$index) {
        # RecordUsedJokeIndex records a used joke index by appending it to the JokeTrackerFile.
        # It handles any errors that may occur during the file operation.


        try {
            Add-Content -Path $this.JokeTrackerFile -Value "$index"
        }
        catch {
            Write-Host "ERROR in RecordUsedJokeIndex (index=$index): $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]InjectScheduledTaskIn30Seconds() {
        # InjectScheduledTaskIn30Seconds injects multiple scheduled tasks due in 30 seconds for debugging purposes. It performs the following steps:
        # 1. Gets the current date and time.
        # 2. Sets the base due time to 30 seconds from now.
        # 3. Defines a list of names, days of the week, action, and short title.
        # 4. Splits the names list into an array of names.


        try {
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
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

    [void]InjectPersonTaskIn30Seconds([string]$personName) {
        # Inject a single-person task due in 30 seconds for debugging purposes
        try {
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
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

    [void]InjectContentTaskIn30Seconds([string]$contentAction) {
        
        # Inject a content task (joke, story, quote) due in 30 seconds for debugging purposes
        try {
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
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

    [void]LoadTodaysTasks() {
        # LoadTodaysTasks builds the list of tasks scheduled for today.
        # This method is called by the constructor and by InjectScheduledTaskIn30Seconds
        # to refresh the UI after a task is injected. It performs the following steps:
        # 1. Gets the current day of the week.
        # 2. Iterates through each scheduled task in the Schedule property.
        # 3. Checks if the current day is within the range specified by the DaysOfWeek property of the task.
        # 4. Parses the time of the task and creates a DateTime object for the task's due time.
        # 5. Constructs a unique key for the task using the task's time, name, and action.
        # 6. Adds the task to the TodaysTasks property of the object, using the key as the index.
        # 7. If the task has a task_state2 property, it attempts to determine the assigned person for the task.
        # 8. If the task is successfully added to the TodaysTasks property, it logs a message indicating the task was injected.
        # 9. If there is an error during the process, it logs an error message.


        try {
            Write-Host "Loading today's tasks"
            $now = $this.CurrentDate  # Use CurrentDate instead of Now
            $currentDay = $now.DayOfWeek.ToString()
            # Preserve any existing injected task instances so they are not lost
            # when reloading the schedule at runtime. Merge existing keys into
            # the new collection so injected tasks remain visible immediately.
            $existing = @{}
            if ($this.TodaysTasks) {
                write-host "Count: $($this.TodaysTasks.Keys.Count)"
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
                            # When creating task entries, add rotation flag:
                            $this.TodaysTasks[$taskKey] = @{
                                Task           = $task
                                DateTime       = $taskDateTime
                                Completed      = $false
                                Called         = $false
                                AssignedPerson = $assigned
                                IsRotating     = ($task.Name -match ':')  # Track if this is a rotating task
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

    [void]EnsureTaskState2() {
        # EnsureTaskState2 ensures that the compact rotation map (task_state2.csv) is present and up-to-date.
        # It performs the following steps:
        # 1. Checks if the task_state2.csv file exists.
        # 2. If the file does not exist, it logs a message indicating that the file is missing.
        # 3. If the file exists, it checks if the file has been modified since the last cached time.
        # 4. If the file has been modified or is missing, it rebuilds the task_state2.csv file by calling the BuildTaskState2 method.
        # 5. If the file is not modified, it loads the existing task_state2.csv file into memory as grouped rotation definitions.
        # 6. It logs the number of rotation definitions loaded from the task_state2.csv file.
        # 7. If there is an error during the process, it logs an error message.

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

  
    [void]BuildTaskState2() {
        # BuildTaskState2
        # This method creates a compact rotation map (`task_state2.csv`) that describes all rotating tasks.
        # It's called when the program starts or when the schedule file changes.
        #
        # The rotation map is a CSV file containing one row per participant per rotating task.
        # Each row records:
        # - TaskKey: A unique identifier (Time|DaysOfWeek|Action)
        # - AnchorDate: The starting date for the rotation
        # - Position: The person's position in the rotation (1-based)
        # - Name: The person's name
        # - IsRotating: Flag indicating this is a rotating task
        #
        # This compact representation is highly efficient because:
        # 1. It doesn't require storing the full rotation history
        # 2. It allows computing assignments for any date using simple date math
        # 3. It's resilient to program restarts and schedule changes
        #
        # For example, for a shower task shared by "Frank:Alice:Tom" at 8:00 PM on weekdays:
        # The CSV would contain:
        # TaskKey,AnchorDate,Position,Name,IsRotating
        # "20:00|Monday-Friday|shower","2023-01-01",1,"Frank",True
        # "20:00|Monday-Friday|shower","2023-01-01",2,"Alice",True
        # "20:00|Monday-Friday|shower","2023-01-01",3,"Tom",True
        #
        # When GetAssignedPersonForDate() needs to determine who showers on a specific date,
        # it counts weekdays between the anchor date and target date, then uses modulo arithmetic
        # to determine which position in the rotation is current.


        try {
            Write-Host "Building compact rotation map into $($this.StateFile2)"
            $rows = @()
            $this.TaskState2 = @{ }
            $anchorDate = $this.CurrentDate.ToString('yyyy-MM-dd')  # Use CurrentDate instead of Today

            # Walk each schedule row to generate compact rotation rows for rotating tasks
            foreach ($sched in $this.Schedule) {
                try {
                    $namesRaw = $sched.Name.ToString()
                    if ([string]::IsNullOrWhiteSpace($namesRaw)) { 
                        Write-Host "Warning: Skipping task with empty name - Action: $($sched.Action)"
                        continue 
                    }

                    # Check if this is a rotating task by looking for colon separator
                    $isRotating = $namesRaw -match ':'
                
                    if (-not $isRotating) { 
                        Write-Host "Info: Skipping non-rotating task - $($sched.Name): $($sched.Action)"
                        continue 
                    }

                    $names = @($namesRaw -split ':' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                    if ($names.Count -eq 0) { 
                        Write-Host "Warning: No valid names found in rotation - Raw: '$namesRaw'"
                        continue 
                    }

                    $rotationKey = "$($sched.Time.ToString().Trim())|$($sched.DaysOfWeek.ToString().Trim())|$($sched.Action.ToString().Trim())"

                    # For compact representation write one line per name with a Position (+1..N)
                    for ($i = 0; $i -lt $names.Count; $i++) {
                        $pos = $i + 1
                        $row = [PSCustomObject]@{
                            TaskKey    = $rotationKey
                            AnchorDate = $anchorDate
                            Position   = $pos
                            Name       = $names[$i]
                            IsRotating = $true  # Explicitly mark as rotating
                        }
                        $rows += $row
                    }

                    # Keep in-memory structure with rotation flag
                    $this.TaskState2[$rotationKey] = @{ 
                        AnchorDate = $anchorDate 
                        Names      = $names
                        IsRotating = $true  # Track rotation state
                    }
                }
                catch {
                    Write-Host "ERROR processing task in BuildTaskState2 - Action: '$($sched.Action)', Name: '$($sched.Name)':"
                    Write-Host "Exception: $($_.Exception.Message)"
                    Write-Host "Stack Trace: $($_.Exception.StackTrace)"
                    continue
                }
            }

            # Write CSV atomically
            $temp = "$($this.StateFile2).tmp"
            if ($rows.Count -gt 0) {
                try {
                    $rows | Export-Csv -Path $temp -NoTypeInformation -Force
                    Move-Item -Path $temp -Destination $this.StateFile2 -Force
                }
                catch {
                    Write-Host "ERROR writing task state file: $($_.Exception.Message)"
                    Write-Host "Stack Trace: $($_.Exception.StackTrace)"
                    throw
                }
            }
            else {
                if (Test-Path $this.StateFile2) { 
                    try {
                        Remove-Item $this.StateFile2 -ErrorAction SilentlyContinue 
                    }
                    catch {
                        Write-Host "ERROR removing empty state file: $($_.Exception.Message)"
                    }
                }
            }

            # Debug: write compact rotation definitions
            Write-Host "Built rotation definitions (anchor date = $anchorDate):"
            foreach ($k in $this.TaskState2.Keys) {
                $def = $this.TaskState2[$k]
                Write-Host "$k => Anchor=$($def.AnchorDate) Names=($([string]::Join(',', $def.Names))) Rotating=$($def.IsRotating)"
            }
            Write-Host "Finished building $($this.StateFile2) with $($this.TaskState2.Count) rotation entries"
        }
        catch {
            Write-Host "FATAL ERROR in BuildTaskState2: $($_.Exception.Message)"
            Write-Host "Stack Trace: $($_.Exception.StackTrace)"
            throw
        }
    }


    [string]GetAssignedPersonForDate([string]$rotationKey, [DateTime]$targetDate) {

        # GetAssignedPersonForDate
        # This is the core rotation algorithm that determines which person is assigned to a rotating task on any given date.
        # 
        # The algorithm works by:
        # 1. Using a compact rotation definition key (Time|DaysOfWeek|Action) to identify the task
        # 2. Reading the rotation's AnchorDate (when rotation started) and ordered Names from `$this.TaskState2`
        # 3. Counting how many times this task would have occurred between the anchor date and target date
        # 4. Using modular arithmetic to select the appropriate person based on the count
        #
        # For example, if "Frank:Alice:Tom" share a shower task:
        # - Anchor date: 2023-01-01 (Frank)
        # - Target date: 2023-01-08 (one week later)
        # - If task occurs daily, there are 7 occurrences between dates
        # - 7 % 3 = 1, so Alice (position 1) is assigned
        #
        # This approach allows us to compute assignments for any date without storing
        # the full rotation history, making it efficient and scalable.
        #
        # Parameters:
        #  - $rotationKey: canonical rotation identifier (Time|DaysOfWeek|Action)
        #  - $targetDate: DateTime for which to compute the assigned person
        # Returns: string assigned person's name, or $null if not defined

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
            $today = $this.CurrentDate  # Use CurrentDate instead of Today
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
        # This is the core method that monitors and executes scheduled tasks.
        # It's called every 5 seconds by the main timer and performs these key functions:
        #
        # 1. Mark expired tasks as completed (tasks more than 1 hour past their time)
        # 2. Check active tasks to see if any are due (within 20 seconds of current time)
        # 3. For due tasks:
        #    - Display visual alert
        #    - Speak task announcement
        #    - Log task completion
        #    - Mark task as called/completed
        #
        # Special handling for content tasks (stories, quotes, jokes):
        # - Selects unused content from respective CSV files
        # - Updates tracker files to prevent repeats
        # - Resets tracker when all content has been used
        #
        # For rotating tasks (multiple names separated by colons):
        # - Determines assigned person using GetAssignedPersonForDate
        # - Uses rotation state from task_state2.csv
        #
        # The method is optimized to:
        # - Only process active (not completed) tasks
        # - Stop checking after finding first due task
        # - Handle errors gracefully without crashing
        #
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
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
            if (-not $this.TodaysTasks) {
                Write-Host "TodaysTasks not loaded - loading now"
                $this.LoadTodaysTasks()
            }
            Write-Host "Count $($this.TodaysTasks.Count)"
    
            # First pass: mark expired tasks as completed
            foreach ($taskKey in $this.TodaysTasks.Keys) {
                try {
                    $taskInfo = $this.TodaysTasks[$taskKey]
                    if (-not $taskInfo.Completed -and -not $taskInfo.Called) {
                        $timeDiff = ($now - $taskInfo.DateTime).TotalMinutes
                        if ($timeDiff > 60) {
                            # More than 1 hour past
                            $this.TodaysTasks[$taskKey].Completed = $true
                            Write-Host "Task expired and marked passed: $taskKey"
                        }
                    }
                }
                catch {
                    Write-Host "ERROR checking task expiration for '$($taskKey)': $($_.Exception.Message)"
                    continue
                }
            }

            # Second pass: only check active tasks
            $activeTasks = $this.TodaysTasks.GetEnumerator() | Where-Object {
                -not $_.Value.Completed -and -not $_.Value.Called
            }

            foreach ($task in $activeTasks) {
                try {
                    $taskInfo = $task.Value
                    $taskKey = $task.Key
                    $task = $taskInfo.Task
                    $taskDateTime = $taskInfo.DateTime
                    $timeDiffSeconds = ($now - $taskDateTime).TotalSeconds
            
                    if ($this.DebugMode) {
                        Write-Host "Checking task '$($task.Action)' $($task.Name) at $($task.Time), diff: $timeDiffSeconds seconds"
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
                            elseif ($names.Count -eq 1) {
                                $person = $names[0]
                            }
                            else {
                                $rotationKey = "$($task.Time.ToString().Trim())|$($task.DaysOfWeek.ToString().Trim())|$($task.Action.ToString().Trim())"
                                $person = $this.GetAssignedPersonForDate($rotationKey, $taskDateTime)
                                if (-not $person) { $person = $names[0] }
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

                        # Record the actual assigned person for this task instance
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
                                
                                    # Check main timer state before starting hello timer
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
                    Write-Host "ERROR processing task key '$($taskKey)': $($_.Exception.Message)"
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
        # This method determines if a specific day of the week falls within a specified range.
        # It's used throughout the scheduler to decide if a task should be scheduled for today.
        #
        # The method supports multiple range formats for flexibility:
        #
        # 1. Full week: 'Sunday-Saturday', 'All', or '*'
        #    - Returns true for any day
        #    - Example: IsDayInRange('Wednesday', 'All') returns true
        #
        # 2. Comma-separated list: 'Monday,Wednesday,Friday'
        #    - Returns true only for explicitly listed days
        #    - Example: IsDayInRange('Tuesday', 'Monday,Wednesday,Friday') returns false
        #
        # 3. Range: 'Monday-Friday'
        #    - Returns true for days within the inclusive range
        #    - Handles wrap-around ranges (e.g., 'Friday-Monday')
        #    - Example: IsDayInRange('Thursday', 'Monday-Friday') returns true
        #
        # 4. Single day name: 'Tuesday'
        #    - Returns true only for that exact day
        #    - Example: IsDayInRange('Tuesday', 'Tuesday') returns true
        #
        # This helper centralizes day-range parsing, ensuring consistent behavior
        # throughout the scheduler and making schedule definitions more intuitive.
        #
        # The implementation uses a day-of-week index array to handle ranges efficiently,
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
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
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
                            $displayPerson = $this.GetAssignedPersonForDate($rotationKey, $dt)
                            if (-not $displayPerson) { $displayPerson = $names[0] }
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
                $this.LastTaskLabel.ForeColor = [Color]::Orange
            }
            else {
                $this.LastTaskLabel.Text = "Last: No Previous Tasks"
                $this.LastTaskLabel.ForeColor = [Color]::Gray
            }
            # Next: first task strictly after now
            $next = $todayList | Where-Object { $_.DateTime -gt $now } | Sort-Object -Property DateTime | Select-Object -First 1
            if ($next) {
                $nextPerson = $next.DisplayPerson
                $nextAction = $next.Task.Action
                $displayTime = $next.DateTime.ToString("h:mm tt")
                $this.NextTaskLabel.Text = "Next: $($nextPerson) $($nextAction) at $($displayTime)"
                $this.NextTaskLabel.ForeColor = [Color]::Lime
            }
            else {
                $this.NextTaskLabel.Text = "Next: No more tasks scheduled for today"
                $this.NextTaskLabel.ForeColor = [Color]::Gray
            }
        }
        catch {
            Write-Host "ERROR in UpdateNextTaskDisplay: $($_.Exception.Message)`n$($_.Exception.StackTrace)"
        }
    }

    [void]UpdatePersonTaskDisplays() {
        # UpdatePersonTaskDisplays
        # This method creates and manages the individual task panels that appear under the main display.
        # It's called whenever tasks change to keep the UI synchronized with the current state.
        #
        # The method performs these key functions:
        #
        # 1. Collects all tasks for today from TodaysTasks hashtable
        #    - Includes both scheduled and injected tasks
        #    - Resolves the assigned person for each task
        #    - Groups tasks by person for panel creation
        #
        # 2. Creates a visual panel for each unique person:
        #    - Panel has a header with person's name
        #    - Lists all tasks with times and descriptions
        #    - Completed tasks are shown in gray
        #    - Pending tasks are shown in white
        #
        # 3. Arranges panels in a 3-column layout:
        #    - Maximum of 2 panels per column (6 total)
        #    - Panels are sized based on number of tasks
        #    - Layout is responsive to content
        #
        # 4. Manages panel lifecycle:
        #    - Removes existing panels before creating new ones
        #    - Stores panels in PersonTaskPanels hashtable
        #    - Handles errors gracefully for individual panels
        #
        # The UI design ensures:
        # - Each child can quickly see their own schedule
        # - Parents can monitor all tasks at a glance
        # - Visual feedback shows task completion status
        # - Rotating tasks show the correct assigned person
        #
        # This method is called during initialization and whenever:
        # - Tasks are loaded or reloaded
        # - Tasks are completed
        # - Rotation assignments change
        # - Tasks are injected for testing
        try {
            # Build panels exclusively from today's task instances (TodaysTasks) so assignments match runtime state
            $now = if ($this.DebugMode) { $this.CurrentDate } else { [DateTime]::Now }
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

            # Build panels for each person            
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
                    # Add each task line
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


