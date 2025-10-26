# Bad_kids
This powershell program helps parents manage their children's daily routines by providing automated, fair, and trustworthy scheduling of tasks and chores.



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
  
