# Bad_kids
This powershell program helps parents manage their children's daily routines by providing automated, spoken, fair, and trustworthy scheduling of tasks and chores.

Why Powershell? Runs on every Windows computer.  You can read the source code and ensure nothing nefarious.
Why written as a powershell class?  It helps to unload all the code from memory by unloading the class on exit.

![alt text](https://github.com/frankb2024/Bad_kids/blob/main/bad_kids.png)


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
- When the program starts running, it immediately lowers the computer volume.
  This allows the program to run overnight without fear of Windows Updates
  or anything else making noise and waking up the family.

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
  
How to Run This Project (Step-by-Step for Beginners)
If you're new to PowerShell or GitHub, don't worry — follow these steps and you'll be up and running in no time!

1. 📥 Download the Project
Go to the GitHub page for this project.
Click the green Code button.
Select Download ZIP.
Once downloaded, unzip the file to a folder on your computer (e.g., your Desktop).

2. 🧭 Open PowerShell
Press Windows + S and type PowerShell.
Click on Windows PowerShell to open it.

3. 📂 Navigate to the Project Folder
In PowerShell, type the following command to go to the folder where you unzipped the project:

powershell
cd "C:\Users\YourName\Desktop\YourProjectFolder"
Replace the path with the actual location of your folder.

4. 📝 Customize Your Schedule
Open the file called schedule.csv using Excel or Notepad.
Add your kids' names and the schedule you want to use.
Save the file when you're done.

5. 🚀 Run the Script
In PowerShell, run the script by typing:
powershell
.\bad_kids.ps1
If you get a security warning, you may need to enable script execution. Type this first:
powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Then run the script again.

6. 🖱️ Interact and Exit
Follow any on-screen instructions.
Click anywhere on the screen to exit the program when you're done.
