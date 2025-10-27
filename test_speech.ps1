# Speech Troubleshooting Script
Write-Host "Starting Speech Troubleshooting..." -ForegroundColor Cyan

try {
    # Load the Speech assembly
    Add-Type -AssemblyName System.Speech
    Write-Host "System.Speech assembly loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to load System.Speech assembly:" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Note: System.Speech only works on Windows (not PowerShell Core on Linux/macOS)" -ForegroundColor Yellow
    exit 1
}

try {
    # Create SpeechSynthesizer
    $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
    Write-Host "SpeechSynthesizer created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create SpeechSynthesizer:" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Get available voices
$voices = $speaker.GetInstalledVoices()
Write-Host "Available voices ($($voices.Count) found):" -ForegroundColor Cyan
foreach ($voice in $voices) {
    $info = $voice.VoiceInfo
    $status = if ($voice.Enabled) { "Enabled" } else { "Disabled" }
    Write-Host "  - $($info.Name) ($status) - $($info.Culture.DisplayName)" -ForegroundColor Gray
}

if ($voices.Count -eq 0) {
    Write-Host "No voices installed! You need to install speech voices in Windows." -ForegroundColor Yellow
    Write-Host "  Go to Settings > Time & Language > Speech > Manage voices" -ForegroundColor Yellow
}

# Configure speaker
$speaker.Volume = 100  # Max volume
$speaker.Rate = 0      # Normal speed

Write-Host "`nCurrent settings:" -ForegroundColor Cyan
Write-Host "  Volume: $($speaker.Volume)%" -ForegroundColor Gray
Write-Host "  Rate: $($speaker.Rate)" -ForegroundColor Gray
Write-Host "  Selected voice: $($speaker.Voice.Name)" -ForegroundColor Gray

# Test phrases to speak
$testPhrases = @(
    "Test",
    "Speech is working",
    "This is a troubleshooting test",
    "If you can hear this, the problem is elsewhere",
    "Volume test at maximum level"
)

Write-Host "`nStarting speech tests..." -ForegroundColor Cyan
Write-Host "Make sure your speakers/headphones are:" -ForegroundColor Yellow
Write-Host "  - Connected and powered on" -ForegroundColor Yellow
Write-Host "  - Not muted" -ForegroundColor Yellow  
Write-Host "  - System volume is turned up" -ForegroundColor Yellow
Write-Host "  - Correct audio output device is selected" -ForegroundColor Yellow

for ($i = 1; $i -le 5; $i++) {
    Write-Host "`nAttempt $i of 5:" -ForegroundColor Magenta
    
    foreach ($phrase in $testPhrases) {
        Write-Host "  Speaking: '$phrase'" -ForegroundColor Blue
        try 
        {
            $speaker.Rate = 0
            $speaker.Volume = 100
            Write-Host $speaker.State
            $speaker.SelectVoice($voices[1].VoiceInfo.Name)
            $speaker.SpeakAsync($phrase)
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Host "Error speaking '$phrase': $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    if ($i -lt 5) {
        Write-Host "  Pausing 2 seconds before next attempt..." -ForegroundColor Gray
        Start-Sleep -Seconds 2
    }
}

# Additional diagnostics
Write-Host "Additional Troubleshooting Steps:" -ForegroundColor Cyan
Write-Host "1. Check Windows Speech Recognition settings:" -ForegroundColor Yellow
Write-Host "   Settings > Time & Language > Speech" -ForegroundColor Yellow
Write-Host "2. Test system audio with other applications" -ForegroundColor Yellow
Write-Host "3. Try running PowerShell as Administrator" -ForegroundColor Yellow
Write-Host "4. Ensure you're using Windows PowerShell (not PowerShell Core) on Windows" -ForegroundColor Yellow
Write-Host "5. Check if Narrator works (Windows + Ctrl + Enter)" -ForegroundColor Yellow

# Clean up
$speaker.Dispose()
Write-Host "Speech troubleshooting completed." -ForegroundColor Cyan
