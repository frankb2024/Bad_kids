
using namespace System.Windows.Forms
using namespace System.Drawing
using namespace System.Speech.Synthesis
using namespace System.Globalization
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech


    [Reflection.Assembly]::LoadFile("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF\System.Speech.dll")

. "C:\Users\frank\Desktop\badkids\bad_kids_class.ps1"
    

# Run the app
$app = $null
$app = [SchedulerScreenSaver]::new($true)



write-host "exiting"

$app = $null
[GC]::Collect()

