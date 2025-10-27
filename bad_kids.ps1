using namespace System.Windows.Forms
using namespace System.Drawing
using namespace System.Speech.Synthesis
using namespace System.Globalization
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class CoreAudio
{
    private enum EDataFlow { eRender = 0, eCapture = 1, eAll = 2 }
    private enum ERole { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class MMDeviceEnumeratorComObject { }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        int NotImpl1();
        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        [PreserveSig]
        int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, uint dwClsCtx, IntPtr pActivationParams, out IntPtr ppInterface);
    }

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAudioEndpointVolume
    {
        // We'll declare only the methods we need
        int RegisterControlChangeNotify(IntPtr pNotify);
        int UnregisterControlChangeNotify(IntPtr pNotify);
        int GetChannelCount(out uint pnChannelCount);
        int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
        int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
        int GetMasterVolumeLevel(out float pfLevelDB);
        int GetMasterVolumeLevelScalar(out float pfLevel);
        // The interface contains more methods, but we don't need them here
    }

    private const uint CLSCTX_ALL = 23;

    public static void SetMasterVolume(float level)
    {
        var enumerator = new MMDeviceEnumeratorComObject();
        var devEnum = (IMMDeviceEnumerator)enumerator;
        IMMDevice device = null;
        IntPtr volumePtr = IntPtr.Zero;
        try
        {
            devEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device);
            Guid IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
            device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, IntPtr.Zero, out volumePtr);
            var volume = (IAudioEndpointVolume)Marshal.GetObjectForIUnknown(volumePtr);
            volume.SetMasterVolumeLevelScalar(level, Guid.Empty);
            Marshal.Release(volumePtr);
        }
        catch (Exception)
        {
            throw;
        }
        finally
        {
            if (volumePtr != IntPtr.Zero)
            {
                Marshal.Release(volumePtr);
            }
        }
    }

    public static float GetMasterVolume()
    {
        var enumerator = new MMDeviceEnumeratorComObject();
        var devEnum = (IMMDeviceEnumerator)enumerator;
        IMMDevice device = null;
        IntPtr volumePtr = IntPtr.Zero;
        try
        {
            devEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device);
            Guid IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
            device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, IntPtr.Zero, out volumePtr);
            var volume = (IAudioEndpointVolume)Marshal.GetObjectForIUnknown(volumePtr);
            float level = 0.0f;
            volume.GetMasterVolumeLevelScalar(out level);
            Marshal.Release(volumePtr);
            return level;
        }
        catch (Exception)
        {
            throw;
        }
        finally
        {
            if (volumePtr != IntPtr.Zero)
            {
                Marshal.Release(volumePtr);
            }
        }
    }
}

public class Audio
{
    [DllImport("winmm.dll")]
    private static extern int waveOutSetVolume(IntPtr hwo, uint dwVolume);

    public static void SetVolume(float level)
    {
        if (level < 0.0f || level > 1.0f)
            throw new ArgumentOutOfRangeException("level", "Volume level must be between 0.0 and 1.0");
        try
        {
            float before = CoreAudio.GetMasterVolume();
            CoreAudio.SetMasterVolume(level);
            float after = CoreAudio.GetMasterVolume();
            // If after doesn't match requested level, attempt one retry
            if (Math.Abs(after - level) > 0.02f)
            {
                System.Threading.Thread.Sleep(150);
                CoreAudio.SetMasterVolume(level);
                after = CoreAudio.GetMasterVolume();
            }
        }
        catch (Exception)
        {
            throw;
        }
    }
}
"@

[Reflection.Assembly]::LoadFile("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF\System.Speech.dll")

$fileName = "bad_kids_class.ps1"
$filePath = Join-Path -Path $PSScriptRoot -ChildPath $fileName
. $filePath

#$fileName = "test_speech.ps1"
#$filePath = Join-Path -Path $PSScriptRoot -ChildPath $fileName
#. $filePath



# Run the app
$app = $null
$app = [SchedulerScreenSaver]::new($false)

write-host "exiting"

$app = $null
[GC]::Collect()
