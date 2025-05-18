# ucSimplePlayer v2.3.10
Simple video player UserControl/ActiveX Control

![image](https://github.com/user-attachments/assets/490b68f4-1ff7-444a-b5ed-31d10542ddc8)

This is a simple video player UserControl for VB6, twinBASIC, and VBA, supporting both 32bit and 64bit. It's just a thin wrapper over the `IMFPMediaPlayer` media player control that's part of Windows Media Foundation. While MS recommends using `IMFEngine`, that doesn't support Windows 7. 

All the basic features are covered:

- Play
- Pause
- Stop
- Volume
- Mute
- Balance
- Seek
- Playback rate
- Duration
- Fullscreen
- Loop
- Choose audio/video track

The VB6 project file and ucSimplePlayDemo.twinproj have basic players implementing the control and its functions using the control as a UserControl.
  
ucSimplePlayer.twinproj is to compile an OCX which you could then use in VB6/tB plus other hosts like VBA 32bit/64bit.

It will automatically toggle full screen when you double click the video, to disable this change `.AllowFullscreen` to `False`. You can still use the manual toggle (`.Fullscreen = True/False`),


**UPDATE - v2.3.10 (18 May 2025)**
```
'Version 2.3.10 (18 May 2025)
'-Album cover is now displayed when you play audio files; you can set
'   ShowAlbumArt to False to disable this display.
'-A default image will be shown as album art if none could be loaded from
'   the file, to disable, set UseDefaultAlbumArt to False, or to customize
'   it, use SetDefaultAlbumArt and pass a byte array of an image file that
'   is compatible with WIC.
'   Tip: You can also use this as an audio only player by setting Visible
'        to False
'-Added LoopPlayback property to automatically loop playback of the current
'   item. The PlaybackEnded and a new start event are still fired at the end
'   of each loop.
'-Added PlayerWheelScroll event. The demo app now shows how to use this
'   to adjust the volume.
'-Player now pauses/unpauses on single left click. Set AllowPauseOnClick to
'   False to disable this behavior.
'-Properties are now either hidden from the designer (settable at run
'   time only), or properly saved/loaded. Ones still visible in the designer
'   now have descriptions.
'-Added HasVideo property get. 
'-Switched CopyMemory variant hack to more proper PropVariantClear.
'-(Bug fix) Duration and playback position not working when an audio-
'           only file was played.
'-(Bug fix) Setting Paused to False did not change the status returned by
'           that property.
'-(Demo) Added FLAC to Open Dialog types.
'-(Demo) File text now also has autocomplete.
'-(Demo) Click to pause/unpause.
'-(Demo) Support for mousewheel on volume and position sliders.
```

**UPDATE - v2.2.5 (29 Mar 2025)**
```
'Version 2.2.5
'-Added ability to select different video and audio streams:
'   Use GetVideoStreams/GetAudioStreams to get the number and their 
'   names/languages, then use ActiveVideoStream/ActiveAudioStream
'   properties to set the 1-based number of the active stream.
'-Added PreserveAspectRatio property (default True)
'-Added PlayerKeyUp and PlayerClick events
'-The Demo projects show how to use the above by showing a context menu
' when the player is right clicked, allowing you to switch tracks and 
' toggle aspect ratio and fullscreen.
'-Added sub GetNativeSize to get original size of video w/o scaling
'-Added PlayTimer event to make it easy for VBA clients to synchronize
'   a progress indicator, since there's no native Timer. Control with:
'      .EnablePlayTimer 
'      .PlayTimerInterval (default 500ms)
```

**Requirements**\
Windows 7 or newer\
VB6, twinBASIC, or VBA
 
**Usage in VBA**\
VBA can only use this project as an OCX. Use twinBASIC (run it as admin or see non-admin section further down) to compile the OCX matching your Office bitness, it will automatically register. (Note: If you don't subscribe to tB, the 64bit build will have a tB splash screen when it's loaded.)

Alternatively, download the OCX matching the bitness of your MS Office version from the [Releases section](https://github.com/fafalone/ucSimplePlayer/releases), and register it with regsvr32. (There is no splash screen as I have a subscription.)
> [!TIP]
> If you don't know whether you have 32bit or 64bit Office, go to File->Account then click 'About Excel/Access/etc'

Once you've done one of the options above, ucSimplePlayer should be available in the Tools->Additional controls dialog under "Simple Video Player Control v1.1", available when you're editing a UserForm in Excel VBA, or 'ActiveX Controls' in the Access form designer-- the menu that pops up from the dropdown button on the righthand side of the built-in controls box.

Tested in MS Office Excel 2021 64bit.

![image](https://github.com/user-attachments/assets/fdd63795-5f52-48a2-9750-60b7d0f15b1f)

**Compiling without admin**\
Admin is required for the OCX for VB6 since it must install to HKLM. But for VBA/twinBASIC you don't need it: From the Project menu, open Project Settings, find the "Project: Register DLLs to HKEY_LOCAL_MACHINE" option and switch it to No. You'll no longer need admin.

**Video or audio not playing?**\
You may need additional codecs for Windows Media Foundation. These are available through the Microsoft Store or by downloading the installer directly and using PowerShell. See https://www.codecguide.com/media_foundation_codecs.htm for some common ones, or get them now from AdGuard Store, searching the given Product Id (9...)

HEVC	9N4WGH0Z6VHQ\
VP9	9n4d0msmp0pt\
AV1	9mvzqvxjbq9v\
MPEG-2	9n95q1zzpmh4\
Web media	9n5tdp8vcmhs\
HEIF image	9pmmsr1cgpwg\
Webp image	9pg2dk419drg\
Raw image	9nctdw2w1bh8\
AC-3/E-AC3	9nvjqjbdkn97\
AC-4	9p7646qph1q0

at https://store.rg-adguard.net/ which generates direct links to Microsoft Store server files. Download the .AppxBundle then install with PowerShell using `add-appxpackage –path "c:\path\to\file.appxbundle"`

The first time you use some of these codecs, if you get error 0x80070426, the "Microsoft Account Sign-in Assistant" service must be enabled (though there's no sign in or need for Store to be installed or to have a MS account/logon). You can disable it again after that first use.
