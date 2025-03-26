# ucSimplePlayer v1.1.3
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

  The VB6 project file and ucSimplePlayDemo.twinproj have basic players implementing the control and its functions using the control as a UserControl.
  
   ucSimplePlayer.twinproj is to compile an OCX which you could then use in VB6/tB plus other hosts like VBA 32bit/64bit.

  It will automatically toggle full screen when you double click the video, to disable this change `.AllowFullscreen` to `False`. You can still use the manual toggle (`.Fullscreen = True/False`),

  **Requirements**\
Windows 7 or newer\
VBA can only use this project as an OCX. Use twinBASIC to compile the OCX matching your Office bitness. Tested in MS Office Excel 2021 64bit.
