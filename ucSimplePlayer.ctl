VERSION 5.00
Begin VB.UserControl ucSimplePlayer 
   BackColor       =   &H80000001&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucSimplePlayer.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   480
   End
End
Attribute VB_Name = "ucSimplePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const DEBUGBLD = 0
'**********************************************************************
'ucSimplePlayer v2.4.12
'by Jon Johnson
'
'This is a simple video player control loosely based on the SimplePlay
'Windows 7 SDK example for using IMFPMediaPlayer, with most basic
'additional features implemented.
'
'While Microsoft recommends IMFMediaEngine now, it's unavailable on
'Windows 7 and obscenely complex for even the simplest playback like
'this control offers.
'
'Requirements:
'-Windows 7 or newer
'
'Version 2.4.11 (19 May 2025)
'-Changes to make VB6 version finally work.  (Thanks again to VanGoghGaming
'   for help with this; graphics not my fort�.)
'Version 2.4.11 (19 May 2025)
'-Now using AlphaBlend instead of ImageList_Draw to render album art, as
'   a workaround to the VB6 issue. (Thanks to VanGoghGaming for tip)
'-Now using SHCreateMemStream instead of CreateStreamOnHGlobal to satisfy
'   VanGoghGaming.
'-(Bug fix) Double-free bug could cause crash on exit in x64.
'
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
'
'Version 2.2.5 (29 Mar 2025)
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
'
'
'Version 1.1.3 (26 Mar 2025)
'-Added Fullscreen support. It can be done either automatically by double
' clicking the control (you can disable this with AllowFullscreen = False),
' or manually via the .Fullscreen property Get/Let.
'- Minor fixes
'Version 1.0.1 - Initial release
'**********************************************************************

Implements IMFPMediaPlayerCallback

#If TWINBASIC Then
Public Event PlaybackStart(ByVal llDuration As LongLong)
Public Event PlayTimer(ByVal llCurrentPos As LongLong, ByVal llDuration As LongLong)
#Else
Public Event PlaybackStart(ByVal cyDuration As Currency)
Public Event PlayTimer(ByVal cyCurrentPos As Currency, ByVal cyDuration As Currency)
#End If
Public Event PlaybackEnded()
Public Event PlayerKeyDown(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal Flags As Long)
Public Event PlayerKeyUp(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal Flags As Long)
Public Event PlayerClick(ByVal Button As Long)
Public Event PlayerWheelScroll(ByVal delta As Integer, ByVal keyState As Integer)
#If DEBUGBLD Then
Public Event zDebugLog(ByVal sMsg As String)
#End If
Private mFile As String
Private mPlaying As Boolean
Private mPaused As Boolean
Private mHasVideo As Boolean
Private mFullscreen As Boolean
Private mAllowFullscreen As Boolean
Private mAllowClickPause As Boolean
Private mPlayTimer As Boolean
Private mPTInterval As Long
Private dwStyleOld As WindowStyles
Private dwStyleExOld As WindowStylesEx
Private hParOld As LongPtr
Private mOldPlacement As WINDOWPLACEMENT
Private mPlayer As IMFPMediaPlayer
Private mItem As IMFPMediaItem
Private mDuration As Variant
Private mLockAR As Boolean
Private mShowCover As Boolean
Private mUseDefCover As Boolean
Private mDefCover As LongPtr
Private pFact As New WICImagingFactory
Private pDecoder As IWICBitmapDecoder
Private pFrame As IWICBitmapFrameDecode
Private pConverter As IWICFormatConverter
Private pScaler As IWICBitmapScaler
Private OnBits(0 To 31) As Long
Private mVidSize As SIZE
Private mLastPos As Variant
Private mSetPos As Variant
Private mLoop As Boolean
Private mWin10 As Boolean 'Win10+
Private mScPb As Boolean
Private Type StreamList
    nVid As Long 'Number of video streams
    idxVidActive As Long 'Active video stream index
    idxVid() As Long 'Indexes of video streames
    strVid() As String 'Names of the video streams
    idVidLang() As String 'Language IDs of video streams
    nAud As Long 'Number of audio streams
    idxAudActive As Long 'Active audio stream index
    idxAud() As Long 'Indexes of audio streames
    strAud() As String 'Names of the audio streams
    idAudLang() As String 'Language IDs of video streams
    ' nSub As Long 'Number of subtitle streams
    ' idxSubActive As Long 'Active subtitle stream index
    ' idxSub() As Long 'Indexes of subtitle streames
    ' strSub() As String 'Names of the subtitle streams
    ' idSubLang() As Long 'Language IDs of subtitle streams
End Type
Private mCurStreams As StreamList

Private Type SetStrmData
    Active As Boolean
    type As Long '1=nv change 2=na change
    lastPos As Variant
    nV As Long
    nA As Long
End Type
Private mResetData As SetStrmData


'In twinBASIC, the WinDevLib package covers both oleexp.tlb interfaces and APIs.
#If TWINBASIC = 0 Then
Private Const vbNullPtr = 0&
Private Const SIZE_MINIMIZED = 1
Private Const COLOR_WINDOW = 5
Private Const CTRUE = 1
Private Const CFALSE = 0
Private Type PAINTSTRUCT
    hDC                  As LongPtr
    fErase               As Long
    rcPaint              As RECT
    fRestore             As Long
    fIncUpdate           As Long
    rgbReserved(0 To 31) As Byte
End Type
Private Enum MonitorInfoFlags
    MONITORINFOF_PRIMARY = &H1
End Enum
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As MonitorInfoFlags
End Type
Public Enum DefaultMonitorValues
    MONITOR_DEFAULTTONULL = &H0
    MONITOR_DEFAULTTOPRIMARY = &H1
    MONITOR_DEFAULTTONEAREST = &H2
End Enum
Private Enum WINDOWPLACEMENT_FLAGS
    WPF_SETMINPOSITION = &H1
    WPF_RESTORETOMAXIMIZED = &H2
    WPF_ASYNCWINDOWPLACEMENT = &H4
End Enum
Private Type WINDOWPLACEMENT
    Length As Long
    Flags As WINDOWPLACEMENT_FLAGS
    showCmd As ShowWindow
    ptMinPosition As Point
    ptMaxPosition As Point
    rcNormalPosition As RECT
End Type
Private Enum SWP_Flags
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOOWNERZORDER = &H200
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSENDCHANGING = &H400
    
    SWP_DEFERERASE = &H2000
    SWP_ASYNCWINDOWPOS = &H4000
End Enum
Private Enum WindowZOrderDefaults
    HWND_DESKTOP = 0&
    HWND_TOP = 0&
    HWND_BOTTOM = 1&
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
End Enum
Private Const SW_RESTORE = 9
Private Enum GDI_BITMAP_COMPRESSION
    BI_RGB = 0
    BI_RLE8 = 1
    BI_RLE4 = 2
    BI_BITFIELDS = 3
    BI_JPEG = 4
    BI_PNG = 5
End Enum
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As GDI_BITMAP_COMPRESSION
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BITMAP
    BMType As Long
    BMWidth As Long
    BMHeight As Long
    BMWidthBytes As Long
    BMPlanes As Integer
    BMBitsPixel As Integer
    BMBits As LongPtr
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 0) As Long
End Type
Private Enum IL_CreateFlags
  ILC_MASK = &H1
  ILC_COLOR = &H0
  ILC_COLORDDB = &HFE
  ILC_COLOR4 = &H4
  ILC_COLOR8 = &H8
  ILC_COLOR16 = &H10
  ILC_COLOR24 = &H18
  ILC_COLOR32 = &H20
  ILC_PALETTE = &H800                  ' (no longer supported...never worked anyway)
  '5.0
  ILC_MIRROR = &H2000
  ILC_PERITEMMIRROR = &H8000&
  '6.0
  ILC_ORIGINALSIZE = &H10000
  ILC_HIGHQUALITYSCALE = &H20000
End Enum
Private Enum IMAGELISTDRAWFLAGS
    ILD_NORMAL = &H0
    ILD_TRANSPARENT = &H1
    ILD_BLEND25 = &H2
    ILD_FOCUS = &H2        'ILD_BLEND25,
    ILD_BLEND50 = &H4
    ILD_SELECTED = &H4        'ILD_BLEND50,
    ILD_BLEND = &H4        'ILD_BLEND50,
    ILD_MASK = &H10
    ILD_IMAGE = &H20
    ILD_ROP = &H40       '(WIN32_IE >= &H300)
    ILD_OVERLAYMASK = &HF00
    ILD_PRESERVEALPHA = &H1000
    ILD_SCALE = &H2000
    ILD_DPISCALE = &H4000
    ILD_ASYNC = &H8000&
End Enum
Private Enum GdiDIBitsColorUse
    DIB_RGB_COLORS = 0 '/* color table in RGBs */
    DIB_PAL_COLORS = 1 '/* color table in palette indices */
End Enum
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER  As Byte = 0
Private Const AC_SRC_ALPHA As Byte = 1
#If VBA7 Then
Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As BOOL
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As BOOL
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal dwFlags As DefaultMonitorValues) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, lpmi As Any) As BOOL
Private Declare PtrSafe Function GetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare PtrSafe Function SetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, Optional ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As SWP_Flags) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function IsEqualGUID Lib "ole32" (ByRef rguid1 As UUID, ByRef rguid2 As UUID) As BOOL
Private Declare PtrSafe Function PropVariantToStringAlloc Lib "propsys" (ByRef propvar1 As Any, ppszOut As LongPtr) As Long
Private Declare PtrSafe Function PropVariantClear Lib "ole32" (ByRef pvarg As Variant) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As BOOL
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As Any, ByVal usage As Long, ByRef ppvBits As Any, ByVal hSection As LongPtr, ByVal offset As Long) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As Any, ByVal usage As Long, ByRef ppvBits As Any, ByVal hSection As LongPtr, ByVal offset As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As BOOL
Private Declare PtrSafe Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As LongPtr, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal ftn As Long) As BOOL
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As BOOL
#Else
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As BOOL
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As BOOL
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal dwFlags As DefaultMonitorValues) As LongPtr
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, lpmi As Any) As BOOL
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, Optional ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SWP_Flags) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function IsEqualGUID Lib "ole32" (ByRef rguid1 As UUID, ByRef rguid2 As UUID) As BOOL
Private Declare Function PropVariantToStringAlloc Lib "propsys" (ByRef propvar1 As Any, ppszOut As LongPtr) As Long
Private Declare Function PropVariantClear Lib "ole32" (ByRef pvarg As Variant) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As BOOL
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As Any, ByVal usage As Long, ByRef ppvBits As Any, ByVal hSection As LongPtr, ByVal offset As Long) As LongPtr
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As BOOL
Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As LongPtr, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As LongPtr, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal ftn As Long) As BOOL
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As BOOL
#End If
#If Win64 Then
Private Const PTR_SIZE = 8
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Const PTR_SIZE = 4
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If

'WinDevLib helpers needed for VB6 only:
Private Type SignedWords
    LOWORD As Integer
    HIWORD As Integer
End Type
Private Function GET_WHEEL_DELTA_WPARAM(ByVal wParam As LongPtr) As Integer
    Dim sw As SignedWords
    CopyMemory sw, wParam, LenB(sw)
    GET_WHEEL_DELTA_WPARAM = sw.HIWORD
End Function
Private Function GET_KEYSTATE_WPARAM(ByVal wParam As LongPtr) As Integer
    GET_KEYSTATE_WPARAM = LOWORD(wParam)
End Function
Private Function SUCCEEDED(hr As Long) As Boolean
    SUCCEEDED = (hr >= 0)
End Function
Private Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String
SysReAllocStringW VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
    lPtr = 0
End If
End Function
Private Function VarTypeEx(pVar As Variant) As VARENUM
    If VarPtr(pVar) = 0 Then VarTypeEx = 0: Exit Function
    CopyMemory VarTypeEx, ByVal VarPtr(pVar), 2
End Function
Private Function VariantSetType(pVar As Variant, ByVal vt As Integer, Optional ByVal vtOnlyIf As Integer = -1) As Boolean
    If VarPtr(pVar) = 0 Then Exit Function
    If vtOnlyIf <> -1 Then
        If VarTypeEx(pVar) <> vtOnlyIf Then
            VariantSetType = False
            Exit Function
        End If
    End If
    CopyMemory pVar, vt, 2
    VariantSetType = True
End Function
Public Function VariantLPWSTRtoSTR(pVar As Variant, pOut As String) As Boolean
    Dim vt As Integer
    If VarPtr(pVar) <> 0 Then
        CopyMemory vt, ByVal VarPtr(pVar), 2
        If (vt = VT_LPWSTR) Then
            Dim lp As LongPtr
            PropVariantToStringAlloc pVar, lp
            If lp Then
                pOut = LPWSTRtoStr(lp, True)
                VariantLPWSTRtoSTR = True
            End If
        End If
    End If
End Function
Private Function CIntToUInt(ByVal Value As Integer) As Long
Const OFFSET_2 As Long = 65536
If Value < 0 Then
    CIntToUInt = Value + OFFSET_2
Else
    CIntToUInt = Value
End If
End Function
Private Function PointerAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr
    #If Win64 Then
    PointerAdd = ((Start Xor &H8000000000000000) + Incr) Xor &H8000000000000000
    #Else
    PointerAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
    #End If
End Function
Private Function LOWORD(ByVal Value As Long) As Integer
If Value And &H8000& Then
    LOWORD = Value Or &HFFFF0000
Else
    LOWORD = Value And &HFFFF&
End If
End Function
Private Function HIWORD(ByVal Value As Long) As Integer
HIWORD = (Value And &HFFFF0000) \ &H10000
End Function
Private Function MFP_POSITIONTYPE_100NS() As UUID
    Static iid As UUID
    MFP_POSITIONTYPE_100NS = iid 'GUID_NULL
End Function
Private Function MF_MT_MAJOR_TYPE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H48EBA18E, &HF8C9, &H4687, &HBF, &H11, &HA, &H74, &HC9, &HF9, &H6A, &H8F)
MF_MT_MAJOR_TYPE = iid
End Function
Private Function MFMediaType_Audio() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73647561, &H0, &H10, &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
MFMediaType_Audio = iid
End Function
Private Function MFMediaType_Video() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H73646976, &H0, &H10, &H80, &H0, &H0, &HAA, &H0, &H38, &H9B, &H71)
MFMediaType_Video = iid
End Function
Private Function MF_SD_LANGUAGE() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF2180, &HBDC2, &H423C, &HAB, &HCA, &HF5, &H3, &H59, &H3B, &HC1, &H21)
MF_SD_LANGUAGE = iid
End Function
Private Function MF_SD_PROTECTED() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAF2181, &HBDC2, &H423C, &HAB, &HCA, &HF5, &H3, &H59, &H3B, &HC1, &H21)
MF_SD_PROTECTED = iid
End Function
Private Function MF_SD_STREAM_NAME() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H4F1B099D, &HD314, &H41E5, &HA7, &H81, &H7F, &HEF, &HAA, &H4C, &H50, &H1F)
MF_SD_STREAM_NAME = iid
End Function
Private Function GUID_WICPixelFormat32bppBGRA() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &HF)
GUID_WICPixelFormat32bppBGRA = iid
End Function
Private Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
End Sub
Private Function UUID_NULL() As UUID
Static bSet As Boolean
Static iid As UUID
If bSet = False Then
  With iid
    .Data1 = 0: .Data2 = 0: .Data3 = 0
    .Data4(0) = 0: .Data4(1) = 0: .Data4(2) = 0: .Data4(3) = 0: .Data4(4) = 0: .Data4(5) = 0: .Data4(6) = 0: .Data4(7) = 0
  End With
End If
bSet = True
UUID_NULL = iid
End Function
Public Function PKEY_ThumbnailStream() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 27)
PKEY_ThumbnailStream = pkk
End Function
Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
  Name.pid = pid
End Sub
#End If
'******************************************************
'END VB ONLY

Private Sub DebugLog(sMsg As String)
    #If DEBUGBLD Then
    Debug.Print sMsg
    RaiseEvent zDebugLog(sMsg)
    #End If
End Sub

Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
    Dim dwMajor As Long
    CopyMemory dwMajor, ByVal &H7FFE026C, 4
    If dwMajor >= 10 Then mWin10 = True
End Sub
Private Sub UserControl_Terminate() 'Handles UserControl.Terminate
    If mScPb Then
        UnSubclass2 Picture1.hwnd, AddressOf ucSimplePlayerHelperProc, Picture1.hwnd
    End If
    If (mPlayer Is Nothing) = False Then
        mPlayer.Shutdown
    End If
    If (pScaler Is Nothing) = False Then Set pScaler = Nothing
    If (pConverter Is Nothing) = False Then Set pConverter = Nothing
    If (pFrame Is Nothing) = False Then Set pFrame = Nothing
    If (pDecoder Is Nothing) = False Then Set pDecoder = Nothing
    If (pFact Is Nothing) = False Then Set pFact = Nothing
    If mDefCover Then
        GlobalFree mDefCover
        mDefCover = 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) 'Handles UserControl.ReadProperties
    mAllowFullscreen = PropBag.ReadProperty("AllowFullscreen", True)
    mAllowClickPause = PropBag.ReadProperty("AllowPauseOnClick", True)
    mLockAR = PropBag.ReadProperty("PreserveAspectRatio", True)
    mShowCover = PropBag.ReadProperty("ShowAlbumArt", True)
    mUseDefCover = PropBag.ReadProperty("UseDefaultAlbumArt", True)
    mPlayTimer = PropBag.ReadProperty("EnablePlayTimer", False)
    mLoop = PropBag.ReadProperty("LoopPlayback", False)
    mPTInterval = PropBag.ReadProperty("PlayTimerInterval", 500)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag) 'Handles UserControl.WriteProperties
    PropBag.WriteProperty "AllowFullscreen", mAllowFullscreen, True
    PropBag.WriteProperty "AllowPauseOnClick", mAllowClickPause, True
    PropBag.WriteProperty "PreserveAspectRatio", mLockAR, True
    PropBag.WriteProperty "ShowAlbumArt", mShowCover, True
    PropBag.WriteProperty "UseDefaultAlbumArt", mUseDefCover, True
    PropBag.WriteProperty "EnablePlayTimer", mPlayTimer, False
    PropBag.WriteProperty "LoopPlayback", mLoop, False
    PropBag.WriteProperty "PlayTimerInterval", mPTInterval, 500
End Sub

Private Sub UserControl_InitProperties() 'Handles UserControl.InitProperties
    mAllowFullscreen = True
    mAllowClickPause = True
    mLockAR = True
    mShowCover = True
    mUseDefCover = True
    mPTInterval = 500
End Sub

Private Sub UserControl_Show() 'Handles UserControl.Show
    If Ambient.UserMode Then
        Subclass2 UserControl.hwnd, AddressOf ucSimplePlayerHelperProc, UserControl.hwnd, ObjPtr(Me)
 
    End If
End Sub

Public Property Get Paused() As Boolean
Attribute Paused.VB_MemberFlags = "400"
    Paused = mPaused
End Property
Public Property Let Paused(ByVal fPaused As Boolean)
    If mPlayer Is Nothing Then Exit Property
 
    If fPaused Then
        mPaused = True
        mPlayer.Pause
    Else
        mPaused = False
        mPlayer.Play
    End If
End Property
Public Property Get FileName() As String
Attribute FileName.VB_MemberFlags = "400"
    FileName = mFile
End Property
Public Property Let FileName(ByVal sFullPath As String)
    mFile = sFullPath
    PlayMediaFile sFullPath
End Property
Public Sub StopPlayback()
    If mPlayer Is Nothing Then Exit Sub
    mPlayer.Stop
End Sub
Public Sub PlayMediaFile(ByVal sFullPath As String)
    On Error GoTo e0
    If mPlayer Is Nothing Then
        Dim hr As Long
        hr = MFPCreateMediaPlayer(0, CFALSE, 0, Me, UserControl.hwnd, mPlayer)
    End If
    If hr < 0 Then
        DebugLog "Failed to create media player, 0x" & Hex$(hr)
        Exit Sub
    End If
    mPlayer.CreateMediaItemFromURL StrPtr(sFullPath), CFALSE, 0, Nothing
    Exit Sub
e0:
    DebugLog "PlayMediaFile error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Sub FrameStep()
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.FrameStep
    End If
    Exit Sub
e0:
    DebugLog "FrameStep error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Property Get Volume() As Single
Attribute Volume.VB_Description = "Adjusts the volume of audio."
Attribute Volume.VB_MemberFlags = "400"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetVolume Volume
    End If
    Exit Property
e0:
    DebugLog "Volume.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Volume(ByVal fVol As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetVolume fVol
    End If
    Exit Property
e0:
    DebugLog "Volume.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get Balance() As Single
Attribute Balance.VB_Description = "Adjusts the left/right balance of audio."
Attribute Balance.VB_MemberFlags = "400"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetBalance Balance
    End If
    Exit Property
e0:
    DebugLog "Balance.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Balance(ByVal fBal As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetBalance fBal
    End If
    Exit Property
e0:
    DebugLog "Balance.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get Muted() As Boolean
Attribute Muted.VB_Description = "Mute or unmute the audio."
Attribute Muted.VB_MemberFlags = "400"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim fMute As BOOL
        mPlayer.GetMute fMute
        Muted = (fMute = CTRUE)
    End If
    Exit Property
e0:
    DebugLog "Muted.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property

Public Property Let Muted(ByVal bMute As Boolean)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        If bMute Then
            mPlayer.SetMute CTRUE
        Else
            mPlayer.SetMute CFALSE
        End If
    End If
    Exit Property
e0:
    DebugLog "Muted.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get Duration() As Currency
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim pv As Variant
        mPlayer.GetDuration MFP_POSITIONTYPE_100NS, pv
        VariantSetType pv, VT_CY, VT_I8
        Duration = CCur(pv)
    End If
    Exit Property
e0:
    DebugLog "Duration.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get Position() As Currency
Attribute Position.VB_MemberFlags = "400"
    ' Debuglog "Enter Position.Get"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim pv As Variant
        mPlayer.GetPosition MFP_POSITIONTYPE_100NS, pv
        VariantSetType pv, VT_CY, VT_I8
        Position = CCur(pv)
    End If
    Exit Property
e0:
    DebugLog "Position.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Position(ByVal cyPos As Currency)
    ' DebugLog "Enter Position.Let " & cyPos
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim pv As Variant
        pv = cyPos
        VariantSetType pv, VT_I8, VT_CY
        mPlayer.SetPosition MFP_POSITIONTYPE_100NS, pv
    End If
    Exit Property
e0:
    DebugLog "Position.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Sub GetNativeVideoSize(pcx As Long, pcy As Long, pcxAspectRatio As Long, pcyAspectRatio As Long)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim sz As SIZE
        Dim szar As SIZE
        mPlayer.GetNativeVideoSize sz, szar
        pcx = sz.cx: pcy = sz.cy
        pcxAspectRatio = szar.cx: pcyAspectRatio = szar.cy
    End If
    Exit Sub
e0:
    DebugLog "GetNativeVideoSize error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Property Get HasVideo() As Boolean
    HasVideo = mHasVideo
End Property
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_MemberFlags = "400"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim pclr As Long
        mPlayer.GetBorderColor pclr
        BorderColor = OleTranslateColor(pclr, 0)
    End If
    Exit Property
e0:
    DebugLog "BorderColor.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let BorderColor(ByVal clr As OLE_COLOR)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        Dim pclr As Long
        pclr = OleTranslateColor(clr, 0)
        mPlayer.SetBorderColor pclr
    End If
    Exit Property
e0:
    DebugLog "BorderColor.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Sub GetSupportedRates(ByVal bForward As Boolean, pfRateMin As Single, pfRateMax As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetSupportedRates IIf(bForward, CTRUE, CFALSE), pfRateMin, pfRateMax
    End If
    Exit Sub
e0:
    DebugLog "GetSupportedRates error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Property Get Rate() As Single
Attribute Rate.VB_MemberFlags = "400"
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetRate Rate
    End If
    Exit Property
e0:
    DebugLog "Rate.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Rate(ByVal fRate As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetRate fRate
    End If
    Exit Property
e0:
    DebugLog "Rate.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get ShowAlbumArt() As Boolean
Attribute ShowAlbumArt.VB_Description = "Sets whether the embedded album art of an audio file like mp3 is displayed."
    ShowAlbumArt = mShowCover
End Property
Public Property Let ShowAlbumArt(bShow As Boolean)
    If bShow = False Then
        Picture1.Visible = False 'hide the cover if it's currently shown and we're turning off
    End If
    mShowCover = bShow
End Property
Public Property Get UseDefaultAlbumArt() As Boolean
Attribute UseDefaultAlbumArt.VB_Description = "Sets whether a default image is displayed if an audio file has no embedded album art but ShowAlbumArt is True."
    UseDefaultAlbumArt = mUseDefCover
End Property
Public Property Let UseDefaultAlbumArt(ByVal bValue As Boolean)
    UseDefaultAlbumArt = bValue
End Property
Public Function SetDefaultAlbumArt(ImageFileBytes() As Byte) As Long
    On Error GoTo e0
    If mDefCover Then
        GlobalFree mDefCover
        mDefCover = 0
        mDefCover = GlobalAlloc(GPTR, UBound(ImageFileBytes) + 1)
    End If
    If mDefCover Then
        CopyMemory ByVal mDefCover, ImageFileBytes(0), UBound(ImageFileBytes) + 1
    End If
    Exit Function
e0:
    SetDefaultAlbumArt = Err.Number
    DebugLog "SetDefaultAlbumArt error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Function
Public Property Get AllowFullscreen() As Boolean
Attribute AllowFullscreen.VB_Description = "Sets whether the player automatically enters fullscreen mode when double-clicked."
    AllowFullscreen = mAllowFullscreen
End Property
Public Property Let AllowFullscreen(bAllow As Boolean)
    mAllowFullscreen = bAllow
End Property
Public Property Get AllowPauseOnClick() As Boolean
Attribute AllowPauseOnClick.VB_Description = "Sets whether the player automatically pauses/unpauses when clicked."
    AllowPauseOnClick = mAllowClickPause
End Property
Public Property Let AllowPauseOnClick(bAllow As Boolean)
    mAllowClickPause = bAllow
End Property
Public Property Get Fullscreen() As Boolean
Attribute Fullscreen.VB_MemberFlags = "400"
    Fullscreen = mFullscreen
End Property
Public Property Let Fullscreen(bFullscreen As Boolean)
    If bFullscreen Then
        If mFullscreen Then Exit Property 'already FS
        If EnterFullscreen() Then
            mFullscreen = True
        End If
    Else
        If mFullscreen = False Then Exit Property 'already windowed
        ExitFullscreen
        mFullscreen = False
    End If
End Property
Public Property Get EnablePlayTimer() As Boolean
Attribute EnablePlayTimer.VB_Description = "Sets whether a timer periodically fires to sync playback. Intended for VBA where there's no built in timer."
    EnablePlayTimer = mPlayTimer
End Property
Public Property Let EnablePlayTimer(ByVal bEnable As Boolean)
    mPlayTimer = bEnable
    If (mPlaying = True) And (bEnable = True) Then
        Timer1.Interval = mPTInterval
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Property
Public Property Get PlayTimerInterval() As Long
Attribute PlayTimerInterval.VB_Description = "Sets the interval for the play timer, in milliseconds."
    EnablePlayTimer = mPTInterval
End Property
Public Property Let PlayTimeInterval(ByVal lInterval As Long)
    mPTInterval = lInterval
End Property
Public Property Get LoopPlayback() As Boolean
Attribute LoopPlayback.VB_Description = "Automatically repeat playback of the current file when it reaches the end."
    LoopPlayback = mLoop
End Property
Public Property Let LoopPlayback(ByVal bValue As Boolean)
    mLoop = bValue
End Property
Public Property Get PreserveAspectRatio() As Boolean
Attribute PreserveAspectRatio.VB_Description = "Sets whether the ratio of width to height is preserved when video is resized to fit the window or monitor."
    PreserveAspectRatio = mLockAR
End Property
Public Property Let PreserveAspectRatio(ByVal bPreserve As Boolean)
    If bPreserve = mLockAR Then Exit Property
    On Error Resume Next
    mLockAR = bPreserve
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetAspectRatioMode IIf(mLockAR, MFVideoARMode_PreservePicture, MFVideoARMode_None)
        If Picture1.Visible = True Then
            Dim rcCur As RECT
            GetClientRect UserControl.hwnd, rcCur
            Picture1.Width = rcCur.Right
            Picture1.Height = rcCur.Bottom
            DrawWicImage
        End If
    End If
End Property
Public Sub GetNativeSize(pcx As Long, pcy As Long)
    pcx = mVidSize.cx
    pcy = mVidSize.cy
End Sub
Public Sub GetVideoStreams(pName() As String, pLang() As String, nStrm As Long)
    nStrm = mCurStreams.nVid
    If nStrm > 0 Then
        pName = mCurStreams.strVid
        pLang = mCurStreams.idVidLang
    End If
End Sub
Public Sub GetAudioStreams(pName() As String, pLang() As String, nStrm As Long)
    nStrm = mCurStreams.nAud
    If nStrm > 0 Then
        pName = mCurStreams.strAud
        pLang = mCurStreams.idAudLang
    End If
End Sub
Public Property Get ActiveVideoStream() As Long
Attribute ActiveVideoStream.VB_MemberFlags = "400"
    ActiveVideoStream = mCurStreams.idxVidActive
End Property
Public Property Let ActiveVideoStream(ByVal nStrm As Long)
    On Error GoTo e0
    If mPlayer Is Nothing Then Exit Property
    If mCurStreams.nVid = 0 Then Exit Property
    
    If nStrm = mCurStreams.idxVidActive Then Exit Property
    
    Dim pItem As IMFPMediaItem
    mPlayer.GetMediaItem pItem
    If (pItem Is Nothing) = False Then
        Dim i As Long
        For i = 0 To UBound(mCurStreams.idxVid)
            If i = (nStrm - 1) Then
                pItem.SetStreamSelection mCurStreams.idxVid(i), CTRUE
            Else
                pItem.SetStreamSelection mCurStreams.idxVid(i), CFALSE
            End If
        Next
        mResetData.Active = True
        mResetData.nV = nStrm
        mResetData.nA = mCurStreams.idxVidActive
        mPlayer.GetPosition MFP_POSITIONTYPE_100NS, mResetData.lastPos
        mPlayer.SetMediaItem pItem
         
    End If
    Exit Property
e0:
    DebugLog "ActiveVideoStream.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get ActiveAudioStream() As Long
Attribute ActiveAudioStream.VB_MemberFlags = "400"
    ActiveAudioStream = mCurStreams.idxAudActive
End Property
Public Property Let ActiveAudioStream(ByVal nStrm As Long)
    On Error GoTo e0
    If mPlayer Is Nothing Then Exit Property
    If mCurStreams.nAud = 0 Then Exit Property
    
    If nStrm = mCurStreams.idxAudActive Then Exit Property
    
    Dim pItem As IMFPMediaItem
    mPlayer.GetMediaItem pItem
    If (pItem Is Nothing) = False Then
        Dim i As Long
        For i = 0 To UBound(mCurStreams.idxAud)
            If i = (nStrm - 1) Then
                pItem.SetStreamSelection mCurStreams.idxAud(i), CTRUE
            Else
                pItem.SetStreamSelection mCurStreams.idxAud(i), CFALSE
            End If
        Next
        mResetData.Active = True
        mResetData.nV = mCurStreams.idxVidActive
        mResetData.nA = nStrm
        mPlayer.GetPosition MFP_POSITIONTYPE_100NS, mResetData.lastPos
        mPlayer.SetMediaItem pItem
    End If
    Exit Property
e0:
    DebugLog "ActiveAudioStream.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
 
#If TWINBASIC Then
Private Sub IMFPMediaPlayerCallback_OnMediaPlayerEvent(pEventHeader As MFP_EVENT_HEADER) 'Implements IMFPMediaPlayerCallback.OnMediaPlayerEvent
    If pEventHeader.hrEvent < 0 Then
        DebugLog "Playback error, 0x" & Hex$(pEventHeader.hrEvent) & "; EventType=" & pEventHeader.eEventType
        'Note: If this error is 0x80070426, you must enable the "Microsoft Account Sign-in Assistant" service
        '      the first time you use a newly installed Microsoft Store codec.
        Exit Sub
    End If
    
    Select Case pEventHeader.eEventType
        Case MFP_EVENT_TYPE_MEDIAITEM_CREATED
            OnMediaItemCreated MFP_GET_MEDIAITEM_CREATED_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_MEDIAITEM_SET
            OnMediaItemSet MFP_GET_MEDIAITEM_SET_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_PLAYBACK_ENDED
            RaiseEvent PlaybackEnded
            If mLoop Then
                mPlayer.Play
            Else
                Timer1.Enabled = False
            End If
             
    End Select
End Sub

#Else
Private Sub IMFPMediaPlayerCallback_OnMediaPlayerEvent(pEventHeader As MFP_EVENT_HEADER) 'Implements IMFPMediaPlayerCallback.OnMediaPlayerEvent
    If pEventHeader.hrEvent < 0 Then
        DebugLog "Playback error, 0x" & Hex$(pEventHeader.hrEvent) & "; EventType=" & pEventHeader.eEventType
        'Note: If this error is 0x80070426, you must enable the "Microsoft Account Sign-in Assistant" service
        '      the first time you use a newly installed Microsoft Store codec.
        Exit Sub

    End If
    
    Select Case pEventHeader.eEventType
        Case MFP_EVENT_TYPE_MEDIAITEM_CREATED
            OnMediaItemCreated MFP_GET_MEDIAITEM_CREATED_EVENT(VarPtr(pEventHeader))
 
        Case MFP_EVENT_TYPE_MEDIAITEM_SET
            OnMediaItemSet MFP_GET_MEDIAITEM_SET_EVENT(VarPtr(pEventHeader))
 
        Case MFP_EVENT_TYPE_PLAYBACK_ENDED
            RaiseEvent PlaybackEnded
            If mLoop Then
                mPlayer.Play
            Else
                Timer1.Enabled = False
            End If
    End Select
End Sub
Private Property Get MFP_GET_MEDIAITEM_CREATED_EVENT_HELPER(MFP_MEDIAITEM_CREATED_EVENT As MFP_MEDIAITEM_CREATED_EVENT, ByVal pEventHeader As LongPtr) As MFP_MEDIAITEM_CREATED_EVENT
    CopyMemory ByVal VarPtr(pEventHeader) - PTR_SIZE, pEventHeader, PTR_SIZE
    MFP_GET_MEDIAITEM_CREATED_EVENT_HELPER = MFP_MEDIAITEM_CREATED_EVENT
End Property

Private Property Get MFP_GET_MEDIAITEM_SET_EVENT_HELPER(MFP_MEDIAITEM_SET_EVENT As MFP_MEDIAITEM_SET_EVENT, ByVal pEventHeader As LongPtr) As MFP_MEDIAITEM_SET_EVENT
    CopyMemory ByVal VarPtr(pEventHeader) - PTR_SIZE, pEventHeader, PTR_SIZE
    MFP_GET_MEDIAITEM_SET_EVENT_HELPER = MFP_MEDIAITEM_SET_EVENT
End Property
Private Property Get MFP_GET_EVENT_HEADER(MFP_EVENT_HEADER As MFP_EVENT_HEADER, ByVal pEventHeader As LongPtr) As MFP_EVENT_HEADER
    CopyMemory ByVal VarPtr(pEventHeader) - PTR_SIZE, pEventHeader, PTR_SIZE
    MFP_GET_EVENT_HEADER = MFP_EVENT_HEADER
End Property
Private Function MFP_GET_MEDIAITEM_CREATED_EVENT(pEventHeader As LongPtr) As MFP_MEDIAITEM_CREATED_EVENT
    If MFP_GET_EVENT_HEADER(MFP_GET_MEDIAITEM_CREATED_EVENT.Header, pEventHeader).eEventType = MFP_EVENT_TYPE_MEDIAITEM_CREATED Then
        MFP_GET_MEDIAITEM_CREATED_EVENT = MFP_GET_MEDIAITEM_CREATED_EVENT_HELPER(MFP_GET_MEDIAITEM_CREATED_EVENT, pEventHeader)
    End If
End Function

Private Function MFP_GET_MEDIAITEM_SET_EVENT(pEventHeader As LongPtr) As MFP_MEDIAITEM_SET_EVENT
    If MFP_GET_EVENT_HEADER(MFP_GET_MEDIAITEM_SET_EVENT.Header, pEventHeader).eEventType = MFP_EVENT_TYPE_MEDIAITEM_SET Then
        MFP_GET_MEDIAITEM_SET_EVENT = MFP_GET_MEDIAITEM_SET_EVENT_HELPER(MFP_GET_MEDIAITEM_SET_EVENT, pEventHeader)
    End If
End Function
 
#End If

 
Private Sub OnMediaItemCreated(pEvent As MFP_MEDIAITEM_CREATED_EVENT)
    DebugLog "OnMediaItemCreated"
    If (mPlayer Is Nothing) = False Then
        Dim hr As Long
        Dim bHasVideo As BOOL, bIsSelected As BOOL
        pEvent.pMediaItem.HasVideo bHasVideo, bIsSelected
        DebugLog "OnMediaItemCreated bHasVideo=" & bHasVideo & ", bIsSelected=" & bIsSelected
        mPlayer.SetMediaItem pEvent.pMediaItem
        mHasVideo = (bHasVideo = CTRUE)
    End If
End Sub
Private Sub OnMediaItemSet(pEvent As MFP_MEDIAITEM_SET_EVENT)
    DebugLog "OnMediaItemSet"
    On Error GoTo e0
    MapStreams pEvent.pMediaItem
    mPlayer.Play
    mPlaying = True
    On Error Resume Next
    If mPlayTimer Then
        Timer1.Interval = mPTInterval
        Timer1.Enabled = True
    End If
    'Can't get duration prior to this
    If mResetData.Active Then
        mResetData.Active = False
        mCurStreams.idxVidActive = mResetData.nV
        mCurStreams.idxAudActive = mResetData.nA
        DebugLog "mResetData.Active"
        mPlayer.SetPosition MFP_POSITIONTYPE_100NS, mResetData.lastPos
    End If
    Dim dr As Currency
    Dim pv As Variant
    mPlayer.GetDuration MFP_POSITIONTYPE_100NS, mDuration
    If mHasVideo Then
        If Picture1.Visible = True Then
            Picture1.Visible = False
            UnSubclass2 Picture1.hwnd, AddressOf ucSimplePlayerHelperProc, Picture1.hwnd
            mScPb = False
        End If
        Dim sz As SIZE
        mPlayer.GetNativeVideoSize mVidSize, sz
    Else
        If mShowCover Then
            'This is a bad hack to display album art.
            'Even though I blocked off all the WM_PAINT handling, it would
            'still glitch out and not draw, even if AutoRedraw was changed.
            Picture1.Visible = True
            Picture1.BackColor = UserControl.BackColor
            Set Picture1.Picture = LoadPicture
            Dim rcCur As RECT
            GetClientRect UserControl.hwnd, rcCur
            ' Picture1.Width = UserControl.ScaleWidth
            ' Picture1.Height = UserControl.ScaleHeight
            Picture1.Width = rcCur.Right
            Picture1.Height = rcCur.Bottom
            DebugLog "OnSet " & UserControl.ScaleWidth & ", " & UserControl.ScaleHeight & ", " & rcCur.Right & ", " & rcCur.Bottom & ", " & Picture1.Width & ", " & Picture1.Height
            Picture1.Cls
            LoadAlbumArt pEvent.pMediaItem
            mScPb = True
            Subclass2 Picture1.hwnd, AddressOf ucSimplePlayerHelperProc, Picture1.hwnd, ObjPtr(Me)
        End If
    End If
    #If TWINBASIC Then
    RaiseEvent PlaybackStart(mDuration)
    #Else
     VariantSetType mDuration, VT_CY, VT_I8
     dr = CCur(mDuration)
     RaiseEvent PlaybackStart(dr)
    #End If
    Exit Sub
e0:
    DebugLog "OnMediaItemSet Error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub

Private Sub MapStreams(ByVal pItem As IMFPMediaItem)
    On Error GoTo e0
    Dim nStm As Long
    Dim i As Long
    Dim pvType As Variant
    Dim pvName As Variant
    Dim pvLang As Variant
    Dim gid As UUID
    Dim ptr As LongPtr
    Dim sTmp As String
    Dim n As Long
    
    If pItem Is Nothing Then Exit Sub
    Erase mCurStreams.idxVid
    Erase mCurStreams.idxAud
    ' Erase mCurStreams.idxSub
    Erase mCurStreams.idVidLang
    Erase mCurStreams.idAudLang
    ' Erase mCurStreams.idSubLang
    Erase mCurStreams.strVid
    Erase mCurStreams.strAud
    ' Erase mCurStreams.strSub
    mCurStreams.nVid = 0
    mCurStreams.nAud = 0
    ' mCurStreams.nSub = 0
    
    mCurStreams.idxVidActive = 0
    mCurStreams.idxAudActive = 0
    
    pItem.GetNumberOfStreams nStm
    If nStm = 0 Then Exit Sub
    
    On Error Resume Next
    For i = 0 To nStm - 1
        n = 0
        pItem.GetStreamAttribute i, MF_MT_MAJOR_TYPE, pvType
        If SUCCEEDED(Err.LastHResult) Then 'really bad form but it works out.
            ' DebugLog "VarTypeT(" & i & ")=" & VarTypeEx(pvType)
            If VarTypeEx(pvType) = VT_CLSID Then
                CopyMemory ptr, ByVal PointerAdd(VarPtr(pvType), 8), LenB(ptr)
                CopyMemory gid, ByVal ptr, LenB(gid)
                DebugLog "Stream type guid " & i & " =" & dbg_GUIDToString(gid)
                If IsEqualGUID(gid, MFMediaType_Video) Then
                    mCurStreams.nVid = mCurStreams.nVid + 1
                    DebugLog "Detected video stream " & mCurStreams.nVid
                    ReDim Preserve mCurStreams.idxVid(mCurStreams.nVid - 1)
                    ReDim Preserve mCurStreams.idVidLang(mCurStreams.nVid - 1)
                    ReDim Preserve mCurStreams.strVid(mCurStreams.nVid - 1)
                    mCurStreams.idxVid(mCurStreams.nVid - 1) = i
                    mCurStreams.idxVidActive = 1
                    n = 1
                ElseIf IsEqualGUID(gid, MFMediaType_Audio) Then
                    mCurStreams.nAud = mCurStreams.nAud + 1
                    DebugLog "Detected audio stream " & mCurStreams.nAud
                    ReDim Preserve mCurStreams.idxAud(mCurStreams.nAud - 1)
                    ReDim Preserve mCurStreams.idAudLang(mCurStreams.nAud - 1)
                    ReDim Preserve mCurStreams.strAud(mCurStreams.nAud - 1)
                    mCurStreams.idxAud(mCurStreams.nAud - 1) = i
                    mCurStreams.idxAudActive = 1
                    n = 2
                End If
            End If
        Else
            DebugLog "Stream " & i & " MF_MT_MAJOR_TYPE error 0x" & Hex$(Err.Number)
        End If
        pItem.GetStreamAttribute i, MF_SD_STREAM_NAME, pvName
        If SUCCEEDED(Err.LastHResult) Then
            ' DebugLog "VarTypeN(" & i & ")=" & VarTypeEx(pvName)
            If VarTypeEx(pvName) = VT_LPWSTR Then
                VariantLPWSTRtoSTR pvName, sTmp
                DebugLog "VT_LPWSTR=" & sTmp
                If n = 1 Then
                    mCurStreams.strVid(mCurStreams.nVid - 1) = sTmp
                ElseIf n = 2 Then
                    mCurStreams.strAud(mCurStreams.nAud - 1) = sTmp
                End If
                sTmp = ""
            End If
        Else
            DebugLog "Stream " & i & " MF_SD_STREAM_NAME error 0x" & Hex$(Err.Number)
        End If
        pItem.GetStreamAttribute i, MF_SD_LANGUAGE, pvLang
        If SUCCEEDED(Err.LastHResult) Then
            ' DebugLog "VarTypeL(" & i & ")=" & VarTypeEx(pvLang)
            If VarTypeEx(pvLang) = VT_LPWSTR Then
                VariantLPWSTRtoSTR pvLang, sTmp
                DebugLog "VT_LPWSTR=" & sTmp
                If n = 1 Then
                    mCurStreams.idVidLang(mCurStreams.nVid - 1) = sTmp
                ElseIf n = 2 Then
                    mCurStreams.idAudLang(mCurStreams.nAud - 1) = sTmp
                End If
                sTmp = ""
            End If
        Else
            DebugLog "Stream " & i & " MF_SD_LANGUAGE error 0x" & Hex$(Err.Number)
        End If
        PropVariantClear pvType
        PropVariantClear pvName
        PropVariantClear pvLang
    Next
Exit Sub
e0:
    DebugLog "MapStreams Error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Private Function dbg_GUIDToString(tg As UUID, Optional bBrack As Boolean = True) As String
'StringFromGUID2 never works, even "working" code from vbaccelerator AND MSDN
dbg_GUIDToString = Right$("00000000" & Hex$(tg.Data1), 8) & "-" & Right$("0000" & Hex$(tg.Data2), 4) & "-" & Right$("0000" & Hex$(tg.Data3), 4) & _
"-" & Right$("00" & Hex$(CLng(tg.Data4(0))), 2) & Right$("00" & Hex$(CLng(tg.Data4(1))), 2) & "-" & Right$("00" & Hex$(CLng(tg.Data4(2))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(3))), 2) & Right$("00" & Hex$(CLng(tg.Data4(4))), 2) & Right$("00" & Hex$(CLng(tg.Data4(5))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(6))), 2) & Right$("00" & Hex$(CLng(tg.Data4(7))), 2)
If bBrack Then dbg_GUIDToString = "{" & dbg_GUIDToString & "}"
End Function
 
Private Sub OnSize(ByVal hwnd As LongPtr, ByVal state As Long, ByVal cx As Long, ByVal cy As Long)
    If state <> SIZE_MINIMIZED Then
        If (mPlayer Is Nothing) = False Then
            mPlayer.UpdateVideo
        End If
    End If
End Sub

Private Sub OnPaint(ByVal hwnd As LongPtr)
    Dim ps As PAINTSTRUCT
    Dim hDC As LongPtr
    hDC = BeginPaint(hwnd, ps)
    If ((mPlayer Is Nothing) = False) And (mHasVideo = True) Then
        On Error Resume Next
        mPlayer.UpdateVideo
    Else
        FillRect hDC, ps.rcPaint, COLOR_WINDOW + 1
    End If
    EndPaint hwnd, ps
End Sub

Private Function EnterFullscreen() As Long
    Dim hMon As LongPtr
    Dim mi As MONITORINFO
    hMon = MonitorFromWindow(UserControl.hwnd, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = LenB(mi)
    If hMon = 0 Then Exit Function
    GetMonitorInfo hMon, mi
        
    GetWindowPlacement UserControl.hwnd, mOldPlacement
    dwStyleOld = CLng(GetWindowLongPtr(UserControl.hwnd, GWL_STYLE))
    dwStyleExOld = CLng(GetWindowLongPtr(UserControl.hwnd, GWL_EXSTYLE))
    hParOld = GetParent(UserControl.hwnd)
    SetParent UserControl.hwnd, GetDesktopWindow()
    SetWindowLongPtr UserControl.hwnd, GWL_STYLE, WS_POPUP Or WS_VISIBLE
    SetWindowPos UserControl.hwnd, HWND_TOP, mi.rcMonitor.Left, mi.rcMonitor.Top, _
                        mi.rcMonitor.Right - mi.rcMonitor.Left, _
                        mi.rcMonitor.Bottom - mi.rcMonitor.Top, SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
    If mHasVideo = False Then
        Dim rcCur As RECT
        GetClientRect UserControl.hwnd, rcCur
        Picture1.Width = rcCur.Right
        Picture1.Height = rcCur.Bottom
        DrawWicImage
    End If
    EnterFullscreen = 1
End Function
 
Private Function ExitFullscreen() As Long
    ShowWindow UserControl.hwnd, SW_RESTORE
    SetWindowLongPtr UserControl.hwnd, GWL_EXSTYLE, dwStyleExOld
    SetWindowLongPtr UserControl.hwnd, GWL_STYLE, dwStyleOld
    SetWindowPos UserControl.hwnd, HWND_NOTOPMOST, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, SWP_SHOWWINDOW
    SetParent UserControl.hwnd, hParOld
    SetWindowPlacement UserControl.hwnd, mOldPlacement
    UserControl.Width = UserControl.Width - 1
    UserControl.Width = UserControl.Width + 1
    If mHasVideo = False Then
        Dim rcCur As RECT
        GetClientRect UserControl.hwnd, rcCur
        Picture1.Width = rcCur.Right
        Picture1.Height = rcCur.Bottom
        DrawWicImage
    End If
    ExitFullscreen = 1
End Function

Private Sub Timer1_Timer() 'Handles Timer1.Timer
    Dim pv As Variant
    mPlayer.GetPosition MFP_POSITIONTYPE_100NS, pv
    #If TWINBASIC Then
    RaiseEvent PlayTimer(pv, mDuration)
    #Else
    VariantSetType pv, VT_CY, VT_I8
    RaiseEvent PlayTimer(CCur(pv), CCur(mDuration))
    #End If
End Sub

Private Function LoadAlbumArt(ByVal pItem As IMFPMediaItem) As Long
    On Error Resume Next
    If mHasVideo = True Then LoadAlbumArt = S_FALSE: Exit Function
    If pItem Is Nothing Then LoadAlbumArt = E_POINTER: Exit Function
    Dim pStore As IPropertyStore
    Dim hr As Long
    Dim bSkipDef As Boolean
    pItem.GetMetadata pStore
    If (pStore Is Nothing) = False Then
        Dim vr As Variant
        Dim pStm As IStream
        hr = pStore.GetValue(PKEY_ThumbnailStream, vr)
        If SUCCEEDED(hr) Then
            If VarTypeEx(vr) = VT_STREAM Then
                'change to VT_UNKNOWN so tB can cast
                Dim vt As Integer
                vt = VT_UNKNOWN
                CopyMemory vr, vt, 2
            Else
                DebugLog "DrawAlbumArt::ThumbStream not VT_STREAM"
            End If
            Set pStm = vr
            If pStm Is Nothing Then
                DebugLog "No IStream, vt=" & VarTypeEx(vr)
            Else
                bSkipDef = True
            End If
        Else
            DebugLog "Couldn't get thumbstream for album art."
        End If
    Else
        DebugLog "Couldn't get property store for album art."
    End If
    'We land here if everything didn't go right; load default
    If Not bSkipDef Then
        If mUseDefCover Then
            If mDefCover = 0 Then
                Dim bt() As Byte
                #If TWINBASIC Then
                bt = LoadResData(201, "PNG")
                #Else
                bt = LoadResData(201, "CUSTOM")
                #End If
                If UBound(bt) Then
                    mDefCover = GlobalAlloc(GPTR, UBound(bt) + 1)
                    If mDefCover Then
                        CopyMemory ByVal mDefCover, bt(0), UBound(bt) + 1
                    Else
                        DebugLog "Error allocating memory for def cover, 0x" & Hex$(Err.LastDllError)
                    End If
                End If
            End If
            Set pStm = SHCreateMemStream(ByVal mDefCover, UBound(bt) + 1)
        End If
    End If
    If (pStm Is Nothing) = False Then
        #If TWINBASIC Then
        Set pDecoder = pFact.CreateDecoderFromStream(pStm, vbNullPtr, WICDecodeMetadataCacheOnDemand)
        #Else
        Set pDecoder = pFact.CreateDecoderFromStream(pStm, UUID_NULL, WICDecodeMetadataCacheOnDemand)
        #End If
        If (pDecoder Is Nothing) = False Then
            Set pFrame = pDecoder.GetFrame(0)
            If (pFrame Is Nothing) = False Then
            Set pConverter = pFact.CreateFormatConverter()
            If pConverter Is Nothing Then
                DebugLog "OpenFile:No converter"
                GoTo done
            End If
            pConverter.Initialize pFrame, GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteTypeCustom
            DrawWicImage
            End If
        Else
            DebugLog "Failed to create bitmap decoder"
        End If
    Else
        DebugLog "No stream for DrawAlbumArt"
    End If
done:
If (pStore Is Nothing) = False Then Set pStore = Nothing

End Function
Private Sub DrawWicImage()
    If pFrame Is Nothing Then Exit Sub
    If mShowCover = False Then Exit Sub
    Dim rcClient As RECT, drawRect As RECT
    Dim cx As Long, cy As Long, dx As Long, dy As Long
    GetClientRect Picture1.hwnd, rcClient
    ' rcClient.Right = Picture1.ScaleWidth
    ' rcClient.Bottom = Picture1.ScaleHeight
    pFrame.GetSize cx, cy
    DebugLog "DrawWicImage " & cx & ", " & cy & ", " & rcClient.Right & ", " & rcClient.Bottom
    'borrowed from ucAniGifEx; scales and centers an image such that the larger dimension touches
    'the edges and it's centered along the smaller dimension.
    If mLockAR Then
        Dim aspectRatio As Single
        aspectRatio = cx / cy
        Dim newWidth As Single, newHeight As Single
        If (rcClient.Right > cx) And (rcClient.Bottom > cx) Then
            'image must be expanded
            Dim clientWidth As Single
            Dim clientHeight As Single
            clientWidth = rcClient.Right - rcClient.Left
            clientHeight = rcClient.Bottom - rcClient.Top

            Dim scaleX As Single, scaleY As Single, sScale As Single
            scaleX = clientWidth / cx
            scaleY = clientHeight / cy
            sScale = scaleX
            If scaleY < scaleX Then sScale = scaleY

            newWidth = cx * sScale
            newHeight = cy * sScale

            drawRect.Left = (clientWidth - newWidth) / 2
            drawRect.Top = (clientHeight - newHeight) / 2
            drawRect.Right = drawRect.Left + newWidth
            drawRect.Bottom = drawRect.Top + newHeight
        Else
            'original handling that works for size <= image size
            drawRect.Left = (rcClient.Right - cx) / 2
            drawRect.Top = (rcClient.Bottom - cy) / 2
            drawRect.Right = drawRect.Left + cx
            drawRect.Bottom = drawRect.Top + cy
            

            If (drawRect.Left < 0) Then
                newWidth = (rcClient.Right)
                newHeight = newWidth / aspectRatio
                drawRect.Left = 0
                drawRect.Top = (rcClient.Bottom - newHeight) / 2
                drawRect.Right = newWidth
                drawRect.Bottom = drawRect.Top + newHeight
            End If
            If (drawRect.Top < 0) Then
                newHeight = (rcClient.Bottom)
                newWidth = newHeight * aspectRatio
                drawRect.Left = (rcClient.Right - newWidth) / 2
                drawRect.Top = 0
                drawRect.Right = drawRect.Left + newWidth
                drawRect.Bottom = newHeight
            End If
        End If
    Else
        drawRect = rcClient
    End If
    dx = drawRect.Right - drawRect.Left
    dy = drawRect.Bottom - drawRect.Top
    DebugLog "DrawWicImage " & cx & "," & cy & "," & dx & "," & dy
    pFact.CreateBitmapScaler pScaler
    If pScaler Is Nothing Then
        DebugLog "No scaler"
        GoTo done
    End If
    pScaler.Initialize pConverter, dx, dy, IIf(mWin10, WICBitmapInterpolationModeHighQualityCubic, WICBitmapInterpolationModeFant)
    RenderWicImage pScaler, Picture1.hDC, drawRect.Left, drawRect.Top, drawRect.Right, drawRect.Bottom
done:
End Sub
Private Sub RenderWicImage(pImage As IWICBitmapSource, hDC As LongPtr, X As Long, Y As Long, cx As Long, cy As Long)
    On Error GoTo e0
    If pImage Is Nothing Then
        DebugLog "Render: No pImage"
        Exit Sub
    End If
    Dim tBMI As BITMAPINFO
    Dim hDCScr As LongPtr
    Dim hDIBBitmap As LongPtr
    Dim pvImageBits As LongPtr
    Dim nImage As Long
    Dim nStride As Long
    Dim hBmpOld As LongPtr
    
    hDCScr = CreateCompatibleDC(hDC)
    
    With tBMI.bmiHeader
        .biSize = LenB(tBMI.bmiHeader)
        .biWidth = cx
        .biHeight = -cy
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    hDIBBitmap = CreateDIBSection(hDC, tBMI, DIB_RGB_COLORS, ByVal VarPtr(pvImageBits), 0&, 0&)
    If hDIBBitmap Then
        hBmpOld = SelectObject(hDCScr, hDIBBitmap)
        nStride = DIB_WIDTHBYTES(cx * 32)
        nImage = nStride * cy
        pImage.CopyPixels vbNullPtr, nStride, nImage, ByVal pvImageBits
        DebugLog nStride & "," & nImage
        #If TWINBASIC Then
        Picture1.Cls
        #End If
        Dim bf As BLENDFUNCTION
        Dim lbf As Long
        With bf
            .BlendOp = AC_SRC_OVER
            .BlendFlags = 0
            .SourceConstantAlpha = 255
            .AlphaFormat = AC_SRC_ALPHA
        End With
        CopyMemory lbf, bf, LenB(lbf)
        AlphaBlend hDC, X, Y, cx, cy, hDCScr, 0, 0, cx, cy, lbf
        SelectObject hDCScr, hBmpOld
        DeleteObject hDIBBitmap
        DeleteDC hDCScr
        Set Picture1.Picture = Picture1.Image
    Else
        DebugLog "Failed to create hDIBBitmap, lastErr=0x" & Hex$(Err.LastDllError)
    End If

Exit Sub

e0:
    DebugLog "cWICImage.Render->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
End Sub

Public Function LShift(ByVal Value As Long, ByVal Shift As Integer) As Long
    MakeOnBits
    If (Value And (2 ^ (31 - Shift))) Then GoTo OverFlow
    LShift = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))
    Exit Function
OverFlow:
    LShift = ((Value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
End Function
Private Sub MakeOnBits()
Dim j As Integer
Dim v As Long
For j = 0 To 30
    v = v + (2 ^ j)
    OnBits(j) = v
Next j
OnBits(j) = v + &H80000000
End Sub
Private Function RShift(ByVal Value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    If (Value And &H80000000) Then hi = &H40000000
  
    RShift = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShift = (RShift Or (hi \ (2 ^ (Shift - 1))))
End Function
Private Function DIB_WIDTHBYTES(bits As Long) As Long
#If TWINBASIC Then
Return (((bits + 31) >> 5) << 2)
#Else
DIB_WIDTHBYTES = LShift(RShift((bits + 31), 5), 2)
#End If
End Function
Private Function Subclass2(hwnd As LongPtr, lpFN As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
    If uId = 0 Then uId = hwnd
    Subclass2 = SetWindowSubclass(hwnd, lpFN, uId, dwRefData):      Debug.Assert Subclass2
End Function
Private Function UnSubclass2(hwnd As LongPtr, ByVal lpFN As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hwnd, lpFN, pid)
End Function
Public Function ucWndProc(ByVal lng_hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr) As LongPtr
    Select Case uMsg
        Case WM_KEYDOWN
            RaiseEvent PlayerKeyDown(CLng(wParam), CTRUE, CIntToUInt(LOWORD(CLng(lParam))), CIntToUInt(HIWORD(CLng(lParam))))
            Exit Function
        Case WM_KEYUP
            RaiseEvent PlayerKeyUp(CLng(wParam), CTRUE, CIntToUInt(LOWORD(CLng(lParam))), CIntToUInt(HIWORD(CLng(lParam))))
            Exit Function
        Case WM_PAINT
            If mHasVideo Then
                OnPaint hwnd
                Exit Function
            End If
        Case WM_SIZE
            If mHasVideo Then
                OnSize hwnd, CLng(wParam), CIntToUInt(LOWORD(CLng(lParam))), CIntToUInt(HIWORD(CLng(lParam)))
                ' Exit Function
            Else
                Set Picture1.Picture = LoadPicture
                Picture1.Width = LOWORD(lParam)
                Picture1.Height = HIWORD(lParam)
                DrawWicImage
            End If
        Case WM_ERASEBKGND
            If mHasVideo Then
                ucWndProc = 1
                Exit Function
            End If
        Case WM_LBUTTONDBLCLK
            If mAllowFullscreen Then
                DebugLog "Toggle FS"
                If mFullscreen Then
                    ExitFullscreen
                    mFullscreen = False
                ElseIf EnterFullscreen() Then
                    mFullscreen = True
                End If
            End If
        Case WM_LBUTTONUP
            RaiseEvent PlayerClick(1)
        Case WM_LBUTTONDOWN
            If (mPlayer Is Nothing) = False Then
                If mAllowClickPause Then
                    Dim dwState As MFP_MEDIAPLAYER_STATE
                    mPlayer.GetState dwState
                    If dwState = MFP_MEDIAPLAYER_STATE_PAUSED Then
                        Me.Paused = False
                    ElseIf dwState = MFP_MEDIAPLAYER_STATE_PLAYING Then
                        Me.Paused = True
                    End If
                End If
            End If
        Case WM_RBUTTONUP
            RaiseEvent PlayerClick(2)
        Case WM_MOUSEWHEEL
            RaiseEvent PlayerWheelScroll(GET_WHEEL_DELTA_WPARAM(wParam), GET_KEYSTATE_WPARAM(wParam))
        Case WM_DESTROY
            UnSubclass2 lng_hWnd, AddressOf ucSimplePlayerHelperProc, lng_hWnd
    End Select
    ucWndProc = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)
End Function
