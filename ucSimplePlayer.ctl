VERSION 5.00
Begin VB.UserControl ucSimplePlayer 
   BackColor       =   &H80000001&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucSimplePlayer.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   840
      Top             =   960
   End
End
Attribute VB_Name = "ucSimplePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const DEBUGMSG = 0
'**********************************************************************
'ucSimplePlayer v2.2.5
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
Public Event PlayerKeyDown(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal flags As Long)
Public Event PlayerKeyUp(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal flags As Long)
Public Event PlayerClick(ByVal Button As Long)
Private mFile As String
Private mPlaying As Boolean
Private mPaused As Boolean
Private mHasVideo As Boolean
Private mFullscreen As Boolean
Private mAllowFullscreen As Boolean
Private mPlayTimer As Boolean
Private dwStyleOld As WindowStyles
Private dwStyleExOld As WindowStylesEx
Private hParOld As LongPtr
Private mOldPlacement As WINDOWPLACEMENT
Private mPlayer As IMFPMediaPlayer
Private mItem As IMFPMediaItem
Private mDuration As Variant
Private mLockAR As Boolean
Private mVidSize As SIZE
Private mLastPos As Variant
 Private mSetPos As Variant
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
    
Private Const SIZE_MINIMIZED = 1
Private Const COLOR_WINDOW = 5
Private Const CTRUE = 1
Private Const CFALSE = 0
Private Type PAINTSTRUCT
    hdc                  As LongPtr
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
    flags As WINDOWPLACEMENT_FLAGS
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
Private DeclareWide PtrSafe Function PropVariantToStringAlloc Lib "propsys" (ByRef propvar1 As Any, ppszOut As LongPtr) As Long
#Else
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As BOOL
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As BOOL
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal dwFlags As DefaultMonitorValues) As LongPtr
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, lpmi As Any) As BOOL
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As BOOL
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, Optional ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SWP_Flags) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function IsEqualGUID Lib "ole32" (ByRef rguid1 As UUID, ByRef rguid2 As UUID) As BOOL
Private Declare Function PropVariantToStringAlloc Lib "propsys" (ByRef propvar1 As Any, ppszOut As LongPtr) As Long
#End If
#If Win64 Then
Private Const PTR_SIZE = 8
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Const PTR_SIZE = 4
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If

'WinDevLib helpers needed for VB6 only:
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
Private Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L: .Data2 = w1: .Data3 = w2: .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
End Sub
#End If
'******************************************************
'END VB ONLY

Private Sub DebugLog(sMsg As String)
    #If DEBUGMSG Then
    Debug.Print sMsg
    #End If
End Sub
Public Property Get Paused() As Boolean
    Paused = mPaused
End Property
Public Property Let Paused(ByVal fPaused As Boolean)
    If mPlayer Is Nothing Then Exit Property
 
    If fPaused Then
        mPaused = True
        mPlayer.Pause
    Else
        mPlayer.Play
    End If
End Property
Public Property Get FileName() As String
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
        hr = MFPCreateMediaPlayer(0, CFALSE, 0, Me, UserControl.hWnd, mPlayer)
    End If
    If hr < 0 Then
        DebugLog "Failed to create media player, 0x" & Hex$(hr)
        Exit Sub
    End If
    mPlayer.CreateMediaItemFromURL StrPtr(sFullPath), CFALSE, 0, Nothing
    #If TWINBASIC Then
    DebugLog "CreateMediaItemFromURL hr=" & Err.LastHResult
    #End If
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
Public Property Get BorderColor() As OLE_COLOR
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
Public Property Get AllowFullscreen() As Boolean
    AllowFullscreen = mAllowFullscreen
End Property
Public Property Let AllowFullscreen(bAllow As Boolean)
    mAllowFullscreen = bAllow
End Property
Public Property Get Fullscreen() As Boolean
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
    EnablePlayTimer = mPlayTimer
End Property
Public Property Let EnablePlayTimer(ByVal bEnable As Boolean)
    mPlayTimer = bEnable
    If (mPlaying = True) And (bEnable = True) Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Property
Public Property Get PlayTimerInterval() As Long
    EnablePlayTimer = Timer1.Interval
End Property
Public Property Let PlayTimeInterval(ByVal lInterval As Long)
    Timer1.Interval = lInterval
End Property
Public Property Get PreserveAspectRatio() As Boolean
    PreserveAspectRatio = mLockAR
End Property
Public Property Let PreserveAspectRatio(ByVal bPreserve As Boolean)
    If bPreserve = mLockAR Then Exit Property
    mLockAR = bPreserve
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetAspectRatioMode (IIf(mLockAR, MFVideoARMode_PreservePicture, MFVideoARMode_None))
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
            Timer1.Enabled = False
            RaiseEvent PlaybackEnded
             
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
    If mPlayTimer Then Timer1.Enabled = True
    'Can't get duration prior to this
    If mResetData.Active Then
        mResetData.Active = False
        mCurStreams.idxVidActive = mResetData.nV
        mCurStreams.idxAudActive = mResetData.nA
        mPlayer.SetPosition MFP_POSITIONTYPE_100NS, mResetData.lastPos
    End If
    Dim dr As Currency
    Dim pv As Variant
    mPlayer.GetDuration MFP_POSITIONTYPE_100NS, mDuration
    Dim sz As SIZE
    mPlayer.GetNativeVideoSize mVidSize, sz
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
    pvType = 0&: pvName = 0&: pvLang = 0&
    On Error Resume Next
    For i = 0 To nStm - 1
        n = 0
        CopyMemory pvType, VT_I4, 2
        CopyMemory pvName, VT_I4, 2
        CopyMemory pvLang, VT_I4, 2
        pItem.GetStreamAttribute i, MF_MT_MAJOR_TYPE, pvType
'        If SUCCEEDED(Err.LastHResult) Then
        If IsEmpty(pvType) = False Then
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
'        If SUCCEEDED(Err.LastHResult) Then
        If IsEmpty(pvName) = False Then
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
        If IsEmpty(pvLang) = False Then
'        If SUCCEEDED(Err.LastHResult) Then
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
    Next
    CopyMemory pvType, VT_I4, 2 'VariantClear will cause a crash trying to improperly free VT_LPWSTR/VT_CLSID
    CopyMemory pvName, VT_I4, 2
    CopyMemory pvLang, VT_I4, 2
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
Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
    mAllowFullscreen = True
    mLockAR = True
    Timer1.Interval = 500
End Sub
Private Sub UserControl_Terminate() 'Handles UserControl.Terminate
    If (mPlayer Is Nothing) = False Then
        mPlayer.Shutdown
    End If
End Sub

Private Sub UserControl_Show() 'Handles UserControl.Show
    If Ambient.UserMode Then
        Subclass2 UserControl.hWnd, AddressOf ucSimplePlayerHelperProc, UserControl.hWnd, ObjPtr(Me)
 
    End If
End Sub
 
Private Sub OnSize(ByVal hWnd As LongPtr, ByVal state As Long, ByVal cx As Long, ByVal cy As Long)
    If state <> SIZE_MINIMIZED Then
        If (mPlayer Is Nothing) = False Then
            mPlayer.UpdateVideo
        End If
    End If
End Sub


Private Sub OnPaint(ByVal hWnd As LongPtr)
    Dim ps As PAINTSTRUCT
    Dim hdc As LongPtr
    hdc = BeginPaint(hWnd, ps)
    If ((mPlayer Is Nothing) = False) And (mHasVideo = True) Then
        ' Dim s As MFP_MEDIAPLAYER_STATE
        ' mPlayer.GetState s
        ' DebugLog "State=" & s
        On Error Resume Next
        mPlayer.UpdateVideo
    Else
        FillRect hdc, ps.rcPaint, COLOR_WINDOW + 1
    End If
    EndPaint hWnd, ps
End Sub

Private Function EnterFullscreen() As Long
    Dim hMon As LongPtr
    Dim mi As MONITORINFO
    hMon = MonitorFromWindow(UserControl.hWnd, MONITOR_DEFAULTTONEAREST)
    mi.cbSize = LenB(mi)
    If hMon = 0 Then Exit Function
    GetMonitorInfo hMon, mi
        
    GetWindowPlacement UserControl.hWnd, mOldPlacement
    dwStyleOld = CLng(GetWindowLongPtr(UserControl.hWnd, GWL_STYLE))
    dwStyleExOld = CLng(GetWindowLongPtr(UserControl.hWnd, GWL_EXSTYLE))
    hParOld = GetParent(UserControl.hWnd)
    SetParent UserControl.hWnd, GetDesktopWindow()
    SetWindowLongPtr UserControl.hWnd, GWL_STYLE, WS_POPUP Or WS_VISIBLE
    SetWindowPos UserControl.hWnd, HWND_TOP, mi.rcMonitor.Left, mi.rcMonitor.Top, _
                        mi.rcMonitor.Right - mi.rcMonitor.Left, _
                        mi.rcMonitor.Bottom - mi.rcMonitor.Top, SWP_NOOWNERZORDER Or SWP_FRAMECHANGED

    EnterFullscreen = 1
End Function
 
Private Function ExitFullscreen() As Long
    ShowWindow UserControl.hWnd, SW_RESTORE
    SetWindowLongPtr UserControl.hWnd, GWL_EXSTYLE, dwStyleExOld
    SetWindowLongPtr UserControl.hWnd, GWL_STYLE, dwStyleOld
    SetWindowPos UserControl.hWnd, HWND_NOTOPMOST, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, SWP_SHOWWINDOW
    SetParent UserControl.hWnd, hParOld
    SetWindowPlacement UserControl.hWnd, mOldPlacement
    UserControl.Width = UserControl.Width - 1
    UserControl.Width = UserControl.Width + 1
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
 
Private Function Subclass2(hWnd As LongPtr, lpFN As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
If uId = 0 Then uId = hWnd
    Subclass2 = SetWindowSubclass(hWnd, lpFN, uId, dwRefData):      Debug.Assert Subclass2
End Function
Private Function UnSubclass2(hWnd As LongPtr, ByVal lpFN As LongPtr, pid As LongPtr) As Boolean
    UnSubclass2 = RemoveWindowSubclass(hWnd, lpFN, pid)
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
            OnPaint hWnd
            Exit Function
        Case WM_SIZE
            OnSize hWnd, CLng(wParam), CIntToUInt(LOWORD(CLng(lParam))), CIntToUInt(HIWORD(CLng(lParam)))
            ' Exit Function
        Case WM_ERASEBKGND
            ucWndProc = 1
            Exit Function
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
        Case WM_RBUTTONUP
            RaiseEvent PlayerClick(2)
        Case WM_DESTROY
            UnSubclass2 lng_hWnd, AddressOf ucSimplePlayerHelperProc, lng_hWnd
    End Select
    ucWndProc = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)
End Function



