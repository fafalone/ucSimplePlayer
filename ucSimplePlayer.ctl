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
End
Attribute VB_Name = "ucSimplePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DEBUGMSG As Boolean = 0
'**********************************************************************
'ucSimplePlayer v1.1.3
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
'Version 1.1.3 (26 Mar 2025)
'-Added Fullscreen support. It can be done either automatically by double
' clicking the control (you can disable this with AllowFullscreen = False),
' or manually via the .Fullscreen property Get/Let.
'- Minor fixes
'Version 1.0.1 - Initial release
'**********************************************************************
 
Implements IMFPMediaPlayerCallback

#If TWINBASIC Then
Public Event PlaybackStart(ByVal cyDuration As LongLong)
#Else
Public Event PlaybackStart(ByVal cyDuration As Currency)
#End If
Public Event PlaybackEnded()
Public Event PlayerKeyDown(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal flags As Long)

Private mFile As String
Private mPlaying As Boolean
Private mPaused As Boolean
Private mHasVideo As Boolean
Private mFullscreen As Boolean
Private mAllowFullscreen As Boolean
Private dwStyleOld As WindowStyles
Private dwStyleExOld As WindowStylesEx
Private hParOld As LongPtr
Private mOldPlacement As WINDOWPLACEMENT
Private mPlayer As IMFPMediaPlayer
Private mItem As IMFPMediaItem

'In twinBASIC, the WinDevLib package covers both oleexp.tlb interfaces and APIs.
#If TWINBASIC = 0 Then
Private Const SIZE_RESTORED = 0
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
    length As Long
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
#Else
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As SWP_Flags) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
#End If
#If Win64 Then
Private Const PTR_SIZE = 8
Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Const PTR_SIZE = 4
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If
'WinDevLib helpers needed for VB6 only:
Private Function VariantSetType(pvar As Variant, ByVal vt As Integer, Optional ByVal vtOnlyIf As Integer = -1) As Boolean
    If VarPtr(pvar) = 0 Then Exit Function
    If vtOnlyIf <> -1 Then
        If VarTypeEx(pvar) <> vtOnlyIf Then
            VariantSetType = False
            Exit Function
        End If
    End If
    CopyMemory pvar, vt, 2
    VariantSetType = True
End Function
Private Function VarTypeEx(pvar As Variant) As VARENUM
    If VarPtr(pvar) = 0 Then Exit Function
    CopyMemory VarTypeEx, ByVal VarPtr(pvar), 2
End Function
Private Function CIntToUInt(ByVal Value As Integer) As Long
Const OFFSET_2 As Long = 65536
If Value < 0 Then
    CIntToUInt = Value + OFFSET_2
Else
    CIntToUInt = Value
End If
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
'    DebugLog "CreateMediaItemFromURL hr=" & Err.LastHResult
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
        pcx = sz.CX: pcy = sz.CY
        pcxAspectRatio = szar.CX: pcyAspectRatio = szar.CY
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


#If TWINBASIC Then
Private Sub IMFPMediaPlayerCallback_OnMediaPlayerEvent(pEventHeader As MFP_EVENT_HEADER) Implements IMFPMediaPlayerCallback.OnMediaPlayerEvent
    If pEventHeader.hrEvent < 0 Then
        DebugLog "Playback error, 0x" & Hex$(pEventHeader.hrEvent) & "; EventType=" & pEventHeader.eEventType
        'Note: If this error is 0x80070426, you must enable the "Microsoft Account Sign-in Assistant" service
        '      the first time you use a newly installed Microsoft Store codec.
        Exit Property
    End If
    
    Select Case pEventHeader.eEventType
        Case MFP_EVENT_TYPE_MEDIAITEM_CREATED
            OnMediaItemCreated MFP_GET_MEDIAITEM_CREATED_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_MEDIAITEM_SET
            OnMediaItemSet MFP_GET_MEDIAITEM_SET_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_PLAYBACK_ENDED
            RaiseEvent PlaybackEnded
    End Select
End Property
 


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
    mPlayer.Play
    'Can't get duration prior to this
    Dim dr As Currency
    Dim pv As Variant
    mPlayer.GetDuration MFP_POSITIONTYPE_100NS, pv
    #If TWINBASIC Then
    RaiseEvent PlaybackStart(pv)
    #Else
     VariantSetType pv, VT_CY, VT_I8
     dr = CCur(pv)
     RaiseEvent PlaybackStart(dr)
    #End If
    Exit Sub
e0:
    DebugLog "OnMediaItemSet Error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub

 
Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
    mAllowFullscreen = True
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
 
Private Sub OnSize(ByVal hWnd As LongPtr, ByVal state As Long, ByVal CX As Long, ByVal CY As Long)
    If state = SIZE_RESTORED Then
        If (mPlayer Is Nothing) = False Then
            mPlayer.UpdateVideo
        End If
    End If
End Sub


Private Sub OnPaint(ByVal hWnd As LongPtr)
    Dim ps As PAINTSTRUCT
    Dim hDC As LongPtr
    hDC = BeginPaint(hWnd, ps)
    If ((mPlayer Is Nothing) = False) And (mHasVideo = True) Then
        ' Dim s As MFP_MEDIAPLAYER_STATE
        ' mPlayer.GetState s
        ' DebugLog "State=" & s
        On Error Resume Next
        mPlayer.UpdateVideo
    Else
        FillRect hDC, ps.rcPaint, COLOR_WINDOW + 1
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
        Case WM_DESTROY
            UnSubclass2 lng_hWnd, AddressOf ucSimplePlayerHelperProc, lng_hWnd
    End Select
    ucWndProc = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)
End Function

 

