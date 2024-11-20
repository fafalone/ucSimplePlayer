VERSION 5.00
Begin VB.UserControl ucSimplePlayer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucSimplePlayer.ctx":0000
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4680
      Y1              =   3480
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   0
      Y2              =   3480
   End
End
Attribute VB_Name = "ucSimplePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************
'ucSimplePlayer v1.0.1
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
'Version 1.0.1 - Initial release
'**********************************************************************

Implements IMFPMediaPlayerCallback

Public Event PlaybackStart(ByVal cyDuration As Currency)
Public Event PlaybackEnded()
Public Event PlayerKeyDown(ByVal vk As Long, ByVal fDown As BOOL, ByVal cRepeat As Long, ByVal flags As Long)

Private mFile As String
Private mPlaying As Boolean
Private mPaused As Boolean
Private mHasVideo As Boolean

Private mPlayer As IMFPMediaPlayer
Private mItem As IMFPMediaItem

'WinDevLib defs for VB6 only:
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
#If VBA7 Then
Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As BOOL
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As BOOL
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Const PTR_SIZE = 8
#Else
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As BOOL
Private Declare Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, Optional ByVal dwRefData As LongPtr) As BOOL
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As BOOL
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Const PTR_SIZE = 4
#End If

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
        Debug.Print "Failed to create media player, 0x" & Hex$(hr)
        Exit Sub
    End If
    mPlayer.CreateMediaItemFromURL StrPtr(sFullPath), CFALSE, 0, Nothing
'    Debug.Print "CreateMediaItemFromURL hr=" & Err.LastHResult
    Exit Sub
e0:
    Debug.Print "PlayMediaFile error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Sub FrameStep()
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.FrameStep
    End If
    Exit Sub
e0:
    Debug.Print "FrameStep error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Property Get Volume() As Single
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetVolume Volume
    End If
    Exit Property
e0:
    Debug.Print "Volume.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Volume(ByVal fVol As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetVolume fVol
    End If
    Exit Property
e0:
    Debug.Print "Volume.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Get Balance() As Single
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetBalance Balance
    End If
    Exit Property
e0:
    Debug.Print "Balance.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Property Let Balance(ByVal fBal As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetBalance fBal
    End If
    Exit Property
e0:
    Debug.Print "Balance.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "Muted.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "Muted.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "Duration.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "Position.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "Position.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "GetNativeVideoSize error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "BorderColor.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
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
    Debug.Print "BorderColor.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Property
Public Sub GetSupportedRates(ByVal bForward As Boolean, pfRateMin As Single, pfRateMax As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetSupportedRates IIf(bForward, CTRUE, CFALSE), pfRateMin, pfRateMax
    End If
    Exit Sub
e0:
    Debug.Print "GetSupportedRates error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub
Public Property Get Rate() As Single
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.GetRate Rate
    End If
    Exit Property
e0:
    Debug.Print "Rate.Get error 0x" & Hex$(Err.Number) & ", " & Err.Description
    End Property
Public Property Let Rate(ByVal fRate As Single)
    On Error GoTo e0
    If (mPlayer Is Nothing) = False Then
        mPlayer.SetRate fRate
    End If
    Exit Property
e0:
    Debug.Print "Rate.Let error 0x" & Hex$(Err.Number) & ", " & Err.Description
    End Property

#If TWINBASIC Then
Private Sub IMFPMediaPlayerCallback_OnMediaPlayerEvent(pEventHeader As MFP_EVENT_HEADER) 'Implements IMFPMediaPlayerCallback.OnMediaPlayerEvent
    If pEventHeader.hrEvent < 0 Then
        Debug.Print "Playback error, 0x" & Hex$(pEventHeader.hrEvent)
        Exit Sub
    End If
    
    Select Case pEventHeader.eEventType
        Case MFP_EVENT_TYPE_MEDIAITEM_CREATED
            OnMediaItemCreated MFP_GET_MEDIAITEM_CREATED_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_MEDIAITEM_SET
            OnMediaItemSet MFP_GET_MEDIAITEM_SET_EVENT(pEventHeader)
 
        Case MFP_EVENT_TYPE_PLAYBACK_ENDED
            RaiseEvent PlaybackEnded
    End Select
End Sub
    Private Function MFP_GET_MEDIAITEM_CREATED_EVENT(pEventHeader As MFP_EVENT_HEADER) As MFP_MEDIAITEM_CREATED_EVENT
        If pEventHeader.eEventType = MFP_EVENT_TYPE_MEDIAITEM_CREATED Then
            CopyMemory MFP_GET_MEDIAITEM_CREATED_EVENT, pEventHeader, LenB(MFP_GET_MEDIAITEM_CREATED_EVENT)
            Dim pUnk As IUnknown
            Set pUnk = MFP_GET_MEDIAITEM_CREATED_EVENT.Header.pMediaPlayer
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
            Set pUnk = MFP_GET_MEDIAITEM_CREATED_EVENT.Header.pMediaPlayer
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
            Set pUnk = MFP_GET_MEDIAITEM_CREATED_EVENT.pMediaItem
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
        End If
    End Function
    
    Private Function MFP_GET_MEDIAITEM_SET_EVENT(pEventHeader As MFP_EVENT_HEADER) As MFP_MEDIAITEM_SET_EVENT
        If pEventHeader.eEventType = MFP_EVENT_TYPE_MEDIAITEM_SET Then
            CopyMemory MFP_GET_MEDIAITEM_SET_EVENT, pEventHeader, LenB(MFP_GET_MEDIAITEM_SET_EVENT)
            Dim pUnk As IUnknown
            Set pUnk = MFP_GET_MEDIAITEM_SET_EVENT.Header.pMediaPlayer
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
            Set pUnk = MFP_GET_MEDIAITEM_SET_EVENT.Header.pMediaPlayer
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
            Set pUnk = MFP_GET_MEDIAITEM_SET_EVENT.pMediaItem
            Call CopyMemory(pUnk, 0&, PTR_SIZE)
        End If
    End Function
#Else
Private Sub IMFPMediaPlayerCallback_OnMediaPlayerEvent(pEventHeader As MFP_EVENT_HEADER) 'Implements IMFPMediaPlayerCallback.OnMediaPlayerEvent
    If pEventHeader.hrEvent < 0 Then
        Debug.Print "Playback error, 0x" & Hex$(pEventHeader.hrEvent)
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
    Debug.Print "OnMediaItemCreated"
    If (mPlayer Is Nothing) = False Then
        Dim hr As Long
        Dim bHasVideo As BOOL, bIsSelected As BOOL
        pEvent.pMediaItem.HasVideo bHasVideo, bIsSelected
'        Debug.Print "OnMediaItemCreated bHasVideo=" & bHasVideo & ", bIsSelected=" & bIsSelected
        mPlayer.SetMediaItem pEvent.pMediaItem
        mHasVideo = (bHasVideo = CTRUE)
    End If
End Sub
Private Sub OnMediaItemSet(pEvent As MFP_MEDIAITEM_SET_EVENT)
'    Debug.Print "OnMediaItemSet"
    On Error GoTo e0
    mPlayer.Play
    'Can't get duration prior to this
    Dim dr As Currency
    Dim pv As Variant
    mPlayer.GetDuration MFP_POSITIONTYPE_100NS, pv
    VariantSetType pv, VT_CY, VT_I8
    dr = CCur(pv)
    RaiseEvent PlaybackStart(dr)
    Exit Sub
e0:
    Debug.Print "OnMediaItemSet Error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Sub

Private Sub UserControl_Initialize() 'Handles UserControl.Initialize
 
End Sub
Private Sub UserControl_Terminate() 'Handles UserControl.Terminate
    If (mPlayer Is Nothing) = False Then
        mPlayer.Shutdown
    End If
End Sub

Private Sub UserControl_Show() 'Handles UserControl.Show
    If Ambient.UserMode Then
        Line1.Visible = False
        Line2.Visible = False
        Subclass2 UserControl.hWnd, AddressOf ucSimplePlayerHelperProc, UserControl.hWnd, ObjPtr(Me)
    Else
        Line1.X1 = 0
        Line1.Y1 = 0
        Line1.X2 = UserControl.ScaleWidth
        Line1.Y2 = UserControl.ScaleHeight
            
        Line2.X1 = 0
        Line2.Y1 = UserControl.ScaleHeight
        Line2.X2 = UserControl.ScaleWidth
        Line2.Y2 = 0
    End If
End Sub
 
Private Sub OnSize(ByVal hWnd As LongPtr, ByVal state As Long, ByVal cx As Long, ByVal cy As Long)
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
        ' Debug.Print "State=" & s
        On Error Resume Next
        mPlayer.UpdateVideo
    Else
        FillRect hDC, ps.rcPaint, COLOR_WINDOW + 1
    End If
    EndPaint hWnd, ps
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
        Case WM_PAINT
            OnPaint hWnd
            Exit Function
        Case WM_SIZE
            OnSize hWnd, CLng(wParam), CIntToUInt(LOWORD(CLng(lParam))), CIntToUInt(HIWORD(CLng(lParam)))
            ' Exit Function
        Case WM_ERASEBKGND
            ucWndProc = 1
            Exit Function
        Case WM_DESTROY
            UnSubclass2 lng_hWnd, AddressOf ucSimplePlayerHelperProc, lng_hWnd
    End Select
    ucWndProc = DefSubclassProc(lng_hWnd, uMsg, wParam, lParam)
End Function

Private Sub UserControl_Resize() 'Handles UserControl.Resize
    If Ambient.UserMode = False Then
        Line1.X1 = 0
        Line1.Y1 = 0
        Line1.X2 = UserControl.ScaleWidth
        Line1.Y2 = UserControl.ScaleHeight
            
        Line2.X1 = 0
        Line2.Y1 = UserControl.ScaleHeight
        Line2.X2 = UserControl.ScaleWidth
        Line2.Y2 = 0
    End If
End Sub


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


