VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{4AD19335-05E8-4F52-AA92-C1FAF1AD8737}#2.3#0"; "ucSimplePlay.ocx"
Begin VB.Form Form1 
   Caption         =   "ucSimplePlayer Demo"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   3  'Windows Default
   Begin ucSimplePlay.ucSimplePlayer ucSimplePlayer1 
      Height          =   5775
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10186
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   600
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Frame Step"
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   375
      Left            =   8160
      TabIndex        =   11
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1085
      _Version        =   327682
      TickStyle       =   2
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Faster 25%"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Slower 25%"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mute"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   7680
      Picture         =   "Form1.frx":4492
      Stretch         =   -1  'True
      Top             =   600
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'In twinBASIC, swap oleexp.tlb for WinDevLib then fix StrPtr in Command1_Click

 #If TWINBASIC = 0 Then
 Private Enum MII_Mask
   MIIM_STATE = &H1
   MIIM_ID = &H2
   MIIM_SUBMENU = &H4
   MIIM_CHECKMARKS = &H8
   MIIM_TYPE = &H10
   MIIM_DATA = &H20
   MIIM_STRING = &H40
   MIIM_BITMAP = &H80
   MIIM_FTYPE = &H100
 End Enum
 Private Enum MenuFlags
   MF_INSERT = &H0
   MF_ENABLED = &H0
   MF_UNCHECKED = &H0
   MF_BYCOMMAND = &H0
   MF_STRING = &H0
   MF_UNHILITE = &H0
   MF_GRAYED = &H1
   MF_DISABLED = &H2
   MF_BITMAP = &H4
   MF_CHECKED = &H8
   MF_POPUP = &H10
   MF_MENUBARBREAK = &H20
   MF_MENUBREAK = &H40
   MF_HILITE = &H80
   MF_CHANGE = &H80
   MF_END = &H80                    ' Obsolete -- only used by old RES files
   MF_APPEND = &H100
   MF_OWNERDRAW = &H100
   MF_DELETE = &H200
   MF_USECHECKBITMAPS = &H200
   MF_BYPOSITION = &H400
   MF_SEPARATOR = &H800
   MF_REMOVE = &H1000
   MF_DEFAULT = &H1000
   MF_SYSMENU = &H2000
   MF_HELP = &H4000
   MF_RIGHTJUSTIFY = &H4000
   MF_MOUSESELECT = &H8000&
 End Enum
 Private Enum MF_Type
   MFT_STRING = MF_STRING
   MFT_BITMAP = MF_BITMAP
   MFT_MENUBARBREAK = MF_MENUBARBREAK
   MFT_MENUBREAK = MF_MENUBREAK
   MFT_OWNERDRAW = MF_OWNERDRAW
   MFT_RADIOCHECK = &H200
   MFT_SEPARATOR = MF_SEPARATOR
   MFT_RIGHTORDER = &H2000
   MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
 End Enum
 Private Enum MF_State
   MFS_GRAYED = &H3
   MFS_DISABLED = MFS_GRAYED
   MFS_CHECKED = MF_CHECKED
   MFS_HILITE = MF_HILITE
   MFS_ENABLED = MF_ENABLED
   MFS_UNCHECKED = MF_UNCHECKED
   MFS_UNHILITE = MF_UNHILITE
   MFS_DEFAULT = MF_DEFAULT
 End Enum
 Private Type MENUITEMINFO
     cbSize As Long
     fMask As MII_Mask
     fType As MF_Type ' used if MIIM_TYPE (4.0) or MIIM_FTYPE (>4.0)
     fState As MF_State ' used if MIIM_STATE
     wID As Long ' used if MIIM_ID
     hSubMenu As LongPtr ' used if MIIM_SUBMENU
     hbmpChecked As LongPtr ' used if MIIM_CHECKMARKS
     hbmpUnchecked As LongPtr ' used if MIIM_CHECKMARKS
     dwItemData As LongPtr ' used if MIIM_DATA
     dwTypeData As LongPtr ' used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
     cch As Long ' used if MIIM_TYPE (4.0) or MIIM_STRING (>4.0)
     hbmpItem As LongPtr ' used if MIIM_BITMAP
 End Type
 Private Enum TPM_wFlags
     TPM_LEFTBUTTON = &H0
     TPM_RECURSE = &H1
     TPM_RIGHTBUTTON = &H2
     TPM_LEFTALIGN = &H0
     TPM_CENTERALIGN = &H4
     TPM_RIGHTALIGN = &H8
     TPM_TOPALIGN = &H0
     TPM_VCENTERALIGN = &H10
     TPM_BOTTOMALIGN = &H20
 
     TPM_HORIZONTAL = &H0         ' Horz alignment matters more
     TPM_VERTICAL = &H40            ' Vert alignment matters more
     TPM_NONOTIFY = &H80           ' Don't send any notification msgs
     TPM_RETURNCMD = &H100
 
     TPM_HORPOSANIMATION = &H400
     TPM_HORNEGANIMATION = &H800
     TPM_VERPOSANIMATION = &H1000
     TPM_VERNEGANIMATION = &H2000
     TPM_NOANIMATION = &H4000
     TPM_LAYOUTRTL = &H8000&
     'Win7+:
     TPM_WORKAREA = &H10000
 End Enum
 Private Enum SHACF
     SHACF_DEFAULT = &H0        ' Currently (SHACF_FILESYSTEM | SHACF_URLALL)
     SHACF_FILESYSTEM = &H1        ' This includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
     SHACF_URLHISTORY = &H2        ' URLs in the User's History
     SHACF_URLMRU = &H4        ' URLs in the User's Recently Used list.
     SHACF_URLALL = &H6        ' (SHACF_URLHISTORY | SHACF_URLMRU)
     SHACF_USETAB = &H8        ' Use the tab to move thru the autocomplete possibilities instead of to the next dialog/window control.
     SHACF_FILESYS_ONLY = &H10       ' Don't AutoComplete non-File System items.
     SHACF_FILESYS_DIRS = &H20       ' Same as SHACF_FILESYS_ONLY except it only includes directories, UNC servers, and UNC server shares.
     SHACF_VIRTUAL_NAMESPACE = &H40       ' Also include the virtual namespace
     SHACF_AUTOSUGGEST_FORCE_ON = &H10000000 ' Ignore the registry default and force the feature on.
     SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000 ' Ignore the registry default and force the feature off.
     SHACF_AUTOAPPEND_FORCE_ON = &H40000000 ' Ignore the registry default and force the feature on. (Also know as AutoComplete)
     SHACF_AUTOAPPEND_FORCE_OFF = &H80000000 ' Ignore the registry default and force the feature off. (Also know as AutoComplete)
 End Enum
 Private Const TBS_DOWNISLEFT = &H400   ' Down=Left and Up=Right (default is Down=Right and Up=Left)
 #If VBA7 Then
 Private Declare PtrSafe Function CreateMenu Lib "user32" () As LongPtr
 Private Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
 Private Declare PtrSafe Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As LongPtr, ByVal first As Long, ByVal last As Long, ByVal check As Long, ByVal flags As MenuFlags) As Long
 Private Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As TPM_wFlags, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As LongPtr, lpRC As Any) As Long
 Private Declare PtrSafe Function InsertMenuItem Lib "user32" Alias "InsertMenuItemW" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPosition As BOOL, lpmii As MENUITEMINFO) As BOOL
 Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As BOOL
 Private Declare PtrSafe Function SHAutoComplete Lib "shlwapi" (ByVal hwndEdit As LongPtr, ByVal dwFlags As SHACF) As Long
 #If Win64 Then
 Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
 Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
 #Else
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
 #End If
 #Else
 Private Declare Function CreateMenu Lib "user32" () As LongPtr
 Private Declare Function CreatePopupMenu Lib "user32" () As LongPtr
 Private Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As LongPtr, ByVal first As Long, ByVal last As Long, ByVal check As Long, ByVal Flags As MenuFlags) As Long
 Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal uFlags As TPM_wFlags, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As LongPtr, lpRC As Any) As Long
 Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemW" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPosition As BOOL, lpmii As MENUITEMINFO) As BOOL
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As BOOL
 Private Declare Function SHAutoComplete Lib "shlwapi" (ByVal hwndEdit As LongPtr, ByVal dwFlags As SHACF) As Long
 Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
 #End If
 
 
 #End If
 
Private mFile As String
Private mPlay As Boolean
Private mNoUpdate As Boolean
Private mDropScroll As Boolean

Private Sub Command1_Click() 'Handles Command1.Click
    Dim fod As IFileOpenDialog
    Set fod = New FileOpenDialog
    Dim lpAbsPath As LongPtr
    Dim lpPath As LongPtr, lpUrl As LongPtr
    Dim siRes As IShellItem
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(1)
    'Note: Must use StrPtr in WinDevLib
    tFilt(0).pszName = "Common media files"
    tFilt(0).pszSpec = "*.mp4; *.mkv; *.m2ts; *.avi; *.asf; *.hevc; *.wmv; *.wma; *.mp3; *.flac; *.m4a; *.m4v; *.mov; *.wav; *.3gp; *.3gpp; *.3gp2; *.3g2; *.aac; *.adts; *.sami; *.smi"
    tFilt(1).pszName = "All Files"
    tFilt(1).pszSpec = "*.*"
    With fod
        .SetTitle "Open media file..."
        .SetOptions FOS_PATHMUSTEXIST
        .SetFileTypes 2, VarPtr(tFilt(0))
        On Error Resume Next
        .Show Me.hwnd
        .GetResult siRes
        On Error GoTo 0
        If (siRes Is Nothing) = False Then
            'siRes.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpAbsPath
             siRes.GetDisplayName SIGDN_URL, lpUrl
             siRes.GetDisplayName SIGDN_FILESYSPATH, lpPath
             mFile = LPWSTRtoStr(lpUrl)
             If lpPath Then
                 Text1.Text = LPWSTRtoStr(lpPath)
             Else
                 Text1.Text = mFile
             End If
             ucSimplePlayer1.PlayMediaFile mFile
         End If
    End With
    
End Sub

Private Sub Command2_Click() 'Handles Command2.Click
    ucSimplePlayer1.PlayMediaFile Text1.Text

    Check1.Value = vbUnchecked
End Sub

#If Win64 Then
Public Function Time2String(ByVal curNanoSec As LongLong, Optional ByVal strTimeFormat As String = "hh:mm:ss") As String
#Else
Public Function Time2String(ByVal curNanoSec As Currency, Optional ByVal strTimeFormat As String = "hh:mm:ss") As String
#End If

    On Error Resume Next

    Dim lngHour As Long
    Dim lngMinu As Long
    Dim dblSeco As Double

    dblSeco = CDbl(curNanoSec / 10000000)

    lngHour = dblSeco \ 3600
    dblSeco = dblSeco Mod 3600
    lngMinu = dblSeco \ 60
    dblSeco = dblSeco Mod 60

    Time2String = Format$(TimeSerial(lngHour, lngMinu, dblSeco), strTimeFormat)

End Function
Public Function TimeSec2String(ByVal dblSeco As Double, Optional ByVal _
    strTimeFormat As String = "hh:mm:ss") As String

    On Error Resume Next

    Dim lngHour As Long
    Dim lngMinu As Long


    lngHour = dblSeco \ 3600
    dblSeco = dblSeco Mod 3600
    lngMinu = dblSeco \ 60
    dblSeco = dblSeco Mod 60

    TimeSec2String = Format$(TimeSerial(lngHour, lngMinu, dblSeco), strTimeFormat)

End Function
Private Sub Check1_Click() 'Handles Check1.Click
    ucSimplePlayer1.Muted = (Check1.Value = vbChecked)
End Sub

Private Sub Form_Resize()
On Error Resume Next
ucSimplePlayer1.Width = Me.ScaleWidth - 30
ucSimplePlayer1.Height = Me.ScaleHeight - 90
End Sub

 #If Win64 Then
 Private Sub ucSimplePlayer1_PlaybackStart(ByVal cyDuration As LongLong) 'Handles ucSimplePlayer1.PlaybackStart
    Dim nSec As LongLong
 #Else
 Private Sub ucSimplePlayer1_PlaybackStart(ByVal cyDuration As Currency) 'Handles ucSimplePlayer1.PlaybackStart
    Dim nSec As Currency
 #End If
    Label4.Caption = Time2String(cyDuration)
    nSec = cyDuration / 10000000
    If nSec < 1 Then nSec = 1
    mPlay = True
    Slider1.Value = 0
    Slider1.Max = CLng(nSec)
    Slider1.TickFrequency = CLng(nSec) \ 10
    Slider1.LargeChange = CLng(nSec) \ 10
    Timer1.Interval = 250
    Timer1.Enabled = True
    ucSimplePlayer1.volume = Slider2.Value / 100
End Sub

  
Private Sub ucSimplePlayer1_PlayerClick(ByVal Button As Long)
           If Button = 2 Then
            Dim nVid As Long, nAud As Long
            Dim sVidL() As String, sVidN() As String
            Dim sAudL() As String, sAudN() As String
            ucSimplePlayer1.GetVideoStreams sVidN, sVidL, nVid
            ucSimplePlayer1.GetAudioStreams sAudN, sAudL, nAud
            Dim hMenu As LongPtr
            Dim hSubV As LongPtr, hSubA As LongPtr
            hMenu = CreatePopupMenu()
            hSubV = CreateMenu()
            hSubA = CreateMenu()
            Dim i As Long
            Dim mii As MENUITEMINFO
            mii.cbSize = LenB(mii)
            With mii
                .fMask = MIIM_ID Or MIIM_STRING Or MIIM_SUBMENU
                If nVid = 0 Then
                    .fMask = .fMask Or MIIM_STATE
                    .fState = MFS_DISABLED
                End If
                .wID = 1000
                .dwTypeData = StrPtr("Video tracks")
                .cch = Len("Video tracks")
                .hSubMenu = hSubV
                Call InsertMenuItem(hMenu, 0, True, mii)
                
                .fMask = MIIM_ID Or MIIM_STRING Or MIIM_SUBMENU
                If nAud = 0 Then
                    .fMask = .fMask Or MIIM_STATE
                    .fState = MFS_DISABLED
                End If
                .wID = 1001
                .dwTypeData = StrPtr("Audio tracks")
                .cch = Len("Audio tracks")
                .hSubMenu = hSubA
                Call InsertMenuItem(hMenu, 1, True, mii)
                
                .fMask = MIIM_ID Or MIIM_TYPE
                .fType = MFT_SEPARATOR
                .wID = 0
                Call InsertMenuItem(hMenu, 2, True, mii)
                
                .fMask = MIIM_ID Or MIIM_STRING
                .wID = 1003
                .dwTypeData = StrPtr("Lock aspect ratio")
                .cch = Len("Lock aspect ratio")
                If ucSimplePlayer1.PreserveAspectRatio Then
                    .fMask = .fMask Or MIIM_STATE
                    .fState = MFS_CHECKED
                End If
                Call InsertMenuItem(hMenu, 3, True, mii)
                
                .fMask = MIIM_ID Or MIIM_STRING
                .wID = 1002
                If ucSimplePlayer1.FullScreen Then
                    .dwTypeData = StrPtr("Exit fullscreen")
                    .cch = Len("Exit fullscreen")
                Else
                    .dwTypeData = StrPtr("Enter fullscreen")
                    .cch = Len("Enter fullscreen")
                End If
                Call InsertMenuItem(hMenu, 4, True, mii)
            End With
            
            Dim sLbl As String
            'Populate track menus
            If nVid > 0 Then
                For i = 0 To nVid - 1
                    With mii
                        .fMask = MIIM_ID Or MIIM_STRING
                        .wID = 2001 + i
                        sLbl = "(" & CStr(i + 1) & ") [" & IIf(sVidL(i) = "", "unk", sVidL(i)) & "]: " & IIf(sVidN(i) = "", "Video track #" & CStr(i + 1), sVidN(i))
                        .dwTypeData = StrPtr(sLbl)
                        .cch = Len(sLbl)

                        Call InsertMenuItem(hSubV, i, True, mii)
                    End With
                Next
                CheckMenuRadioItem hSubV, 0, i, ucSimplePlayer1.ActiveVideoStream - 1, MF_BYPOSITION
            End If
            If nAud > 0 Then
                For i = 0 To nAud - 1
                    With mii
                        .fMask = MIIM_ID Or MIIM_STRING
                        .wID = 9001 + i
                        sLbl = "(" & CStr(i + 1) & ") [" & IIf(sAudL(i) = "", "unk", sAudL(i)) & "]: " & IIf(sAudN(i) = "", "Audio track #" & CStr(i + 1), sAudN(i))
                        .dwTypeData = StrPtr(sLbl)
                        .cch = Len(sLbl)

                        Call InsertMenuItem(hSubA, i, True, mii)
                    End With
                Next
                CheckMenuRadioItem hSubA, 0, i, ucSimplePlayer1.ActiveAudioStream - 1, MF_BYPOSITION
            End If
            
            Dim pt As Point
            GetCursorPos pt
            Dim idCmd As Long
            Dim n As Long
            idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, pt.X, pt.Y, 0, Me.hwnd, 0)
            
            'Be careful adding commands with how this is processed
            Select Case idCmd
                Case 1002
                    If ucSimplePlayer1.FullScreen Then
                        ucSimplePlayer1.FullScreen = False
                    Else
                        ucSimplePlayer1.FullScreen = True
                    End If
                    
                Case 1003
                    If ucSimplePlayer1.PreserveAspectRatio Then
                        ucSimplePlayer1.PreserveAspectRatio = False
                    Else
                        ucSimplePlayer1.PreserveAspectRatio = True
                    End If
                    
                Case Is > 9000 'audio
                    n = idCmd - 9000
                    If n <> ucSimplePlayer1.ActiveAudioStream Then
                        ucSimplePlayer1.ActiveAudioStream = n
                    End If
                    
                Case Is > 2000 'video
                    n = idCmd - 2000
                    If n <> ucSimplePlayer1.ActiveVideoStream Then
                        ucSimplePlayer1.ActiveVideoStream = n
                    End If
            End Select
            
        End If
End Sub

Private Sub Command3_Click() 'Handles Command3.Click
    If ucSimplePlayer1.Paused = True Then
        ucSimplePlayer1.Paused = False
    Else
        ucSimplePlayer1.Paused = True
    End If
End Sub

Private Sub Command4_Click() 'Handles Command4.Click
    Timer1.Enabled = False
    mPlay = False
    ucSimplePlayer1.StopPlayback
End Sub

Private Sub Timer1_Timer() 'Handles Timer1.Timer
    Dim nSec As Currency
    nSec = ucSimplePlayer1.position / 1000 '0000
    If mNoUpdate = False Then Slider1.Value = nSec
    Label1.Caption = TimeSec2String(nSec)
    ' Debug.Print "updatepost " & nSec & " / " & Slider1.Max
End Sub

Private Sub Slider1_Click() 'Handles Slider1.Click
    If mPlay Then
        Debug.Print "SetPos Click"
        ucSimplePlayer1.position = CCur(Slider1.Value) * 1000
        mDropScroll = True
    End If
End Sub
Private Sub Slider1_Scroll() 'Handles Slider1.Scroll
    If mPlay Then
        If mDropScroll Then
            mDropScroll = False
        Else
            If mNoUpdate = False Then
                Debug.Print "SetPos Scroll"
                ucSimplePlayer1.position = CCur(Slider1.Value) * 1000
            End If
        End If
    End If
End Sub
Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Handles Slider1.MouseDown
    mNoUpdate = True
End Sub
Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Handles Slider1.MouseUp
    mNoUpdate = False
End Sub


Private Sub Slider2_Click() 'Handles Slider2.Click
    ucSimplePlayer1.volume = Slider2.Value / 100
End Sub

Private Sub Slider1_Validate(Cancel As Boolean) 'Handles Slider1.Validate
    
End Sub
 

Private Sub Command5_Click() 'Handles Command5.Click
    If ucSimplePlayer1.Rate > 0.24 Then
        ucSimplePlayer1.Rate = ucSimplePlayer1.Rate - 0.25
    End If
End Sub

Private Sub Command6_Click() 'Handles Command6.Click
    ucSimplePlayer1.Rate = ucSimplePlayer1.Rate + 0.25
End Sub

Private Sub Command7_Click() 'Handles Command7.Click
    ucSimplePlayer1.FrameStep
End Sub
Private Sub ucSimplePlayer1_PlayerWheelScroll(ByVal delta As Integer, ByVal keyState As Integer) 'Handles ucSimplePlayer1.PlayerWheelScroll
If delta < 0 Then
    If Slider2.Value >= 5 Then
        Slider2.Value = Slider2.Value - 5
    ElseIf Slider2.Value > 0 Then '1-4, set to 0
        Slider2.Value = 0
    End If
Else
    If Slider2.Value <= 95 Then
        Slider2.Value = Slider2.Value + 5
    ElseIf Slider2.Value > 95 Then '96-99, set to 100
        Slider2.Value = 100
    End If
End If
ucSimplePlayer1.volume = Slider2.Value / 100
End Sub

Private Sub Form_Load() 'Handles Form.Load
SHAutoComplete Text1.hwnd, SHACF_FILESYSTEM
Dim dwStyle As LongPtr
dwStyle = GetWindowLong(Slider1.hwnd, GWL_STYLE)
dwStyle = dwStyle Or TBS_DOWNISLEFT
SetWindowLong Slider1.hwnd, GWL_STYLE, dwStyle
dwStyle = GetWindowLong(Slider2.hwnd, GWL_STYLE)
dwStyle = dwStyle Or TBS_DOWNISLEFT
SetWindowLong Slider2.hwnd, GWL_STYLE, dwStyle
'the above makes using the mousewheel much more natural
'and matches how the control does it
End Sub
'VB6-only helpers
Private Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String

SysReAllocStringW VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
    lPtr = 0
End If

End Function
