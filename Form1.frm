VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
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
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   600
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Frame Step"
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin ucSimplePlayerDemo.ucSimplePlayer ucSimplePlayer1 
      Height          =   5655
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   9615
      _ExtentX        =   15266
      _ExtentY        =   10398
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

   Private mFile As String
   Private bPause As Boolean
   Private mPlay As Boolean
   Private mNoUpdate As Boolean
   
   Private Sub Command1_Click() 'Handles Command1.Click
       Dim fod As IFileOpenDialog
       Set fod = New FileOpenDialog
       Dim lpAbsPath As LongPtr
       Dim lpPath As LongPtr
       Dim siRes As IShellItem
       Dim tFilt() As COMDLG_FILTERSPEC
       ReDim tFilt(1)
       'Note: Must use StrPtr in WinDevLib
       tFilt(0).pszName = "Common media files"
       tFilt(0).pszSpec = "*.mp4; *.mkv; *.m2ts; *.avi; *.asf; *.hevc; *.wmv; *.wma; *.mp3; *.m4a; *.m4v; *.mov; *.wav; *.3gp; *.3gpp; *.3gp2; *.3g2; *.aac; *.adts; *.sami; *.smi"
       tFilt(1).pszName = "All Files"
       tFilt(1).pszSpec = "*.*"
       With fod
           .SetTitle "Open media file..."
           .SetOptions FOS_PATHMUSTEXIST
           .SetFileTypes 2, VarPtr(tFilt(0))
           On Error Resume Next
           .Show Me.hWnd
           .GetResult siRes
           On Error GoTo 0
           If (siRes Is Nothing) = False Then
               'siRes.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpAbsPath
               siRes.GetDisplayName SIGDN_URL, lpPath
               mFile = LPWSTRtoStr(lpPath)
               Text1.Text = mFile
           End If
           
       End With
       
   End Sub
   
   Private Sub Command2_Click() 'Handles Command2.Click
       ucSimplePlayer1.PlayMediaFile Text1.Text

       Check1.Value = vbUnchecked
       Command3.Caption = "Pause"
       bPause = False
   End Sub
   Public Function Time2String(ByVal curNanoSec As Currency, Optional ByVal _
       strTimeFormat As String = "hh:mm:ss") As String

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
Debug.Print Me.Width, ucSimplePlayer1.Width
End Sub

   Private Sub ucSimplePlayer1_PlaybackStart(ByVal cyDuration As Currency) 'Handles ucSimplePlayer1.PlaybackStart
       Label4.Caption = Time2String(cyDuration)
       Dim nSec As Currency
       nSec = cyDuration / 10000000
       mPlay = True
       Slider1.Value = 0
       Slider1.Max = nSec
       Slider1.TickFrequency = nSec / 10
       Timer1.Interval = 250
       Timer1.Enabled = True
       ucSimplePlayer1.Volume = Slider2.Value / 100
   End Sub
   
   Private Sub Command3_Click() 'Handles Command3.Click
       If bPause Then
           Command3.Caption = "Pause"
           bPause = False
           ucSimplePlayer1.Paused = False
       Else
           Command3.Caption = "Unpause"
           bPause = True
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
       nSec = ucSimplePlayer1.Position / 1000 '0000
       If mNoUpdate = False Then Slider1.Value = nSec
       Label1.Caption = TimeSec2String(nSec)
       ' Debug.Print "updatepost " & nSec & " / " & Slider1.Max
   End Sub
   
   Private Sub Slider1_Click() 'Handles Slider1.Click
       If mPlay Then
           ucSimplePlayer1.Position = CCur(Slider1.Value) * 1000 '0000
       End If
       
   End Sub

   Private Sub Slider2_Click() 'Handles Slider2.Click
       ucSimplePlayer1.Volume = Slider2.Value / 100
   End Sub
   
   Private Sub Slider1_Validate(Cancel As Boolean) 'Handles Slider1.Validate
       
   End Sub
   
   Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Handles Slider1.MouseDown
       mNoUpdate = True
   End Sub
   
   Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'Handles Slider1.MouseUp
       mNoUpdate = False
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
    
 'VB6-only helpers
 Private Function LPWSTRtoStr(lPtr As LongPtr, Optional ByVal fFree As Boolean = True) As String

SysReAllocStringW VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
    lPtr = 0
End If

 End Function
