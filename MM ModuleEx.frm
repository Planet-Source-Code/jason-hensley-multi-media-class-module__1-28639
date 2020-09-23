VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   " MM Module Example"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   0.931
      TabIndex        =   12
      Top             =   1800
      Width           =   3255
      Begin VB.PictureBox P2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   0
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   2520
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   1590
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Playing: "
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Position:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00\00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFC0&
      X1              =   120
      X2              =   3600
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      X1              =   120
      X2              =   3600
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MM Module Example"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFC0&
      Height          =   2655
      Left            =   120
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00\00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFC0&
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MM As New MusicModule
Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
MM.StopPlay
End Sub

Private Sub Label2_Click()
On Error Resume Next
'If a song is not selected then exit sub
If List2.text = "" Then MsgBox "Please select a song to play!", , "Error": Exit Sub
'Stop if a song is playing so you don't play 2 at a time
MM.StopPlay
'Load the filename of the song
MM.FileName = List2

'Play the song
MM.Play
'Load the current song playing
Text1 = "Playing: " & List1
'Make sure the song has had enough time to load
'to get status information
MM.TimeOut 0.5
'Load the duration in seconds
P1.ScaleWidth = MM.DurationInSec
End Sub

Private Sub Label3_Click()
MM.StopPlay
End Sub

Private Sub Label4_Click()
With Label4
If .Caption = "Pause" Then
.Caption = "Resume"
MM.Pause
Else
.Caption = "Pause"
MM.ResumePlay
End If
End With
End Sub

Private Sub Label5_Click()
C.Filter = "M3U Playlist (*.m3u)|*.m3u|MP3 Files (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|Midi Files (*.mid)|*.mid|All Files (*.*)|*.*"
C.ShowOpen
If C.FileName = "" Then Exit Sub
If C.FileName = " " Then Exit Sub
If LCase(Right(C.FileName, 3)) = LCase("m3u") Then
List1.Clear
List2.Clear
Call MM.OpenPlaylist(C.FileName, List2)
Call MM.ListNoChar(List1, List2)
Else
List2.AddItem C.FileName
Call MM.ListSingleNoChar(List1, List2)
End If
C.FileName = ""
End Sub

Private Sub Label6_Click()
C.Filter = "M3U Playlist (*.m3u)|*.m3u"
C.ShowSave
If C.FileName = "" Then Exit Sub
If C.FileName = " " Then Exit Sub

Call MM.SavePlaylist(C.FileName, List2)

C.FileName = ""
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MM.FormMove Me
End Sub

Private Sub Label8_Click()
Me.WindowState = 1
End Sub

Private Sub Label9_Click()
Unload Me
End
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
List2.ListIndex = List1.ListIndex
Label2_Click
End Sub

Private Sub P1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
P1.CurrentX = x
P2.Left = P1.CurrentX
MM.ChangePosition P1.CurrentX
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If MM.IsPlaying = False Then Exit Sub
Label1.Caption = MM.FormatPosition & "\" & MM.FormatDuration
Label10.Caption = MM.FormatTimeRemaining & "\" & MM.FormatDuration
P1.CurrentX = MM.PositioninSec
P2.Left = P1.CurrentX
If MM.EndOfSong = True Then
If List1.ListCount = 1 Then
Exit Sub
Else
List1.ListIndex = Val(List1.ListIndex) + 1
Label2_Click
End If
End If
End Sub
