VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bob Can Speak  - by Tangent UFS"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7680
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   7680
      TabIndex        =   8
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdShutUp 
      Caption         =   "Shut &Up Bob"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3720
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animated Eyes"
      Height          =   3615
      Left            =   7680
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   5
         Left            =   2520
         Picture         =   "Form1.frx":030A
         Top             =   2520
         Width           =   2205
      End
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   4
         Left            =   2520
         Picture         =   "Form1.frx":08B6
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   3
         Left            =   120
         Picture         =   "Form1.frx":0E7D
         Top             =   2520
         Width           =   2205
      End
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   2
         Left            =   120
         Picture         =   "Form1.frx":141F
         Top             =   360
         Width           =   2205
      End
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   1
         Left            =   2520
         Picture         =   "Form1.frx":19C3
         Top             =   360
         Width           =   2205
      End
      Begin VB.Image imgEyes 
         Height          =   960
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":1F61
         Top             =   1440
         Width           =   2205
      End
   End
   Begin VB.PictureBox picBkgr 
      Height          =   5775
      Left            =   120
      Picture         =   "Form1.frx":24F7
      ScaleHeight     =   5715
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   360
      Width           =   7335
      Begin HTTSLibCtl.TextToSpeech SS 
         Height          =   2895
         Left            =   2325
         OleObjectBlob   =   "Form1.frx":89635
         TabIndex        =   3
         Top             =   2275
         Width           =   2895
      End
      Begin VB.Image imgX 
         Height          =   1095
         Left            =   2695
         Top             =   785
         Width           =   2055
      End
   End
   Begin VB.TextBox txtSpeak 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6240
      Width           =   6375
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "&Speak Bob Speak"
      Height          =   1575
      Left            =   6600
      TabIndex        =   0
      Top             =   6240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()
'== open a speak file ===
On Error Resume Next
Dim BS, xT, F1, F2, Fltr As String
txtSpeak = ""
CD.DefaultExt = "bob"
F1 = "Bob Files (*.bob)|*.bob"
F2 = "All Files (*.*)|*.*"
Fltr = F1 & "|" & F2
CD.Filter = Fltr
CD.ShowOpen
Open CD.FileName For Input As #1
Do While Not EOF(1)
   Input #1, xT
   BS = BS + xT
Loop
txtSpeak = BS
Close #1
End Sub

Private Sub cmdSave_Click()
Dim BS, xT, F1, F2, Fltr As String
CD.DefaultExt = "bob"
F1 = "Bob Files (*.bob)|*.bob"
F2 = "All Files (*.*)|*.*"
Fltr = F1 & "|" & F2
CD.Filter = Fltr
CD.ShowSave
Open CD.FileName For Output As #1
Write #1, Trim(txtSpeak.Text)
Close #1
End Sub

Private Sub cmdShutUp_Click()
SS.StopSpeaking
txtSpeak = ""
End Sub

Private Sub cmdSpeak_Click()
List1.AddItem SS.Style(5)
SS.Speak txtSpeak ' <== say it
End Sub


Private Sub Form_Activate()
Me.Top = 120
End Sub

Private Sub Form_Load()
Dim t As Integer
For t = 1 To 10
List1.AddItem "Mode ID:: " & SS.ModeID(t)
List1.AddItem "Mode Name:: " & SS.ModeName(t)
List1.AddItem "Age:: " & SS.Age(t)
List1.AddItem "Dialect:: " & SS.dialect(t)
List1.AddItem "Mode:: " & SS.CurrentMode
List1.AddItem "IntrFace:: " & SS.Interfaces(t)
List1.AddItem "Eng Feat:: " & SS.EngineFeatures(t)
List1.AddItem "Features:: " & SS.Features(t)
List1.AddItem "Gender:: " & SS.Gender(t)
List1.AddItem "Pname:: " & SS.ProductName(t)
List1.AddItem "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
Next
End Sub

Private Sub SS_Speak(ByVal Text As String, ByVal App As String, ByVal thetype As Long)
List1.AddItem thetype
End Sub

Private Sub Timer1_Timer()
'=== eyeball animation ===
Static IV As Integer
Dim x As Integer
x = Int(Rnd(1) * 5) + 0
imgX = imgEyes(x)
IV = Int(Rnd(1) * 2050) + 575
Timer1.Interval = IV
End Sub
