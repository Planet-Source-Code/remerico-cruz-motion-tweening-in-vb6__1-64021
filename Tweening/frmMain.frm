VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Motion Tween Sample"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Move This Form!"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Red"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "Go figure"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbAnimType 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0013
      TabIndex        =   2
      Text            =   "easeoutquad"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move Me"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer tmrTween 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   2040
      Picture         =   "frmMain.frx":0052
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hello"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
  TweenThis Check1
End Sub

Private Sub Command1_Click()
  TweenThis Command1
End Sub

Private Sub Command2_Click()
  TweenThis Command2, Frame1.Width, Frame1.Height
End Sub

Private Sub Command3_Click()
  TweenThis Command3
End Sub

Private Sub Command4_Click()
  If Me.WindowState = vbNormal Then TweenThis Me, Screen.Width, Screen.Height
End Sub

Private Sub Frame1_Click()
  TweenThis Frame1
End Sub

Private Sub Image1_Click()
  TweenThis Image1
End Sub

Private Sub Label1_Click()
  TweenThis Label1
End Sub

Private Sub Text1_Click()
  TweenThis Text1
End Sub

Private Sub tmrTween_Timer()
  ' Calls the timer_tick subroutine from the tweening module.
  Timer_Tick
End Sub

Sub TweenThis(xControl As Object, Optional xWidth As Long, Optional xHeight As Long)
  ' Calls the main tweening subroutine from the tweening module.
  
  If xWidth = 0 Then xWidth = Me.ScaleWidth
  If xHeight = 0 Then xHeight = Me.ScaleHeight
  
  ' Starts the tweening
  StartTween tmrTween, xControl, GoRandom(0, xWidth - xControl.Width), GoRandom(0, xHeight - xControl.Height), cmbAnimType.Text, 40
End Sub

Function GoRandom(xDown As Long, xUp As Long)
  ' Generates a random number
  Randomize Timer
  GoRandom = Int((xUp - 1 + 1) * Rnd + xDown)
End Function
