VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   120
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      TickStyle       =   3
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Stopped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   8520
      X2              =   8520
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   8640
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu asseria 
      Caption         =   "shit"
      Visible         =   0   'False
      Begin VB.Menu gvgvggvybgy 
         Caption         =   "&Options"
      End
      Begin VB.Menu break37463127856134 
         Caption         =   "-"
      End
      Begin VB.Menu ghgghfgfg 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub Form_Load()
    ontop.MakeTopMost hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Me.PopupMenu asseria
                End If
End Sub

Private Sub ghgghfgfg_Click()
End
End Sub

Private Sub gvgvggvybgy_Click()
frmOptions.Show
End Sub

Private Sub Label1_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Me.PopupMenu asseria
                End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Me.PopupMenu asseria
        End If
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Me.PopupMenu asseria
                End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "" & frmMain.Label1.Caption
Slider1.Max = frmMain.Slider3.Max
Slider1.Value = frmMain.Slider3.Value
End Sub

Private Sub Timer2_Timer()
If frmMain.MediaPlayer1.PlayState = mpPlaying Then Timer1.Enabled = True
If frmMain.MediaPlayer1.PlayState = mpStopped Then Timer1.Enabled = False
If frmMain.MediaPlayer1.PlayState = mpStopped Then Label1.Caption = "Stopped"
End Sub
