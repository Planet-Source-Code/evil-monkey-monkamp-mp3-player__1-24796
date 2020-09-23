VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FFBEC4C3-839E-11D1-85FE-0020AFE4DE54}#1.0#0"; "MP3ENC.OCX"
Begin VB.Form Main 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monkamp MP3 Encoder"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton hElPeRoNy 
      BackColor       =   &H00C00000&
      Caption         =   "Help"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton image11 
      BackColor       =   &H00C00000&
      Caption         =   "Stop"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton image9 
      BackColor       =   &H00C00000&
      Caption         =   "Encode"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6000
      Top             =   1320
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C00000&
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Text            =   "0"
      Top             =   4200
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5880
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Output"
      Filter          =   "MP3 Files (*.mp3)|*.mp3"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Input"
      Filter          =   "WAV Files (*.mp3)|*.wav"
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C00000&
      Caption         =   "Info:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1215
      TabIndex        =   13
      Top             =   1515
      Width           =   4725
      Begin VB.TextBox txtpod 
         BackColor       =   &H00C00000&
         Height          =   285
         Left            =   525
         TabIndex        =   19
         Top             =   1260
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.Label Lbl1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   765
         Visible         =   0   'False
         Width           =   4380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frames:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label LblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   105
         TabIndex        =   16
         Top             =   450
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label LblActFrame 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   810
         TabIndex        =   15
         Top             =   225
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.TextBox OutPutPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   270
      Width           =   3630
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Bitrate:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   105
      TabIndex        =   4
      Top             =   1515
      Width           =   1095
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "32 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   255
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "56 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   10
         Top             =   450
         Width           =   915
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "64 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   45
         TabIndex        =   9
         Top             =   645
         Width           =   900
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C00000&
         Caption         =   "96 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   8
         Top             =   840
         Width           =   915
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C00000&
         Caption         =   "112 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   7
         Top             =   1035
         Width           =   1005
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C00000&
         Caption         =   "128 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   6
         Top             =   1230
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C00000&
         Caption         =   "256 Kbps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   45
         TabIndex        =   5
         Top             =   1425
         Width           =   990
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   3075
         Picture         =   "Main.frx":030A
         Top             =   1620
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   3990
      TabIndex        =   2
      Top             =   15
      Width           =   3825
      Begin VB.CommandButton image5 
         BackColor       =   &H00C00000&
         Caption         =   "Browse"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   120
      TabIndex        =   1
      Top             =   15
      Width           =   3825
      Begin VB.CommandButton image2 
         BackColor       =   &H00C00000&
         Caption         =   "Browse"
         Height          =   495
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox InputPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   255
         Width           =   3630
      End
   End
   Begin MP3ENCLib.Mp3Enc Mp3 
      Height          =   450
      Left            =   6720
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
      _Version        =   65536
      _ExtentX        =   794
      _ExtentY        =   794
      _StockProps     =   0
   End
   Begin VB.Menu BrowseMnu 
      Caption         =   "Browse Menu"
      Visible         =   0   'False
      Begin VB.Menu Single 
         Caption         =   "&Single File (Input)"
      End
      Begin VB.Menu BrowseDir 
         Caption         =   "&Directory (Input)"
      End
   End
   Begin VB.Menu BrowseMnu2 
      Caption         =   "Browse Menu"
      Visible         =   0   'False
      Begin VB.Menu SingleOut 
         Caption         =   "&Single File (Output)"
      End
      Begin VB.Menu BrowseDirOut 
         Caption         =   "&Directory (Output)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
hWndOwner      As Long
pIDLRoot       As Long
pszDisplayName As Long
lpszTitle      As Long
ulFlags        As Long
lpfnCallback   As Long
lParam         As Long
iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Dim kbits As Variant
Private Sub BrowseDir_Click()
DirBox "Select Input Directory", ID$
InputPath.Text = ID$
Dim count As Long
Dim filecount As Boolean
Dim I
On Error Resume Next
Pod = ID$ 'Name of my Directory
If filecount = True Then
For I = 0 To 9
    For a = 0 To 9
        fname = Dir(Pod & I & "\" & a & "\*.mp3")
        While Not fname = ""
            fname = Dir
            Counter = Counter + 1
        Wend
    Next a
Next I
txtpod.Text = Counter 'name For my textbox or richtextbox
End If
End Sub
Private Sub BrowseDirOut_Click()
DirBox "Select OutPut Directory", O$
OutPutPath.Text = O$
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ontop.MakeTopMost hWnd
Mp3.Authorize "James Green", "1279606713"
Mp3.BitRate = 128000
Mp3.AllowDownSample = True
Mp3.DownMix = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Mp3.Stop
End Sub
Private Sub InBrowse_Click()

End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Visible = False
image2.Visible = True
End Sub


Private Sub hElPeRoNy_Click()
MsgBox "You need to select an input then save an output.  Then choose you bitrate and click encode.", vbOKOnly, "Help"
End Sub

Private Sub image2_Click()
CommonDialog1.ShowOpen
InputPath.Text = CommonDialog1.Filename
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image4.Visible = False
image5.Visible = True
End Sub

Private Sub image5_Click()
CommonDialog2.ShowSave
OutPutPath.Text = CommonDialog2.Filename
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.Visible = False
Image7.Visible = True
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.Visible = True
Image7.Visible = False
About.Show
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image8.Visible = False
image9.Visible = True
End Sub

Private Sub image9_Click()
On Error GoTo WaveError
'Do Checks First
If InputPath = "" Then: MsgBox ("Input Path Empty"), vbExclamation, "Error": Exit Sub
If OutPutPath = "" Then: MsgBox ("Output Path Empty"), vbExclamation, "Error": Exit Sub
'Encode It
Mp3.Open InputPath.Text, OutPutPath.Text
Text3.Text = Mp3.Encode
Label1.Visible = True
LblActFrame.Visible = True
LblTotal.Visible = True
Lbl1.Visible = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
InBrowse.Enabled = False
OutBrowse.Enabled = False
BtnEncode.Enabled = False
About.Show
Exit Sub
WaveError:
On Error Resume Next
Exit Sub
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Mp3.Stop
End Sub
Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
Private Sub InputPath_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
InputPath.ToolTipText = InputPath.Text
End Sub

Private Sub Mp3_ActFrame(ByVal ActFrame As Long)
Dim Name As String
Name = StripPath(InputPath.Text)
LblActFrame.Caption = ActFrame
LblTotal.Caption = (ActFrame * 100 / Mp3.GetFrameCount) \ 1 & " %"
LblTotal.Caption = LblTotal.Caption & " Complete: " & (Timer - start_time) / (ActFrame * 0.026)
Lbl1.Caption = "Encoding " & Name & " @ " & Mp3.BitRate & " Kbps"
End Sub
Private Sub OutPutPath_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
OutPutPath.ToolTipText = OutPutPath.Text
End Sub
Private Sub OutBrowse_Click()
Me.PopupMenu Me.BrowseMnu2
End Sub
'******************************************************************
Private Sub DirBox(Msg As String, Directory As String)
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    'Change this to set what info is displayed.
    szTitle = Msg
    With tBrowseInfo
       .hWndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       Directory = sBuffer
    End If
End Sub
Private Sub Option1_Click()
On Error Resume Next
kbits = 32000
Mp3.BitRate = kbits
End Sub
Private Sub Option2_Click()
On Error Resume Next
kbits = 56000
Mp3.BitRate = kbits
End Sub
Private Sub Option3_Click()
On Error Resume Next
kbits = 64000
Mp3.BitRate = kbits
End Sub
Private Sub Option4_Click()
On Error Resume Next
kbits = 96000
Mp3.BitRate = kbits
End Sub
Private Sub Option5_Click()
On Error Resume Next
kbits = 112000
Mp3.BitRate = kbits
End Sub
Private Sub Option6_Click()
On Error Resume Next
kbits = 128000
Mp3.BitRate = kbits
End Sub
Private Sub Option7_Click()
On Error Resume Next
kbits = 256000
Mp3.BitRate = kbits
End Sub
Private Sub Single_Click()
CommonDialog1.ShowOpen
InputPath.Text = CommonDialog1.Filename
End Sub
Private Sub SingleOut_Click()
CommonDialog2.ShowSave
OutPutPath.Text = CommonDialog2.Filename
End Sub
Function StripPath(T$) As String
    Dim x%, ct%
    StripPath$ = T$
    x% = InStr(T$, "\")


    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, T$, "\")
    Loop
    If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function

Private Sub Timer1_Timer()
If LblTotal.Caption = "97 % Complete:" Then LblTotal.Caption = "Complete"
End Sub

