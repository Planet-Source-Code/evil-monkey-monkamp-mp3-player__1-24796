VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   5580
   ClientLeft      =   75
   ClientTop       =   -495
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C00000&
      Caption         =   "ID3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Edit an MP3s ID3 tag"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C00000&
      Caption         =   "Encode"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Encode MP3s"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   360
      Top             =   960
   End
   Begin VB.CommandButton imgPlay 
      BackColor       =   &H00C00000&
      Caption         =   "Play"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Play Selected Song"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton imgPrev 
      BackColor       =   &H00C00000&
      Caption         =   "Back"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Previous Song"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton imgNext 
      BackColor       =   &H00C00000&
      Caption         =   "Next"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Next Song"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton imgPause 
      BackColor       =   &H00C00000&
      Caption         =   "Pause"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Pause Current Song"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton imgStop 
      BackColor       =   &H00C00000&
      Caption         =   "Stop"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Stop Playing"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2640
      Top             =   1200
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Song Progress"
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Max             =   1000
      TickStyle       =   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "_"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Minimize"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C00000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clear the Playlist"
      Top             =   4920
      Width           =   615
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CheckBox cbShuffle 
      BackColor       =   &H00C00000&
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Shuffle Playlist"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox cbRepeat 
      BackColor       =   &H00C00000&
      Caption         =   "Repeat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Repeat Playlist"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C00000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save the playlist"
      Top             =   4920
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.ListBox lstFilenames 
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C00000&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add files to the playlist"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2400
      Top             =   3960
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   135
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Balance"
      Top             =   2760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393216
      Min             =   -3000
      Max             =   3000
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   135
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Volume"
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   238
      _Version        =   393216
      Min             =   -3000
      Max             =   0
      TickStyle       =   3
   End
   Begin MSComctlLib.ListView lstPlayList1 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Playlist"
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   32768
      BackColor       =   12582912
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   661
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SongTitle"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Artist"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Proggress/Control"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ToolTipText     =   "DRAG ME!!!"
      Top             =   0
      Width           =   3420
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   2880
      X2              =   2880
      Y1              =   600
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   0
      X2              =   3000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3240
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   0
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Song Status/Progress"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "No song currently playing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Song/Artist"
      Top             =   720
      Width           =   3015
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHidePlaylist 
         Caption         =   "&Hide Playlist/Options"
      End
      Begin VB.Menu mnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortBy 
         Caption         =   "Sort &By"
         Begin VB.Menu mnuArtist 
            Caption         =   "&Artist/Songtitle"
         End
         Begin VB.Menu mnuTitle 
            Caption         =   "&Songtitle"
         End
         Begin VB.Menu mnuFilename 
            Caption         =   "&Filename"
         End
         Begin VB.Menu mnuBar4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRandom 
            Caption         =   "&Randomize"
         End
         Begin VB.Menu mnuReverse 
            Caption         =   "R&everse Order"
         End
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "&Preferences"
         Begin VB.Menu mnuFont 
            Caption         =   "&Font"
         End
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Visible         =   0   'False
      Begin VB.Menu mnuHidePlaylist2 
         Caption         =   "&Hide Playlist/Options"
      End
      Begin VB.Menu mnuPref2 
         Caption         =   "&Preferences"
         Begin VB.Menu mnuFont2 
            Caption         =   "&Font"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "&Options"
         End
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Option Explicit
Option Compare Text

Const gMISSINGREGDATA = "$$EMPTY$$"
Const gEMPTYSTRING = ""
Const gKEYNAME = "SoftWare\MP3Player"   'reg settings

Public bSaveSettings As Boolean
Public bSavePlaylist As Boolean
Public bScrollTitle As Boolean
Public bConfirmDelete As Boolean
Public sPlayerName As String
Public sFormTitle As String

Enum enFontObject
    Playlist = 2
    Other = 4
End Enum


Dim Mp3Info As New clsMP3Info
Dim bMove As Boolean
Dim iOldX As Integer
Dim iOldY As Integer
Dim iCurrentIndex As Integer
Dim iPrevIndex As Integer
Dim sImageDir As String
Dim bStopPressed As Boolean
Dim sStartDir As String
Dim gFontColor1 As Long
Dim gFontColor2 As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Sub cbRepeat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub cbShuffle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub cmdClear_Click()
'Purpose:  Clear the playlists
    lstFilenames.Clear
    lstPlayList1.ListItems.Clear
End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub cmdOpen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

'For saving playlists
Private Sub cmdSave_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.Filter = "Playlist (*.mpl)|*.mpl"
    CommonDialog1.ShowSave
    If CommonDialog1.Filename <> "" Then
        SavePlaylist CommonDialog1.Filename
    End If
End Sub


Private Sub cmdOpen_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.DefaultExt = ".mpl"
    CommonDialog1.Filter = "MP3 Audio (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|MIDI Files (*.mid)|*.mid|PlayList (*.mpl)|*.mpl|All Files (*.*)|*.*"
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.ShowOpen
    'Check to see if it's a playlist or just a song and import accordingly
    If CommonDialog1.Filename <> "" And Right(CommonDialog1.Filename, 3) <> "mpl" Then
        If ParseFiles(CommonDialog1.Filename) Then
            lstFilenames.AddItem CommonDialog1.Filename
        End If
    ElseIf CommonDialog1.Filename <> "" And Right(CommonDialog1.Filename, 3) = "mpl" Then
        LoadPlayList CommonDialog1.Filename
    End If
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Command2_Click()
frmMain.Hide
Form1.Show
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Command3_Click()
Main.Show
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Command4_Click()
id3.Show
End Sub

Private Sub fuck_Click()
frmOptions.Show
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgStop_Click()
On Error GoTo penis
    MediaPlayer1.Stop
    MediaPlayer1.CurrentPosition = 0
    bStopPressed = True 'A hack that fixes a bug where the next song wouldn't play
                        'because the total time had a small decimal after it.  Mediaplayer
                        'would stop before the program realized it was at the end of the song
penis:
End Sub


Private Sub imgPlay_Click()
On Error GoTo pizza
    PlaySong lstPlayList1.SelectedItem
pizza:
End Sub

Private Sub Form_Load()
    ontop.MakeTopMost hWnd
    Timer2.Enabled = False
Me.Height = 5685
    Dim iFilenum As Integer
    Dim sTemp As String
    Dim iIndex As Integer
    
    'Set initial properties in the MediaPlayer control
    MediaPlayer1.AutoStart = False
    MediaPlayer1.AutoRewind = True
    MediaPlayer1.ShowAudioControls = True
    MediaPlayer1.Volume = 0
    'Set default values for program variables
    bStopPressed = True
    iCurrentIndex = 1
    gFontColor1 = -1
    gFontColor2 = -1
    
    'Sets the directory where the MP3 player is located (used for saving, loading, etc.)

    'Format the sStartDir so I know it has a '\' at the end

    'Sets the directory where the images are kept

    
    'Loads preferences from and ini file

    
    'Loads the images from the image directory into an imagelist control

    'Sets the images from the imagesList into the appropriate places on the form

    
    'Checks for the existance of the playlist that is automatically saved on close

        'loads the saved playlist's songs
 
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMove Then
        'moves the form
        frmMain.Move frmMain.Left + (x - iOldX), frmMain.Top + (y - iOldY)
    Else
        iOldX = x
        iOldY = y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        bMove = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Me.bSavePlaylist Then
        'Save songs to a file to be loaded when re-opened
        SavePlaylist sStartDir & "\playlist.mpl"
    Else
        'If they choose not to have the playlist automatically saved, and an old
        'playlist exists, it is deleted
        If Dir(sStartDir & "\playlist.mpl") <> "" Then
            Kill sStartDir & "\playlist.mpl"
        End If
    End If
    
    'Save the settings if the option is checked
    'Otherwise, delete the ini file...it is recreated with default values next
    'time the player is opened.
    If Me.bSaveSettings Then
        SaveINIFile
    Else
        If Dir(sImageDir & "skin.ini") <> "" Then
            Kill sImageDir & "skin.ini"
        End If
    End If
    
    'Destroy the ID3 parsing object
    Set Mp3Info = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
FormDrag Me
End If
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub imgNext_Click()
'Purpose:  Plays the next song, taking shuffle/repeat into account,
'          as well as the position in the playlist (if it's at the end, it moves
'          to the beginning)
    Dim iIndex As Integer
    
    'Check to see if shuffle is on
    If cbShuffle.Value = vbChecked Then
        Randomize 'Initialize random number generator with system time
        iIndex = GetRandomIndex(lstFilenames.ListCount - 1) 'Get random index
        While iIndex = iCurrentIndex 'Make sure random index isn't the current song
            iIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        Wend
        PlaySong iIndex 'Play randomly generated song
    Else
        'Check to see if it's at the end of the list
        If iCurrentIndex = lstFilenames.ListCount Then
            'If it's set to repeat, play the first song,
            'otherwise, just load it and stop
            If cbRepeat.Value = vbChecked Then
                PlaySong 1
            Else
                bStopPressed = True
                LoadSong 1
            End If
        Else
            'If it's not at the end, simply play the next song
            PlaySong iCurrentIndex + 1
        End If
    End If
End Sub

Private Sub imgPause_Click()
On Error GoTo cockmaster
    If MediaPlayer1.PlayState = mpPlaying Then
        MediaPlayer1.Pause
    Else
        MediaPlayer1.Play
    End If
cockmaster:
End Sub

Private Sub imgPrev_Click()
    Dim iIndex As Integer
        
    If cbShuffle.Value = vbChecked Then
        Randomize
        iIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        PlaySong iIndex
    Else
        If iCurrentIndex = 1 Then
            PlaySong lstPlayList1.ListItems.count
        Else
            PlaySong iCurrentIndex - 1
        End If
    End If
End Sub


Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub lstPlayList1_DblClick()
    PlaySong lstPlayList1.SelectedItem
    Timer2.Enabled = True
End Sub

Private Sub lstPlayList1_KeyDown(KeyCode As Integer, Shift As Integer)
'Purpose:  Removes a song from the playlist
    Dim iIndex As Integer
    Dim iCnt As Integer

    If KeyCode = vbKeyDelete Then
        DeleteFile CInt(lstPlayList1.SelectedItem.Text)
    End If
End Sub

Private Sub lstPlayList1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Set lstPlayList1.SelectedItem = lstPlayList1.HitTest(x, y)
        frmMain.PopupMenu mnuSort
    End If
End Sub

Private Sub lstPlayList1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCnt As Integer
    
    For iCnt = 1 To Data.Files.count
        If InStr(1, Data.Files(iCnt), ".") <> 0 Then
            If ParseFiles(Data.Files(iCnt)) Then
                lstFilenames.AddItem Data.Files(iCnt)
            End If
        Else
            AddDirectory Data.Files(iCnt)
        End If
    Next iCnt
End Sub


Private Sub mnuArtist_Click()
    Dim sInfo As New clsMP3Info
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sArtist As String
    Dim sTitle As String
    Dim sTemp As String
    
    For iCnt = 1 To lstPlayList1.ListItems.count
        sArtist = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        sTitle = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        iFoundAt = iCnt - 1
        sTemp = lstFilenames.List(iCnt - 1)
        For iPos = iCnt To lstPlayList1.ListItems.count
            If sArtist > lstPlayList1.ListItems(iPos).ListSubItems(2).Text Then
                sTemp = lstFilenames.List(iPos - 1)
                sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text
                sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                iFoundAt = iPos - 1
            ElseIf sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text Then
                If sTitle > lstPlayList1.ListItems(iPos).ListSubItems(1).Text Then
                    sTemp = lstFilenames.List(iPos - 1)
                    sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                    iFoundAt = iPos - 1
                End If
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt - 1)
        lstFilenames.List(iCnt - 1) = sTemp
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(1).Text = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(2).Text = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        lstPlayList1.ListItems(iCnt).ListSubItems(1).Text = sTitle
        lstPlayList1.ListItems(iCnt).ListSubItems(2).Text = sArtist
    Next iCnt
    
'    lstPlayList1.ListItems.Clear
'    For iCnt = 0 To lstFilenames.ListCount - 1
'        ParseFiles lstFilenames.List(iCnt)
'    Next iCnt
End Sub

Private Sub mnuBackColor_Click()
    CommonDialog1.Color = lstPlayList1.BackColor
    CommonDialog1.ShowColor
    lstPlayList1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuBackColor2_Click()
    CommonDialog1.Color = frmMain.BackColor
    CommonDialog1.ShowColor
    frmMain.BackColor = CommonDialog1.Color
    cbRepeat.BackColor = CommonDialog1.Color
    cbShuffle.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuClear_Click()
    lstFilenames.Clear
    lstPlayList1.ListItems.Clear
End Sub

Private Sub mnuDelete_Click()
    DeleteFile CInt(lstPlayList1.SelectedItem.Text)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFilename_Click()
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sTemp As String
    
    For iCnt = 0 To lstFilenames.ListCount - 1
        sTemp = lstFilenames.List(iCnt)
        iFoundAt = iCnt
        For iPos = iCnt To lstFilenames.ListCount - 1
            If sTemp > lstFilenames.List(iPos) Then
                sTemp = lstFilenames.List(iPos)
                iFoundAt = iPos
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt)
        lstFilenames.List(iCnt) = sTemp
    Next iCnt
    
    lstPlayList1.ListItems.Clear
    For iCnt = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iCnt)
    Next iCnt
End Sub

Private Sub mnuFont_Click()
    CommonDialog1.Flags = cdlCFBoth
    
    CommonDialog1.FontName = lstPlayList1.Font.Name
    CommonDialog1.FontSize = lstPlayList1.Font.Size
    CommonDialog1.ShowFont
    
    If CommonDialog1.FontSize > 8 Then
        CommonDialog1.FontSize = 8
    End If
    
    ChangeFont Playlist, CommonDialog1.FontName, CommonDialog1.FontSize, CommonDialog1.FontBold, CommonDialog1.FontItalic, gFontColor1
End Sub

Private Sub mnuFont2_Click()
    CommonDialog1.Flags = cdlCFBoth
    
    CommonDialog1.FontName = frmMain.cbRepeat.Font.Name
    CommonDialog1.FontSize = frmMain.cbRepeat.Font.Size
    CommonDialog1.ShowFont
    
    If CommonDialog1.FontSize > 8 Then
        CommonDialog1.FontSize = 8
    End If
    
    ChangeFont Other, CommonDialog1.FontName, CommonDialog1.FontSize, CommonDialog1.FontBold, CommonDialog1.FontItalic, gFontColor2
End Sub

Private Sub mnuFontColor_Click()
    CommonDialog1.Color = lstPlayList1.ForeColor
    CommonDialog1.ShowColor
    gFontColor1 = CommonDialog1.Color
    
    ChangeFont Playlist, , , , , gFontColor1
End Sub

Private Sub mnuFontColor2_Click()
    CommonDialog1.Color = frmMain.cbRepeat.ForeColor
    CommonDialog1.ShowColor
    gFontColor2 = CommonDialog1.Color
    
    ChangeFont Other, , , , , gFontColor2
End Sub

Private Sub mnuHidePlaylist_Click()
    If Me.Height = 3240 Then
        Me.Height = 5610
        mnuHidePlaylist.Caption = "&Hide Playlist/Options"
        mnuHidePlaylist2.Caption = "&Hide Playlist/Options"
    Else
        Me.Height = 3240
                mnuHidePlaylist.Caption = "S&how Playlist/Options"
                mnuHidePlaylist2.Caption = "S&how Playlist/Options"
    End If
    Me.Width = 3300
End Sub

Private Sub mnuHidePlaylist2_Click()
    If Me.Height = 3240 Then
        Me.Height = 5610
        mnuHidePlaylist2.Caption = "&Hide Playlist/Options"
        mnuHidePlaylist.Caption = "&Hide Playlist/Options"
    Else
        Me.Height = 3240
                mnuHidePlaylist2.Caption = "S&how Playlist/Options"
                mnuHidePlaylist.Caption = "S&how Playlist/Options"
    End If
    Me.Width = 3300
End Sub

Private Sub mnuOpen_Click()
    cmdOpen_Click
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuRandom_Click()
    Dim iCnt As Integer
    Dim sTemp As String
    Dim iNewIndex As Integer
    
    For iCnt = 0 To lstFilenames.ListCount - 1
        iNewIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        sTemp = lstFilenames.List(iNewIndex)
        lstFilenames.List(iNewIndex) = lstFilenames.List(iCnt)
        lstFilenames.List(iCnt) = sTemp
    Next iCnt
    
    lstPlayList1.ListItems.Clear
    For iCnt = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iCnt)
    Next iCnt
End Sub

Private Sub mnuReverse_Click()
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim sTemp As String
    
    iTop = 0
    iBottom = lstFilenames.ListCount - 1
    
    While iTop <= iBottom
        
        sTemp = lstFilenames.List(iBottom)
        lstFilenames.List(iBottom) = lstFilenames.List(iTop)
        lstFilenames.List(iTop) = sTemp
        
        iTop = iTop + 1
        iBottom = iBottom - 1
    Wend
    
    lstPlayList1.ListItems.Clear
    For iTop = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iTop)
    Next iTop
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub mnuTitle_Click()
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sArtist As String
    Dim sTitle As String
    Dim sTemp As String
    
    For iCnt = 1 To lstPlayList1.ListItems.count
        sArtist = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        sTitle = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        iFoundAt = iCnt - 1
        sTemp = lstFilenames.List(iCnt - 1)
        For iPos = iCnt To lstPlayList1.ListItems.count
            If sTitle > lstPlayList1.ListItems(iPos).ListSubItems(1).Text Then
                sTemp = lstFilenames.List(iPos - 1)
                sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text
                sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                iFoundAt = iPos - 1
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt - 1)
        lstFilenames.List(iCnt - 1) = sTemp
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(1).Text = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(2).Text = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        lstPlayList1.ListItems(iCnt).ListSubItems(1).Text = sTitle
        lstPlayList1.ListItems(iCnt).ListSubItems(2).Text = sArtist
    Next iCnt
End Sub

Private Sub shittyballz_Click()
End
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Slider1_Scroll()
    MediaPlayer1.Volume = Slider1.Value
End Sub

Private Sub Slider2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Slider2_Scroll()
    MediaPlayer1.Balance = Slider2.Value
End Sub

Private Sub Slider3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        frmMain.PopupMenu mnuForm
    End If
End Sub

Private Sub Slider3_Scroll()
MediaPlayer1.CurrentPosition = Slider3.Value
End Sub

Private Sub Timer1_Timer()
'Purpose:  Updates labels and keeps songs playing continuously
    Dim sStatus As String
    
    Select Case MediaPlayer1.PlayState
    
    Case mpClosed
        sStatus = "No song playing         "
    Case mpPaused
        sStatus = "Paused                  "
    Case mpPlaying
        sStatus = "Playing                 "
    Case mpStopped
        If Not bStopPressed Then
            imgNext_Click
        End If
        sStatus = "Stopped                 "
    End Select
    
    If frmMain.bScrollTitle Then
        If sFormTitle <> "" Then
            sFormTitle = Right(sFormTitle, Len(sFormTitle) - 1) & Left(sFormTitle, 1)
        End If
        
        frmMain.Caption = sFormTitle
    End If
    
    If MediaPlayer1.PlayState <> mpClosed And MediaPlayer1.PlayState <> mpStopped Then
        Label2.Caption = sStatus & SecondsToTime(MediaPlayer1.CurrentPosition) & "/" & SecondsToTime(MediaPlayer1.Duration)
    Else
        Label2.Caption = sStatus & "0:00/0:00"
    End If
    
    If Round(MediaPlayer1.CurrentPosition, 0) >= Fix(MediaPlayer1.Duration) Then
        imgNext_Click
    End If
    
End Sub

Private Sub DeleteFile(iPlaylistIndex As Integer)
    Dim iIndex As Integer
    Dim iCnt As Integer
    Dim iResult As Integer
    
    If bConfirmDelete Then
        iResult = MsgBox("Are you sure you want to remove this song from your playlist?" & vbCrLf & "The song will still be on your computer at:" & vbCrLf & lstFilenames.List(iPlaylistIndex - 1), vbYesNo + vbQuestion, "Are you sure?")
    Else
        iResult = vbYes
    End If
    
    If iResult = vbYes Then
        lstFilenames.RemoveItem iPlaylistIndex - 1
        lstPlayList1.ListItems.Remove iPlaylistIndex
            
        For iCnt = 1 To lstPlayList1.ListItems.count
            lstPlayList1.ListItems(iCnt).Text = iCnt
        Next iCnt
    End If
End Sub

Private Sub ChangeFont(Object As enFontObject, Optional sFontName As String, Optional sFontSize As Integer, Optional sFontBold As Boolean, Optional sFontItalic As Boolean, Optional sFontColor As Long)
    
    If Object = Playlist Then
        If sFontName <> "" Then
            lstPlayList1.Font.Name = sFontName
        End If
        
        If sFontSize > 0 Then
            lstPlayList1.Font.Size = sFontSize
        End If
        
        If Not IsMissing(sFontBold) Then
            lstPlayList1.Font.Bold = sFontBold
        End If
        
        If Not IsMissing(sFontItalic) Then
            lstPlayList1.Font.Italic = sFontItalic
        End If
        
        If Not IsMissing(sFontColor) And sFontColor <> -1 Then
            lstPlayList1.ForeColor = sFontColor
        End If
        
    ElseIf Object = Other Then
    
        If sFontName <> "" Then
            Label1.Font.Name = sFontName
            Label2.Font.Name = sFontName
            cbRepeat.Font.Name = sFontName
            cbShuffle.Font.Name = sFontName
        End If
        
        If sFontSize > 0 Then
            Label1.Font.Size = sFontSize
            Label2.Font.Size = sFontSize
            cbRepeat.Font.Size = sFontSize
            cbShuffle.Font.Size = sFontSize
        End If
        
        If Not IsMissing(sFontBold) Then
            Label1.Font.Bold = sFontBold
            Label2.Font.Bold = sFontBold
            cbRepeat.Font.Bold = sFontBold
            cbShuffle.Font.Bold = sFontBold
        End If
        
        If Not IsMissing(sFontItalic) Then
            Label1.Font.Italic = sFontItalic
            Label2.Font.Italic = sFontItalic
            cbRepeat.Font.Italic = sFontItalic
            cbShuffle.Font.Italic = sFontItalic
        End If
        
        If Not IsMissing(sFontColor) And sFontColor <> -1 Then
            Label1.ForeColor = sFontColor
            Label2.ForeColor = sFontColor
            cbRepeat.ForeColor = sFontColor
            cbShuffle.ForeColor = sFontColor
        End If
        
    End If
    
End Sub

Private Function SecondsToTime(lSeconds As Double) As String
'Purpose:  Changes seconds into mintues:seconds (ex.  140 becomes 2:20)
    Dim sTime As String
    Dim iSeconds As Integer
    Dim iMinutes As Integer
    
    iSeconds = Abs(Fix(lSeconds)) Mod 60
    iMinutes = Fix(Abs(Fix(lSeconds)) / 60)
    
    sTime = iMinutes & ":" & IIf(iSeconds < 10, "0", "") & iSeconds
    
    SecondsToTime = sTime
End Function

Private Function ParseFiles(sFilename As String) As Boolean
'Purpose:  Adds a song to the playlist with the file's info
    Dim sTitle As String
    Dim sArtist As String
    Dim sPlaylistName As String
    Dim sName As String
    Dim iPos As Integer
    
    If LCase(Right(sFilename, 3)) = "mp3" Then
        Mp3Info.Filename = sFilename
        
        sTitle = Mp3Info.Title
        sArtist = Mp3Info.Artist
        
        If sTitle = "" Then
            iPos = InStrRev(sFilename, "\")
            sTitle = Mid(sFilename, iPos + 1, Len(sFilename) - iPos - 4)
        End If
        
        sPlaylistName = sTitle
        
        If sArtist <> "" Then
            sPlaylistName = sPlaylistName & " - " & sArtist
        End If
        
        'lstPlayList.AddItem lstPlayList.ListCount + 1 & ". " & sPlaylistName
        lstPlayList1.ListItems.Add lstPlayList1.ListItems.count + 1, , lstPlayList1.ListItems.count + 1
        lstPlayList1.ListItems(lstPlayList1.ListItems.count).ListSubItems.Add 1, , sTitle
        lstPlayList1.ListItems(lstPlayList1.ListItems.count).ListSubItems.Add 2, , sArtist
        ParseFiles = True
    ElseIf LCase(Right(sFilename, 3)) = "mid" Or LCase(Right(sFilename, 3)) = "wav" Then
        sTitle = Mid(sFilename, iPos + 1, Len(sFilename) - iPos - 4)
        
        'lstPlayList.AddItem lstPlayList.ListCount + 1 & ". " & sTitle
        lstPlayList1.ListItems.Add lstPlayList1.ListItems.count + 1, , lstPlayList1.ListItems.count + 1
        lstPlayList1.ListItems(lstPlayList1.ListItems.count).ListSubItems.Add 1, , sTitle
        lstPlayList1.ListItems(lstPlayList1.ListItems.count).ListSubItems.Add 2, , "Unknown"
        
        ParseFiles = True
    Else
        ParseFiles = False
    End If
End Function

Private Sub PlaySong(ByVal Index As Integer)
'Purpose:  Load a song, then play it
    If Index >= 1 Then
        LoadSong Index
    Else
        LoadSong 1
    End If
    
    If MediaPlayer1.Filename <> "" Then
        MediaPlayer1.Play
        bStopPressed = False
    End If
End Sub

Private Sub LoadSong(ByVal Index As Integer)
'Purpose:  Loads a song into the mediaplayer and sets properties

    If Index <= lstPlayList1.ListItems.count Then
        lstPlayList1.ListItems(iCurrentIndex).Bold = False
        lstPlayList1.ListItems(iCurrentIndex).ListSubItems(1).Bold = False
        lstPlayList1.ListItems(iCurrentIndex).ListSubItems(2).Bold = False
        iCurrentIndex = Index
        MediaPlayer1.Filename = lstFilenames.List(Index - 1)
        
        Label1.Caption = "" & lstPlayList1.ListItems(Index).ListSubItems(2).Text & " - " & lstPlayList1.ListItems(Index).ListSubItems(1).Text
        'Label1.Caption = "Current Song: " & Right(lstPlayList.List(Index), Len(lstPlayList.List(Index)) - Len(Str(Index)) - 1)
        sFormTitle = frmMain.sPlayerName & " - [" & lstPlayList1.ListItems(Index).ListSubItems(1).Text & "-" & lstPlayList1.ListItems(Index).ListSubItems(2).Text & "]  "
        lstPlayList1.ListItems(Index).Selected = True
        lstPlayList1.ListItems(Index).EnsureVisible
        lstPlayList1.ListItems(Index).Bold = True
        lstPlayList1.ListItems(Index).ListSubItems(1).Bold = True
        lstPlayList1.ListItems(Index).ListSubItems(2).Bold = True
        lstPlayList1.Refresh
    End If
End Sub

Private Sub LoadPlayList(sFilename As String)
'Purpose:  Load a playlist from a file
    Dim iFilenum As Integer
    Dim sTemp As String
    
    If Dir(sFilename) <> "" Then
        iFilenum = FreeFile
        
        Open sFilename For Input As #iFilenum
        While Not EOF(iFilenum)
            Line Input #iFilenum, sTemp
            If ParseFiles(sTemp) Then
                lstFilenames.AddItem sTemp
            End If
        Wend
        Close #iFilenum
        
        If lstPlayList1.ListItems.count >= 1 Then
            LoadSong 1
        End If
    End If
    
End Sub

Private Sub SavePlaylist(sFilename As String)
'Purpose:  Saves the songs in the playlist to a file
    Dim iFilenum As Integer
    Dim iCnt As Integer
    
    iFilenum = FreeFile
    
    Open sFilename For Output As #iFilenum
    
    For iCnt = 1 To lstFilenames.ListCount
        Print #iFilenum, lstFilenames.List(iCnt - 1)
    Next iCnt
    
    Close #iFilenum
End Sub

Private Function GetRandomIndex(iNumOfSongs As Integer) As Integer
'Purpose:  Generates a random number for the next song
    GetRandomIndex = Int((iNumOfSongs - 0 + 1) * Rnd + 0)
End Function

Private Sub ParseMultipleFiles(sAllFiles As String)
    Dim sDir As String
    Dim iPos As String
    
    'iPos = instr(1,sallfiles,
    
End Sub

Private Sub AddDirectory(sPath As String)
    Dim p As Integer
    Dim I As Integer

    Dir1.Path = sPath
    File1.Path = sPath

    For p = 0 To Dir1.ListCount - 1
        If p < Dir1.ListCount Then
            AddDirectory Dir1.List(p)
        End If
    Next p
    For I = 0 To File1.ListCount - 1
        If LCase(Right(File1.List(I), 3)) = "mp3" _
        Or LCase(Right(File1.List(I), 3)) = "mid" _
        Or LCase(Right(File1.List(I), 3)) = "wav" Then
            If ParseFiles(Dir1.Path & "\" & File1.List(I)) Then
                lstFilenames.AddItem Dir1.Path & "\" & File1.List(I)
            End If
        End If
    Next I

    Dir1.Path = UpOneDir(Dir1.Path)
    File1.Path = Dir1.Path
End Sub

Public Function UpOneDir(sPathName As String) As String
    Dim q As Integer
    Dim num As Integer
        
    For q = 1 To Len(sPathName)
        If Mid(sPathName, q, 1) = "\" Then
            num = q
        End If
    Next q
    If Len(Mid(sPathName, 1, num - 1)) < 3 Then
        UpOneDir = Mid(sPathName, 1, num - 1) & "\"
    Else
        UpOneDir = Mid(sPathName, 1, num - 1)
    End If
End Function

Private Sub LoadINIFile()
    Dim iFilenum As Integer
    Dim sColors() As String
    Dim sTemp As String
    Dim iPos As String

    On Error GoTo EH
    
    iFilenum = FreeFile
    
    If Dir(sImageDir & "skin.ini") <> "" Then
    
        Open sImageDir & "skin.ini" For Input As #iFilenum
        
        While Not EOF(iFilenum)
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            lstPlayList1.ForeColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            Label1.ForeColor = CLng(sTemp)
            Label2.ForeColor = CLng(sTemp)
            cbRepeat.ForeColor = CLng(sTemp)
            cbShuffle.ForeColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            lstPlayList1.Font.Name = Trim$(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Label1.Font.Name = Trim$(sTemp)
            frmMain.Label2.Font.Name = Trim(sTemp)
            frmMain.cbRepeat.Font.Name = Trim(sTemp)
            frmMain.cbShuffle.Font.Name = Trim(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.lstPlayList1.BackColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.BackColor = CLng(sTemp)
            cbRepeat.BackColor = CLng(sTemp)
            cbShuffle.BackColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bConfirmDelete = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bSavePlaylist = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bSaveSettings = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bScrollTitle = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.cbRepeat.Value = Val(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.cbShuffle.Value = Val(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Slider1.Value = CInt(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Slider2.Value = CInt(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.sPlayerName = Trim(sTemp)
        Wend
        
        Close #iFilenum
    End If
    Exit Sub
EH:
    MsgBox Err.Description & " in LoadINIFile"
    Close #iFilenum
End Sub

Private Sub SaveINIFile()
    Dim iFilenum As Integer
    Dim sTemp As String
    
    iFilenum = FreeFile()
    
    If Dir(sImageDir & "skin.ini") <> "" Then
        Kill sImageDir & "skin.ini"
    End If
    
    Open sImageDir & "skin.ini" For Output As #iFilenum
    
        Print #iFilenum, lstPlayList1.ForeColor & " //Playlist Font Color"
        Print #iFilenum, frmMain.cbRepeat.ForeColor & " //All Other Font Color"
        Print #iFilenum, lstPlayList1.Font.Name & " //Playlist Font Name"
        Print #iFilenum, frmMain.cbRepeat.Font.Name & " //Other Font Name"
        Print #iFilenum, lstPlayList1.BackColor & " //Playlist Background Color"
        Print #iFilenum, frmMain.BackColor & " //All Other Background Color"
        Print #iFilenum, frmMain.bConfirmDelete & " //Confirm Delete"
        Print #iFilenum, frmMain.bSavePlaylist & " //Save Playlist"
        Print #iFilenum, frmMain.bSaveSettings & " //Save Settings"
        Print #iFilenum, frmMain.bScrollTitle & " //Scroll Title"
        Print #iFilenum, frmMain.cbRepeat.Value & " //Repeat"
        Print #iFilenum, frmMain.cbShuffle.Value & " //Shuffle"
        Print #iFilenum, frmMain.Slider1.Value & " //Volume"
        Print #iFilenum, frmMain.Slider2.Value & " //Balance"
        Print #iFilenum, frmMain.sPlayerName & " //Player Title"
    Close #iFilenum
    
    
End Sub

Function RegistryQuery(sValue As String, Optional vPrompt As Variant) As String
'Purpose: sets a value in the registry by means of input box from user
'Parameters: sValue - value to get in registry, string
'            vPrompt - text of input box, varaint, optional
    On Error GoTo ErrorHandler
    Dim sTemp As String
1    sTemp = Registry.QueryValue(HKEY_LOCAL_MACHINE, gKEYNAME, sValue, gMISSINGREGDATA)
2    If sTemp = gMISSINGREGDATA Then
3        If Not IsMissing(vPrompt) Then
4            sTemp = InputBox(vPrompt, "Monkamp")
5            Registry.SetKeyValue HKEY_LOCAL_MACHINE, gKEYNAME, sValue, sTemp, REG_SZ
6        Else
7            sTemp = gEMPTYSTRING
        End If
    End If
8    RegistryQuery = sTemp
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, "RegistryQuery " & Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Sub Timer2_Timer()
Slider3.Max = (MediaPlayer1.Duration)
Slider3.Value = (MediaPlayer1.CurrentPosition)
End Sub

Private Sub Timer3_Timer()
If MediaPlayer1.PlayState = mpPlaying Then Timer2.Enabled = True
If MediaPlayer1.PlayState = mpStopped Then Timer2.Enabled = False
If MediaPlayer1.PlayState = mpPaused Then Timer2.Enabled = True
End Sub
