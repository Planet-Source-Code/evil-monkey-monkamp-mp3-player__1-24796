VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form id3 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monkamp ID3 Editor"
   ClientHeight    =   2070
   ClientLeft      =   540
   ClientTop       =   3435
   ClientWidth     =   4230
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "id3 stuff.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4230
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C00000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C00000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mp3"
      DialogTitle     =   "Mp3 Files"
      Filter          =   "Mp3 files (*.mp3)|*.mp3"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "id3 stuff.frx":030A
      Left            =   2280
      List            =   "id3 stuff.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaxLength       =   30
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Title"
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
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
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
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "id3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop

Private Sub Command1_Click()
On Error GoTo shit
CommonDialog1.ShowOpen
GetId3 CommonDialog1.Filename           ' Get the filename
Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
Text3 = RTrim(id3Info.Album)                   ' trailing bytes
Text4 = RTrim(id3Info.sYear)
Text5 = RTrim(id3Info.Comments)
Combo1.ListIndex = id3Info.Genre        ' fill in all the correct info.
Command2.Enabled = True
shit:
End Sub

Private Sub Command2_Click()
On Error GoTo penis
id3Info.Title = Text1           ' just filling in the information into the type
id3Info.Artist = Text2
id3Info.Album = Text3
id3Info.sYear = Text4
id3Info.Comments = Text5
id3Info.Genre = Combo1.ListIndex
On Error GoTo ErrHandle             ' If the file is writeprotected
SaveId3 CommonDialog1.Filename, id3Info     ' Calling the Saveid3 function

Exit Sub


ErrHandle:
If Err.Number = 75 Then

Else
MsgBox Err.Description
End If
penis:

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ontop.MakeTopMost hWnd
GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
For I = LBound(GenreArray) To UBound(GenreArray)
Combo1.AddItem GenreArray(I)        ' now fill the Combobox with the array, and voila, the code you
                                    ' you recieve form the Genre part of the Type, represents the combobox Listindex =)
Next


End Sub


