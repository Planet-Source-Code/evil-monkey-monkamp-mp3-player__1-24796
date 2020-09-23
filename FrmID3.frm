VERSION 5.00
Begin VB.Form FrmID3 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monkamp ID3 Editor"
   ClientHeight    =   2640
   ClientLeft      =   4575
   ClientTop       =   3855
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C00000&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C00000&
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   2640
      MaxLength       =   30
      TabIndex        =   11
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtAlbum 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox txtArtist 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   3
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtSong 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtFilename 
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Comments:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Year:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Album:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Artist:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Song:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Filename:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Option Explicit

Dim mvarFilename As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sInfo As Id3
    
    sInfo.Album = Left(txtAlbum.Text, 30)
    sInfo.Artist = Left(txtArtist.Text, 30)
    sInfo.Title = Left(txtSong.Text, 30)
    sInfo.sYear = Left(txtYear.Text, 4)
    sInfo.Comments = Left(txtComments.Text, 30)
    
    Id3Module.SaveId3 mvarFilename, sInfo
    
    Unload Me
End Sub

Private Sub Form_Load()
    ontop.MakeTopMost hWnd
    Dim sInfo As New clsMP3Info
    
    sInfo.Filename = mvarFilename
    
    txtFilename.Text = mvarFilename
    txtSong.Text = sInfo.Title
    txtArtist.Text = sInfo.Artist
    txtAlbum.Text = sInfo.Album
    txtYear.Text = sInfo.Year
    txtComments.Text = sInfo.Comment
    
    Set sInfo = Nothing
End Sub

Public Property Let Filename(ByVal sData As String)
    mvarFilename = sData
End Property

Public Property Get Filename() As String
    Filename = mvarFilename
End Property

