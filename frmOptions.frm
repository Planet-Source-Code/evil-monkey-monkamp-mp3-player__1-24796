VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monkamp Options"
   ClientHeight    =   1680
   ClientLeft      =   6195
   ClientTop       =   4650
   ClientWidth     =   2280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2280
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1800
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00008000&
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   120
      List            =   "frmOptions.frx":0013
      TabIndex        =   1
      Text            =   "Choose a Color"
      ToolTipText     =   "Changes Between the Colors in Monkamp"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox cbDelete 
      BackColor       =   &H00C00000&
      Caption         =   "Confirm Song Deletion"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Change Color"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Option Explicit

Private Sub cbSave_Click()

End Sub

Private Sub Command1_Click()
    frmMain.bConfirmDelete = Me.cbDelete
    Me.Hide
End Sub

Private Sub Form_Load()
    ontop.MakeTopMost hWnd
    If frmMain.bConfirmDelete Then
        cbDelete.Value = vbChecked
    End If
End Sub

Private Sub Timer1_Timer()
'The Forms that gets changed...
        'Blue
    If Combo1.Text = "Blue" Then Me.BackColor = &HC00000
    If Combo1.Text = "Blue" Then Form1.BackColor = &HC00000
    If Combo1.Text = "Blue" Then frmMain.BackColor = &HC00000
    If Combo1.Text = "Blue" Then Main.BackColor = &HC00000
    If Combo1.Text = "Blue" Then id3.BackColor = &HC00000
        'Green
        If Combo1.Text = "Green" Then Me.BackColor = &HC000&
        If Combo1.Text = "Green" Then Form1.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.BackColor = &HC000&
        If Combo1.Text = "Green" Then Main.BackColor = &HC000&
        If Combo1.Text = "Green" Then id3.BackColor = &HC000&
        'Red
            If Combo1.Text = "Red" Then Me.BackColor = &HC0&
            If Combo1.Text = "Red" Then Form1.BackColor = &HC0&
            If Combo1.Text = "Red" Then frmMain.BackColor = &HC0&
            If Combo1.Text = "Red" Then Main.BackColor = &HC0&
            If Combo1.Text = "Red" Then id3.BackColor = &HC0&
        'Yellow
                If Combo1.Text = "Yellow" Then Me.BackColor = &HFFFF&
                If Combo1.Text = "Yellow" Then Form1.BackColor = &HFFFF&
                If Combo1.Text = "Yellow" Then frmMain.BackColor = &HFFFF&
                If Combo1.Text = "Yellow" Then Main.BackColor = &HFFFF&
                If Combo1.Text = "Yellow" Then id3.BackColor = &HFFFF&
        'Windows Standard
                    If Combo1.Text = "Windows Standard" Then Me.BackColor = &H8000000A
                    If Combo1.Text = "Windows Standard" Then Form1.BackColor = &H8000000A
                    If Combo1.Text = "Windows Standard" Then frmMain.BackColor = &H8000000A
                    If Combo1.Text = "Windows Standard" Then Main.BackColor = &H8000000A
                    If Combo1.Text = "Windows Standard" Then id3.BackColor = &H8000000A
'All the other crap that gets changed...
    'Blue
        'Options Form
    If Combo1.Text = "Blue" Then Me.Label1.BackColor = &HC00000
    If Combo1.Text = "Blue" Then Me.cbDelete.BackColor = &HC00000
    If Combo1.Text = "Blue" Then Me.cbDelete.ForeColor = &HC0&
    If Combo1.Text = "Blue" Then Me.Label1.ForeColor = &HC0&
    If Combo1.Text = "Blue" Then Me.Combo1.ForeColor = &H8000&
    If Combo1.Text = "Blue" Then Me.Combo1.BackColor = &HC00000
    If Combo1.Text = "Blue" Then Me.Command1.BackColor = &HC00000
        'Main Form
        If Combo1.Text = "Blue" Then frmMain.Label1.ForeColor = &H8000&
        If Combo1.Text = "Blue" Then frmMain.Label1.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.Label2.ForeColor = &H8000&
        If Combo1.Text = "Blue" Then frmMain.Label2.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.imgPlay.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.imgStop.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.imgPrev.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.imgNext.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.imgPause.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.lstPlayList1.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.lstPlayList1.ForeColor = &H8000&
        If Combo1.Text = "Blue" Then frmMain.cmdOpen.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.cmdSave.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.cmdClear.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.cbRepeat.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.cbShuffle.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.cbRepeat.ForeColor = &H8000&
        If Combo1.Text = "Blue" Then frmMain.cbShuffle.ForeColor = &H8000&
        If Combo1.Text = "Blue" Then frmMain.Frame1.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.Command3.BackColor = &HC00000
        If Combo1.Text = "Blue" Then frmMain.Command4.BackColor = &HC00000
        'Smalled Form
               If Combo1.Text = "Blue" Then Form1.Label1.BackColor = &HC00000
               If Combo1.Text = "Blue" Then Form1.Label2.BackColor = &HC00000
               If Combo1.Text = "Blue" Then Form1.Label2.ForeColor = &HFFFFFF
               If Combo1.Text = "Blue" Then Form1.Label1.ForeColor = &HC0C0&
               If Combo1.Text = "Blue" Then Form1.Line2.BorderColor = &HC0&
               If Combo1.Text = "Blue" Then Form1.Line1.BorderColor = &HC0&
        'ID3 Form
                    If Combo1.Text = "Blue" Then id3.Label1.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Label2.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Label3.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Label4.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Label5.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Label6.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Command1.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Command2.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Command3.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Text1.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Text2.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Text3.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Text4.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Text5.BackColor = &HC00000
                    If Combo1.Text = "Blue" Then id3.Combo1.BackColor = &HC00000
        'Encoder Form
                        If Combo1.Text = "Blue" Then Main.image9.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.image11.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.hElPeRoNy.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Command1.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Frame1.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Frame2.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Frame3.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Frame5.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.image2.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.image5.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.OutPutPath.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.InputPath.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option1.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option2.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option3.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option4.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option5.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option6.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Option7.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Label1.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.Lbl1.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.LblActFrame.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.LblTotal.BackColor = &HC00000
                        If Combo1.Text = "Blue" Then Main.txtpod.BackColor = &HC00000
    'Green
        'Options Form
    If Combo1.Text = "Green" Then Me.Label1.BackColor = &HC000&
    If Combo1.Text = "Green" Then Me.cbDelete.BackColor = &HC000&
    If Combo1.Text = "Green" Then Me.cbDelete.ForeColor = &HC0&
    If Combo1.Text = "Green" Then Me.Label1.ForeColor = &HC0&
    If Combo1.Text = "Green" Then Me.Combo1.ForeColor = &HC00000
    If Combo1.Text = "Green" Then Me.Combo1.BackColor = &HC000&
    If Combo1.Text = "Green" Then Me.Command1.BackColor = &HC000&
        'Main Form
        If Combo1.Text = "Green" Then frmMain.Label1.ForeColor = &HC00000
        If Combo1.Text = "Green" Then frmMain.Label1.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.Label2.ForeColor = &HC00000
        If Combo1.Text = "Green" Then frmMain.Label2.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.imgPlay.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.imgStop.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.imgPrev.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.imgNext.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.imgPause.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.lstPlayList1.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.lstPlayList1.ForeColor = &HC00000
        If Combo1.Text = "Green" Then frmMain.cmdOpen.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.cmdSave.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.cmdClear.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.cbRepeat.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.cbShuffle.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.cbRepeat.ForeColor = &HC00000
        If Combo1.Text = "Green" Then frmMain.cbShuffle.ForeColor = &HC00000
        If Combo1.Text = "Green" Then frmMain.Frame1.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.Command3.BackColor = &HC000&
        If Combo1.Text = "Green" Then frmMain.Command4.BackColor = &HC000&
        'Smalled Form
               If Combo1.Text = "Green" Then Form1.Label1.BackColor = &HC000&
               If Combo1.Text = "Green" Then Form1.Label2.BackColor = &HC000&
               If Combo1.Text = "Green" Then Form1.Label2.ForeColor = &HFFFFFF
               If Combo1.Text = "Green" Then Form1.Label1.ForeColor = &HC000C0
               If Combo1.Text = "Green" Then Form1.Line2.BorderColor = &HC0&
               If Combo1.Text = "Green" Then Form1.Line1.BorderColor = &HC0&
        'ID3 Form
                    If Combo1.Text = "Green" Then id3.Label1.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Label2.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Label3.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Label4.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Label5.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Label6.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Command1.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Command2.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Command3.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Text1.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Text2.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Text3.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Text4.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Text5.BackColor = &HC000&
                    If Combo1.Text = "Green" Then id3.Combo1.BackColor = &HC000&
        'Encoder Form
                        If Combo1.Text = "Green" Then Main.image9.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.image11.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.hElPeRoNy.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Command1.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Frame1.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Frame2.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Frame3.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Frame5.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.image2.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.image5.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.OutPutPath.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.InputPath.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option1.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option2.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option3.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option4.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option5.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option6.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Option7.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Label1.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.Lbl1.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.LblActFrame.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.LblTotal.BackColor = &HC000&
                        If Combo1.Text = "Green" Then Main.txtpod.BackColor = &HC000&
    'Red
        'Options Form
    If Combo1.Text = "Red" Then Me.Label1.BackColor = &HC0&
    If Combo1.Text = "Red" Then Me.cbDelete.BackColor = &HC0&
    If Combo1.Text = "Red" Then Me.cbDelete.ForeColor = &H8000000A
    If Combo1.Text = "Red" Then Me.Label1.ForeColor = &H8000000A
    If Combo1.Text = "Red" Then Me.Combo1.ForeColor = &H8000&
    If Combo1.Text = "Red" Then Me.Combo1.BackColor = &HC0&
    If Combo1.Text = "Red" Then Me.Command1.BackColor = &HC0&
        'Main Form
        If Combo1.Text = "Red" Then frmMain.Label1.ForeColor = &H8000&
        If Combo1.Text = "Red" Then frmMain.Label1.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.Label2.ForeColor = &H8000&
        If Combo1.Text = "Red" Then frmMain.Label2.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.imgPlay.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.imgStop.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.imgPrev.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.imgNext.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.imgPause.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.lstPlayList1.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.lstPlayList1.ForeColor = &H8000&
        If Combo1.Text = "Red" Then frmMain.cmdOpen.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.cmdSave.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.cmdClear.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.cbRepeat.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.cbShuffle.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.cbRepeat.ForeColor = &H8000&
        If Combo1.Text = "Red" Then frmMain.cbShuffle.ForeColor = &H8000&
        If Combo1.Text = "Red" Then frmMain.Frame1.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.Command3.BackColor = &HC0&
        If Combo1.Text = "Red" Then frmMain.Command4.BackColor = &HC0&
        'Smalled Form
               If Combo1.Text = "Red" Then Form1.Label1.BackColor = &HC0&
               If Combo1.Text = "Red" Then Form1.Label2.BackColor = &HC0&
               If Combo1.Text = "Red" Then Form1.Label2.ForeColor = &HFFFFFF
               If Combo1.Text = "Red" Then Form1.Label1.ForeColor = &HC0C0&
               If Combo1.Text = "Red" Then Form1.Line2.BorderColor = &HC000&
               If Combo1.Text = "Red" Then Form1.Line1.BorderColor = &HC000&
        'ID3 Form
                    If Combo1.Text = "Red" Then id3.Label1.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Label2.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Label3.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Label4.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Label5.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Label6.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Command1.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Command2.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Command3.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Text1.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Text2.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Text3.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Text4.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Text5.BackColor = &HC0&
                    If Combo1.Text = "Red" Then id3.Combo1.BackColor = &HC0&
        'Encoder Form
                        If Combo1.Text = "Red" Then Main.image9.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.image11.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.hElPeRoNy.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Command1.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Frame1.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Frame2.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Frame3.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Frame5.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.image2.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.image5.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.OutPutPath.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.InputPath.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option1.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option2.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option3.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option4.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option5.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option6.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Option7.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Label1.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.Lbl1.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.LblActFrame.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.LblTotal.BackColor = &HC0&
                        If Combo1.Text = "Red" Then Main.txtpod.BackColor = &HC0&
    'Yellow
        'Options Form
    If Combo1.Text = "Yellow" Then Me.Label1.BackColor = &HFFFF&
    If Combo1.Text = "Yellow" Then Me.cbDelete.BackColor = &HFFFF&
    If Combo1.Text = "Yellow" Then Me.cbDelete.ForeColor = &HC0&
    If Combo1.Text = "Yellow" Then Me.Label1.ForeColor = &HC0&
    If Combo1.Text = "Yellow" Then Me.Combo1.ForeColor = &H8000&
    If Combo1.Text = "Yellow" Then Me.Combo1.BackColor = &HFFFF&
    If Combo1.Text = "Yellow" Then Me.Command1.BackColor = &HFFFF&
        'Main Form
        If Combo1.Text = "Yellow" Then frmMain.Label1.ForeColor = &H8000&
        If Combo1.Text = "Yellow" Then frmMain.Label1.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.Label2.ForeColor = &H8000&
        If Combo1.Text = "Yellow" Then frmMain.Label2.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.imgPlay.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.imgStop.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.imgPrev.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.imgNext.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.imgPause.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.lstPlayList1.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.lstPlayList1.ForeColor = &H8000&
        If Combo1.Text = "Yellow" Then frmMain.cmdOpen.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.cmdSave.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.cmdClear.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.cbRepeat.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.cbShuffle.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.cbRepeat.ForeColor = &H8000&
        If Combo1.Text = "Yellow" Then frmMain.cbShuffle.ForeColor = &H8000&
        If Combo1.Text = "Yellow" Then frmMain.Frame1.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.Command3.BackColor = &HFFFF&
        If Combo1.Text = "Yellow" Then frmMain.Command4.BackColor = &HFFFF&
        'Smalled Form
               If Combo1.Text = "Yellow" Then Form1.Label1.BackColor = &HFFFF&
               If Combo1.Text = "Yellow" Then Form1.Label2.BackColor = &HFFFF&
               If Combo1.Text = "Yellow" Then Form1.Label2.ForeColor = &HFFFFFF
               If Combo1.Text = "Yellow" Then Form1.Label1.ForeColor = &HC000C0
               If Combo1.Text = "Yellow" Then Form1.Line2.BorderColor = &HC0&
               If Combo1.Text = "Yellow" Then Form1.Line1.BorderColor = &HC0&
        'ID3 Form
                    If Combo1.Text = "Yellow" Then id3.Label1.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Label2.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Label3.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Label4.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Label5.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Label6.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Command1.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Command2.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Command3.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Text1.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Text2.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Text3.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Text4.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Text5.BackColor = &HFFFF&
                    If Combo1.Text = "Yellow" Then id3.Combo1.BackColor = &HFFFF&
        'Encoder Form
                        If Combo1.Text = "Yellow" Then Main.image9.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.image11.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.hElPeRoNy.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Command1.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Frame1.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Frame2.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Frame3.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Frame5.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.image2.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.image5.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.OutPutPath.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.InputPath.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option1.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option2.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option3.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option4.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option5.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option6.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Option7.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Label1.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.Lbl1.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.LblActFrame.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.LblTotal.BackColor = &HFFFF&
                        If Combo1.Text = "Yellow" Then Main.txtpod.BackColor = &HFFFF&
    'Windows Standard
        'Options Form
    If Combo1.Text = "Windows Standard" Then Me.Label1.BackColor = &H8000000F
    If Combo1.Text = "Windows Standard" Then Me.cbDelete.BackColor = &H8000000F
    If Combo1.Text = "Windows Standard" Then Me.cbDelete.ForeColor = &H80000008
    If Combo1.Text = "Windows Standard" Then Me.Label1.ForeColor = &H80000008
    If Combo1.Text = "Windows Standard" Then Me.Combo1.ForeColor = &H80000008
    If Combo1.Text = "Windows Standard" Then Me.Combo1.BackColor = &H80000005
    If Combo1.Text = "Windows Standard" Then Me.Command1.BackColor = &H8000000F
        'Main Form
        If Combo1.Text = "Windows Standard" Then frmMain.Label1.ForeColor = &H80000008
        If Combo1.Text = "Windows Standard" Then frmMain.Label1.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.Label2.ForeColor = &H80000008
        If Combo1.Text = "Windows Standard" Then frmMain.Label2.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.imgPlay.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.imgStop.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.imgPrev.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.imgNext.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.imgPause.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.lstPlayList1.BackColor = &H80000005
        If Combo1.Text = "Windows Standard" Then frmMain.lstPlayList1.ForeColor = &H80000008
        If Combo1.Text = "Windows Standard" Then frmMain.cmdOpen.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.cmdSave.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.cmdClear.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.cbRepeat.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.cbShuffle.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.cbRepeat.ForeColor = &H80000008
        If Combo1.Text = "Windows Standard" Then frmMain.cbShuffle.ForeColor = &H80000008
        If Combo1.Text = "Windows Standard" Then frmMain.Frame1.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.Command3.BackColor = &H8000000F
        If Combo1.Text = "Windows Standard" Then frmMain.Command4.BackColor = &H8000000F
        'Smalled Form
               If Combo1.Text = "Windows Standard" Then Form1.Label1.BackColor = &H8000000F
               If Combo1.Text = "Windows Standard" Then Form1.Label2.BackColor = &H8000000F
               If Combo1.Text = "Windows Standard" Then Form1.Label2.ForeColor = &H80000008
               If Combo1.Text = "Windows Standard" Then Form1.Label1.ForeColor = &H80000008
               If Combo1.Text = "Windows Standard" Then Form1.Line2.BorderColor = &H80000008
               If Combo1.Text = "Windows Standard" Then Form1.Line1.BorderColor = &H80000008
        'ID3 Form
                    If Combo1.Text = "Windows Standard" Then id3.Label1.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Label2.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Label3.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Label4.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Label5.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Label6.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Command1.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Command2.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Command3.BackColor = &H8000000F
                    If Combo1.Text = "Windows Standard" Then id3.Text1.BackColor = &H80000005
                    If Combo1.Text = "Windows Standard" Then id3.Text2.BackColor = &H80000005
                    If Combo1.Text = "Windows Standard" Then id3.Text3.BackColor = &H80000005
                    If Combo1.Text = "Windows Standard" Then id3.Text4.BackColor = &H80000005
                    If Combo1.Text = "Windows Standard" Then id3.Text5.BackColor = &H80000005
                    If Combo1.Text = "Windows Standard" Then id3.Combo1.BackColor = &H80000005
        'Encoder Form
                        If Combo1.Text = "Windows Standard" Then Main.image9.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.image11.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.hElPeRoNy.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Command1.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Frame1.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Frame2.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Frame3.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Frame5.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.image2.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.image5.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.OutPutPath.BackColor = &H80000005
                        If Combo1.Text = "Windows Standard" Then Main.InputPath.BackColor = &H80000005
                        If Combo1.Text = "Windows Standard" Then Main.Option1.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option2.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option3.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option4.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option5.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option6.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Option7.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Label1.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.Lbl1.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.LblActFrame.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.LblTotal.BackColor = &H8000000F
                        If Combo1.Text = "Windows Standard" Then Main.txtpod.BackColor = &H80000005
End Sub
