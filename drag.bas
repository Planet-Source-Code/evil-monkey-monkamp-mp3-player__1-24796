Attribute VB_Name = "drag"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
'
' Example:
'
' Private Sub BLAH_MouseDown(STUFF)
' Call FormDrag Me or (X)
' End Sub
'
'--------------------------------------------------------------------------
'
' BLAH = the thing being pressed ( i.e. a command button or
' a picture or something )
'
' STUFF = all the crap thats in that thing when u make it MouseDown.
' Don't worry about STUFF.  Its does it automaticially.
'
' If you type me, VB recognizes it as the current form.  So if you
' type Me, Me is the current form.  Usually you want to type Me instead
' of the form name.
'
' X = the name of the form u want to drag
'
Public Sub FormDrag(FormName As Form)
Call ReleaseCapture
Call SendMessage(FormName.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
 
