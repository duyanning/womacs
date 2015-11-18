VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Activate event for UserForm1
Private Sub UserForm_Activate()
    UserForm1.Caption = "Click my client area"
'    hwnd = FindWindow("ThunderDFrame", Me.Caption)
'    SetCapture (hwnd)
End Sub

' Click event for UserForm1
Private Sub UserForm_Click()
    Load UserForm2
    UserForm2.StartUpPosition = 3
    UserForm2.Show vbModeless
End Sub

' Deactivate event for UserForm1
Private Sub UserForm_Deactivate()
    UserForm1.Caption = "I just lost the focus!"
    UserForm2.Caption = "Focus just left UserForm1 and came to me"
End Sub
 

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'MsgBox "hi"
    Debug.Print "hi"
End Sub
