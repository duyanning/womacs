VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Type POINTAPI
    X As Long
    Y As Long
End Type
 
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function GetCapture Lib "user32" () As Long

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As LongPtr) As LongPtr
Private Declare PtrSafe Function SetCapture Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function GetCapture Lib "user32" () As LongPtr

 
Private FrameHwnd As Long
 
Private Function GetFrameHwnd() As Long
    Dim PT As POINTAPI
    GetCursorPos PT
    GetFrameHwnd = WindowFromPoint(PT.X, PT.Y)
End Function
 
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Label1.Caption = "down"
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If FrameHwnd = 0 Then FrameHwnd = GetFrameHwnd()
     
    If (X < 0) Or (Y < 0) Or (X > UserForm3.Width) Or (Y > UserForm3.Height) Then
        ReleaseCapture
        Call MouseExit
    ElseIf GetCapture() <> FrameHwnd Then
        SetCapture FrameHwnd
        Call MouseEnter
    End If
End Sub
 
Private Sub MouseExit()
    Label1.Caption = "MouseExit"
End Sub
 
Private Sub MouseEnter()
    Label1.Caption = "MouseEnter"
End Sub
