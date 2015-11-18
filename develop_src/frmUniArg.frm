VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUniArg 
   Caption         =   "universal argument"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   OleObjectBlob   =   "frmUniArg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUniArg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' womacs
Option Explicit

'Dim num_arg_value As Long
'Dim self_ins_string As String

Public Enum DialogQuitReason
    dqrCancel
    dqrEnter
End Enum


Dim quit_reason As DialogQuitReason


Function get_quit_reason() As DialogQuitReason
    get_quit_reason = quit_reason
End Function


Sub get_two_parts(num_arg_value As Long, self_ins_string As String)
    '将输入分解为前后两部分，前边是数值部分，后边是自插入字符串部分
    Dim raw_input As String
    raw_input = tbInput.Text

    Dim space_pos As Long
    Dim num_arg_part As String
    Dim self_ins_part As String

    space_pos = InStr(raw_input, " ")
    Debug.Assert (space_pos = 0 Or space_pos > 1)

    If space_pos > 1 Then
        num_arg_part = Left(raw_input, space_pos - 1)
        self_ins_part = Mid(raw_input, space_pos + 1)

    Else
        num_arg_part = raw_input
        self_ins_part = ""
    End If

    num_arg_value = Val(num_arg_part)
    self_ins_string = self_ins_part

End Sub


Sub inc_arg_num()
    Dim num_arg_value As Long
    Dim self_ins_string As String
    
    get_two_parts num_arg_value, self_ins_string
    num_arg_value = num_arg_value * 4
    
    Dim num_arg_part As String
    num_arg_part = LTrim(str(num_arg_value))
    
    tbInput.Text = num_arg_part & " " & self_ins_string
End Sub


Sub when_key_down(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        quit_reason = dqrEnter
        Me.hide
        Exit Sub
    End If

    ' 2 is fmCtrlMask, I don't know why fmCtrlMask is Empty
    If Not Shift = 2 Then
        Exit Sub
    End If


    '如果按下了C-u
        '数值部分*4
    '如果按下C-g
        '关闭对话框
    Select Case KeyCode
    Case vbKeyG
        quit_reason = dqrEnter
        Me.hide
    Case vbKeyU
        inc_arg_num
    Case Else
        'do nothing
    End Select

End Sub


Private Sub tbInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    when_key_down KeyCode, Shift

End Sub


Private Sub UserForm_Initialize()
    tbInput.Text = "4 "
End Sub
