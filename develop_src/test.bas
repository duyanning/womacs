Attribute VB_Name = "test"
Option Explicit
Option Private Module

Public xyz As Integer
Dim abcxyz As Integer
Public abc As Sth
'Static fakel As New FakeLambda

 Dim Doc As Document

Sub ensure_proc_exists(module As CodeModule, proc_name As String, proc_def As String)
'    Dim c As VBComponent
'    Set c = ActiveDocument.VBProject.VBComponents.Add(vbext_ct_StdModule)
'    c.Name = "keymap_installers"
    
    'Dim c As VBComponent
    
    'Set c = ActiveDocument.VBProject.VBComponents("keymap_installers")
'    Set c = doc.VBProject.VBComponents("keymap_installers")
'
'    If is_proc_exists(doc, proc_name) Then
'        Exit Sub
'    End If
'
'
'
'    c.CodeModule.AddFromString (proc_def)

    If is_proc_exists(module, proc_name) Then
        Exit Sub
    End If

    module.AddFromString (proc_def)
    
    
End Sub

Function is_proc_exists(module As CodeModule, proc_name As String)
'    Dim c As VBComponent
'    Set c = ActiveDocument.VBProject.VBComponents("keymap_installers")
    
    On Error GoTo handleError
    Dim n As Long
    'n = c.CodeModule.ProcStartLine(proc_name, vbext_pk_Proc)
    n = module.ProcStartLine(proc_name, vbext_pk_Proc)
    Debug.Print n
    
    is_proc_exists = True
    Exit Function
    
handleError:
    is_proc_exists = False

End Function


Sub test()
    
    Dim foo_def As String
    foo_def = "Sub foo1()" & vbCrLf & _
                vbTab & "MsgBox ""haha""" & vbCrLf & _
                "End Sub"
    
    ensure_proc_exists ThisDocument.VBProject.VBComponents("keymap_installers").CodeModule, "foo1", foo_def

End Sub


Sub test2()
    foo1
End Sub

Sub test_ptr()
    Dim o As New Sth
    o.Value = 3
    Debug.Print o.Value
    
    Debug.Print CStr(ObjPtr(o))
    Dim p As LongPtr
    p = ObjPtr(o)
    Dim n As Long
    n = p
    Debug.Print n
    Dim v
    v = p
    Debug.Print v
    Debug.Print CStr(ObjPtr(o))
    Dim addr As String
    addr = p
    Debug.Print "addr: " & addr
    
    Dim ob As Sth
    Set ob = AsObject(p)
    ob.Value
End Sub
Sub setxyz()
    xyz = 9
    Set abc = New Sth
    abc.Value = 8
    ThisDocument.Variables("fake_lambda_proc_id") = abc
End Sub

Sub showxyz()
    Debug.Print xyz
    'Debug.Print abc.Value
    Dim v As Sth
    v = ThisDocument.Variables("fake_lambda_proc_id")
    Debug.Print v.Value
End Sub

Sub test_fakelambda()
    'Dim fakel As New FakeLambda
    'fakel.make "asdfsda"
    'Debug.Print ThisDocument.Variables("fake_lambda_proc_id")
    'fake_lambda_proc_id = 6
    make_fake_lambda ("Debug.Print 1")
End Sub

Sub aaaaaaaaaaaaaa()
    ActiveDocument.Save
End Sub

Sub test_fakelambda2()
    Dim fakel As New FakeLambdaMgr
    fakel.clearAll
    'Debug.Print ThisDocument.Variables("fake_lambda_proc_id")
End Sub

Sub aadfadfa()
    Dim o As Sth
    If o Is Nothing Then MsgBox "aa"
End Sub





Sub 브10()
    'Debug.Print CBool("true")
    Debug.Print CLng("123")
    'cbool
End Sub

Sub dsfaddf()
    Documents("匡도12").ActiveWindow.Visible = True
    'doc.Close wdDoNotSaveChanges
End Sub

Public Sub taoyan()
    MsgBox "taoyan"
End Sub

Sub aaaaa()
   
    Set Doc = Documents.Add
    Doc.ActiveWindow.Visible = False
    
    
    'doc.name = "for lambda"
    
    Dim proc_name As String
    Dim proc_body As String
    Dim proc_def As String
    
    proc_name = "qux"
    proc_body = "  Application.Run ""taoyan""  "
    proc_def = "Sub " & proc_name & "()" & vbCrLf & _
                proc_body & _
                vbCrLf & "End Sub"
    
    
    Dim c As VBComponent
    Set c = Doc.VBProject.VBComponents.Add(vbext_ct_StdModule)
    c.name = "lambdas"

    c.CodeModule.AddFromString proc_def
    

End Sub

Sub 브11()
'
' 브11 브
'
'
    Application.Run "qux"

    MsgBox abcxyz
End Sub


Function foobar(t As String, i As Long) As String
    't = "xyz"
    i = 99
    foobar = t
End Function


Sub test_return()
    Dim s As String
    s = "abc"
    Dim ss As String
    Dim i As Long
    i = 2
    ss = foobar(s, i)
    ss = "xyz"
End Sub
