Attribute VB_Name = "FakeLambda"
Option Explicit
Option Private Module

Public fake_lambda_proc_id As Integer

Public Function make_fake_lambda(proc_body As String)
    Dim id As Integer
    id = fake_lambda_proc_id
    
    Dim proc_name As String
    proc_name = "fake_lambda_proc_" & id
    
    Dim proc_def As String
    proc_def = "Sub " & proc_name & "()" & vbCrLf & _
                proc_body & _
                vbCrLf & "End Sub"
                
    Debug.Print "before: " & fake_lambda_proc_id
    'fake_lambda_proc_id = 7
    ThisDocument.VBProject.VBComponents("fake_lambda_procs").CodeModule.AddFromString proc_def
    Debug.Print "immd after: " & fake_lambda_proc_id
    
    
    id = id + 1
    fake_lambda_proc_id = id
    Debug.Print "after: " & fake_lambda_proc_id
End Function

