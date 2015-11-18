VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeLambdaMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_module As CodeModule
'Private m_id As Long

Private Sub Class_Initialize()
    Set m_module = ThisDocument.VBProject.VBComponents("fake_lambda_procs").CodeModule
    'm_id = 0
    'ThisDocument.Variables.Add Name:="fake_lambda_proc_id", Value:="0"

End Sub

Public Function make(proc_body As String)
    Dim id As Integer
    id = ThisDocument.Variables("fake_lambda_proc_id")
    
    Dim proc_name As String
    proc_name = "fake_lambda_proc_" & id
    
    Dim proc_def As String
    proc_def = "Sub " & proc_name & "()" & vbCrLf & _
                proc_body & _
                vbCrLf & "End Sub"
                
    
    m_module.AddFromString proc_def
    
    'ensure_proc_exists m_module, "", proc_def
    id = id + 1
    ThisDocument.Variables("fake_lambda_proc_id") = id
End Function

Public Sub clearAll()
    ThisDocument.Variables("fake_lambda_proc_id") = 0
    ThisDocument.Variables.Remove "fake_lambda_proc_id"
    
    ThisDocument.VBProject.VBComponents.Remove m_module.Parent
   
    Dim c As VBComponent
    Set c = ThisDocument.VBProject.VBComponents.Add(vbext_ct_StdModule)
    c.name = "fake_lambda_procs"
End Sub

