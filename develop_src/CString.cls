VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private str As String

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = str
End Property

Public Property Let Value(ByVal vNewValue As String)
    str = vNewValue
End Property
