VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private X As Integer

Public Property Get Value() As Integer
Attribute Value.VB_UserMemId = 0
    Value = X
End Property

Public Property Let Value(ByVal vNewValue As Integer)
    X = vNewValue
End Property
