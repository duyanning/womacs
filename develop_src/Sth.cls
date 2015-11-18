VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_value As Long
Public Property Get Value() As String
    Value = m_value
End Property

Public Property Let Value(ByVal vNewValue As String)
    m_value = vNewValue
End Property
