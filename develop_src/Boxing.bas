Attribute VB_Name = "Boxing"
Option Explicit


Public Function BoxInteger(ByVal X As Integer) As CInteger
    Dim i As New CInteger
    i = X
    Set BoxInteger = i
End Function

Public Function BoxString(ByVal X As String) As CString
    Dim s As New CString
    s = X
    Set BoxString = s
End Function


