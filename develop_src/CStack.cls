VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICollection
Implements ICloneable

Private al As CArrayList

Public Property Get Count() As Integer
    Let Count = al.Count
End Property

Public Sub push(ByVal obj As Object)
    al.Add obj
End Sub

Public Function pop() As Object
    Dim top As Object
    Set top = al(al.Count - 1)
    al.RemoveAt (al.Count - 1)
    Set pop = top
End Function

Public Sub Clear()
    al.Clear
End Sub

Public Function Peek() As Object
    Set Peek = al(al.Count - 1)
End Function

Private Sub Class_Initialize()
    Set al = New CArrayList
End Sub


Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = al.Clone
End Function

Private Property Get ICollection_Count() As Integer
    ICollection_Count = al.Count
End Property
