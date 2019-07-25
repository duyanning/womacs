VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' womacs
Option Explicit

Implements ICloneable
Implements IList
Implements ICollection

Private m_data() As Object
Private m_count As Integer
Private m_capacity As Integer

Public Property Get Capacity() As Integer
    Let Capacity = m_capacity
End Property

Public Property Let Capacity(ByVal newcapacity As Integer)
    If newcapacity = 0 Then newcapacity = 16
    m_capacity = newcapacity
    ReDim Preserve m_data(m_capacity - 1)
End Property

Public Property Get Count() As Integer
    Let Count = actual_Count
End Property

Public Sub Clear()
    actual_Clear
End Sub

Public Function Add(ByVal Value As Object) As Integer
    Let Add = actual_Add(Value)
End Function

Public Property Get item(ByVal index As Integer) As Object
Attribute item.VB_UserMemId = 0
    Set item = actual_item(index)
End Property

Public Function NewEnum() As IUnknown
   Set NewEnum = m_data.[_NewEnum]
End Function

Public Function Equals(Object) As Boolean

End Function

Private Sub actual_Insert(ByVal index As Integer, ByVal Value As Object)
    If index < 0 Or index > m_count Then Debug.Assert False
    
    ' if it is inserted at end, call Add to do
    If index = m_count Then
        Add (Value)
        Exit Sub
    End If
    
    ' if there is no enough space
    If m_count = m_capacity Then
        m_capacity = m_capacity * 2
        ReDim Preserve m_data(m_capacity - 1)
    End If
    
    ' each elem from end to insertion point is moved in turn
    Dim i As Integer
    For i = m_count - 1 To index Step -1
        Set m_data(i + 1) = m_data(i)
    Next
    
    ' place new elem at insertion point
    Set m_data(index) = Value
    
    Let m_count = m_count + 1
End Sub

Public Sub Insert(ByVal index As Integer, ByVal Value As Object)
    actual_Insert index, Value
End Sub

Public Sub RemoveAt(ByVal index As Integer)
    IList_RemoveAt (index)
End Sub

Public Sub TrimToSize()
    Let m_capacity = m_count + 1
    ReDim m_data(m_capacity - 1)
End Sub

Private Sub Class_Initialize()
    Let m_capacity = 16
    ReDim m_data(m_capacity - 1)
    Let m_count = 0
End Sub

Private Function actual_Count() As Integer
    Let actual_Count = m_count
End Function

Private Property Get ICollection_Count() As Integer
    Let ICollection_Count = actual_Count
End Property

Private Function actual_Add(ByVal Value As Object) As Integer
    If m_count = m_capacity Then
        m_capacity = 2 * m_capacity
        ReDim Preserve m_data(m_capacity - 1)
    End If
    
    Set m_data(m_count) = Value
    Let m_count = m_count + 1
End Function

Private Function IList_Add(ByVal Value As Object) As Integer
    actual_Add Value
End Function

Private Function actual_Clone() As Object
    Dim al As New CArrayList

    Dim i As Integer
    For i = 0 To m_count - 1
        al.Add m_data(i)
    Next

    Set actual_Clone = al
End Function

Public Function Clone() As Object
    Set Clone = actual_Clone
End Function

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = actual_Clone
End Function

Private Sub actual_Clear()
    Dim o As Variant
    For Each o In m_data
        Set o = Nothing
    Next
    Let m_count = 0
End Sub

Private Sub IList_Clear()
    actual_Clear
End Sub

Private Function actual_Contains(ByVal item As Object) As Boolean
    If m_count = 0 Then
        Let Contains = False
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To m_count - 1
        If m_data(i) = obj Then
            Contains = True
            Exit Function
        End If
    Next
        
    Let Contains = False
End Function

Private Function IList_Contains(ByVal item As Object) As Boolean
    actual_Contains
End Function

Private Function actual_item(ByVal index As Integer) As Object
    If index < 0 Or index >= m_count Then Debug.Assert False
    Set actual_item = m_data(index)
End Function

Private Property Get IList_Count() As Integer
    IList_Count = actual_Count
End Property

Private Sub IList_Insert(ByVal index As Integer, ByVal Value As Object)
    actual_Insert index, Value
End Sub

Private Property Get IList_item(ByVal index As Integer) As Object
    Set IList_item = actual_item(index)
End Property

Private Sub actual_Remove(ByVal obj As Object)
    If m_count = 0 Then Debug.Assert False
    
    Dim i As Integer
    For i = 0 To m_count - 1
        If m_data(i) = obj Then Exit For
    Next
    
    If Not i = m_count Then RemoveAt (i)
End Sub

Private Sub IList_Remove(ByVal obj As Object)
    actual_Remove obj
End Sub

Private Sub actual_RemoveAt(ByVal index As Integer)
    If index < 0 Or index >= m_count Then Exit Sub
    If index = m_count - 1 Then
        Set m_data(m_count - 1) = Nothing
        m_count = m_count - 1
        Exit Sub
    End If
    
    Dim i As Integer
    For i = index To m_count - 2
        Set m_data(i) = m_data(i + 1)
    Next
    
    Let m_count = m_count - 1
End Sub

Private Sub IList_RemoveAt(ByVal index As Integer)
    actual_RemoveAt index
End Sub
