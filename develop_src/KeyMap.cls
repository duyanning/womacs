VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' womacs
Option Explicit

Private al As New CArrayList
Private m_name As String
Private m_parent As KeyMap

Public Property Get name() As String
    name = m_name
End Property


Public Property Let name(ByVal vNewValue As String)
    m_name = vNewValue
End Property


Public Property Get Parent() As KeyMap
    Set Parent = m_parent
End Property


Public Property Set Parent(ByVal vNewValue As KeyMap)
    Set m_parent = vNewValue
End Property


Public Property Get FullName() As String
    '如果我是顶级keymap
    If m_parent Is Nothing Then
        '返回空字符串
        FullName = ""
    '否则
    Else
        '返回父亲全名&自己的名字
        FullName = m_parent.FullName & " " & m_name
    End If
End Property


Public Function add_cmd(key_code As Long, cmd As String) As KeyMapEntry
    Dim entry As New KeyMapEntry
    
    entry.key_code = key_code
    entry.is_keymap = False
    entry.command = cmd
    Set entry.map = Nothing
    
    al.Add entry
    
    Set add_cmd = entry
End Function

Public Function add_map(key_code As Long, map As KeyMap, Optional cmd As String = "") As KeyMapEntry
    Dim entry As New KeyMapEntry
    
    entry.key_code = key_code
    entry.is_keymap = True
    entry.command = cmd
    Set entry.map = map
    
    al.Add entry
    
    Set add_map = entry
End Function

'Public Sub Remove(keyCode As Long)
'
'End Sub

Public Sub bind(ByVal Doc As Document)

    '只有global_map中的被绑定
    Debug.Assert m_name = "global"

    push_saved_state Doc

'    Set current_keymap = Me

    CustomizationContext = Doc
    
    'Application.Visible = False
    turn_on_off_screen_updating False

    KeyBindings.clearAll


    Dim i As Integer
    'keymap中，同一个键码可能有多个条目，靠后的优先级高，所以绑定时要从前开始
    For i = 0 To al.Count - 1
        Dim entry As KeyMapEntry
        Set entry = al(i)
        
        Debug.Assert entry.command <> ""
        
        KeyBindings.Add KeyCode:=entry.key_code, _
            KeyCategory:=wdKeyCategoryCommand, command:=entry.command
        
        
    Next i


    'Application.Visible = True
    turn_on_off_screen_updating True
    pop_saved_state Doc

End Sub

Public Sub Clear()
    al.Clear
End Sub

Public Function lookup(code As Long) As KeyMapEntry
'    Dim found As Boolean
'    found = False
    
    Dim entry As KeyMapEntry
    Set entry = Nothing
    
    '依次对比，查找code对应的条目
    Dim i As Integer
    For i = al.Count - 1 To 0 Step -1
        Set entry = al(i)
        
        '如果找到了
        If entry.key_code = code Then
            '停止查找
            Exit For
        End If
    Next i
    
    Set lookup = entry

End Function
