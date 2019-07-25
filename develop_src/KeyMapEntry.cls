VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' womacs
Option Explicit

Private m_key_code As Long
Private m_is_keymap As Boolean
Private m_command As String
Private m_map As KeyMap

Public Property Get is_keymap() As Boolean
    is_keymap = m_is_keymap
End Property

Public Property Let is_keymap(ByVal vNewValue As Boolean)
    m_is_keymap = vNewValue
End Property

Public Property Get command() As String
    command = m_command
End Property

Public Property Let command(ByVal vNewValue As String)
    m_command = vNewValue
End Property

Public Property Get key_code() As Long
    key_code = m_key_code
End Property

Public Property Let key_code(ByVal vNewValue As Long)
    m_key_code = vNewValue
End Property

Public Property Get map() As KeyMap
    Set map = m_map
End Property

Public Property Set map(ByVal vNewValue As KeyMap)
    Set m_map = vNewValue
End Property


