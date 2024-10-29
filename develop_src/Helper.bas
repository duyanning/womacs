Attribute VB_Name = "Helper"
' womacs
Option Explicit

' Application.ScreenUpdating = False��������
' ֻ����Application.Visible = False����
' �����������ΪTrue�ͱ���word�����ڵļ���/ʧ���ǳ�����Ϊ��ģ�⣨ScreenUpdating��������������
' ������֤�����ַ�ʽ�в�ͨ
Public as_screen_updating As Boolean
'Public event_count As Integer


Dim old_doc_saved_state As Boolean
Dim push_num As Long    ' for debug

Sub turn_on_off_screen_updating(ByVal flag As Boolean)
    Application.ScreenUpdating = flag
    'Application.ScreenUpdating = False��������
    'ֻ����Application.Visible = False������
    '������һЩ����
    '����������Ҫ��һ����������������Application.Visible�ı仯��Ϊ��ģ��Application.ScreenUpdating�ı仯
    as_screen_updating = True
'    event_count = event_count + 1  '�������н����������¼�������˵����һ���л����word���ڴ���
    Application.Visible = flag
'    DoEvents
    as_screen_updating = False
    'Ȼ������WindowDeactivate/WindowActivate�¼�����ʱ����⵽��as_screen_updating����False
    'Ҳ����˵Application.Visible = flag�Ƿ������ؼ�����WindowDeactivate/WindowActivate�¼�
    '��WindowDeactivate/WindowActivate�¼�������򱻵���ʱ��as_screen_updating = False�Ѿ�����
    '�������Ǽ�⵽��ֵΪFalse
    
    
    '������1��C-x C-sʱ��������C-xʱ��word���ڻ��������
    '��Ҫ�û��Ե�Ƭ�̣���word��������֮���ٰ���C-s
    '����������ڰ���C-s����û��һ����word�����ڵò����ð���
    '���word��������󣬳�Ϊǰ̨���ڵ�Ҳ��һ��word���ڣ�������Ҳ����womacsģʽ��
    'C-s�����µ�ǰ̨�����в�������
    
    '������2����������󣬿��ܻ����µ��ĵ���ΪActiveDocument
    'ʹ��ActiveDocumentʱһ��Ҫ����
End Sub


Sub push_saved_state(ByVal Doc As Document)
    Debug.Assert push_num = 0
    old_doc_saved_state = Doc.Saved
    push_num = push_num + 1
End Sub


Sub pop_saved_state(ByVal Doc As Document)
    Debug.Assert push_num = 1
    Doc.Saved = old_doc_saved_state
    push_num = push_num - 1
End Sub


Sub set_word_app_caption()
    Application.Caption = ""
End Sub


Sub set_emacs_app_caption()
    Application.Caption = ""
    Application.Caption = Application.Caption & " (Emacs)"
End Sub


Sub unset_all_modified_keys(ByVal map As KeyMap)
    
    Dim i As Long
    Dim KeyCode As Long

    ' For i = 1 To 254
    For i = wdKey0 To wdKey9
        KeyCode = BuildKeyCode(i, wdKeyControl)
        map.add_cmd KeyCode, "prompt_undefined"

        KeyCode = BuildKeyCode(i, wdKeyAlt)
        map.add_cmd KeyCode, "prompt_undefined"

    Next i

    For i = wdKeyA To wdKeyZ
        KeyCode = BuildKeyCode(i, wdKeyControl)
        map.add_cmd KeyCode, "prompt_undefined"

        KeyCode = BuildKeyCode(i, wdKeyAlt)
        map.add_cmd KeyCode, "prompt_undefined"

    Next i

    For i = wdKeySemiColon To wdKeyBackSingleQuote
        KeyCode = BuildKeyCode(i, wdKeyControl)
        map.add_cmd KeyCode, "prompt_undefined"

        KeyCode = BuildKeyCode(i, wdKeyAlt)
        map.add_cmd KeyCode, "prompt_undefined"

    Next i

    For i = wdKeyOpenSquareBrace To wdKeySingleQuote
        KeyCode = BuildKeyCode(i, wdKeyControl)
        map.add_cmd KeyCode, "prompt_undefined"

        KeyCode = BuildKeyCode(i, wdKeyAlt)
        map.add_cmd KeyCode, "prompt_undefined"

    Next i


'    For i = wdKeyF1 To wdKeyF16
'        KeyCode = BuildKeyCode(i, wdKeyControl)
'        map.add_cmd KeyCode, "prompt_undefined"
'
'        KeyCode = BuildKeyCode(i, wdKeyAlt)
'        map.add_cmd KeyCode, "prompt_undefined"
'
'    Next i
    
    ' C-g
    map.add_cmd BuildKeyCode(wdKeyControl, wdKeyG), "keyboard_quit"
    'map.add_cmd BuildKeyCode(wdKeyControl, wdKeyG), "foo"

End Sub


Sub unset_all_sigle_keys(ByVal map As KeyMap)
    Dim i As Long
    
    For i = wdKey0 To wdKey9
        map.add_cmd BuildKeyCode(i), "prompt_undefined"
    Next i

    For i = wdKeyA To wdKeyZ
        map.add_cmd BuildKeyCode(i), "prompt_undefined"
    Next i

    For i = wdKeySemiColon To wdKeyBackSingleQuote
        map.add_cmd BuildKeyCode(i), "prompt_undefined"
    Next i
    
    For i = wdKeyOpenSquareBrace To wdKeySingleQuote
        map.add_cmd BuildKeyCode(i), "prompt_undefined"
    Next i

    ' C-g
    map.add_cmd BuildKeyCode(wdKeyControl, wdKeyG), "keyboard_quit"

End Sub


Sub unset_all_keys(ByVal map As KeyMap)
    unset_all_sigle_keys map
    unset_all_modified_keys map
End Sub


Sub disable_all_key_bindings(Doc As Document)
    
    Dim i As Long
    Dim KeyCode
    
    For i = 1 To 254
        KeyCode = BuildKeyCode(i, wdKeyControl)
        FindKey(KeyCode).Disable
            
        KeyCode = BuildKeyCode(i, wdKeyAlt)
        FindKey(KeyCode).Disable
            
    Next i
End Sub


Sub use_default_key_bindings()
'    Debug.Print "beg====================================="
'    show_all_key_bindings ActiveDocument

    CustomizationContext = ActiveDocument
    
    ' Disable before clearAll to avoid a bug of MS Word, which causes clearAlled KeyBindings not work
    Dim kbLoop As KeyBinding
    For Each kbLoop In KeyBindings
        kbLoop.Disable
    Next kbLoop
   
    KeyBindings.clearAll
   
    
'    Debug.Print "mid====================================="
'    show_all_key_bindings ActiveDocument
'    Debug.Print "end====================================="

End Sub


Sub showctrlf()
    CustomizationContext = ActiveDocument
    Dim KeyCode As Long
    KeyCode = BuildKeyCode(wdKeyF, wdKeyControl)
    Debug.Print womacs_on, FindKey(KeyCode).command
End Sub


Sub show_all_key_bindings(Doc As Document)
    Dim kbLoop As KeyBinding
    
    CustomizationContext = Doc
    
    For Each kbLoop In KeyBindings
        Debug.Print kbLoop.command & vbTab _
            & kbLoop.KeyString
    Next kbLoop

End Sub


'Sub show_all_key_bindings_in_normal()
'    Dim kbLoop As KeyBinding
'
'    CustomizationContext = NormalTemplate
'
'    For Each kbLoop In KeyBindings
'        Selection.InsertAfter kbLoop.Command & vbTab _
'            & kbLoop.KeyString & vbCr
'        Selection.Collapse Direction:=wdCollapseEnd
'    Next kbLoop
'
'End Sub


'is_prefix_of("abc", "") = False
'is_prefix_of("", "abc") = False
'is_prefix_of("", "") = False
'is_prefix_of("abc", "abc") = True
Public Function is_prefix_of(str_to_search_in As String, str_to_search_for As String) As Boolean
    If str_to_search_for = "" Then
        is_prefix_of = False
        Exit Function
    End If
    
    Dim pos As Long
    pos = InStr(1, str_to_search_in, str_to_search_for)
    is_prefix_of = pos <> 0
End Function


' �㲻��������
Public Sub ensure_selection_in_view()
    Dim sel_left As Long
    Dim sel_top As Long
    Dim sel_width As Long
    Dim sel_height As Long

    ActiveWindow.GetPoint sel_left, sel_top, sel_width, sel_height, Selection.Range
    
    Dim visible_top As Long
    '��ΪRibbon��Ruler�ĸ߶ȶ�������UsableHeight�ڣ�Ribbon�·�QAT�ĸ߶Ȳ�����UsableHeight����
    '����Ҫ�ȹر����ǲ���׼ȷ����ɼ�����λ��
    'ActiveWindow.ToggleRibbon
    ActiveWindow.ActivePane.DisplayRulers = False
    
    '���㲻�������ϱߵĴ��룬û�ȹص�Ribbon��Ruler���±ߵĴ����ִ����
    '��Spy++����һ�£�word 2010�м���Ƭ��ɫ�ġ�������ʾ�ĵ����ݵĴ��ڣ�����Ϊ��Microsoft Word Document��
    '��������ý���API��
    
    visible_top = ActiveWindow.top + (ActiveWindow.Height - ActiveWindow.UsableHeight)
    
    Dim visible_bottom As Long
    visible_bottom = visible_top + ActiveWindow.UsableHeight
    
    Debug.Print "selection top is " & sel_top
    Debug.Print "visible top is " & visible_top
    Debug.Print "visible bot is " & visible_bottom

    'ActiveWindow.ToggleRibbon
    'ActiveWindow.Panes(1).DisplayRulers = True

End Sub
