Attribute VB_Name = "KeyHelpers"
' womacs

Option Explicit


Function key_desc_to_key_code(key_desc As String) As Long
    Dim code As WdKey
    Dim ascii As Long
    Dim Shift As Boolean
    
    Shift = False
    
    'convert key_desc to upper case
    key_desc = UCase(key_desc)
    
    'split key_desc to an array according to '-'
    Dim part_array As Variant
    part_array = Split(key_desc, "-")
    
    'how many parts are there?
    Dim part_num As Integer
    part_num = UBound(part_array) + 1
    
    'the last part is the main part
    Dim main_part As String
    main_part = part_array(part_num - 1)
    
    'only can take form like x, C-x, M-x, C-M-x.
    ' x may be single character of forms like SPC, DEL
    Debug.Assert part_num = 1 Or part_num = 2 Or part_num = 3
    
    'get the ascii code of main part
    ascii = Asc(main_part)
        
    
    'following should be corresponding to bind_all_key_procs
    'if the main part is a single character
    'get its code, only 0-9, A-Z has key code same as its ascii of upper case, else not.
    If Len(main_part) = 1 And ((ascii >= 48 And ascii <= 57) Or (ascii >= 65 And ascii <= 90)) Then
        code = ascii
    ElseIf key_desc = "[" Then
        code = wdKeyOpenSquareBrace
    ElseIf key_desc = "]" Then
        code = wdKeyCloseSquareBrace
    ElseIf key_desc = "\" Then
        code = wdKeyBackSlash
    ElseIf key_desc = ";" Then
        code = wdKeySemiColon
    ElseIf key_desc = "'" Then
        code = wdKeySingleQuote
    ElseIf key_desc = """" Then
        code = wdKeySingleQuote
        Shift = True
    ElseIf key_desc = "," Then
        code = wdKeyComma
    ElseIf key_desc = "." Then
        code = wdKeyPeriod
    ElseIf key_desc = "/" Then
        code = wdKeySlash
    ElseIf key_desc = "`" Then
        code = wdKeyBackSingleQuote
    ElseIf key_desc = "-" Then
        code = wdKeyHyphen
    ElseIf key_desc = "=" Then
        code = wdKeyEquals
    ElseIf key_desc = "SPC" Then
        code = wdKeySpacebar
    ElseIf key_desc = "DEL" Then
        code = wdKeyBackspace
    Else
        Debug.Assert False
    End If
    
    'if it is a single char
    If part_num = 1 Then
        If Shift Then
            key_desc_to_key_code = BuildKeyCode(wdKeyShift, code)
        Else
            key_desc_to_key_code = BuildKeyCode(code)
        End If
        
        'MsgBox code
        Exit Function
    End If
    
    'get the first modifier
    Dim modifier_desc_1 As String
    modifier_desc_1 = part_array(0)
    
    Dim modifier_code_1 As WdKey
    'if it is 'C', modifier_code is wdKeyControl
    If modifier_desc_1 = "C" Then
        modifier_code_1 = wdKeyControl
    'if it is 'M', modifier_code is wdKeyAlt
    ElseIf modifier_desc_1 = "M" Then
        modifier_code_1 = wdKeyAlt
    Else
        Debug.Assert False
    End If
    
    
    'if there is only one modifier
    If part_num = 2 Then
        If Shift Then
            key_desc_to_key_code = BuildKeyCode(wdKeyShift, modifier_code_1, code)
        Else
            key_desc_to_key_code = BuildKeyCode(modifier_code_1, code)
        End If
        
        Exit Function
    End If

    'get the second modifier
    Dim modifier_desc_2 As String
    modifier_desc_2 = part_array(1)
    
    Dim modifier_code_2 As WdKey
    'if it is 'C', modifier_code is wdKeyControl
    If modifier_desc_2 = "C" Then
        modifier_code_2 = wdKeyControl
    'if it is 'M', modifier_code is wdKeyAlt
    ElseIf modifier_desc_2 = "M" Then
        modifier_code_2 = wdKeyAlt
    Else
        Debug.Assert False
    End If

    If Shift Then
        key_desc_to_key_code = BuildKeyCode(wdKeyShift, modifier_code_1, modifier_code_2, code)
    Else
        key_desc_to_key_code = BuildKeyCode(modifier_code_1, modifier_code_2, code)
    End If

End Function


' desc of sigle key to proc name
Function key_desc_to_key_proc_name(ByVal key_desc As String) As String
    'todo
    Dim name As String
    name = key_desc
    'name = Replace(name, " -", " Hyphen")
    name = Replace(name, "--", "-Hyphen")
    name = Replace(name, "_", "Underscore")
    name = Replace(name, "<f1>", "F1")
    
    name = Replace(name, "-", "_")
    
    key_desc_to_key_proc_name = "key_proc_" & name
End Function


Sub bind_range(l As Long, u As Long)
    Dim code As Long

    For code = l To u
        'get the text form of key code
        Dim key_desc As String
        key_desc = Chr(code)
        
        
        Dim key_proc_name As String

        'like x
        key_proc_name = "key_proc_" & key_desc
        KeyBindings.Add KeyCode:=code, _
            KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name

        'like C-x
        key_proc_name = "key_proc_C_" & key_desc
        KeyBindings.Add KeyCode:=BuildKeyCode(code, wdKeyControl), _
            KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name

        'like M-x
        key_proc_name = "key_proc_M_" & key_desc
        KeyBindings.Add KeyCode:=BuildKeyCode(code, wdKeyAlt), _
            KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name
        
    Next code

End Sub


Sub bind(code As Long, key_desc As String, Optional Shift As Boolean = False)

    Dim key_proc_name As String
    Dim final_key_code As Long
    
    'like x
    key_proc_name = "key_proc_" & key_desc
    If Shift Then
        final_key_code = BuildKeyCode(wdKeyShift, code)
    Else
        final_key_code = BuildKeyCode(code)
    End If
    KeyBindings.Add KeyCode:=final_key_code, _
        KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name
        
    'like C-x
    key_proc_name = "key_proc_C_" & key_desc
    If Shift Then
        final_key_code = BuildKeyCode(wdKeyControl, wdKeyShift, code)
    Else
        final_key_code = BuildKeyCode(wdKeyControl, code)
    End If
    KeyBindings.Add KeyCode:=final_key_code, _
        KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name
        
    'like M-x
    key_proc_name = "key_proc_M_" & key_desc
    If Shift Then
        final_key_code = BuildKeyCode(wdKeyAlt, wdKeyShift, code)
    Else
        final_key_code = BuildKeyCode(wdKeyAlt, code)
    End If
    KeyBindings.Add KeyCode:=final_key_code, _
        KeyCategory:=wdKeyCategoryCommand, command:=key_proc_name
        
End Sub


Sub bind_all_key_procs(ByVal Doc As Document)
    push_saved_state Doc

    CustomizationContext = Doc
    
    'MsgBox "begin"
    turn_on_off_screen_updating False
    

    KeyBindings.clearAll '关掉这行貌似能快一点
    
    'following should be corresponding to key_desc_to_key_code
    bind_range wdKey0, wdKey9
    'DoEvents
    bind_range wdKeyA, wdKeyZ
    'DoEvents
    'MsgBox "1" '从begin到此处耗时很久，比从此处到end还久
    bind wdKeyOpenSquareBrace, "OpenSquareBrace"
    bind wdKeyCloseSquareBrace, "CloseSquareBrace"
    bind wdKeyBackSlash, "BackSlash"
    bind wdKeySemiColon, "SemiColon"
    bind wdKeySingleQuote, "SingleQuote"
    bind wdKeySingleQuote, "DoubleQuote", True
    bind wdKeyComma, "Comma"
    bind wdKeyPeriod, "Period"
    bind wdKeySlash, "Slash"
    bind wdKeyBackSingleQuote, "BackSingleQuote"
    bind wdKeyHyphen, "Hyphen"
    bind wdKeyEquals, "Equals"
    bind wdKeyBackspace, "Backspace"
    bind wdKeySpacebar, "Spacebar"
    
    turn_on_off_screen_updating True

    '经实测，从上面的begin到下面的end，两个消息框之间时间很长，光标闪烁
    '但怪异的是，如果在vba IDE中，在当前过程入口和出口各设一个断点，
    '这两个断点之间只要很短的时间就能完成，begin跟end两个消息也间隔很短
    '为什么断点的存在会影响程序行为呢？
    '因为vba IDE挡住了word主窗口，所以界面不会刷新。
    '所以我怀疑是word界面刷新导致。光标忽闪忽闪就是明证
    '按理说Application.ScreenUpdating = False应该有效的，但实测无效
    '估计是word2010之后的版本给改坏了
    '这就启发我用Application.Visible = False代替Application.ScreenUpdating = False
    '果然成功了
    'MsgBox "end"
    
    pop_saved_state Doc

End Sub

