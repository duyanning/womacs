Attribute VB_Name = "KeymapRelated"
' womacs

Option Explicit

' key_desc_seq only accept the kbd form string of emacs, such as "C-c [ p"
Sub global_set_key(key_desc_seq As String, cmd_name As String)
    
    '目前是根据keymap tree来设置word的key binding，本过程修改keymap tree，
    '所以必须在设置word的key binding之前调用，以后会取消这一项限制。
    '届时只要修改了global_keymap（无论是在global_keymap中增加除C-x、C-c之外的keymap，
    '还是在global_keymap中增加新的command），
    '都要修改相应的word key binding。
    '这要求可以从键代码得到键过程的名字。此过程先从键描述到键代码，再从键代码到键过程名字，
    '如果能把键描述跟键过程名字统一起来就方便多了。可以减少一次转换。
    Debug.Assert Not womacs_initialized
    
    Debug.Assert key_desc_seq <> ""
    
    Dim key_proc_name As String

    
    'split key_desc_seq into an array corresponding to ' ', thereofore every elem of the array is a key desc
    Dim key_desc_array As Variant
    key_desc_array = Split(key_desc_seq)
    
    'convert key desc to key code
    Dim key_code_array() As Long
    ReDim key_code_array(UBound(key_desc_array))
    Dim i As Integer
    For i = LBound(key_desc_array) To UBound(key_desc_array)
        key_code_array(i) = key_desc_to_key_code((key_desc_array(i)))
        'MsgBox key_desc_array(i) & " is " & key_code_array(i)
    Next i
    
    'there are at least one key in key_desc_seq
    Debug.Assert UBound(key_code_array) >= LBound(key_code_array)
    
    Dim this_keymap As KeyMap
    Set this_keymap = global_keymap
    
    Dim new_keymap As KeyMap
    
    For i = LBound(key_code_array) To UBound(key_code_array)
        Dim code As Long
        code = key_code_array(i)
        
        Dim entry As KeyMapEntry
        Set entry = this_keymap.lookup(code)
        
        'if there is no entry corresponding to i in this_keymap
        If entry Is Nothing Then
            'if i is the last key in key_desc_seq
            If i = UBound(key_code_array) Then
                'new an entry corresponding to i, which has type 'command' and content cmd_name
                this_keymap.add_cmd code, cmd_name
            'else
            Else
                'create a new keymap, i.e. new_keymap
                Set new_keymap = New KeyMap
                
                unset_all_keys new_keymap
                
                'set the name of new_keymap to corresponding key
                new_keymap.name = key_desc_array(i)
                
                'new an entry corresponding to i, which has type 'keymap' and content new_keymap
                Set entry = this_keymap.add_map(code, new_keymap)
                
                '如果this_keymap是global_keymap
                If i = LBound(key_code_array) Then
                    ' 从键描述得到对应的键过程名
                    
                    key_proc_name = key_desc_to_key_proc_name(key_desc_array(i))
                    
                    '设置相应的cmd
                    entry.command = key_proc_name
                End If
                
                'hook new_keymap on this_keymap
                Set new_keymap.Parent = this_keymap
                
                'make new_keymap this_keymap
                Set this_keymap = new_keymap
            End If
        'else (entry corresponding to i is found)
        Else
            'if i is the last key in key_desc_seq
            If i = UBound(key_code_array) Then
            
                'Debug.Assert Not entry.is_keymap
                
                'change the type of corresponding entry to 'command and content to cmd_name
                entry.is_keymap = False
                Set entry.map = Nothing
                entry.command = cmd_name
            'else
            Else
                'Debug.Assert entry.is_keymap
                
                'if the entry is not keymap
                If Not entry.is_keymap Then
                    'create a new keymap, i.e. new_keymap
                    Set new_keymap = New KeyMap
                    
                    
                    unset_all_keys new_keymap
                    
                    'set the name of new_keymap to corresponding key
                    new_keymap.name = key_desc_array(i)
                    
                    'hook it on this_keymap
                    Set new_keymap.Parent = this_keymap
                    
                    'change the type of the entry to 'keymap' and content to new_keymap
                    entry.is_keymap = True
                    Set entry.map = new_keymap
                    
                    
                End If
                
                Debug.Assert Not entry.map Is Nothing
                
                '如果this_keymap是global_keymap
                If i = LBound(key_code_array) Then
                    ' 从键描述得到对应的键过程名
                    key_proc_name = key_desc_to_key_proc_name(key_desc_array(i))
                    
                    '设置相应的cmd
                    entry.command = key_proc_name
                End If
                
                
                'make the keymap recorded in the entry this_keymap
                Set this_keymap = entry.map
            End If
        End If
    Next i
    
End Sub


Sub global_unset_key(key_desc_seq As String)
    Debug.Assert key_desc_seq <> ""
    
    'split key_desc_seq into an array according to ' ', therefore every elem in the array is a key desc
    Dim key_desc_array As Variant
    key_desc_array = Split(key_desc_seq)
    
    'convert key desc to key code
    Dim key_code_array() As Long
    ReDim key_code_array(UBound(key_desc_array))
    Dim i As Integer
    For i = LBound(key_desc_array) To UBound(key_desc_array)
        key_code_array(i) = key_desc_to_key_code((key_desc_array(i)))
        'MsgBox key_desc_array(i) & " is " & key_code_array(i)
    Next i
    
    'there are at least one key in key_desc_seq
    Debug.Assert UBound(key_code_array) >= LBound(key_code_array)
    
    Dim this_keymap As KeyMap
    Set this_keymap = global_keymap
    
    For i = LBound(key_code_array) To UBound(key_code_array)
        Dim code As Long
        code = key_code_array(i)
        
        Dim entry As KeyMapEntry
        Set entry = this_keymap.lookup(code)
        
        'if there is no entry corresponding to i in this_keymap
        If entry Is Nothing Then
            Exit For
        End If
        
        'entry corresponding to i is found
            
        'if i is the last key in key_desc_seq
        If i = UBound(key_code_array) Then
        
            'change the type of the entry to 'command', content to prompt_undefined
            entry.is_keymap = False
            Set entry.map = Nothing
            entry.command = "prompt_undefined"
        'else
        Else
            
            'the the entry is a keymap
            If Not entry.is_keymap Then
                'do not process remaining keys
                Exit For
            End If
            
            'make this_keymap pointing to the keymap pointed by the entry
            Set this_keymap = entry.map
            
        End If
        
    Next i
    
End Sub


'look up the entry corresponding to code in current_keymap and do sth according to entry type (command or keymap?)
Sub common_key_proc(code As Long)
    'look up the entry corresponding to code in current_keymap
    Dim entry As KeyMapEntry
    Set entry = current_keymap.lookup(code)
    
    'if it is not found
    If entry Is Nothing Then
        'prompt undefined
        MsgBox "undefinded key"
    End If
    
    'if the entry type is 'command'
    If Not entry.is_keymap Then
    
        'if executing a command
        If Not is_desc_key Then
            'execute the command recorded in the entry
            Application.Run macroname:=entry.command
        
        'else (describing a key seq)
        Else
            Dim doc_string_proc_name As String
            doc_string_proc_name = "doc_of_" & entry.command
            
            On Error Resume Next ' there may not be corresponding doc_string_proc
            Application.Run macroname:=doc_string_proc_name
            
            'show information about the command
            MsgBox doc_string, Title:=entry.command
            
            reset_doc_string
            
            is_desc_key = False
            statusbar_prefix = ""
        End If
    
'        'reset num_arg to 1 (because keys in global_keymap are bound directly, does not reset here)
'        num_arg = 1
        
        'resotre to global_keymap
        Set current_keymap = global_keymap
        
        'if not describing key seq
        If Not is_desc_key Then
        
            'restore key bindings of global_keymap
            global_keymap.bind ActiveDocument
            
            'the command is complete, restore the IME status
            ImeRestoreStatus
        'else (if describing key seq)
        Else
            'do nothing
        End If
    
    'else (entry type is 'keymap')
    Else

        ' if current_keymap is global_keymap and not describing key seq
        'note: if describing key seq, the IME status had been saved when C-h k and can be saved again
        'note: and the key bindings is in normal state (should not be in global_keymap speedup state)
        If current_keymap.name = "global" And Not is_desc_key Then
            'when the first prefix key is pressed, turn off IME
            ImeSaveAndSetStatus False
            
            'bind all keys to corresponding key proc (i.e. normal state)
            bind_all_key_procs ActiveDocument
        End If
        
        
        'make the keymap recorded in the enry current_keymap
        Set current_keymap = entry.map
        
        'prompt prefix key has been pressed in the status bar
        StatusBar = statusbar_prefix & current_keymap.FullName
    End If
End Sub

