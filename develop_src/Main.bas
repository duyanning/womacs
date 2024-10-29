Attribute VB_Name = "Main"
' womacs
Option Explicit

Public womacs_initialized As Boolean
Public isearch_running As Boolean
Public doing_search As Boolean

Dim X As New EventClassModule

Private Sub register_event_handler()
    Set X.App = Word.Application
End Sub


Public Sub toggle_womacs()
    'MsgBox "00000: " & womacs_on
    
    If Not womacs_initialized Then
        register_event_handler
        build_global_keymap
        
        num_arg = 1
        reset_doc_string
        user_customize
        
        womacs_initialized = True
    End If
    
    If Application.Documents.Count < 1 Then
        MsgBox "There are no documents open!", vbOKOnly, "Womacs"
        Exit Sub
    End If
    
    'Debug.Print "before: " & womacs_on
    
    'If Not womacs_on Then
    If Not get_womacs_status(ActiveDocument) Then
        'womacs_on = True
        set_womacs_status ActiveDocument, True
        use_emacs_key_bindings
        
        'Debug.Print "before: " & womacs_on
        
        MsgBox "GNU Emacs mode", vbOKOnly, "Womacs" '我感觉弹出消息框都会让word主窗口失活
        set_emacs_app_caption
    Else
        'womacs_on = False
        set_womacs_status ActiveDocument, False
        use_word_key_bindings
        
        MsgBox "MS Word mode", vbOKOnly, "Womacs"
        set_word_app_caption
    End If

    'MsgBox "after: " & womacs_on

End Sub

