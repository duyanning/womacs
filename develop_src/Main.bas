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
    
    If Not womacs_on Then
        use_emacs_key_bindings
        womacs_on = True
        
        set_emacs_app_caption
        MsgBox "GNU Emacs mode", vbOKOnly, "Womacs"
    Else
        use_word_key_bindings
        womacs_on = False
        
        set_word_app_caption
        MsgBox "MS Word mode", vbOKOnly, "Womacs"
    End If
End Sub

