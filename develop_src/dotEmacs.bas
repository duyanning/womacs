Attribute VB_Name = "dotEmacs"
' womacs

Option Explicit




' This is your .emacs

Sub user_customize()

    'kill_whole_line = False
    delete_word = False
    
    'global_unset_key "C-c c r"
    global_set_key "C-c r", "word_toggle_font_color_red"
    global_set_key "C-c a", "word_toggle_font_color_auto"
    
    global_set_key "C-c [ '", "word_single_opening_quote"
    global_set_key "C-c ] '", "word_single_closing_quote"

    'according the syntax of VBA, in string, " should be written as ""
    global_set_key "C-c [ """, "word_double_opening_quote"
    global_set_key "C-c ] """, "word_double_closing_quote"
    
    global_set_key "C-c u", "upcase_prev_word"
    
    global_set_key "C-c s p", "word_superscript"
    global_set_key "C-c s b", "word_subscript"
    global_set_key "C-c s n", "word_normalscript"
    
    If StrComp(Application.version, "14.0") > -1 Then
        'uncomment out the following two lines to use the
        'NavPaneSearch of MS-Word 2010 instead of the Emacs style I-search
        'global_set_key "C-s", "NavPaneSearch"
        'global_set_key "C-r", "NavPaneSearch"
    End If
    
    'Examples:
    'global_set_key "C-c c `", "foo"
    'global_set_key "C-q", "foo"
    'global_set_key "C-q a", "foo"
    
End Sub


Sub upcase_prev_word()
    num_arg = -1
    
    upcase_word
End Sub

'An example of user defined command.
Sub foo() 'Command name is foo, no arguments
    MsgBox "hi"
    complete 'Every command MUST call 'complete' before exiting
End Sub
'You can add a doc string to a command by defining corresponding 'doc_of_' procedure.
Sub doc_of_foo()
    'Assign the doc to the variable 'doc_string'
    doc_string = "foo is a user defined command."
End Sub

