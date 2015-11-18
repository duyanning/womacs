Attribute VB_Name = "EmacsLikeKeyBindings"
' womacs

Option Explicit


' Key Binding Conventions
' http://www.gnu.org/software/emacs/elisp/html_node/Key-Binding-Conventions.html


Public current_keymap As KeyMap ' should be per doc

Public global_keymap As KeyMap
Public C_x_keymap As KeyMap
Public C_c_keymap As KeyMap
'Public C_c_OpenSquareBrace_keymap As KeyMap
'Public C_c_CloseSquareBrace_keymap As KeyMap
Public C_c_c_keymap As KeyMap
Public C_c_o_keymap As KeyMap
Public C_c_s_keymap As KeyMap
Public M_j_keymap As KeyMap
Public C_h_keymap As KeyMap


'Sub build_all_keymaps()
'    build_global_keymap
'End Sub


Sub use_word_key_bindings()
    push_saved_state ActiveDocument
    use_default_key_bindings
    pop_saved_state ActiveDocument

'    Set current_keymap = Nothing
    
'    MsgBox "word"
'    Debug.Assert current_keymap Is Nothing

End Sub


Sub use_emacs_key_bindings()
'    Debug.Assert current_keymap Is Nothing
    Set current_keymap = global_keymap
    global_keymap.bind ActiveDocument
    
End Sub



Sub build_global_keymap()
    Set global_keymap = New KeyMap
    global_keymap.name = "global"
    Set global_keymap.Parent = Nothing

    global_keymap.Clear
    
    ' there may be more than one entry corresponing to a key
    ' an entriy near the end of a keymap has higher priority than an entry near the begin
    
    unset_all_modified_keys global_keymap
    
    ' C-z
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyZ), "undo"
    
    ' C-/
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeySlash), "undo"
    
    ' C-SPC
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeySpacebar), "set_mark_command"

    ' C-@
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyShift, wdKey2), "set_mark_command"

    ' C-d
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyD), "delete_char"
    
    ' M-d
'    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyD), "DeleteWord"
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyD), "kill_word"
    
    ' M-backspace
'    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyBackspace), "DeleteBackWord"
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyBackspace), "backward_kill_word"
    
'    ' C-g
'    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyG), "keyboard_quit"
    
    ' C-w
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyW), "kill_region"
    
    ' M-w
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyW), "kill_ring_save"
    
    ' C-k
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyK), "kill_line"
    
    ' C-y
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyY), "yank"
    
    ' M-u
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyU), "upcase_word"
    
    ' M-l
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyL), "downcase_word"
    
    ' M-c
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyC), "capitalize_word"
    
    ' C-t
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyT), "transpose_chars"
    
    ' M-t
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyT), "transpose_words"
    
    ' C-m
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyM), "newline"
    
    ' C-u
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyU), "universal_argument"
    
    ' M-x
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyX), "execute_extended_command"
    
    ' M-f
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyF), "forward_word"
    
    ' M-b
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyB), "backward_word"

    ' C-f
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyF), "forward_char"
    
    ' C-b
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyB), "backward_char"
    
    ' C-n
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyN), "next_line"
    
    ' C-p
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyP), "previous_line"

    ' C-e
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyE), "move_end_of_line"
    
    ' C-a
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyA), "move_beginning_of_line"
    
    ' M-e
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyE), "forward_sentence"
    
    ' M-a
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyA), "backward_sentence"
    
    ' C-s
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyS), "isearch_forward"
    
    ' C-r
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyR), "isearch_backward"
    
    ' C-M-s
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), "isearch_forward_regexp"
    
    ' C-M-r
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyR), "isearch_backward_regexp"
    
    
    ' M-}
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyCloseSquareBrace), "forward_paragraph"
    
    ' M-{
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyOpenSquareBrace), "backward_paragraph"
    
    ' C-v
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyV), "scroll_up_command"
    
    ' M-v
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyV), "scroll_down_command"
    
    ' C-l
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyL), "recenter"
        
    ' C-o
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyO), "open_line"
    
    ' C-home
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyHome), "beginning_of_buffer"

    ' C-end
    global_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyEnd), "end_of_buffer"

    ' M-\
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeyBackSlash), "delete_horizontal_space"

    ' M-SPC
    global_keymap.add_cmd BuildKeyCode(wdKeyAlt, wdKeySpacebar), "just_one_space"

    ' C-x
    build_C_x_keymap
    global_keymap.add_map BuildKeyCode(wdKeyControl, wdKeyX), C_x_keymap, "key_proc_C_X"

    ' C-c
    build_C_c_keymap
    global_keymap.add_map BuildKeyCode(wdKeyControl, wdKeyC), C_c_keymap, "key_proc_C_C"

    ' M-j
    build_M_j_keymap
    global_keymap.add_map BuildKeyCode(wdKeyAlt, wdKeyJ), M_j_keymap, "key_proc_M_J"

    ' C-h
    build_C_h_keymap
    global_keymap.add_map BuildKeyCode(wdKeyControl, wdKeyH), C_h_keymap, "key_proc_C_H"

End Sub


Sub build_C_x_keymap()

    Set C_x_keymap = New KeyMap
    C_x_keymap.name = "C-x"
    Set C_x_keymap.Parent = global_keymap

    C_x_keymap.Clear
    unset_all_sigle_keys C_x_keymap
    
  
'    ' C-g
'    C_x_keymap.Add BuildKeyCode(wdKeyControl, wdKeyG), "prompt_undefined"
    
    ' C-l
    C_x_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyL), "downcase_region"

    ' C-u
    C_x_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyU), "upcase_region"
    
    ' C-s
    C_x_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyS), "save_buffer"
    
    ' C-x
    C_x_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyX), "exchange_point_and_mark"
    
    ' h
    C_x_keymap.add_cmd BuildKeyCode(wdKeyH), "mark_whole_buffer"
    
    ' [
    C_x_keymap.add_cmd BuildKeyCode(wdKeyOpenSquareBrace), "backward_page"
    
    ' ]
    C_x_keymap.add_cmd BuildKeyCode(wdKeyCloseSquareBrace), "forward_page"
    
    ' 2
    C_x_keymap.add_cmd BuildKeyCode(wdKey2), "split_window_vertically"
    
    ' 1
    C_x_keymap.add_cmd BuildKeyCode(wdKey1), "delete_other_windows"
    
    ' 0
    C_x_keymap.add_cmd BuildKeyCode(wdKey0), "delete_window"
    
    ' o
    C_x_keymap.add_cmd BuildKeyCode(wdKeyO), "other_window"
    
    
End Sub


Sub build_C_c_keymap()
    Set C_c_keymap = New KeyMap
    C_c_keymap.name = "C-c"
    Set C_c_keymap.Parent = global_keymap
    
    
    C_c_keymap.Clear
    
    unset_all_sigle_keys C_c_keymap
    
'    ' C-g
'    C_c_keymap.Add BuildKeyCode(wdKeyControl, wdKeyG), "prompt_undefined"
    
    ' C-y
    C_c_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyY), "word_yank_plain_text"
    
    ' C-c
    C_c_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyC), "word_show_all"
    
    ' b
    C_c_keymap.add_cmd BuildKeyCode(wdKeyB), "word_toggle_bold"
    
    ' i
    C_c_keymap.add_cmd BuildKeyCode(wdKeyI), "word_toggle_italic"

    ' x
    C_c_keymap.add_cmd BuildKeyCode(wdKeyX), "word_toggle_strike_through"

    ' C-l
    C_c_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyL), "word_toggle_subscript"

    ' C-u
    C_c_keymap.add_cmd BuildKeyCode(wdKeyControl, wdKeyU), "word_toggle_superscript"

    ' w
    C_c_keymap.add_cmd BuildKeyCode(wdKeyW), "delete_region"
    
    ' h
    C_c_keymap.add_cmd BuildKeyCode(wdKeyH), "word_toggle_highlight"

'    ' r
'    C_c_keymap.add_cmd BuildKeyCode(wdKeyR), "word_toggle_font_color_red"

    ' n
    C_c_keymap.add_cmd BuildKeyCode(wdKeyN), "word_style_normal"

    ' d
    C_c_keymap.add_cmd BuildKeyCode(wdKeyD), "word_style_normal_indent"

    ' t
    C_c_keymap.add_cmd BuildKeyCode(wdKeyT), "word_style_title"
    
    ' 1
    C_c_keymap.add_cmd BuildKeyCode(wdKey1), "word_style_heading1"

    ' 2
    C_c_keymap.add_cmd BuildKeyCode(wdKey2), "word_style_heading2"

    ' 3
    C_c_keymap.add_cmd BuildKeyCode(wdKey3), "word_style_heading3"

    ' 4
    C_c_keymap.add_cmd BuildKeyCode(wdKey4), "word_style_heading4"

    ' 5
    C_c_keymap.add_cmd BuildKeyCode(wdKey5), "word_style_heading5"

    ' 6
    C_c_keymap.add_cmd BuildKeyCode(wdKey6), "word_style_heading6"

    ' 7
    C_c_keymap.add_cmd BuildKeyCode(wdKey7), "word_style_heading7"

    ' 8
    C_c_keymap.add_cmd BuildKeyCode(wdKey8), "word_style_heading8"

    ' 9
    C_c_keymap.add_cmd BuildKeyCode(wdKey9), "word_style_heading9"

'    ' [
'    build_C_c_OpenSquareBrace_keymap
'    C_c_keymap.add_map BuildKeyCode(wdKeyOpenSquareBrace), C_c_OpenSquareBrace_keymap
'
'    ' ]
'    build_C_c_CloseSquareBrace_keymap
'    C_c_keymap.add_map BuildKeyCode(wdKeyCloseSquareBrace), C_c_CloseSquareBrace_keymap
    
    ' c
    build_C_c_c_keymap
    C_c_keymap.add_map BuildKeyCode(wdKeyC), C_c_c_keymap
    
    ' o
    build_C_c_o_keymap
    C_c_keymap.add_map BuildKeyCode(wdKeyO), C_c_o_keymap
    
    ' s
    build_C_c_s_keymap
    C_c_keymap.add_map BuildKeyCode(wdKeyS), C_c_s_keymap
End Sub


'Sub build_C_c_OpenSquareBrace_keymap()
'    Set C_c_OpenSquareBrace_keymap = New KeyMap
'    C_c_OpenSquareBrace_keymap.name = "["
'    Set C_c_OpenSquareBrace_keymap.Parent = C_c_keymap
'
'
'    C_c_OpenSquareBrace_keymap.Clear
'
'    unset_all_sigle_keys C_c_OpenSquareBrace_keymap
'
'    ' '
'    C_c_OpenSquareBrace_keymap.add_cmd BuildKeyCode(wdKeySingleQuote), "word_single_opening_quote"
'
'    ' "
'    C_c_OpenSquareBrace_keymap.add_cmd BuildKeyCode(wdKeyShift, wdKeySingleQuote), "word_double_opening_quote"
'
'
'End Sub


'Sub build_C_c_CloseSquareBrace_keymap()
'    Set C_c_CloseSquareBrace_keymap = New KeyMap
'    C_c_CloseSquareBrace_keymap.name = "]"
'    Set C_c_CloseSquareBrace_keymap.Parent = C_c_keymap
'
'
'    C_c_CloseSquareBrace_keymap.Clear
'
'    unset_all_sigle_keys C_c_CloseSquareBrace_keymap
'
'    ' '
'    C_c_CloseSquareBrace_keymap.add_cmd BuildKeyCode(wdKeySingleQuote), "word_single_closing_quote"
'
'    ' "
'    C_c_CloseSquareBrace_keymap.add_cmd BuildKeyCode(wdKeyShift, wdKeySingleQuote), "word_double_closing_quote"
'
'
'End Sub


Sub build_C_c_c_keymap()
    Set C_c_c_keymap = New KeyMap
    C_c_c_keymap.name = "c"
    Set C_c_c_keymap.Parent = C_c_keymap
    
    
    C_c_c_keymap.Clear
    
    unset_all_sigle_keys C_c_c_keymap
 
    ' a
    C_c_c_keymap.add_cmd BuildKeyCode(wdKeyA), "word_toggle_font_color_auto"

    ' r
    C_c_c_keymap.add_cmd BuildKeyCode(wdKeyR), "word_toggle_font_color_red"

    ' g
    C_c_c_keymap.add_cmd BuildKeyCode(wdKeyG), "word_toggle_font_color_green"

    ' b
    C_c_c_keymap.add_cmd BuildKeyCode(wdKeyB), "word_toggle_font_color_blue"


End Sub


Sub build_C_c_o_keymap()
    Set C_c_o_keymap = New KeyMap
    C_c_o_keymap.name = "o"
    Set C_c_o_keymap.Parent = C_c_keymap
    
    
    C_c_o_keymap.Clear
    
    unset_all_sigle_keys C_c_o_keymap
 
    ' 1
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey1), "word_outline_level1"
    
    ' 2
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey2), "word_outline_level2"

    ' 3
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey3), "word_outline_level3"
    
    ' 4
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey4), "word_outline_level4"

    ' 5
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey5), "word_outline_level5"
    
    ' 6
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey6), "word_outline_level6"

    ' 7
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey7), "word_outline_level7"

    ' 8
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey8), "word_outline_level8"

    ' 9
    C_c_o_keymap.add_cmd BuildKeyCode(wdKey9), "word_outline_level9"

    ' b
    C_c_o_keymap.add_cmd BuildKeyCode(wdKeyB), "word_outline_bodytext"

End Sub


Sub build_C_c_s_keymap()
    Set C_c_s_keymap = New KeyMap
    C_c_s_keymap.name = "s"
    Set C_c_s_keymap.Parent = C_c_keymap
    
    
    C_c_s_keymap.Clear
    
    unset_all_sigle_keys C_c_s_keymap
 
    ' c
    C_c_s_keymap.add_cmd BuildKeyCode(wdKeyC), "caption_on_the_spot"


End Sub


Sub build_M_j_keymap()
    Set M_j_keymap = New KeyMap
    M_j_keymap.name = "M-j"
    Set M_j_keymap.Parent = global_keymap
    
    
    M_j_keymap.Clear
    unset_all_sigle_keys M_j_keymap
    

    ' l
    M_j_keymap.add_cmd BuildKeyCode(wdKeyL), "set_justification_left"
    
    ' c
    M_j_keymap.add_cmd BuildKeyCode(wdKeyC), "set_justification_center"
    
    ' r
    M_j_keymap.add_cmd BuildKeyCode(wdKeyR), "set_justification_right"
    
    ' b
    M_j_keymap.add_cmd BuildKeyCode(wdKeyB), "set_justification_full"
    
    
End Sub


Sub build_C_h_keymap()
    Set C_h_keymap = New KeyMap
    C_h_keymap.name = "C-h"
    Set C_h_keymap.Parent = global_keymap
    
    
    C_h_keymap.Clear
    unset_all_sigle_keys C_h_keymap
    
 
    ' k
    C_h_keymap.add_cmd BuildKeyCode(wdKeyK), "describe_key"
    
    ' f
    C_h_keymap.add_cmd BuildKeyCode(wdKeyF), "describe_function"
    
End Sub

