Attribute VB_Name = "EmacsLikeCommands"
' womacs

Option Explicit


Public is_desc_key As Boolean
Public statusbar_prefix As String

Public search_fwd As Boolean

' doc local
Public womacs_on As Boolean
Public mark_set As Boolean     ' has mark been set?
Public mark_pos As Long        ' position of the mark




Public num_arg As Long  ' numeric argument
' template to process num_arg
'    While num_arg < 0
'        do_a()
'
'        num_arg = num_arg + 1
'    Wend
'
'    While num_arg > 0
'        do_b()
'
'        num_arg = num_arg - 1
'    Wend
'    num_arg = 1  ' reset num_arg to 1

Dim action_record_balance As Integer    ' for debug

'Public kill_whole_line As Boolean
Public delete_word As Boolean


Sub store_doc_locals(ByVal Doc As Document)
    push_saved_state Doc
    
    Doc.Variables("womacs_on") = CStr(womacs_on)
    Doc.Variables("mark_set") = CStr(mark_set)
    Doc.Variables("mark_pos") = CStr(mark_pos)
    
    pop_saved_state Doc
End Sub


Sub load_doc_locals(ByVal Doc As Document)
    womacs_on = False
    mark_set = False
    mark_pos = -1

    On Error Resume Next
    womacs_on = CBool(Doc.Variables("womacs_on"))
    mark_set = CBool(Doc.Variables("mark_set"))
    mark_pos = CLng(Doc.Variables("mark_pos"))
End Sub


Sub delete_all_womacs_vars(ByVal Doc As Document)
    push_saved_state Doc
    
    On Error Resume Next

    Doc.Variables("womacs_on").Delete
    Doc.Variables("mark_set").Delete
    Doc.Variables("mark_pos").Delete
    
    pop_saved_state Doc

End Sub


Sub begin_action(name As String)
    ' need word 2010
    If StrComp(Application.version, "14.0") = -1 Then
        Exit Sub
    End If
    
    Debug.Assert action_record_balance = 0
    action_record_balance = action_record_balance + 1
    Application.UndoRecord.StartCustomRecord "Undo " & name
End Sub


Sub end_action()
    ' need word 2010
    If StrComp(Application.version, "14.0") = -1 Then
        Exit Sub
    End If
    
    Debug.Assert action_record_balance = 1
    action_record_balance = action_record_balance - 1
    Application.UndoRecord.EndCustomRecord
End Sub


Sub doc_of_set_mark_command()
    doc_string = "Set the mark where point is."
End Sub
Sub set_mark_command()
'C-SPC

    Dim old_mark_pos As Long

    If mark_set Then
        select_nothing
        ' if point has not been moved since mark is set, then cancel the mark
        If mark_pos = Selection.Start Then
            mark_set = False
            Application.StatusBar = "Mark deactivated"
            Exit Sub
        End If
    End If

    old_mark_pos = mark_pos

    mark_set = True
    mark_pos = Selection.Start
    'Debug.Print mark_pos

    If mark_pos = old_mark_pos Then
        Application.StatusBar = "Mark activated"
    Else
        Application.StatusBar = "Mark set"
    End If


    complete
End Sub


Sub select_nothing()
    ' if there is no highlight
    If Selection.type = wdSelectionIP Then
        Exit Sub
    End If

    ' move cursor to make highlight disapper
'    If mark_pos <= Selection.Start Then
'        Selection.Collapse Direction:=wdCollapseEnd
'    Else
'        Selection.Collapse Direction:=wdCollapseStart
'    End If
    
    If Selection.StartIsActive Then
        Selection.Collapse direction:=wdCollapseStart
    Else
        Selection.Collapse direction:=wdCollapseEnd
    End If
    

End Sub


Sub doc_of_keyboard_quit()
    doc_string = "Signal a `quit' condiction."
End Sub
Sub keyboard_quit()
'C-g
    Application.StatusBar = "Quit"

    mark_set = False
    select_nothing
    
    complete
    
    'restore to the global_keymap
    Set current_keymap = global_keymap
    global_keymap.bind ActiveDocument
    
    'restore the IME status
    ImeRestoreStatus

End Sub


Sub doc_of_kill_line()
    doc_string = "Move point right one character."
End Sub
Sub kill_line()
'C-k

    mark_set = False
    select_nothing
    
    ' move to the end of line
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend

    If Selection.type <> wdSelectionIP _
            And Asc(Selection.Characters.last.Text) = 13 _
            And Asc(Selection.Characters.First.Text) <> 13 Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If
    

    'Selection.StartIsActive = Not Selection.StartIsActive
    
    If Selection.type <> wdSelectionIP Then
        Selection.Cut
    End If


    complete
End Sub


Sub doc_of_kill_region()
    doc_string = "Kill (""cut"") text between point and mark."
End Sub
Sub kill_region()
'C-w
    On Error Resume Next
    'Selection.Cut
    Application.Run macroname:="EditCut"
    
    mark_set = False

    complete
End Sub


Sub doc_of_delete_region()
    doc_string = "Delete the text between point and mark."
End Sub
Sub delete_region()
'C-c w
    On Error Resume Next
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    mark_set = False

    complete
End Sub


'    While num_arg < 0
'        do_a()
'
'        num_arg = num_arg + 1
'    Wend
'
'    While num_arg > 0
'        do_b()
'
'        num_arg = num_arg - 1
'    Wend
Sub doc_of_kill_word()
    doc_string = "Kill a word forward, skipping over intervening delimiters."
End Sub
Sub kill_word()
'M-d
    On Error Resume Next

    mark_set = False
    select_nothing
    
    begin_action "kill word"
    
    If delete_word Then
        While num_arg > 0
            Application.Run macroname:="DeleteWord"
            num_arg = num_arg - 1
        Wend
    Else
        Selection.MoveRight Unit:=wdWord, Count:=num_arg, Extend:=wdExtend
        Selection.StartIsActive = Not Selection.StartIsActive   ' when select_nothing in undo, move point to a correct position
        Selection.Cut
    End If
    
    end_action
    
    complete
End Sub


Sub doc_of_backward_kill_word()
    doc_string = "Kill characters backward until encountering the beginning of a word."
End Sub
Sub backward_kill_word()
'M-backspace
    begin_action "backward kill word"
    
    If delete_word Then
        While num_arg > 0
            Application.Run macroname:="DeleteBackWord"
            num_arg = num_arg - 1
        Wend
    Else
        Selection.MoveLeft Unit:=wdWord, Count:=num_arg, Extend:=wdExtend
        Selection.StartIsActive = Not Selection.StartIsActive
        Selection.Cut
    End If
    
    end_action
    
    complete
    
End Sub


Sub doc_of_kill_ring_save()
    doc_string = "Copy"
End Sub
Sub kill_ring_save()
'M-w
    On Error Resume Next
    'Selection.Copy
    Application.Run macroname:="EditCopy"

    select_nothing
    mark_set = False


    complete
End Sub


Sub doc_of_yank()
    doc_string = "Paste"
End Sub
Sub yank()
'C-y
    mark_set = False
    select_nothing
    
    'Selection.PasteAndFormat (wdFormatPlainText)
    'Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Application.Run macroname:="EditPaste"

    complete
End Sub


Sub doc_of_delete_char()
    doc_string = "Delete the following one character."
End Sub
Sub delete_char()
'C-d
    mark_set = False
    select_nothing
    
    Selection.Delete Unit:=wdCharacter, Count:=num_arg

    complete
End Sub


Sub doc_of_upcase_word()
    doc_string = "Convert following word to upper case, moving over."
End Sub
Sub upcase_word()
'M-u
    select_nothing
    mark_set = False

    begin_action "upcase word"
    
    If num_arg < 0 Then
        Dim point_pos As Long
        point_pos = Selection.Start
        Selection.MoveLeft Unit:=wdWord, Count:=-num_arg

        While Selection.Start < point_pos
            Selection.Range.Case = wdUpperCase
            Selection.MoveRight Unit:=wdWord, Count:=1
        Wend


    End If

    While num_arg > 0
        Selection.Range.Case = wdUpperCase
        Selection.MoveRight Unit:=wdWord, Count:=1

        num_arg = num_arg - 1
    Wend
    
    end_action

    complete
End Sub


Sub doc_of_downcase_word()
    doc_string = "Convert following word to lower case, moving over."
End Sub
Sub downcase_word()
'M-l

    select_nothing
    mark_set = False
    
    begin_action "downcase word"

    If num_arg < 0 Then
        Dim point_pos As Long
        point_pos = Selection.Start
        Selection.MoveLeft Unit:=wdWord, Count:=-num_arg

        While Selection.Start < point_pos
            Selection.Range.Case = wdLowerCase
            Selection.MoveRight Unit:=wdWord, Count:=1
        Wend


    End If

    While num_arg > 0
        Selection.Range.Case = wdLowerCase
        Selection.MoveRight Unit:=wdWord, Count:=1

        num_arg = num_arg - 1
    Wend

    end_action
    
    complete
End Sub


Sub doc_of_capitalize_word()
    doc_string = "Capitalize the following word (or ARG words), moving over."
End Sub
Sub capitalize_word()
'M-c
    select_nothing
    mark_set = False
    
    begin_action "capitalize word"

    If num_arg < 0 Then
        Dim point_pos As Long
        point_pos = Selection.Start
        Selection.MoveLeft Unit:=wdWord, Count:=-num_arg

        While Selection.Start < point_pos
            Selection.Range.Case = wdTitleWord
            Selection.MoveRight Unit:=wdWord, Count:=1
        Wend


    End If

    While num_arg > 0
        Selection.Range.Case = wdTitleWord
        Selection.MoveRight Unit:=wdWord, Count:=1

        num_arg = num_arg - 1
    Wend

    end_action
    
    complete
End Sub


Sub doc_of_newline()
    doc_string = "Insert a newline."
End Sub
Sub newline()
'C-m
    select_nothing
    mark_set = False
    
    begin_action "newline"

    While num_arg > 0
        Selection.TypeParagraph

        num_arg = num_arg - 1
    Wend

    end_action
    complete
End Sub


Sub doc_of_open_line()
    doc_string = "Insert a newline and leave point before it."
End Sub
Sub open_line()
'C-o
    select_nothing
    mark_set = False
    
    begin_action ("open line")

    Dim i As Integer
    For i = 1 To num_arg
        Selection.TypeParagraph
    Next i
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=num_arg
    
    end_action

    complete

End Sub


Sub doc_of_forward_word()
    doc_string = "Move point forward one word."
End Sub
Sub forward_word()
'M-f
'    While num_arg < 0
'        do_a()
'
'        num_arg = num_arg + 1
'    Wend
'

'    While num_arg > 0
'
'        num_arg = num_arg - 1
'    Wend

    If mark_set Then
        Selection.MoveRight Unit:=wdWord, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveRight Unit:=wdWord, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_backward_word()
    doc_string = "Move backward until encountering the beginning of a word."
End Sub
Sub backward_word()
'M-b
    If mark_set Then
        Selection.MoveLeft Unit:=wdWord, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveLeft Unit:=wdWord, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_forward_char()
    doc_string = "Move point right one character."
End Sub
Sub forward_char()
'C-f
    If mark_set Then
        Selection.MoveRight Unit:=wdCharacter, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveRight Unit:=wdCharacter, Count:=num_arg, Extend:=wdMove
    End If
    
    complete
End Sub


Sub doc_of_backward_char()
    doc_string = "Move point left one character."
End Sub
Sub backward_char()
'C-b
    If mark_set Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveLeft Unit:=wdCharacter, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_previous_line()
    doc_string = "Move cursor vertically up one line."
End Sub
Sub previous_line()
'C-p
    If mark_set Then
        Selection.MoveUp Unit:=wdLine, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveUp Unit:=wdLine, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_next_line()
    doc_string = "Move cursor vertically down one line."
End Sub
Sub next_line()
'C-n
    If mark_set Then
        Selection.MoveDown Unit:=wdLine, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveDown Unit:=wdLine, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


' yet not support C-u
Sub doc_of_move_beginning_of_line()
    doc_string = "Move point to beginning of current line."
End Sub
Sub move_beginning_of_line()
'C-a
    If mark_set Then
        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Else
        Selection.HomeKey Unit:=wdLine, Extend:=wdMove
    End If

    complete
End Sub


' yet not support C-u
Sub doc_of_move_end_of_line()
    doc_string = "Move point to end of current line."
End Sub
Sub move_end_of_line()
'C-e
    If mark_set Then
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Else
        Selection.EndKey Unit:=wdLine, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_backward_sentence()
    doc_string = "Move backward to start of sentence."
End Sub
Sub backward_sentence()
'M-a
    If mark_set Then
        Selection.MoveLeft Unit:=wdSentence, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveLeft Unit:=wdSentence, Count:=num_arg, Extend:=wdMove
    End If
    
    complete
End Sub


Sub doc_of_forward_sentence()
    doc_string = "Move forward to next end of sentence."
End Sub
Sub forward_sentence()
'M-e
    If mark_set Then
        Selection.MoveRight Unit:=wdSentence, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveRight Unit:=wdSentence, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_beginning_of_buffer()
    doc_string = "Move point to the beginning of the buffer."
End Sub
Sub beginning_of_buffer()
'C-home
    If mark_set Then
        Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
    Else
        Selection.HomeKey Unit:=wdStory
    End If
    
    complete
End Sub


Sub doc_of_end_of_buffer()
    doc_string = "Move point to the end of the buffer."
End Sub
Sub end_of_buffer()
'C-end
    If mark_set Then
        Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Else
        Selection.EndKey Unit:=wdStory
    End If
    
    complete
End Sub


Sub doc_of_forward_page()
    doc_string = "Move forward to page boundary."
End Sub
Sub forward_page()
'C-x ]
    Selection.GoToNext What:=wdGoToPage
    
    complete
End Sub


Sub doc_of_backward_page()
    doc_string = "Move backward to page boundary."
End Sub
Sub backward_page()
'C-x [
    Selection.GoToPrevious What:=wdGoToPage
    
    complete
End Sub


Sub doc_of_forward_paragraph()
    doc_string = "Move forward to end of paragraph."
End Sub
Sub forward_paragraph()
'M-}
    If mark_set Then
        Selection.MoveDown Unit:=wdParagraph, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveDown Unit:=wdParagraph, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_backward_paragraph()
    doc_string = "Move backward to start of paragraph."
End Sub
Sub backward_paragraph()
'M-{
    If mark_set Then
        Selection.MoveUp Unit:=wdParagraph, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveUp Unit:=wdParagraph, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_scroll_up_command()
    doc_string = "Scroll text of selected window upward ARG lines; or near full screen if no ARG."
End Sub
Sub scroll_up_command()
'C-v
    If mark_set Then
        Selection.MoveDown Unit:=wdScreen, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveDown Unit:=wdScreen, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


Sub doc_of_scroll_down_command()
    doc_string = "Scroll text of selected window down ARG lines; or near full screen if no ARG."
End Sub
Sub scroll_down_command()
'M-v
    If mark_set Then
        Selection.MoveUp Unit:=wdScreen, Count:=num_arg, Extend:=wdExtend
    Else
        Selection.MoveUp Unit:=wdScreen, Count:=num_arg, Extend:=wdMove
    End If

    complete
End Sub


' after LTrim
' "NUM str"
' "str"
' "NUM "
' "NUM"
' ""
Sub doc_of_universal_argument()
    doc_string = "Begin a numeric argument for the following command."
End Sub
Sub universal_argument()
'C-u
    frmUniArg.Show
    
    If frmUniArg.get_quit_reason() = dqrCancel Then
        Unload frmUniArg
        Exit Sub
    End If
    
    ' deal with self insert command
    'Dim num_arg_value As Long
    Dim self_ins_string As String
    
    frmUniArg.get_two_parts num_arg, self_ins_string
    
    'num_arg = num_arg_value
    
    If self_ins_string <> "" Then
        While num_arg > 0
            Selection.TypeText Text:=self_ins_string

            num_arg = num_arg - 1
        Wend
        num_arg = 1
    End If
    
    Unload frmUniArg
End Sub


Sub doc_of_execute_extended_command()
    doc_string = "Read function name, then read its arguments and call it."
End Sub
Sub execute_extended_command()
'M-x
    Dim cmd As String
    Dim arg As String
    cmd = InputBox("enter command name", "M-x")
    
    cmd = Replace(cmd, "-", "_")

    ' call read_args_for_XXX to prepare arg

    'Application.Run(cmd, arg)
    On Error Resume Next
    Application.Run cmd
End Sub


Sub doc_of_transpose_chars()
    doc_string = "Interchange characters around point, moving forward one character."
End Sub
Sub transpose_chars()
'C-t

    begin_action "transpose chars"

    Dim c
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    c = Selection
    Selection.Delete
    Selection.Collapse direction:=wdCollapseEnd
    Selection.MoveRight Unit:=wdCharacter, Count:=num_arg
    Selection.InsertAfter c
    Selection.Collapse
    
    end_action

    complete
End Sub


Sub doc_of_transpose_words()
    doc_string = "Interchange words around point, leaving point at end of them."
End Sub
Sub transpose_words()
'M-t

    begin_action "transpose words"

    Dim w
    Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    w = Selection
    Selection.Delete
    Selection.Collapse direction:=wdCollapseEnd
    Selection.MoveRight Unit:=wdWord, Count:=num_arg
    Selection.InsertAfter w
    Selection.Collapse
    
    end_action
    
    complete
End Sub


Sub doc_of_exchange_point_and_mark()
    doc_string = "Put the mark where point is now, and point where the mark is now."
End Sub
Sub exchange_point_and_mark()
'C-x C-x
    
    ' if region is not highlight, highlight it first.
    If Selection.type = wdSelectionIP Then
        Dim point_pos As Long
        point_pos = Selection.Start
        
        If mark_pos <= point_pos Then
            Selection.Start = mark_pos
            Selection.StartIsActive = False
        Else
            Selection.End = mark_pos
            Selection.StartIsActive = True
            
        End If
        
        mark_set = True
    End If
    
    
    Selection.StartIsActive = Not Selection.StartIsActive
    
    If Selection.StartIsActive Then
        mark_pos = Selection.End
    Else
        mark_pos = Selection.Start
    End If

    complete
End Sub


Sub doc_of_downcase_region()
    doc_string = "Convert the region to lower case."
End Sub
Sub downcase_region()
'C-x C-l
    Selection.Range.Case = wdLowerCase
    
    select_nothing
    mark_set = False

    complete
End Sub


Sub doc_of_upcase_region()
    doc_string = "Convert the region to upper case."
End Sub
Sub upcase_region()
'C-x C-u
    Selection.Range.Case = wdUpperCase

    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_mark_whole_buffer()
    doc_string = "Put point at beginning and mark at end of buffer."
End Sub
Sub mark_whole_buffer()
'C-x h
    ' Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    mark_set = True
    mark_pos = Selection.Start
    
    Selection.HomeKey Unit:=wdStory, Extend:=wdExtend

    complete
End Sub


Sub doc_of_save_buffer()
    doc_string = "Save current buffer in visited file if modified."
End Sub
Sub save_buffer()
'C-x C-s

    On Error Resume Next
    ActiveDocument.Save
    'Application.Run macroname:="FileSave"

    complete
End Sub


Sub doc_of_recenter()
    doc_string = "Move current buffer line to the specified window line."
End Sub
Sub recenter()
'C-l

    Application.ScreenUpdating = False
    
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    Application.ScreenUpdating = True
    
    complete
End Sub


Sub doc_of_undo()
    doc_string = "Undo some previous changes."
End Sub
Sub undo()
'C-z
    On Error Resume Next
    Application.Run "EditUndo"
    select_nothing
    complete
End Sub


Function other_pane_index(this_pane_index As Integer)
    Debug.Assert this_pane_index = 1 Or this_pane_index = 2
    If this_pane_index = 1 Then
        other_pane_index = 2
    Else
        other_pane_index = 1
    End If
    
End Function


Sub doc_of_split_window_vertically()
    doc_string = "Split the document window."
End Sub
Sub split_window_vertically()
'C-x 2

    Application.ScreenUpdating = False
    
    ' macro DocSplit is interactive
    'Application.Run MacroName:="DocSplit"
    
    ActiveDocument.ActiveWindow.Split = True
    
    ActiveWindow.Panes(1).Activate
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    
    Debug.Print Selection.Start, Selection.End
    Dim sel_start As Long
    Dim sel_end As Long
    sel_start = Selection.Start
    sel_end = Selection.End
    

    ActiveWindow.Panes(other_pane_index(ActiveWindow.ActivePane.index)).Activate
    Selection.Start = sel_start
    Selection.End = sel_end
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    
    ActiveWindow.Panes(1).Activate
    'ActiveWindow.ScrollIntoView Selection.Range, True
    
    Application.ScreenUpdating = True
    

    complete
End Sub


Sub doc_of_other_window()
    doc_string = "Activate another pane."
End Sub
Sub other_window()
'C-x o
    If ActiveWindow.Panes.Count > 1 Then
        ActiveWindow.Panes(other_pane_index(ActiveWindow.ActivePane.index)).Activate
    End If
    
    complete
End Sub


Sub doc_of_delete_window()
    doc_string = "Close current pane."
End Sub
Sub delete_window()
'C-x 0
    If ActiveWindow.Panes.Count > 1 Then
        ActiveWindow.ActivePane.Close
    End If
    
    complete
End Sub


Sub doc_of_delete_other_windows()
    doc_string = "Remove the document window split."
End Sub
Sub delete_other_windows()
'C-x 1
    'Application.Run MacroName:="ClosePane"
    
'   if we use Split = Flase, we have no idea about which pane will remains.
'    ActiveDocument.ActiveWindow.Split = False

    While ActiveWindow.Panes.Count > 1
        ActiveWindow.ActivePane.Next.Activate
        ActiveWindow.ActivePane.Close
    Wend
    

    complete
End Sub


Sub doc_of_set_justification_left()
    doc_string = "Align text to the left."
End Sub
Sub set_justification_left()
'M-j l
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    complete
End Sub


Sub doc_of_set_justification_center()
    doc_string = "Center text."
End Sub
Sub set_justification_center()
'M-j c
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    complete
End Sub


Sub doc_of_set_justification_right()
    doc_string = "Align text to the right."
End Sub
Sub set_justification_right()
'M-j r
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    complete
End Sub


Sub doc_of_set_justification_full()
    doc_string = "Justify."
End Sub
Sub set_justification_full()
'M-j b
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    complete
End Sub


Sub doc_of_describe_key()
    doc_string = "Display documentation of the function invoked by KEY."
End Sub
Sub describe_key()
'C-h k
    is_desc_key = True
    statusbar_prefix = "Describe key: "
    StatusBar = statusbar_prefix
    complete
End Sub


Sub doc_of_describe_function()
    doc_string = "Display the documentation of FUNCTION (a symbol)."
End Sub
Sub describe_function()
'C-h f

    Dim cmd As String
    cmd = InputBox("enter function name", "Describe function")
    
    cmd = Replace(cmd, "-", "_")

    Dim doc_string_proc_name As String
    doc_string_proc_name = "doc_of_" & cmd
    
    On Error Resume Next ' there may NOT be a corresponding doc string porc
    Application.Run macroname:=doc_string_proc_name
    
    'display information about this command
    MsgBox doc_string, Title:=cmd
    
    reset_doc_string

    complete
End Sub


Sub doc_of_isearch_forward()
    doc_string = "Do incremental search forward."
End Sub
Sub isearch_forward()
'C-s
    search_fwd = True
    'isearch_running = True
    doing_search = True
    frmISearch.cbWildcards = False
    doing_search = False
    
    frmISearch.Show vbModeless
End Sub


Sub doc_of_isearch_backward()
    doc_string = "Do incremental search backward."
End Sub
Sub isearch_backward()
'C-r
    search_fwd = False
    frmISearch.cbWildcards = False
    frmISearch.Show vbModeless
End Sub


Sub doc_of_isearch_forward_regexp()
    doc_string = "Do incremental search forward for regular expression."
End Sub
Sub isearch_forward_regexp()
'C-M-s
    search_fwd = True
    frmISearch.cbWildcards = True
    frmISearch.Show vbModeless
End Sub


Sub doc_of_isearch_backward_regexp()
    doc_string = "Do incremental search backward for regular expression."
End Sub
Sub isearch_backward_regexp()
'C-M-r
    search_fwd = False
    frmISearch.cbWildcards = True
    frmISearch.Show vbModeless
End Sub


Sub doc_of_delete_horizontal_space()
    doc_string = "Delete all spaces and tabs around point."
End Sub
Sub delete_horizontal_space()
'M-\
    Application.ScreenUpdating = False
    
    Selection.MoveStartWhile cset:=" ", Count:=wdBackward
    Selection.MoveEndWhile cset:=" ", Count:=wdForward
    If Selection.type = wdSelectionNormal Then
        Selection.Delete
    End If
    
    Application.ScreenUpdating = True
End Sub


Sub doc_of_just_one_space()
    doc_string = "Delete all spaces and tabs around point, leaving one space (or N spaces)."
End Sub
Sub just_one_space()
'M-SPC
    Application.ScreenUpdating = False
    
    Selection.MoveStartWhile cset:=" ", Count:=wdBackward
    Selection.MoveEndWhile cset:=" ", Count:=wdForward
    If Selection.type = wdSelectionNormal Then
        Selection.Delete
    End If
    
    Selection.TypeText " "
    
    Application.ScreenUpdating = True
End Sub


Sub complete()
    num_arg = 1
End Sub


Sub prompt_undefined()
    StatusBar = "key undefined"
    
    complete
End Sub
