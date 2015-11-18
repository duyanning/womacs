VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmISearch 
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6195
   OleObjectBlob   =   "frmISearch.frx":0000
End
Attribute VB_Name = "frmISearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' womacs
Option Explicit


'执行isearch之前的光标位置，即查找起点
Dim pos_before_isearch As Long

'查找是否成功
Dim match_found As Boolean
'只有match_found为真，以下的match_start跟match_end才有意义。
Dim match_start As Long
Dim match_end As Long

'最近一次成功匹配处的头和尾
Dim prev_ok_match_start As Long
Dim prev_ok_match_end As Long

'最近一次成功匹配处的目标串
Dim prev_ok_match_text As String

'上次查找失败后，在同一方向上继续查找时将wrapped设为true。
Dim wrapped As Boolean

'True表示查找是被C-g中止的，False表示是enter结束的
Dim search_canceled As Boolean


'更新UI提示
Sub update_UI_prompt()

    Dim found_str As String
    found_str = IIf(match_found, "", "Failing ")
    
    If match_found Then
        tbText.ForeColor = &H80000008
    Else
        tbText.ForeColor = RGB(255, 0, 0)
    End If

    Dim wrapped_str As String
    wrapped_str = IIf(wrapped, "Overwrapped ", "")
    
    Dim fwd_str As String
    fwd_str = IIf(search_fwd, ":", " backward:")
    
    'prompt = found_str & wrapped_str & "I-search" & fwd_str
    Me.Caption = found_str & wrapped_str & "I-search" & fwd_str
End Sub


Sub highlight_all_matches()
    '在全文中高亮匹配处，而不是在选定区域中高亮匹配处
    Selection.Collapse
    
    'HitHighlight不受Selection.Find.Wrap的影响，遇到文档末尾会提示你：
    'Word has reached the end of the document.
    'Do you want to continue searching at the beginnig?
    '为了避免出现提示，在调用之前把光标放在文档起始处
    
    'HomeKey不仅改变了插入点位置，还使得屏幕移动，让我们能看见插入点
    'Selection.HomeKey Unit:=wdStory
    Selection.SetRange 0, 0
    
    Selection.Find.HitHighlight FindText:=tbText.Text, _
                                MatchCase:=cbMatchCase.Value, _
                                MatchWholeWord:=cbWholeWord.Value, _
                                MatchWildcards:=cbWildcards.Value
End Sub


Sub clear_all_matches_hightlight()
    Selection.Collapse
'    'Selection.HomeKey Unit:=wdStory
'    'HomeKey不仅改变了插入点位置，还使得屏幕移动，让我们能看见插入点
'    Selection.SetRange 0, 0
    Selection.Find.ClearHitHighlight
End Sub


Sub search_match(fwd As Boolean, wrap As WdFindWrap)
    Debug.Assert Selection.type = wdSelectionIP
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = tbText.Text
        .Replacement.Text = ""
        .forward = fwd
        .wrap = wrap
        .Format = False
        .MatchCase = cbMatchCase.Value
        .MatchWholeWord = cbWholeWord
        .MatchWildcards = cbWildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute

    
    match_found = Selection.Find.found
    If match_found Then
        match_start = Selection.Start
        match_end = Selection.End
    Else
        match_start = 0
        match_end = 0
    End If
End Sub


Sub update_match_result(new_cursor_pos As Long)
    '如果找到了
    If match_found Then
        '将本次匹配状态为最近一次成功匹配状态
        prev_ok_match_text = tbText.Text
        prev_ok_match_start = match_start
        prev_ok_match_end = match_end
    '否则（即没找到）
    Else
        '重新设定光标位置，即下次查找的起点
        Selection.SetRange new_cursor_pos, new_cursor_pos
        'ActiveWindow.ScrollIntoView Selection.Range
    End If

    '更新UI提示
    update_UI_prompt
End Sub


Sub when_change()
    '在对话框显现之前，我们可能会更变对话框上的一些复选框，从而导致本过程被调用。
    If Not Me.Visible Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
'    Dim start_before_hightlight As Long
'    Dim end_before_hightlight As Long
'
'    start_before_hightlight = Selection.Start
'    end_before_hightlight = Selection.End

    '先清除上次的高亮
    clear_all_matches_hightlight
    highlight_all_matches
    
'    Selection.SetRange start_before_hightlight, end_before_hightlight
'    ActiveWindow.ScrollIntoView Selection.Range
    
    Application.ScreenUpdating = True
    
    '对话框不关闭之前不会自动刷新屏幕，所以我们要求刷新
    Application.ScreenRefresh
    
    Selection.Collapse
    
    If search_fwd Then
        '如果最近一次成功匹配的目标串是新目标串的真子串
        '或者，新目标串是最近一次成功匹配的目标串的真子串
        If is_prefix_of(tbText.Text, prev_ok_match_text) _
                Or is_prefix_of(prev_ok_match_text, tbText.Text) Then
            '从上次成功匹配的起点开始向下查找
            Selection.SetRange prev_ok_match_start, prev_ok_match_start
            search_match True, wdFindStop
            
            update_match_result prev_ok_match_start
            Application.ScreenRefresh
            '结束
            Exit Sub
        End If
        
        '能到这里，表示新目标串跟最近一次成功匹配的目标串之间无任何关系
        '从查找起点开始向下查找
        Selection.SetRange pos_before_isearch, pos_before_isearch
        search_match True, wdFindStop
        
        update_match_result pos_before_isearch
        Application.ScreenRefresh
        '结束
        Exit Sub
        
    Else
        '如果最近一次成功匹配的目标串是新目标串的真子串
        '或者，新目标串是最近一次成功匹配的目标串的真子串
        If is_prefix_of(tbText.Text, prev_ok_match_text) _
                Or is_prefix_of(prev_ok_match_text, tbText.Text) Then
            '从上次匹配处的起点开始查找（向上？向下？）
            Selection.SetRange prev_ok_match_start, prev_ok_match_start
            
            '先向下找一下
            search_match True, wdFindStop
    
            '如果找到了，匹配起点必须同最近一次成功匹配的起点，
            '且终点不越过查找起点（如果先C-s几下，再C-r几下，查找起点会位于匹配处之前）
            If match_found _
                    And match_start = prev_ok_match_start _
                    And (pos_before_isearch <= match_start _
                            Or match_end <= pos_before_isearch) Then
                update_match_result 0
                Exit Sub
            Else
                '不管是因为上面的哪个子式为false导致进入else分支，我们都认为是失败
                match_found = False
                update_match_result prev_ok_match_start
                Debug.Assert Selection.type = wdSelectionIP
            End If
            
            '到这里说明向下没找到，再向上找
            search_match False, wdFindStop
            update_match_result prev_ok_match_start
            Exit Sub
        
        End If

        '能到这里，表示新目标串跟最近一次成功匹配的目标串之间无任何关系
        '从查找起点开始向上查找
        Selection.SetRange pos_before_isearch, pos_before_isearch
        search_match False, wdFindStop
        
        update_match_result pos_before_isearch
        
        '结束
        Exit Sub

    End If

End Sub


'C-s或C-r按下后调用
Sub isearch_next(fwd As Boolean)

    search_fwd = fwd
    
    If tbText.Text = "" Then
'        '拾取光标下文字作为目标字符串
'        Selection.Expand
'        tbText.Text = Selection.Text
        Exit Sub
    End If
    
    
    If Not match_found Then
        'search_fwd为上次查找方向，在该方向上失败了，你还要继续在该方向上查，就是折回
        If fwd = search_fwd Then
            wrapped = True
        End If
        
        If fwd Then
            Selection.SetRange prev_ok_match_end, prev_ok_match_end
        Else
            Selection.SetRange prev_ok_match_start, prev_ok_match_start
        End If
        search_match fwd, wdFindContinue
        update_match_result prev_ok_match_start
        
        Exit Sub
    End If

    '在既定方向上先查一下
    If fwd Then
        Selection.SetRange prev_ok_match_end, prev_ok_match_end
    Else
        Selection.SetRange prev_ok_match_start, prev_ok_match_start
    End If
    
    search_match fwd, wdFindStop
    update_match_result prev_ok_match_start
    

End Sub


Private Sub cbMatchCase_Change()
    doing_search = True
    when_change
    doing_search = False
End Sub


Private Sub cbWholeWord_Change()
    doing_search = True
    when_change
    doing_search = False
End Sub

Private Sub cbWholeWord_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    doing_search = True
    when_key_down KeyCode, Shift
    doing_search = False
End Sub


Private Sub cbWildcards_Change()
    doing_search = True
    cbMatchCase.Enabled = Not cbWildcards.Value
    cbWholeWord.Enabled = Not cbWildcards.Value

    when_change
    doing_search = False
End Sub


Private Sub cbWildcards_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    doing_search = True
    when_key_down KeyCode, Shift
    doing_search = False
End Sub


Private Sub tbText_Change()
    doing_search = True
    when_change
    doing_search = False
End Sub


Sub when_key_down(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        search_canceled = False
        Unload Me
        'Me.hide
        Exit Sub
    End If
    
    ' 2 is fmCtrlMask, I don't know why fmCtrlMask is Empty
    If Not Shift = 2 Then
        Exit Sub
    End If
    
    Select Case KeyCode
    Case vbKeyG
        search_canceled = True
        Unload Me
        'Me.hide
    Case vbKeyS
        isearch_next True
    Case vbKeyR
        isearch_next False
    Case Else
        'do nothing
    End Select
    
End Sub


Private Sub tbText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    doing_search = True
    when_key_down KeyCode, Shift
    doing_search = False
End Sub


Private Sub cbMatchCase_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    doing_search = True
    when_key_down KeyCode, Shift
    doing_search = False
End Sub


Sub prelude()
    
    Me.StartUpPosition = 0 ' 0-Manual
    Me.Left = GetSetting("Womacs", "Settings", "Left", "380")
    Me.top = GetSetting("Womacs", "Settings", "Top", "15")
    
    cbMatchCase.Value = GetSetting("Womacs", "Settings", "MatchCase", "0")
    cbWholeWord.Value = GetSetting("Womacs", "Settings", "WholeWord", "0")
    cbWildcards.Value = GetSetting("Womacs", "Settings", "Wildcards", "0")

    'Selection.Collapse
    pos_before_isearch = Selection.Start
    match_start = pos_before_isearch
    match_end = pos_before_isearch
    match_found = True
    
    prev_ok_match_text = ""
    prev_ok_match_start = pos_before_isearch
    prev_ok_match_end = pos_before_isearch
    
    search_canceled = False
    wrapped = False
    
    update_UI_prompt
End Sub


Sub finale()

    SaveSetting "Womacs", "Settings", "Left", Me.Left
    SaveSetting "Womacs", "Settings", "Top", Me.top

    SaveSetting "Womacs", "Settings", "MatchCase", cbMatchCase.Value
    SaveSetting "Womacs", "Settings", "WholeWord", cbWholeWord.Value
    SaveSetting "Womacs", "Settings", "Wildcards", cbWildcards.Value

    clear_all_matches_hightlight
    Application.ScreenRefresh

    '空，取消=>回之前位置
    '失败，取消=>回之前位置
    '成功，取消=>回之前位置
    '空，确认=>回之前位置，非增量查找
    '失败，确认=>回之前位置
    '成功，确认=>什么也不做
    If search_canceled Then
        Selection.SetRange pos_before_isearch, pos_before_isearch
        ActiveWindow.ScrollIntoView Selection.Range
        Exit Sub
    End If
    
    Debug.Assert Not search_canceled
    
    If Not match_found Then
        Selection.SetRange pos_before_isearch, pos_before_isearch
        ActiveWindow.ScrollIntoView Selection.Range
        Exit Sub
    End If
    
    If tbText.Text = "" Then
        Selection.SetRange pos_before_isearch, pos_before_isearch
        ActiveWindow.ScrollIntoView Selection.Range
        
        'Me.Hide
        'Application.Run macroname:="EditFind"
        'Dialogs(wdDialogEditFind).Show
        Selection.Find.forward = search_fwd
        '要得到一个一个modeless的内建对话框只能这样做
        '来源：http://help.wugnet.com/office/Find-Replace-dialog-ftopict527791.html
        CommandBars.FindControl(id:=141).Execute

        Exit Sub
    End If
    
    Selection.SetRange IIf(search_fwd, match_end, match_start), _
                        IIf(search_fwd, match_end, match_start)

    'ActiveWindow.ScrollIntoView Selection.Range
    
    '如果设置了mark，还要选中一片
    If Not mark_set Then
        Exit Sub
    End If
    
    If Selection.Start >= mark_pos Then
        Selection.Start = mark_pos
        Selection.StartIsActive = False
    Else
        Selection.End = mark_pos
        Selection.StartIsActive = True
    End If
    
End Sub


'Private Sub UserForm_Deactivate()
'    Debug.Print "hih"
'End Sub

Private Sub UserForm_Initialize()
    doing_search = True
    isearch_running = True
    prelude
    doing_search = False
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    doing_search = True
    finale
    isearch_running = False
    doing_search = False
End Sub

