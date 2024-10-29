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


'ִ��isearch֮ǰ�Ĺ��λ�ã����������
Dim pos_before_isearch As Long

'�����Ƿ�ɹ�
Dim match_found As Boolean
'ֻ��match_foundΪ�棬���µ�match_start��match_end�������塣
Dim match_start As Long
Dim match_end As Long

'���һ�γɹ�ƥ�䴦��ͷ��β
Dim prev_ok_match_start As Long
Dim prev_ok_match_end As Long

'���һ�γɹ�ƥ�䴦��Ŀ�괮
Dim prev_ok_match_text As String

'�ϴβ���ʧ�ܺ���ͬһ�����ϼ�������ʱ��wrapped��Ϊtrue��
Dim wrapped As Boolean

'True��ʾ�����Ǳ�C-g��ֹ�ģ�False��ʾ��enter������
Dim search_canceled As Boolean


'����UI��ʾ
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
    '��ȫ���и���ƥ�䴦����������ѡ�������и���ƥ�䴦
    Selection.Collapse
    
    'HitHighlight����Selection.Find.Wrap��Ӱ�죬�����ĵ�ĩβ����ʾ�㣺
    'Word has reached the end of the document.
    'Do you want to continue searching at the beginnig?
    'Ϊ�˱��������ʾ���ڵ���֮ǰ�ѹ������ĵ���ʼ��
    
    'HomeKey�����ı��˲����λ�ã���ʹ����Ļ�ƶ����������ܿ��������
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
'    'HomeKey�����ı��˲����λ�ã���ʹ����Ļ�ƶ����������ܿ��������
'    Selection.SetRange 0, 0
    Selection.Find.ClearHitHighlight
End Sub


Sub search_match(fwd As Boolean, wrap As WdFindWrap)
    Debug.Assert Selection.Type = wdSelectionIP
    
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
    '����ҵ���
    If match_found Then
        '������ƥ��״̬Ϊ���һ�γɹ�ƥ��״̬
        prev_ok_match_text = tbText.Text
        prev_ok_match_start = match_start
        prev_ok_match_end = match_end
    '���򣨼�û�ҵ���
    Else
        '�����趨���λ�ã����´β��ҵ����
        Selection.SetRange new_cursor_pos, new_cursor_pos
        'ActiveWindow.ScrollIntoView Selection.Range
    End If

    '����UI��ʾ
    update_UI_prompt
End Sub


Sub when_change()
    '�ڶԻ�������֮ǰ�����ǿ��ܻ����Ի����ϵ�һЩ��ѡ�򣬴Ӷ����±����̱����á�
    If Not Me.Visible Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
'    Dim start_before_hightlight As Long
'    Dim end_before_hightlight As Long
'
'    start_before_hightlight = Selection.Start
'    end_before_hightlight = Selection.End

    '������ϴεĸ���
    clear_all_matches_hightlight
    highlight_all_matches
    
'    Selection.SetRange start_before_hightlight, end_before_hightlight
'    ActiveWindow.ScrollIntoView Selection.Range
    
    Application.ScreenUpdating = True
    
    '�Ի��򲻹ر�֮ǰ�����Զ�ˢ����Ļ����������Ҫ��ˢ��
    Application.ScreenRefresh
    
    Selection.Collapse
    
    If search_fwd Then
        '������һ�γɹ�ƥ���Ŀ�괮����Ŀ�괮�����Ӵ�
        '���ߣ���Ŀ�괮�����һ�γɹ�ƥ���Ŀ�괮�����Ӵ�
        If is_prefix_of(tbText.Text, prev_ok_match_text) _
                Or is_prefix_of(prev_ok_match_text, tbText.Text) Then
            '���ϴγɹ�ƥ�����㿪ʼ���²���
            Selection.SetRange prev_ok_match_start, prev_ok_match_start
            search_match True, wdFindStop
            
            update_match_result prev_ok_match_start
            Application.ScreenRefresh
            '����
            Exit Sub
        End If
        
        '�ܵ������ʾ��Ŀ�괮�����һ�γɹ�ƥ���Ŀ�괮֮�����κι�ϵ
        '�Ӳ�����㿪ʼ���²���
        Selection.SetRange pos_before_isearch, pos_before_isearch
        search_match True, wdFindStop
        
        update_match_result pos_before_isearch
        Application.ScreenRefresh
        '����
        Exit Sub
        
    Else
        '������һ�γɹ�ƥ���Ŀ�괮����Ŀ�괮�����Ӵ�
        '���ߣ���Ŀ�괮�����һ�γɹ�ƥ���Ŀ�괮�����Ӵ�
        If is_prefix_of(tbText.Text, prev_ok_match_text) _
                Or is_prefix_of(prev_ok_match_text, tbText.Text) Then
            '���ϴ�ƥ�䴦����㿪ʼ���ң����ϣ����£���
            Selection.SetRange prev_ok_match_start, prev_ok_match_start
            
            '��������һ��
            search_match True, wdFindStop
    
            '����ҵ��ˣ�ƥ��������ͬ���һ�γɹ�ƥ�����㣬
            '���յ㲻Խ��������㣨�����C-s���£���C-r���£���������λ��ƥ�䴦֮ǰ��
            If match_found _
                    And match_start = prev_ok_match_start _
                    And (pos_before_isearch <= match_start _
                            Or match_end <= pos_before_isearch) Then
                update_match_result 0
                Exit Sub
            Else
                '��������Ϊ������ĸ���ʽΪfalse���½���else��֧�����Ƕ���Ϊ��ʧ��
                match_found = False
                update_match_result prev_ok_match_start
                Debug.Assert Selection.Type = wdSelectionIP
            End If
            
            '������˵������û�ҵ�����������
            search_match False, wdFindStop
            update_match_result prev_ok_match_start
            Exit Sub
        
        End If

        '�ܵ������ʾ��Ŀ�괮�����һ�γɹ�ƥ���Ŀ�괮֮�����κι�ϵ
        '�Ӳ�����㿪ʼ���ϲ���
        Selection.SetRange pos_before_isearch, pos_before_isearch
        search_match False, wdFindStop
        
        update_match_result pos_before_isearch
        
        '����
        Exit Sub

    End If

End Sub


'C-s��C-r���º����
Sub isearch_next(fwd As Boolean)

    search_fwd = fwd
    
    If tbText.Text = "" Then
'        'ʰȡ�����������ΪĿ���ַ���
'        Selection.Expand
'        tbText.Text = Selection.Text
        Exit Sub
    End If
    
    
    If Not match_found Then
        'search_fwdΪ�ϴβ��ҷ����ڸ÷�����ʧ���ˣ��㻹Ҫ�����ڸ÷����ϲ飬�����ۻ�
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

    '�ڼȶ��������Ȳ�һ��
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

    '�գ�ȡ��=>��֮ǰλ��
    'ʧ�ܣ�ȡ��=>��֮ǰλ��
    '�ɹ���ȡ��=>��֮ǰλ��
    '�գ�ȷ��=>��֮ǰλ�ã�����������
    'ʧ�ܣ�ȷ��=>��֮ǰλ��
    '�ɹ���ȷ��=>ʲôҲ����
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
        'Ҫ�õ�һ��һ��modeless���ڽ��Ի���ֻ��������
        '��Դ��http://help.wugnet.com/office/Find-Replace-dialog-ftopict527791.html
        CommandBars.FindControl(id:=141).Execute

        Exit Sub
    End If
    
    Selection.SetRange IIf(search_fwd, match_end, match_start), _
                        IIf(search_fwd, match_end, match_start)

    'ActiveWindow.ScrollIntoView Selection.Range
    
    '���������mark����Ҫѡ��һƬ
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

