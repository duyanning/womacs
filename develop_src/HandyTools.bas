Attribute VB_Name = "HandyTools"
' handy tools
Option Explicit

' when you copy text from a pdf to word, there is a return on the end of every line
' so we need to get rid of them.
Sub handy_tools_paste_and_remove_retrun()
    Dim oldLoc As Long
    oldLoc = Selection.Start
    Selection.PasteSpecial DataType:=wdPasteText
    
    Dim t As Range
    Set t = Selection.Range
    t.SetRange Start:=oldLoc, End:=Selection.End
    t.Select
    
    ' replace 'end of paragraph' with single space
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p"
        If MsgBox("replace return with space?", vbYesNo) = vbYes Then
            .Replacement.Text = " "
        Else
            .Replacement.Text = ""
        End If
        .forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace continuous two spaces with single space
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub


Private Sub replace_punct(s As String, t As String)
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = s
        .Replacement.Text = t
        .forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Private Sub replace_punct_regex(s As String, t As String)
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = s
        .Replacement.Text = t
        .forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub handy_tools_punct_ch_to_en()
    If Selection.type <> wdSelectionNormal Then
        Exit Sub
    End If

    begin_action "convert punct from ch to en"
    
    replace_punct "，", ","
    replace_punct "。", "."
    replace_punct "（", "("
    replace_punct "）", ")"
    replace_punct "；", ";"
    replace_punct "：", ":"
    replace_punct "！", "!"
    replace_punct "？", "?"
    replace_punct ChrW(8226), "·"
    
'    Selection.LanguageID = wdEnglishUS
'    Selection.NoProofing = False
'    Application.CheckLanguage = True
    end_action
End Sub

Sub handy_tools_punct_en_to_ch()
    If Selection.type <> wdSelectionNormal Then
        Exit Sub
    End If

    begin_action "convert punct from en to ch"
    
    replace_punct ",", "，"
    'replace_punct ".", "。"
    '前面是汉字的英文句点
    replace_punct_regex "([一-" & ChrW(40891) & "])\.", "\1。"
    '后面是汉字的英文句点
    replace_punct_regex "\.([一-" & ChrW(40891) & "])", "。\1"
    '后边是空格的英文句点
    replace_punct_regex "\. ", "。"
    '后边是换行的英文句点(用regex的时候无法用^p之流)
    replace_punct ".^p", "。^p"
    
    'replace_punct "(", "（"
    '后边紧跟汉字的英文(
    replace_punct_regex "\(([一-" & ChrW(40891) & "])", "（\1"
    
    'replace_punct ")", "）"
    '紧跟汉字的英文)
    replace_punct_regex "([一-" & ChrW(40891) & "])\)", "\1）"
    
    replace_punct ";", "；"
    replace_punct ":", "："
    replace_punct "!", "！"
    replace_punct "?", "？"

'    Selection.LanguageID = wdSimplifiedChinese
'    Selection.NoProofing = False
'    Application.CheckLanguage = True
    end_action
End Sub

Sub handy_tools_add_space_between_punct_and_char()

    Dim found As Boolean
    Dim t As Range
    
    If Selection.type <> wdSelectionNormal Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    
    Application.UndoRecord.StartCustomRecord "add_space_between_punct_and_char"

    ' there is a bug with replaceAll and regex.
    ' it causes selection change to IP after the first match and replace.
    
    Set t = Selection.Range ' t will extend automatically
    
    Do
        With Selection.Find
'            .ClearFormatting
'            .Replacement.ClearFormatting
            .Text = "([,.;:])([a-zA-Z])"
            .Replacement.Text = "\1 \2"   ' replacing 2 chars with 3 chars
            .forward = True
            .wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            '.Execute Replace:=wdReplaceAll
            found = .Execute(Replace:=wdReplaceOne)
        End With
        Selection.End = t.End
    Loop While found
    
    Selection.Start = t.Start
    Selection.End = t.End
    
    Application.UndoRecord.EndCustomRecord
    
    Application.ScreenUpdating = True

End Sub

Sub handy_tools_add_line_no()
    Dim nLineNum As Integer
    Dim sLineNum As String
    Dim selRge As Range
    Dim i As Integer
    
    Set selRge = Selection.Range
    For nLineNum = 1 To selRge.Paragraphs.Count
        sLineNum = str(nLineNum)
        sLineNum = Trim(sLineNum)
        
        Dim length As Integer
        length = Len(sLineNum)
        'MsgBox length
        
        For i = 1 To (3 - length)
            sLineNum = "0" + sLineNum
        Next i
        
        sLineNum = "#" & sLineNum
        sLineNum = sLineNum + " "
        
        selRge.Paragraphs(nLineNum).Range.InsertBefore (sLineNum)
    Next nLineNum
        
End Sub


' http://support.microsoft.com/kb/290978/
' it has bug
' workaround: if select text in table cell, insert a blank before end, then select text before the blank
Sub handy_tools_convert_symbol()

   Dim dlg As Object
   Dim NoFC As Integer
   Dim SCP As Integer
   Dim StartRange As Range
   Dim UniCodeNum As Integer

   ' Temporarily disable Screen Updating
   Application.ScreenUpdating = False

   ' Temporarily disable Smart Cut & Paste
   If Options.SmartCutPaste = True Then
      SCP = 1
      Options.SmartCutPaste = False
   End If

   ' Temporarily display field text
   If ActiveWindow.View.ShowFieldCodes = False Then
      NoFC = 1
      ActiveWindow.View.ShowFieldCodes = True
   End If

   ' Set StartRange variable to current selection's range
   Set StartRange = Selection.Range
   Selection.Collapse

   ' Select first, then each next character in user-defined selection
   Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
   While Selection.End <= StartRange.End And _
      ActiveDocument.Content.End > Selection.End

      ' If the character is a space, then move to next character
      Set dlg = Dialogs(wdDialogInsertSymbol)
      UniCodeNum = dlg.charnum

      If UniCodeNum = 32 Then
         Selection.Collapse
         Selection.MoveRight Unit:=wdCharacter, Extend:=wdMove
         Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
      End If

      ' Loop, converting symbol Unicode characters to ASCII characters
      Set dlg = Dialogs(wdDialogInsertSymbol)
      UniCodeNum = dlg.charnum

      While UniCodeNum < 0 And Selection.End <= StartRange.End _
         And ActiveDocument.Content.End > Selection.End
            Selection.Delete
            Selection.InsertAfter (ChrW(UniCodeNum + 4096))
            Selection.Collapse (wdCollapseEnd)
            Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
            Set dlg = Dialogs(wdDialogInsertSymbol)
            UniCodeNum = dlg.charnum
      Wend

      Selection.Collapse (wdCollapseEnd)
      Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
   Wend

   ' Reset Word document settings
   If SCP = 1 Then Options.SmartCutPaste = True
   If NoFC = 1 Then ActiveWindow.View.ShowFieldCodes = False
      Selection.Collapse (wdCollapseStart)
      Selection.MoveLeft Unit:=wdCharacter
      Application.ScreenUpdating = True

End Sub



' In word 2010, if you insert a visio drawing object into a table,
' and then right click on the object, in the context menu
' you can NOT find the option "Format Object"
Sub handy_tools_format_drawing_object()
    On Error Resume Next
    Application.Run macroname:="FormatDrawingObject"
End Sub


' 参考http://software-solutions-online.com/word-vba-resize-pictures/
Sub change_pictures_scale()
    Dim iShape
    Dim preferredScale As Integer
    preferredScale = InputBox("input x for x%", "scale")
    
    
    For Each iShape In ActiveDocument.InlineShapes
        iShape.LockAspectRatio = msoTrue
        'iShape.ScaleHeight = 50
        iShape.ScaleHeight = preferredScale
        
    Next iShape

End Sub

