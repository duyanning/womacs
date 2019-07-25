Attribute VB_Name = "WordSpecificCommands"
' womacs
Option Explicit


Sub doc_of_word_toggle_bold()
    doc_string = "Toggle bold."
End Sub
Sub word_toggle_bold()
    Selection.Font.Bold = wdToggle
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_word_toggle_italic()
    doc_string = "Toggle italic"
End Sub
Sub word_toggle_italic()
    Selection.Font.Italic = wdToggle
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_word_toggle_strike_through()
    doc_string = "Toggle strike through"
End Sub
Sub word_toggle_strike_through()
    Selection.Font.StrikeThrough = wdToggle
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_word_toggle_subscript()
    doc_string = "Toggle subscript"
End Sub
Sub word_toggle_subscript()
    Selection.Font.Subscript = wdToggle
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_word_toggle_superscript()
    doc_string = "Toggle superscript"
End Sub
Sub word_toggle_superscript()
    Selection.Font.Superscript = wdToggle
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub doc_of_word_toggle_highlight()
    doc_string = "Toggle highlight"
End Sub
Sub word_toggle_highlight()
    If Selection.Range.HighlightColorIndex = wdNoHighlight Then
        Selection.Range.HighlightColorIndex = Options.DefaultHighlightColorIndex
    Else
        Selection.Range.HighlightColorIndex = wdNoHighlight
    End If
    
    select_nothing
    mark_set = False
    
    complete
End Sub


Sub word_font_color(color As WdColorIndex)
    'Selection.Range.Font.ColorIndex = color
    Selection.Font.ColorIndex = color
    
    select_nothing
    mark_set = False
    
    complete

End Sub


Sub doc_of_word_toggle_font_color_red()
    doc_string = "Toggle font color RED"
End Sub
Sub word_toggle_font_color_red()
    word_font_color wdRed
End Sub


Sub doc_of_word_toggle_font_color_green()
    doc_string = "Toggle font color GREEN"
End Sub
Sub word_toggle_font_color_green()
    word_font_color wdGreen
End Sub


Sub doc_of_word_toggle_font_color_blue()
    doc_string = "Toggle font color BLUE"
End Sub
Sub word_toggle_font_color_blue()
    word_font_color wdBlue
End Sub


Sub doc_of_word_toggle_font_color_auto()
    doc_string = "Toggle font color AUTO"
End Sub
Sub word_toggle_font_color_auto()
    word_font_color wdAuto
End Sub


Sub doc_of_word_yank_plain_text()
    doc_string = "Paste as plain text."
End Sub
Sub word_yank_plain_text()
    ' reasons causing error:
    ' text being copied is hidden
    On Error Resume Next
    Selection.PasteAndFormat (wdFormatPlainText)

    complete
End Sub


Sub doc_of_word_show_all()
    doc_string = "Toggle displaying all nonprinting characters."
End Sub
Sub word_show_all()
    ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View.ShowAll
    'ActiveWindow.View.ShowBookmarks = False
'    ActiveWindow.View.FieldShading = wdFieldShadingWhenSelected
'    ActiveWindow.View.ShowTextBoundaries = ActiveWindow.View.ShowTextBoundaries
    complete
End Sub


Sub doc_of_word_style_normal()
    doc_string = "Style Normal"
End Sub
Sub word_style_normal()
    Selection.Style = ActiveDocument.Styles(wdStyleNormal)
    
    complete
End Sub


Sub doc_of_word_style_normal_indent()
    doc_string = "Style Normal Indent"
End Sub
Sub word_style_normal_indent()
    Selection.Style = ActiveDocument.Styles(wdStyleNormalIndent)
    
    complete
End Sub


Sub doc_of_word_style_title()
    doc_string = "Style Title."
End Sub
Sub word_style_title()
    Selection.Style = ActiveDocument.Styles(wdStyleTitle)
    
    complete
End Sub


Sub doc_of_word_style_heading1()
    doc_string = "Style Heading 1"
End Sub
Sub word_style_heading1()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading1)
    
    complete
End Sub


Sub doc_of_word_style_heading2()
    doc_string = "Style Heading 2"
End Sub
Sub word_style_heading2()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading2)
    
    complete
End Sub


Sub doc_of_word_style_heading3()
    doc_string = "Style Heading 3"
End Sub
Sub word_style_heading3()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading3)
    
    complete
End Sub


Sub doc_of_word_style_heading4()
    doc_string = "Style Heading 4"
End Sub
Sub word_style_heading4()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading4)
    
    complete
End Sub


Sub doc_of_word_style_heading5()
    doc_string = "Style Heading 5"
End Sub
Sub word_style_heading5()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading5)
    
    complete
End Sub


Sub doc_of_word_style_heading6()
    doc_string = "Style Heading 6"
End Sub
Sub word_style_heading6()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading6)
    
    complete
End Sub


Sub doc_of_word_style_heading7()
    doc_string = "Style Heading 7"
End Sub
Sub word_style_heading7()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading7)
    
    complete
End Sub


Sub doc_of_word_style_heading8()
    doc_string = "Style Heading 8"
End Sub
Sub word_style_heading8()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading8)
    
    complete
End Sub


Sub doc_of_word_style_heading9()
    doc_string = "Style Heading 9"
End Sub
Sub word_style_heading9()
    Selection.Style = ActiveDocument.Styles(wdStyleHeading9)
    
    complete
End Sub


Sub doc_of_word_single_opening_quote()
    doc_string = "Insert Single Opening Quote."
End Sub
Sub word_single_opening_quote()
    Selection.TypeText Text:=ChrW(8216)
    'Selection.TypeText Text:="‘"
    complete
End Sub


Sub doc_of_word_single_closing_quote()
    doc_string = "Insert Single Closing Quote."
End Sub
Sub word_single_closing_quote()
    Selection.TypeText Text:=ChrW(8217)
    'Selection.TypeText Text:="’"
    complete
End Sub


Sub doc_of_word_double_opening_quote()
    doc_string = "Insert Double Opening Quote."
End Sub
Sub word_double_opening_quote()
    Selection.TypeText Text:=ChrW(8220)
    complete
End Sub


Sub doc_of_word_double_closing_quote()
    doc_string = "Insert Double Closing Quote."
End Sub
Sub word_double_closing_quote()
    Selection.TypeText Text:=ChrW(8221)
    complete
End Sub


Sub doc_of_word_outline_level1()
    doc_string = "Outline Level 1"
End Sub
Sub word_outline_level1()
    ' property Paragraphs.OutlineLevel is NOT available in outline view
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel1
    complete
End Sub


Sub doc_of_word_outline_level2()
    doc_string = "Outline Level 2"
End Sub
Sub word_outline_level2()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel2
    complete
End Sub


Sub doc_of_word_outline_level3()
    doc_string = "Outline Level 3"
End Sub
Sub word_outline_level3()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel3
    complete
End Sub


Sub doc_of_word_outline_level4()
    doc_string = "Outline Level 4"
End Sub
Sub word_outline_level4()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel4
    complete
End Sub


Sub doc_of_word_outline_level5()
    doc_string = "Outline Level 5"
End Sub
Sub word_outline_level5()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel5
    complete
End Sub


Sub doc_of_word_outline_level6()
    doc_string = "Outline Level 6"
End Sub
Sub word_outline_level6()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel6
    complete
End Sub


Sub doc_of_word_outline_level7()
    doc_string = "Outline Level 7"
End Sub
Sub word_outline_level7()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel7
    complete
End Sub


Sub doc_of_word_outline_level8()
    doc_string = "Outline Level 8"
End Sub
Sub word_outline_level8()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel8
    complete
End Sub


Sub doc_of_word_outline_level9()
    doc_string = "Outline Level 9"
End Sub
Sub word_outline_level9()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevel9
    complete
End Sub


Sub doc_of_word_outline_bodytext()
    doc_string = "Outline Body Text"
End Sub
Sub word_outline_bodytext()
    Selection.Paragraphs(1).OutlineLevel = wdOutlineLevelBodyText
    complete
End Sub


Sub doc_of_caption_on_the_spot()
    doc_string = "Caption on the spot"
End Sub
Sub caption_on_the_spot()
    If num_arg > 1 Then
        Selection.Collapse
        
        Selection.HomeKey Unit:=wdLine, Extend:=wdMove
        Selection.InsertAfter "c" ' 临时锚点
        
        
        'Dialogs(wdDialogInsertCaption).Position = wdCaptionPositionAbove
        With Dialogs(wdDialogInsertCaption)
            .Position = wdCaptionPositionAbove ' 放在with语句内部才有效
            '.Display
            .Show
        End With
        'Selection.InsertCaption Label:="操", TitleAutoText:="", Title:="", Position :=wdCaptionPositionAbove, ExcludeLabel:=0
        Selection.Delete Count:=2 '删除题注后的回车与临时锚点
    Else
        Selection.Style = ActiveDocument.Styles(wdStyleCaption)
    End If
    
    complete
End Sub

Sub doc_of_word_subscript()
    doc_string = "Subscript"
End Sub
Sub word_subscript()
    With Selection.Font
        .Superscript = False
        .Subscript = True
    End With
    
    select_nothing
    mark_set = False
    complete
End Sub

Sub doc_of_word_superscript()
    doc_string = "Superscript"
End Sub
Sub word_superscript()
    With Selection.Font
        .Superscript = True
        .Subscript = False
    End With
    
    select_nothing
    mark_set = False
    complete
End Sub

Sub doc_of_word_normalscript()
    doc_string = "Normal script"
End Sub
Sub word_normalscript()
    With Selection.Font
        .Superscript = False
        .Subscript = False
    End With
    
    select_nothing
    mark_set = False
    complete
End Sub


