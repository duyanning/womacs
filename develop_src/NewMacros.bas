Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro1"
'
' Macro1 Macro
'
'
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    With Selection.Font
        .NameFarEast = "ËÎÌו"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 14
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = True
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With Selection.Font
        .NameFarEast = "ËÎÌו"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 14
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .color = wdColorAutomatic
        .Engrave = False
        .Superscript = True
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro3"
'
' Macro3 Macro
'
'
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Macro4"
'
' Macro4 Macro
'
'
    CommandBars("Navigation").Visible = False
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ChrW(8226)
        .Replacement.Text = ""
        .forward = True
        .wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute
End Sub
