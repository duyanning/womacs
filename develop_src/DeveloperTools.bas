Attribute VB_Name = "DeveloperTools"
Option Explicit
Option Private Module

' for developers of womacs.dotm's use only

' note:
' path = dirname/basename


Private Sub export_vba_source_code_from_this_project()
    Dim src_dirname As String
    
    src_dirname = ThisDocument.Path & Application.PathSeparator & "develop_src"
    
    If Dir(src_dirname, vbDirectory) = "" Then
        MsgBox src_dirname & " must exist before exporting!"
        Exit Sub
    End If

    Dim c As VBComponent
    
    For Each c In ThisDocument.VBProject.VBComponents
        export_vba_source_code c, src_dirname
    Next c

End Sub

'---------------------------------------------------------------------------------------------------------------
' dangerous !!!!! all changes unsaved will lost.
Private Sub import_vba_source_code_to_this_project()

    Dim answer As Integer
    Dim c As VBComponent
    Dim src_dirname As String
    
    Debug.Print ThisDocument.name
    Debug.Assert (ThisDocument.name = "womacs-dev.docm")
    
    answer = MsgBox("are you sure?", vbOKCancel)
    If answer = vbCancel Then
        Exit Sub
    End If
    
    For Each c In ThisDocument.VBProject.VBComponents
        Debug.Print c.name
        
        If c.name <> "ThisDocument" And c.name <> "DeveloperTools" Then
            ThisDocument.VBProject.VBComponents.Remove c
        End If
        
    Next c
    
    
    src_dirname = ThisDocument.Path & Application.PathSeparator & "develop_src"
    
    If Dir(src_dirname, vbDirectory) = "" Then
        MsgBox src_dirname & " must exist before exporting!"
        Exit Sub
    End If

    import_vba_source_code ThisDocument, src_dirname
    
End Sub
'-------------------------------------------------------------------------------------------------------------------

' note:
' to substitute $WCREV$ with correct number, we MUST commit first, then call this sub
Private Sub generate_template_for_publishing()
    Dim publish_dirname As String
    Dim publish_src_dirname As String
    
    publish_dirname = ThisDocument.Path & Application.PathSeparator & "publish"
    publish_src_dirname = publish_dirname & Application.PathSeparator & "src"

    
    If Dir(publish_dirname, vbDirectory) = "" Then
        MkDir publish_dirname
    End If
    
    If Dir(publish_src_dirname, vbDirectory) = "" Then
        MkDir publish_src_dirname
    End If
    
    Dim Doc As Document
    'Set doc = Documents.Add(NewTemplate:=True)
    Set Doc = Documents.Add
    ', Visible:=False)
    
    
    
    ' export
    Dim c_develop As VBComponent
    
    Dim tag_womacs As String
    tag_womacs = "' womacs"
    
    Dim tag_handy_tools As String
    tag_handy_tools = "' handy tools"
    
    
    For Each c_develop In ThisDocument.VBProject.VBComponents
        Dim first_line As String
        
        first_line = c_develop.CodeModule.Lines(1, 1)
        
        If first_line = tag_womacs Or first_line = tag_handy_tools Then
            export_vba_source_code c_develop, publish_src_dirname
        End If
    Next c_develop
    
    
    ' pause to avoid bug with vba
    MsgBox "Exporting source code"
 
 
    
    ' use SubWCRev.exe from TortoiseSVN to substitute $WCREV$ in Version.bas
    ' shell is NOT asynchronous, see also http://support.microsoft.com/kb/129796
    Dim wc_dirname As String
    wc_dirname = ThisDocument.Path & Application.PathSeparator & "develop_src"
    
    Dim version_file_path As String
    version_file_path = publish_src_dirname & Application.PathSeparator & "Version.bas"
    
    Dim cmd_line_to_subwcrev As String
    cmd_line_to_subwcrev = "subwcrev " & wc_dirname & " " & version_file_path & " " & version_file_path
    Debug.Print cmd_line_to_subwcrev
    
    ' womacs has been changed from sourceforge(svn) to bitbucket(hg)
    'Shell cmd_line_to_subwcrev, 1
    
    ' pause to wait for subwcrev.exe to exit
    MsgBox "Please wait for subwcrev.exe to exit"
    
    
    ' import
    import_vba_source_code Doc, publish_src_dirname
    MsgBox "Importing source code to womacs.dotm"
    
    
    ' save
    Dim publish_basename As String
    publish_basename = "womacs.dotm"
    
    Dim publish_path As String
    publish_path = publish_dirname & Application.PathSeparator & publish_basename


    If Dir(publish_path) <> "" Then
        Kill publish_path
        ' pause to avoid bug with vba
        MsgBox "Deleting old womacs.dotm"
    End If

    
    Doc.SaveAs2 FileName:=publish_path, FileFormat:=wdFormatXMLTemplateMacroEnabled
    'wdFormatTemplate dot
    
    MsgBox "Saving new womacs.dotm"
    Doc.Close
        

    MsgBox "Finished"
    
End Sub

'-------------------------------------------------------------------------------------------------------------------
Sub show_all_key_bindings_in_current_doc()
    Dim kbLoop As KeyBinding
  
    CustomizationContext = ThisDocument
    For Each kbLoop In KeyBindings
        Selection.InsertAfter kbLoop.command & vbTab _
            & kbLoop.KeyString & vbCr
        Selection.Collapse direction:=wdCollapseEnd
    Next kbLoop
End Sub

'===================================================================================================================

Private Sub export_vba_source_code(c As VBComponent, src_dirname As String)
    Dim suffix As String
    'Debug.Print c.Name, c.CodeModule.Name, c.CodeModule.CountOfLines
    
    Select Case c.type
    Case vbext_ct_ClassModule
        suffix = ".cls"
    Case vbext_ct_MSForm
        suffix = ".frm"
    Case vbext_ct_StdModule
        suffix = ".bas"
    Case Else
        suffix = ""
    End Select
    
    If suffix <> "" Then
        Dim src_abspath
        
        src_abspath = src_dirname & Application.PathSeparator & c.name & suffix
            
        Debug.Print src_abspath
        
        c.Export src_abspath
    
    End If

End Sub


Private Sub import_vba_source_code(Doc As Document, src_dirname As String)
    Dim src_basename As String
    Dim src_abspath As String
    
    src_basename = Dir(src_dirname & Application.PathSeparator & "*")
    Do
        If src_basename = "" Then
            Exit Do
        End If
        
        If src_basename <> "DeveloperTools.bas" And InStr(src_basename, ".frx") = 0 Then
            src_abspath = src_dirname & Application.PathSeparator & src_basename
            
            Debug.Print src_abspath
            Doc.VBProject.VBComponents.Import src_abspath
        End If
        
        src_basename = Dir
        
    Loop

End Sub

