VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' womacs
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1


Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)

    If Documents.Count = 0 Then
        Debug.Assert Doc.name = "Normal.dotm"
        'ActiveDocument: If there are no documents open, an error occurs.
        Exit Sub
    End If
    
    If womacs_on Then
    
        use_word_key_bindings
        delete_all_womacs_vars Doc
        
        If SaveAsUI Then
            On Error Resume Next ' if an user chooses 'Cancel' in Save As dialog, an error is raised.
            Dialogs(wdDialogFileSaveAs).Show
        Else
            On Error Resume Next ' if an user chooses 'Cancel' in Save As dialog, an error is raised.
            Doc.Save
        End If
        
        use_emacs_key_bindings
        Cancel = True
        
    End If
End Sub

' App_WindowDeactivate is called before App_WindowActivate
Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
 
    load_doc_locals Doc
    
    If womacs_on Then
        set_emacs_app_caption
    Else
        set_word_app_caption
    End If

End Sub


Private Sub App_WindowDeactivate(ByVal Doc As Document, ByVal Wn As Window)
    
    store_doc_locals Doc
    Set current_keymap = global_keymap

End Sub

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)

    If Not isearch_running = True Then
        Exit Sub
    End If

    If doing_search Then
        Exit Sub
    End If
    
    Debug.Print "fuck"
End Sub

