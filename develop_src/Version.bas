Attribute VB_Name = "Version"
' womacs

Option Explicit



Sub version()
    MsgBox "version svn" & "$WCREV$", vbOKOnly, "Womacs"
    
    complete
End Sub

