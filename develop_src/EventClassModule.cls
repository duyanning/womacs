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
    
    'If womacs_on Then
    If get_womacs_status(Doc) = True Then
    
        use_word_key_bindings
        delete_all_womacs_vars Doc
        
        If SaveAsUI Then
            On Error Resume Next ' if an user chooses 'Cancel' in Save As dialog, an error is raised.
            Dialogs(wdDialogFileSaveAs).Show
        Else
            On Error Resume Next ' if an user chooses 'Cancel' in Save As dialog, an error is raised.
            Doc.Save
        End If
        
        set_womacs_status Doc, True '文件保存到硬盘前，会删除文档对象中的变量，现在得恢复。
        use_emacs_key_bindings
        Cancel = True
        
    End If
End Sub

'我忘了以前为啥要整下面这两个函数了
'可能是因为同时打开几个文档的情况下，就有好几个word主窗口。
'给一个word主窗口标题栏加上womacs后缀之后，切换到其他word主窗口时，womacs这个后缀仍在。
'而我们womacs应当是针对各个文档独立开启/关闭的。
'所以就要管理各个word主窗口的标题。
'如果该文档已经启用womacs，才给该word窗口的标题加上womacs后缀
'那如何才能知道一个文档是否已经开启了womacs呢？
'只能将womacs是否开启的状态保存在文档当中
'下面这俩函数就是在word主窗口激活时，根据文档中保存的这个状态来调整标题栏
'word主窗口失活时，保存womacs的状态信息。
'但现在有个问题是：为了克服在word2013+版本中Application.ScreenUpdating = False无效的问题
'我改用Application.Visible = False了，然而操作窗口的可见与否，会激发下面两个函数
'所以，我要用另外一个变量来记录，窗口激活/失活是否是因为我在模拟“Application.ScreenUpdating = False”


' App_WindowDeactivate is called before App_WindowActivate
Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    'MsgBox "App_WindowActivate as_screen_updating: " & as_screen_updating
'    Debug.Print "App_WindowActivate as_screen_updating: " & as_screen_updating
'    If as_screen_updating Then
'        Exit Sub
'    End If

    load_doc_locals Doc

    ' 根据文档是否开启了womacs来决定word窗口的标题栏里是否带womacs字样
    'If womacs_on Then
    If get_womacs_status(Doc) = True Then
        set_emacs_app_caption
    Else
        set_word_app_caption
    End If

End Sub


Private Sub App_WindowDeactivate(ByVal Doc As Document, ByVal Wn As Window)
    'MsgBox "App_WindowDeactivate as_screen_updating: " & as_screen_updating
'    Debug.Print "App_WindowDeactivate as_screen_updating: " & as_screen_updating
'    If as_screen_updating Then
'        Exit Sub
'    End If

    store_doc_locals Doc
    '下面这行不能要，因为现在处理C-x C-s时，当按下C-x时，就会主动激活/失活窗口
    ' 导致本过程被调用，下面这行的执行会话破坏刚刚建立起来的C-x keymap
    'Set current_keymap = global_keymap

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

