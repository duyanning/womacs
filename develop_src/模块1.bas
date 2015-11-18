Attribute VB_Name = "Ä£¿é1"
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
Declare PtrSafe Function SetCapture Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function GetCapture Lib "user32" () As LongPtr
Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr


