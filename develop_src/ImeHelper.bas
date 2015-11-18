Attribute VB_Name = "ImeHelper"
' womacs

#Const has_imm = 1

#If has_imm Then
' Office 2010 Help Files: Win32API_PtrSafe with 64-bit Support
' You can download it from: http://www.microsoft.com/download/en/details.aspx?id=9970
' or view it at: http://www.tudosobrexcel.com/vba/vba_32bits/Win32API_PtrSafe.TXT

Const SM_DBCSENABLED = 42
Const SM_IMMENABLED = 82
Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'The IMM is only enabled on East Asian (Chinese, Japanese, Korean)
'localized Windows operating systems. On these systems, the application
'calls GetSystemMetrics with SM_DBCSENABLED to determine if the IMM is enabled.
'Windows 2000: Full-featured IMM support is provided in all localized language
'versions. However, the IMM is enabled only when an Asian
'language pack is installed.
'An IME-aware application can call GetSystemMetrics with SM_IMMENABLED
'to determine if the IMM is enabled.
'Nonzero if User32.dll supports DBCS; otherwise, 0.


Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Declare PtrSafe Function ImmGetOpenStatus Lib "imm32.dll" (ByVal hIMC As LongPtr) As Long
Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal hIMC As LongPtr, ByVal b As Long) As Long
Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal hIMC As LongPtr) As Long

#Else
#End If


Dim ime_status As Boolean


Public Sub ImeSetStatus(fOpen As Boolean)
#If has_imm Then
    Dim hIMC As LongPtr
    hIMC = ImmGetContext(GetActiveWindow())
    Debug.Assert hIMC <> 0
    
    
    Dim ok As Boolean
    ok = ImmSetOpenStatus(hIMC, fOpen)
    Debug.Assert ok = True
    
    ImmReleaseContext GetActiveWindow(), hIMC
#End If
End Sub


Public Function ImeGetStatus() As Boolean
#If has_imm Then
    Dim hIMC As LongPtr
    hIMC = ImmGetContext(GetActiveWindow())
    Debug.Assert hIMC <> 0
    
    Dim fOpen As Long
    fOpen = ImmGetOpenStatus(hIMC)
    
    ImmReleaseContext GetActiveWindow(), hIMC
    
    ImeGetStatus = fOpen
#Else
    ImeGetStatus = False
#End If
End Function


Public Sub ImeSaveAndSetStatus(fOpen As Boolean)
#If has_imm Then
    If GetSystemMetrics(SM_IMMENABLED) = 0 Then
        Exit Sub
    End If
    
    ime_status = ImeGetStatus()
    Debug.Print "before set: " & ime_status
    ImeSetStatus fOpen
#End If
End Sub


Public Sub ImeRestoreStatus()
#If has_imm Then
    If GetSystemMetrics(SM_IMMENABLED) = 0 Then
        Exit Sub
    End If
    
    Debug.Print "before restore: " & ime_status
    ImeSetStatus ime_status
#End If
End Sub
