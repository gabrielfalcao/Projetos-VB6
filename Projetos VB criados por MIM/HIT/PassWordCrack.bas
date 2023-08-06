Attribute VB_Name = "crackPassword"
' ===================================================================
' crackPassword Source Code
' Version 1.0 (FREEWARE)
' Copyright (C) August/2000 Khaery Rida
' Special thanks to VBCode.com
' ===================================================================
'
'
' Thank you for using crackPassword!
' Please log on http://www.geocities.com/lio889 for more great VB programs.
' Comments or questions? Please do NOT hesitate at emailing me:
' lio_889@ziplip.com
' ===================================================================

' User defined type
Public Type point
    x As Long
    y As Long
End Type

' Public Variables
Public Targeting As Boolean
Public RetVal As Long
Public CursorPosition As POINTAPI
' Global Constants
Global Const MainTitle = "H1T - Password Cracker "
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

' Declare Windows' API functions
'Public Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As POINT) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Function GetTopLevelParent(ByVal hwndNum As Long) As Long
    'returns highest-level parent window of hWnd
    'if hWnd is the parent it just returns the input hWnd
    Dim ParentHwnd As Long
    Dim tmpHwnd As Long
    
    tmpHwnd = hwndNum
 If 0 <> IsWindow(tmpHwnd) Then ' make sure the input hWnd refers to a window
            ParentHwnd = GetParent(tmpHwnd)
            tmpHwnd = ParentHwnd
   
 End If
    
    GetTopLevelParent = hwndNum 'suposed to be parent hwnd
End Function


