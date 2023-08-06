Attribute VB_Name = "Module2"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
 ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As _
String, ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" _
Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function GetNextWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) _
As Long

Declare Function GetTickCount& Lib "kernel32" ()


Public Const WM_CLOSE = &H10


