Attribute VB_Name = "Module1"
Option Explicit
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


Public Const EWX_REBOOT As Long = 2
Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
     
     Public Declare Function GetCurrentProcessId Lib _
   "kernel32.dll" () As Long

Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'''''''''''''''''''''''''''''''''''''''

Public Type ProcData
    AppHwnd As Long
    title As String
    Placement As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As _
Any, ByVal lParam As Long) As Long

Public Function EditInfo(window_hwnd As Long) As String
Dim txt As String
Dim buf As String
Dim buflen As Long
Dim child_hwnd As Long
Dim children() As Long
Dim num_children As Integer
Dim i As Integer

    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)
''''

  If buf = "Edit" Then
        EditInfo = WindowText(window_hwnd)
       Exit Function
   End If
  '
  '
    num_children = 0
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)
    Do While child_hwnd <> 0
        num_children = num_children + 1
        ReDim Preserve children(1 To num_children)
        children(num_children) = child_hwnd
        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop
    
    
    For i = 1 To num_children
        txt = EditInfo(children(i))
        If txt <> "" Then Exit For
    Next i

    EditInfo = txt
End Function

Public Function WindowText(window_hwnd As Long) As String
Dim txtlen As Long
Dim txt As String

    WindowText = ""
    If window_hwnd = 0 Then Exit Function
    
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Function
    
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)
End Function
Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
Dim buf As String * 1024
Dim title As String
Dim Length As Long
  
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    title = Left$(buf, Length)

   
    If Right$(title, 30) = " - Microsoft Internet Explorer" Then
        
        frmChatCliente.Caption = EditInfo(app_hwnd)

       
        EnumProc = 0
    Else
        
        EnumProc = 1
    End If
End Function







