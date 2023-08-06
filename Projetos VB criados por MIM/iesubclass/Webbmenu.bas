Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
    
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_RBUTTONDOWN = &H204

Public origWndProc As Long

Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_MOUSEACTIVATE
            Dim C As Integer
            Call CopyMemory(C, ByVal VarPtr(lParam) + 2, 2)
            If C = WM_RBUTTONDOWN Then
                Form1.PopupMenu Form1.mnuBrowser
                SendKeys "{ESC}"
            End If
        Case WM_CONTEXTMENU
            Form1.PopupMenu Form1.mnuBrowser
            SendKeys "{ESC}"
    End Select
    AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
End Function
