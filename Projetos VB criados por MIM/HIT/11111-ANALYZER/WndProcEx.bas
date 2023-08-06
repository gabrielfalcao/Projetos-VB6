Attribute VB_Name = "modWndProcEx"
Option Explicit
'
' Copyright © 1997-2000 Brad Martinez, http://www.mvps.org
'
' Code was written in and formatted for 8pt MS San Serif

' allocated buffer (pointer) containing OLDWNDPROC string, for faster GetProp retrieval
Private m_lpszOldWndProc As Long
Private m_lpszObjPtr As Long

' count of subclassed windows, incremented in SubClass proc where m_lpszOldWndProc is allocated when 0,
' decremented in UnSubClass proc where m_lpszOldWndProc is deallocated when 0
Private m_nSubClasses As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

Public Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5
Private Const WM_SHOWWINDOW = &H18
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115

Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

' LocalAlloc uFlags
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As Any) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As Any, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As Any) As Long

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

#If DEBUGWINDOWPROC Then
  ' maintains a WindowProcHook reference for each subclassed window.
  ' the window's handle is the collection item's key string.
  Private m_colWPHooks As New Collection
#End If
'

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As ScrollView = Nothing) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out
  
  If GetProp(hWnd, m_lpszOldWndProc) Then
    SubClass = True
    Exit Function
  End If
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)

#Else
    Dim objWPHook As WindowProcHook
    
    Set objWPHook = CreateWindowProcHook
    m_colWPHooks.Add objWPHook, CStr(hWnd)
    
    With objWPHook
      Call .SetMainProc(lpfnNew)
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
      Call .SetDebugProc(lpfnOld)
    End With

#End If
  
  If lpfnOld Then
    
    ' allocate m_lpszOldWndProc and m_lpszObjPtr if first subclass
    If (m_lpszOldWndProc = 0) Then   ' (m_nSubClasses = 0) Then
      m_lpszOldWndProc = LocalAlloc(LPTR, Len(OLDWNDPROC))
      If m_lpszOldWndProc Then
        Call lstrcpyA(ByVal m_lpszOldWndProc, ByVal OLDWNDPROC)
      End If
    End If
    
    m_lpszObjPtr = LocalAlloc(LPTR, Len(OBJECTPTR))
    If m_lpszObjPtr Then
      Call lstrcpyA(ByVal m_lpszObjPtr, ByVal OBJECTPTR)
    End If
    
    If m_lpszOldWndProc Then m_nSubClasses = m_nSubClasses + 1
    fSuccess = m_lpszOldWndProc
    
    fSuccess = fSuccess And SetProp(hWnd, m_lpszOldWndProc, lpfnOld)
    If (objNotify Is Nothing) = False Then
      fSuccess = fSuccess And SetProp(hWnd, m_lpszObjPtr, ObjPtr(objNotify))
    End If
  End If   ' lpfnOld
  
Out:
  If fSuccess Then
    SubClass = True
  
  Else
    If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
                  "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, m_lpszOldWndProc)
  If lpfnOld Then
    
    If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
      Call RemoveProp(hWnd, m_lpszOldWndProc)
      Call RemoveProp(hWnd, m_lpszObjPtr)
      
      ' deallocate and zero m_lpszOldWndProc and m_lpszObjPtr if last unsubclass
      m_nSubClasses = m_nSubClasses - 1
      If (m_nSubClasses = 0) Then
        If m_lpszOldWndProc Then
          If LocalFree(m_lpszOldWndProc) Then m_lpszOldWndProc = 0
        End If
        If m_lpszObjPtr Then
          If LocalFree(m_lpszObjPtr) Then m_lpszObjPtr = 0
        End If
      End If
      
#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      UnSubClass = True
    
    End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

Public Function SVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, m_lpszOldWndProc)
  If (lpfnOld = 0) Then
    Call UnSubClass(hWnd)
    Exit Function
  End If
  
  Select Case uMsg
    
    ' ============================================================
    Case WM_SHOWWINDOW, WM_SIZE, WM_HSCROLL, WM_VSCROLL, WM_DESTROY
      Dim objSV As ScrollView
      Dim pobjSV As Long
      
      pobjSV = GetProp(hWnd, m_lpszObjPtr)
      If pobjSV Then
        MoveMemory objSV, pobjSV, 4
        If (objSV Is Nothing) = False Then
          SVWndProc = objSV.UCWndProc(lpfnOld, hWnd, uMsg, wParam, lParam)
          MoveMemory objSV, 0&, 4
        End If
      End If
      
      If (pobjSV = 0) Then
        SVWndProc = CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
        Call UnSubClass(hWnd)
      End If
    
    ' ============================================================
    Case Else
      SVWndProc = CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
      
  End Select   ' uMsg
    
End Function
