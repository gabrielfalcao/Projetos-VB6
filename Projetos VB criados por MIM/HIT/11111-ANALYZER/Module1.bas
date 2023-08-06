Attribute VB_Name = "modLib"
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public HMOD As Long
Public Const LOAD_LIBRARY_AS_DATAFILE = &H2
Public Const DONT_RESOLVE_DLL_REFERENCES = &H1
Public Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111

Public Type DLGTEMPLATE
    style As Long
    dwExtendedStyle As Long
    cdit As Integer
    X As Integer
    Y As Integer
    cx As Integer
    cy As Integer
End Type

Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm" Alias _
"mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, _
ByVal uLength As Long) As Long

Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, _
ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Const WS_CHILD = &H40000000
Public Const DS_MODALFRAME = &H80
Public Const DS_SYSMODAL = &H2&

Public MHDL As Long
Public HHDL As Long

Private Declare Function EndDialog Lib "user32" ( _
    ByVal hDlg As Long, _
    ByVal nResult As Long _
) As Long

Public OtherData() As Byte 'General LOADDATA
Public LangID As Integer
Public TypePtr As Long
Public TrueType() As Byte
Public TrueName As Long
Public TrueBuffer() As Byte
Public ResTotLen As Long

Public Const MF_OWNERDRAW = &H100&
Public Const MF_BYPOSITION = &H400&
Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long

Sub PlayAVIPictureBox(Filename As String, ByVal Window As PictureBox)
Dim RetVal As Long
Dim CommandString As String
Dim ShortFileName As String * 260
Dim deviceIsOpen As Boolean
RetVal = GetShortPathName(Filename, ShortFileName, Len(ShortFileName))
Filename = left$(ShortFileName, RetVal)
CommandString = "Open " & Filename & " type AVIVideo alias AVIFile parent " _
& CStr(Window.hWnd) & " style " & CStr(WS_CHILD)
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal Then GoTo Error
deviceIsOpen = True
CommandString = "put AVIFile window at 0 0 " & CStr(Window.ScaleWidth / _
Screen.TwipsPerPixelX) & " " & CStr(Window.ScaleHeight / _
Screen.TwipsPerPixelY)
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo Error
CommandString = "Play AVIFile wait"
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo Error
CommandString = "Close AVIFile"
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo Error
Exit Sub
Error:
Dim ErrorString As String
ErrorString = Space$(256)
mciGetErrorString RetVal, ErrorString, Len(ErrorString)
ErrorString = left$(ErrorString, InStr(ErrorString, vbNullChar) - 1)
If deviceIsOpen Then
CommandString = "Close AVIFile"
mciSendString CommandString, vbNullString, 0, 0&
End If
MsgBox ErrorString, vbCritical, "Error"
End Sub

Public Sub LoadIntoMemory(ByVal nameX As Long, ByVal typeX As Long)
Dim HOBJ As Long
Dim TTB As Long
Dim TTF As Long
Dim SZ As Long
HOBJ = FindResource(HMOD, nameX, typeX)
TTB = LoadResource(HMOD, HOBJ)
TTF = LockResource(TTB)
SZ = SizeofResource(HMOD, HOBJ)
ReDim OtherData(SZ - 1)
CopyMemory OtherData(0), ByVal TTF, SZ
ResTotLen = SZ
End Sub

Public Sub SetWinPosByCursor(ByVal hWnd As Long, ByVal stat As Long)
Dim ppt As POINTAPI
GetCursorPos ppt
SetWindowPos hWnd, 0, ppt.X, ppt.Y, 0, 0, stat
End Sub


Public Function dialogProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'If uMsg = WM_COMMAND Then
'EndDialog hwnd, 0
'dialogProc = 0
'End If

End Function
