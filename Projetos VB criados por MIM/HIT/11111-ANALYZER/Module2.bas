Attribute VB_Name = "modResource"
Public RESTYPE As New Collection
Public RESNAME As New Collection
Public RESTYPENAME As New Collection


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Const DIFFERENCE = 11
Const RT_ACCELERATOR = 9&
Const RT_ANICURSOR = (21)
Const RT_ANIICON = (22)
Const RT_BITMAP = 2&
Const RT_CURSOR = 1&
Const RT_DIALOG = 5&
Const RT_DLGINCLUDE = (17)
Const RT_FONT = 8&
Const RT_FONTDIR = 7&
Const RT_ICON = 3&
Const RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
Const RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
Const RT_HTML = (23)
Const RT_MENU = 4&
Const RT_MESSAGETABLE = (11)
Const RT_PLUGPLAY = (19)
Const RT_RCDATA = 10&
Const RT_STRING = 6&
Const RT_VERSION = (16)
Const RT_VXD = (20)
Public Xcnt As Long

Function EnumRSLang(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lpszname As Long, ByVal lpszID As Integer, ByVal lParam As Long) As Long
LangID = lpszID
EnumRSLang = 1
End Function



Function EnumRSType(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lParam As Long) As Long
Call EnumResourceNames(hModule, lpsztype, AddressOf EnumRSName, Xcnt)
EnumRSType = 1
End Function
Function EnumRSName(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lpszname As Long, ByVal lParam As Long) As Long
Dim typeX As String
Dim nameX As String


If lpsztype = RT_STRING Then
SetStringName hModule, lpsztype, lpszname
Else
typeX = GetStringFromPointer(lpsztype)
nameX = GetStringFromPointer(lpszname)
SetPropName lpsztype, typeX
RESTYPE.Add typeX
RESNAME.Add nameX
End If

EnumRSName = 1
End Function

Public Sub ClearCOLLECTION()
Set RESTYPE = Nothing
Set RESNAME = Nothing
Set RESTYPENAME = Nothing
End Sub
Function GetStringFromPointer(ByVal point As Long) As String
Dim llen As Long
If (point > &HFFFF&) Or (point < 0) Then
llen = lstrlen(point)
GetStringFromPointer = Space(llen)
CopyMemory ByVal GetStringFromPointer, ByVal point, llen
Else
GetStringFromPointer = CStr(point)
End If
End Function

Private Sub SetPropName(ByVal nameX As Long, ByVal typeX As String)
Dim nname As String
Select Case nameX
Case RT_ACCELERATOR
nname = "Accelerator Table"
Case RT_ANICURSOR
nname = "Animated Cursor"
Case RT_ANIICON
nname = "Animated Icon"
Case RT_BITMAP
nname = "Bitmap"
Case RT_CURSOR
nname = "Hardware Cursor"
Case RT_DIALOG
nname = "Dialog Box"
Case RT_DLGINCLUDE
nname = "Header file that contains menu and dialog box define statements"
Case RT_FONT
nname = "Font"
Case RT_FONTDIR
nname = "Font directory"
Case RT_ICON
nname = "Hardware Icon"
Case RT_GROUP_CURSOR
nname = "Cursor"
Case RT_GROUP_ICON
nname = "Icon"
Case RT_HTML
nname = "HTML document"
Case RT_MENU
nname = "Menu"
Case RT_MESSAGETABLE
nname = "Message Table"
Case RT_PLUGPLAY
nname = "Plug and Play"
Case RT_RCDATA
nname = "Raw Data"
Case RT_VERSION
nname = "Version Info"
Case RT_VXD
nname = "VXD"
Case Else
If IsNumeric(typeX) Then
nname = "Custom Defined"
Else
nname = typeX
End If
End Select
RESTYPENAME.Add nname
End Sub
Public Sub SetStringName(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lpszname As Long)
Dim typeX As String
Dim buffer As String
Dim ret As Long
For i = ((lpszname - 1) * 16) To (lpszname * 16)
buffer = Space(255)
ret = LoadString(hModule, i, buffer, Len(buffer))
If ret <> 0 Then
buffer = left(buffer, ret)
typeX = GetStringFromPointer(lpsztype)
nameX = i
RESTYPE.Add typeX
RESNAME.Add nameX & "  :" & buffer
RESTYPENAME.Add "String"
End If
Next i


End Sub



