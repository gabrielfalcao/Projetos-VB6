Attribute VB_Name = "modOpenSave"


Public Xhwnd As Long
Public PXhwnd As Long
Public Xpos As Long
Public Ypos As Long
Private prevID As Long
Private WWdt As Long

Private OldProcedura As Long
Private CLL As Long
Private DIALOGHWND As Long

Const GWL_ID = (-12)

Type POINTAPI
     x As Long
     y As Long
End Type

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXBORDER = 5
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYCAPTION = 4
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYBORDER = 6
Public Const SM_CYMENU = 15
Public Const SM_CYMENUSIZE = 55
Public Const SM_CXMENUSIZE = 54

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&)
Private TMR As Long
Private TMRHWND As Long

Const WM_NOTIFY = &H4E
Const WM_PAINT = &HF
Const WM_DRAWITEM = &H2B
Const WM_SETTEXT = &HC
Const WM_SETREDRAW = &HB

Public Const WM_CLOSE = &H10
' ============================================================================
' GetOpen/SaveFileName
Const WM_INITDIALOG = &H110

Const WM_CREATE = &H1
Const WM_PARENTNOTIFY = &H210
Const WM_NCCREATE = &H81

Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Const MAX_PATH = 260

Public Type OPENFILENAME  '  ofn
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As OFN_Flags
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

' File Open/Save Dialog Flags
Public Enum OFN_Flags
  OFN_READONLY = &H1
  OFN_OVERWRITEPROMPT = &H2
  OFN_HIDEREADONLY = &H4
  OFN_NOCHANGEDIR = &H8
  OFN_SHOWHELP = &H10
  OFN_ENABLEHOOK = &H20
  OFN_ENABLETEMPLATE = &H40
  OFN_ENABLETEMPLATEHANDLE = &H80
  OFN_NOVALIDATE = &H100
  OFN_ALLOWMULTISELECT = &H200
  OFN_EXTENSIONDIFFERENT = &H400
  OFN_PATHMUSTEXIST = &H800
  OFN_FILEMUSTEXIST = &H1000
  OFN_CREATEPROMPT = &H2000
  OFN_SHAREAWARE = &H4000
  OFN_NOREADONLYRETURN = &H8000&
  OFN_NOTESTFILECREATE = &H10000
  OFN_NONETWORKBUTTON = &H20000
  OFN_NOLONGNAMES = &H40000               ' force no long names for 4.x modules
  OFN_EXPLORER = &H80000                       ' new look commdlg
  OFN_NODEREFERENCELINKS = &H100000
  OFN_LONGNAMES = &H200000                 ' force long names for 3.x modules
  ' ===============================
  ' Win98/NT5 only...
  OFN_ENABLEINCLUDENOTIFY = &H400000           ' send include message to callback
  OFN_ENABLESIZING = &H800000
  ' ===============================
End Enum

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
(ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As _
Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'
'Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Property Let SetWidth(ByVal newwidth As Long)
WWdt = newwidth
End Property

Public Sub InsertCtrl(ByVal Ctrlhwnd As Long, ByVal Parhwnd As Long, ByVal x As Long, ByVal y As Long)
Xhwnd = Ctrlhwnd
PXhwnd = Parhwnd
Xpos = x + 4
Ypos = y + 23
End Sub

Public Function GetOpenFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim OFN As OPENFILENAME
  
  With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .nFilterIndex = iFilter
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_SHAREAWARE
    .lpfnHook = GetAddress(AddressOf HookX)

  End With
  

  
  If GetOpenFileName(OFN) Then
    iFilter = OFN.nFilterIndex
    sFile = Mid$(OFN.lpstrFile, OFN.nFileOffset + 1, InStr(OFN.lpstrFile, vbNullChar) - (OFN.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(OFN.lpstrFile)
    GetOpenFilePath = True
  End If

End Function

Public Function GetSaveFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sDefExt As String, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim OFN As OPENFILENAME
  With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .lpstrDefExt = sDefExt
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_SHAREAWARE
    .lpfnHook = GetAddress(AddressOf HookX)
    
  End With
  
  If GetSaveFileName(OFN) Then
    iFilter = OFN.nFilterIndex
    sFile = Mid$(OFN.lpstrFile, OFN.nFileOffset + 1, InStr(OFN.lpstrFile, vbNullChar) - (OFN.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(OFN.lpstrFile)
    GetSaveFilePath = True
  End If

End Function

' Returns the string before first null char (if any) in an ANSII string.

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else

    GetStrFromBufferA = szA
  End If
End Function
Public Function GetAddress(ByVal address As Long) As Long
GetAddress = address
End Function

Public Function HookX(ByVal hDlg As Long, ByVal uiMsg As _
Long, ByVal wParam As Long, ByVal lParam As Long) As _
Long

Select Case uiMsg

Case WM_INITDIALOG
Dim przRECT As RECT
Dim Parhwnd As Long
Dim ctrlX As Long

Parhwnd = GetParent(hDlg) 'Uzmi Handle Dialoga
DIALOGHWND = Parhwnd
ctrlX = GetDlgItem(Parhwnd, &H1)  'Uzmi handle ID Item=1

'*****Poziv Timera
'TMRHWND = ctrlX
'TMR = SetTimer(TMRHWND, 2, 0, AddressOf Provjera)
'******************
CLL = ctrlX
OldProcedura = SetWindowLong(ctrlX, -4, AddressOf Provjera2)
SetWindowText CLL, "Visualizar"


SetParent Xhwnd, Parhwnd 'Promjeni vlasnika
prevID = GetDlgCtrlID(Xhwnd) 'Zapamti Stari ID
SetWindowLong Xhwnd, GWL_ID, &H6000 'Promjeni ID PictureBoxa da ga se ne dupla ID(proizvoljna vrijednost)

'Dim tmphwnd As Long
'tmphwnd = GetDlgItem(Parhwnd, 1)
'SetWindowLong tmphwnd, GWL_ID, &H6001
'SetWindowText tmphwnd, "Otvori"
'Pokušaj promjene ID ---prozor ne reagira

'*************************

Dim XY As POINTAPI
Dim RC1 As RECT
GetWindowRect Parhwnd, RC1
XY.x = RC1.Left
XY.y = RC1.Top
ScreenToClient Parhwnd, XY
GetWindowRect Xhwnd, RC1
Dim MTY As Long
MTY = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER)
Dim MTX As Long
MTX = GetSystemMetrics(SM_CXBORDER)

MoveWindow Xhwnd, XY.x + Xpos + MTX, XY.y + Ypos + MTY, RC1.Right - RC1.Left, RC1.Bottom - RC1.Top, 1
ShowWindow Xhwnd, 1
'********Postavi Picture Box

'Dim ID As Long
'Dim tmphwnd As Long
'tmphwnd = GetDlgItem(Parhwnd, 1)
'ID = GetDlgCtrlID(Xhwnd)
'Ovim gore je utvrdeno da je PictureBox došao na DLGItem=1

SetDlgItemText Parhwnd, &H2, "Cancelar"
SetDlgItemText Parhwnd, &H443, "Em:"
SetDlgItemText Parhwnd, &H442, "Nome:"
SetDlgItemText Parhwnd, &H441, "Tipo:"

GetWindowRect Parhwnd, RC1
MoveWindow Parhwnd, 0, 0, RC1.Right - RC1.Left, WWdt, 1 'Promjeni velicinu prozora



'******************************
GetWindowRect Parhwnd, przRECT
Dim x As Long
Dim y As Long
x = (Screen.Width / 15 - (przRECT.Right - przRECT.Left)) / 2
y = (Screen.Height / 15 - (przRECT.Bottom - przRECT.Top)) / 2
SetWindowPos Parhwnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
'*******Centriraj prozor





Case WM_DESTROY
ShowWindow Xhwnd, 0 'Sakrij kontrolu
'KillTimer TMRHWND, TMR 'Zgazi timer
SetParent Xhwnd, Parhwnd 'Vrati kontrolu na pravog vlasnika
SetWindowLong Xhwnd, GWL_ID, prevID 'Vrati stari ID
Call SetWindowLong(CLL, -4, OldProcedura)

End Select
End Function
Public Function Provjera2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Select Case uMsg





Case WM_SETTEXT

Dim TTX As String
Dim TX() As Byte
TTX = "Visualizar" & Chr(CByte(0))
TX = StrConv(TTX, vbFromUnicode)
'CopyMemory ByVal lparam, TX(0), Len(TTX)
lParam = VarPtr(TX(0))

'******Može i ovako......................
'Dim TX() As Byte 'Ansi
'Dim TTX As String 'Unicode
'TTX = "Snimi" & Chr(CByte(0))
'ReDim TX(Len(TTX-1))
'CopyMemory TX(0), ByVal TTX, Len(TTX)
'CopyMemory ByVal lparam, TX(0), Len(TTX)
'***************************************

'********Po tvome sam primjeru skužio dosta stvar...A to je da na lparam dolazi pointer
'na ANSI string... i umjesto da šaljem SetWindowText jednostavno sam kopirao strukturu,ne UNICODE veæ ANSI...

'********Nema razloga da se koristi VBACCELERATOROV Subclassing...(Dobar je za kontrole,da se ne ruši VB)

End Select


Provjera2 = CallWindowProc(OldProcedura, hWnd, uMsg, wParam, lParam)
End Function


Public Sub Provjera(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
'Konstantno mijenjaj Text na buttonu ako je promjenjen.
Dim TXT1 As String
Dim ltxt1 As Long
TXT1 = Space(20)
ltxt1 = GetWindowText(hWnd, TXT1, Len(TXT1))
TXT1 = Left(TXT1, ltxt1)
If TXT1 = "&Open" Then
SetWindowText hWnd, "Otvori Folder"
ElseIf TXT1 = "&Save" Then
SetWindowText hWnd, "Snimi"
End If

End Sub

Public Sub CloseDLG()
PostMessage DIALOGHWND, WM_CLOSE, 0, 0
End Sub




