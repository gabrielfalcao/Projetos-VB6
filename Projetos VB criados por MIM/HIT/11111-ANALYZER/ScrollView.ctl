VERSION 5.00
Begin VB.UserControl ScrollView 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   5175
End
Attribute VB_Name = "ScrollView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-2000 Brad Martinez, http://www.mvps.org
'
' Demonstrates how to create a real (and native) scrollable viewport with a
' UserControl, when displaying either a Picture, or child contained controls.
'
' ========================================================
' This project uses subclassing, and utilizes the services of the "Debug
' Object for AddressOf Subclassing" ActiveX server, Dbgwproc.dll, which
' allows unencumbered code execution when stepping through code in
' the VB IDE. This server is freely distributable and can be obtained from
' Microsoft at http://msdn.microsoft.com/vbasic/downloads/controls.asp.

' Set the conditional compilation argument:   DEBUGWINDOWPROC = 1
' in the project properties dialog/Make tab to enable the server's services.
' ========================================================

' Set to non-zero to have significant proc names printed to the Immediate window
' note: this may affect scrollbar operation (SB_THUMBTRACK may not work [???])
'#Const PROCLOG = 1

' ============================================================================
' public member definitions

Private m_stdpic As New StdPicture   ' Picture property

Public Enum DisabledScrollbarTypes   ' ShowDisabledScrollbars property
  dstNone = 0
  dstHorizontal = 1
  dstVertical = 2
  dstBoth = 3
End Enum
Private m_dwDisabledScrollbars As DisabledScrollbarTypes

Private m_vTag As Variant   ' Tag property

Public Enum ViewTypes   ' View property
  vtEmpty = 0
  vtPicture = 1
  vtControl = 2
End Enum
Private m_dwView As ViewTypes

Public Event ViewChange(ByVal OldView As ViewTypes, ByVal NewView As ViewTypes)

' ============================================================================
' private control definitions

Private m_cxTPP As Long   ' Screen.TwipsPerPixelX
Private m_cyTPP As Long   ' Screen.TwipsPerPixelY

Private m_cxChar As Long   ' average font width, in pixels
Private m_cyChar As Long   ' average font height, in pixels

Private m_cxView As Long   ' width of current view, in pixels
Private m_cyView As Long   ' height of current view, in pixels

' used to determine average font width and height (m_cxChar, m_cyChar)
Private Const CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"

Private Const TWIPSPERINCH = 1440           ' TWIPSPERINCH / m_cx/yTPP = GetDeviceCaps(LOGPIXELSX/Y) (96 for small fonts)
Private Const HIMETRICSPERINCH = 2540   ' HIMETRIC units (0.01 millimeters) per inch

' see
' Q166473: PRB: CScrollView Scroll Range Limited to 32K:
' "NOTE: Windows 98, Windows 95 and Win32s only support logical and device coordinates up to 32K."
' and
' Q136989: INFO: Developing Win32-Based GDI Apps for Windows 95 and Windows NT or Windows2000:
' "If you pass 32-bit coordinates to GDI functions in Windows 95, the system truncates the upper 16 bits of
'  the coordinates before actually performing the function."
Private Const MAX_DC = &H7FFF   ' user-defined

Private m_hwndUC As Long
Private m_fDesigntime As Boolean
Private m_fShown As Boolean   ' set on WM_SHOWWINDOW for designtime unsubclass
Private m_fSizing As Boolean   ' prevents WM_SIZE recursion when calling SWP

' BackColor brush, for FillRect in SetPictureView, used instead of UserControl.Cls, is much faster
Private m_brshUCBack As Long

' ============================================================================
' various api definitions

Private Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5
Private Const WM_SETREDRAW = &HB
Private Const WM_SHOWWINDOW = &H18
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, lpPoint As Any) As Long
'Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, pPoint As POINTAPI) As Long
Private Declare Function OffsetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOffset As Long, ByVal nYOffset As Long, pPoint As Any) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
 
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbsh As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

' ============================================================================
' SetWindowPos

Private Enum SWP_hWndInsertAfter
  HWND_TOP = 0
  HWND_BOTTOM = 1
  HWND_TOPMOST = -1
  HWND_NOTOPMOST = -2
End Enum

Private Enum SWP_Flags
  SWP_NOSIZE = &H1
  SWP_NOMOVE = &H2
  SWP_NOZORDER = &H4
  SWP_NOREDRAW = &H8
  SWP_NOACTIVATE = &H10
  SWP_FRAMECHANGED = &H20      ' The frame changed: send WM_NCCALCSIZE
  SWP_SHOWWINDOW = &H40
  SWP_HIDEWINDOW = &H80
  SWP_NOCOPYBITS = &H100
  SWP_NOOWNERZORDER = &H200    ' Don't do owner Z ordering
  SWP_NOSENDCHANGING = &H400   ' Don't send WM_WINDOWPOSCHANGING

  SWP_DRAWFRAME = SWP_FRAMECHANGED
  SWP_NOREPOSITION = SWP_NOOWNERZORDER

  SWP_DEFERERASE = &H2000
  SWP_ASYNCWINDOWPOS = &H4000
End Enum

Private Declare Function SetWindowPos Lib "user32" _
                              (ByVal hWnd As Long, _
                              ByVal hwndInsertAfter As SWP_hWndInsertAfter, _
                              ByVal x As Long, ByVal y As Long, _
                              ByVal cx As Long, ByVal cy As Long, _
                              ByVal uFlags As SWP_Flags) As Long

' ============================================================================
' scrollbar

Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Private Const WS_VISIBLE = &H10000000

Private Type SCROLLINFO
  cbSize As Long
  fMask As Long 'SIF_Mask
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type

Private Enum SIF_Mask
  SIF_RANGE = &H1
  SIF_PAGE = &H2
  SIF_POS = &H4
  SIF_DISABLENOSCROLL = &H8
  SIF_TRACKPOS = &H10
  SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
End Enum

Private Enum SB_Type
  SB_HORZ = 0
  SB_VERT = 1
  SB_CTL = 2
  SB_BOTH = 3
End Enum

Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO) As Boolean
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO, ByVal fRedraw As Long) As Boolean
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type) As Boolean

' Scroll Bar Commands for WM_H/VSCROLL
Private Enum SB_Commands
  SB_LINEUP = 0
  SB_LINELEFT = 0
  SB_LINEDOWN = 1
  SB_LINERIGHT = 1
  SB_PAGEUP = 2
  SB_PAGELEFT = 2
  SB_PAGEDOWN = 3
  SB_PAGERIGHT = 3
  SB_THUMBPOSITION = 4
  SB_THUMBTRACK = 5
  SB_TOP = 6
  SB_LEFT = 6
  SB_BOTTOM = 7
  SB_RIGHT = 7
  SB_ENDSCROLL = 8
End Enum

'Private Const WM_HSCROLL = &H114
'Private Const WM_VSCROLL = &H115

Private Enum SW_Flags
  SW_SCROLLCHILDREN = &H1   ' Scroll children within *lprcScroll.
  SW_INVALIDATE = &H2              ' Invalidate after scrolling
  SW_ERASE = &H4                       ' If SW_INVALIDATE, don't send WM_ERASEBACKGROUND
#If (WINVER >= &H500) Then
  SW_SMOOTHSCROLL = &H10    ' Use smooth scrolling
#End If
End Enum

Private Enum SW_RtnVals   ' ScrollWindowEx
  RGN_ERROR = 0
  NULLREGION = 1
  SIMPLEREGION = 2
  COMPLEXREGION = 3
End Enum

' all rects are in client coords.
'int ScrollWindowEx(
'    HWND hWnd,  // handle of window to scroll
'    int dx, // amount of horizontal scrolling
'    int dy, // amount of vertical scrolling
'    CONST RECT *prcScroll,  // address of structure with scroll rectangle, NULL for enitre client area
'    CONST RECT *prcClip,  // address of structure with clip rectangle, can be NULL
'    HRGN hrgnUpdate,  // handle of update region, can be NULL
'    LPRECT prcUpdate, // address of structure to receive update rectangle, can be NULL
'    UINT flags  // scrolling flags
'   );
Private Declare Function ScrollWindowEx Lib "user32" _
                              (ByVal hWnd As Long, _
                              ByVal dx As Long, _
                              ByVal dy As Long, _
                              prcScroll As Any, _
                              prcClip As Any, _
                              ByVal hrgnUpdate As Long, _
                              prcUpdate As Any, _
                              ByVal flags As SW_Flags) As SW_RtnVals
'

' ============================================================================
' UserControl

Private Sub UserControl_Initialize()
  
  With UserControl
'    .BorderStyle = vbFixedSingle
'    .ClipControls = False
'    .ControlContainer = True
    .KeyPreview = True   ' for accelerator key receipt in the Key* events
'    .Name = "ScrollView"
    .ScaleMode = vbPixels
  End With

  ' can't be changed w/o a reboot
  m_cxTPP = Screen.TwipsPerPixelX
  m_cyTPP = Screen.TwipsPerPixelY

End Sub

' UserControl.ScaleWidth/Height are not valid until after the UC gets
' its first WM_SHOWWINDOW, where the UC is finally sited on the
' client, see UCWndProc/WM_SHOWWINDOW

Private Sub UserControl_InitProperties()
  ' assign to properties, have to create m_brshUCBack and set mod level variables
  BackColor = vbWindowBackground
  ForeColor = vbWindowText
  Set m_stdpic = Nothing
  m_dwDisabledScrollbars = dstNone
  m_vTag = ""
  m_dwView = vtEmpty
  Call InitControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
  ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
  Set m_stdpic = PropBag.ReadProperty("Picture", Nothing)
  m_dwDisabledScrollbars = PropBag.ReadProperty("ShowDisabledScrollbars", dstNone)
  m_vTag = PropBag.ReadProperty("Tag", "")
  m_dwView = PropBag.ReadProperty("View", vtEmpty)
  Call InitControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbWindowBackground)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, vbWindowText)
  ' m_stdpic may not be Nothing and still hold a null picture, which gets written to prop bag...
  If m_stdpic Then Call PropBag.WriteProperty("Picture", m_stdpic, 0)
  Call PropBag.WriteProperty("ShowDisabledScrollbars", m_dwDisabledScrollbars, dstNone)
  Call PropBag.WriteProperty("Tag", m_vTag, "")
  Call PropBag.WriteProperty("View", m_dwView, vtEmpty)
End Sub

Private Sub InitControl()
  
  On Error Resume Next
  m_fDesigntime = (Ambient.UserMode = False)
  On Error GoTo 0
  
  m_cxChar = UserControl.TextWidth(CHARS) / Len(CHARS)
  m_cyChar = UserControl.TextHeight(CHARS)
  
  m_hwndUC = UserControl.hWnd
  Call SubClass(m_hwndUC, AddressOf SVWndProc, Me)
  
End Sub

Private Sub UserControl_Terminate()
  Call UnSubClass(m_hwndUC)
  If m_brshUCBack Then Call DeleteObject(m_brshUCBack)
End Sub

' adjusts the scroll view only in designtime, WM_SIZE is handled at runtime

Private Sub UserControl_Resize()
  If (m_fShown And m_fDesigntime) Then Call SetScrollbars(False)
End Sub

' ============================================================================
' public members

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(oleclr As OLE_COLOR)
  
#If PROCLOG Then
  Debug.Print "Let BackColor"
#End If
  
  If (UserControl.BackColor <> oleclr) Then
    UserControl.BackColor = oleclr
    Call PropertyChanged("BackColor")
      
    If m_brshUCBack Then Call DeleteObject(m_brshUCBack)
    m_brshUCBack = CreateSolidBrush(OleColorToColorRef(oleclr))
    
    ' redraw the current view if necessary...
    If (m_dwView = vtPicture) Then Call SetPictureView(m_stdpic)
  End If
  
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(oleclr As OLE_COLOR)
  
#If PROCLOG Then
  Debug.Print "Let ForeColor"
#End If
  
  UserControl.ForeColor = oleclr
  Call PropertyChanged("ForeColor")

End Property

Public Property Get hWnd() As Long
  hWnd = m_hwndUC
End Property

Public Property Get Picture() As StdPicture
  Set Picture = m_stdpic
End Property

Public Property Set Picture(pic As StdPicture)
  
#If PROCLOG Then
  Debug.Print "Set Picture"
#End If
  
  Set m_stdpic = pic
  Call PropertyChanged("Picture")
  
  ' reset the current view if necessary...
  If (m_dwView And vtPicture) Then View = m_dwView

End Property

Public Property Get ShowDisabledScrollbars() As DisabledScrollbarTypes
  ShowDisabledScrollbars = m_dwDisabledScrollbars
End Property

Public Property Let ShowDisabledScrollbars(dw As DisabledScrollbarTypes)
  
#If PROCLOG Then
  Debug.Print "Let ShowDisabledScrollbars"
#End If
  
  m_dwDisabledScrollbars = dw
  Call PropertyChanged("ShowDisabledScrollbars")
  Call SetScrollbars(False)

End Property

Public Property Get Tag() As Variant
  Tag = m_vTag
End Property

Public Property Let Tag(v As Variant)
  m_vTag = v
  Call PropertyChanged("Tag")
End Property

Public Property Set Tag(v As Variant)
  Set m_vTag = v
  Call PropertyChanged("Tag")
End Property

Public Property Get View() As ViewTypes
  View = m_dwView
End Property

Public Property Let View(dw As ViewTypes)
  Dim dwViewOld As ViewTypes
  
#If PROCLOG Then
  Debug.Print "Let View", vbTab, vbTab, vbTab, dw
#End If
  
'  Screen.MousePointer = vbHourglass
  
  dwViewOld = m_dwView
  
  m_dwView = dw
  Call PropertyChanged("View")

  Select Case dw
    
    ' ============================================================
    Case vtEmpty
      
      m_cxView = 0
      m_cyView = 0
    
      Call SetControlView(False)
      Set UserControl.Picture = Nothing   ' no need to adjust the update region
    
    ' ============================================================
    Case vtPicture
      
      ' get the view dimensions in pixels (convert m_stdpic.Width and Height from HIMETRIC)
      m_cxView = Min((m_stdpic.Width * (TWIPSPERINCH / m_cxTPP)) / HIMETRICSPERINCH, MAX_DC)
      m_cyView = Min((m_stdpic.Height * (TWIPSPERINCH / m_cyTPP)) / HIMETRICSPERINCH, MAX_DC)
      
      Call SetControlView(False)
      Call SetPictureView(m_stdpic)
      
    ' ============================================================
    Case vtControl
      
      Call SetControlView(True, m_cxView, m_cyView)   ' sets m_cxView, m_cyView
      Call SetPictureView(Nothing)   ' have to set the update region
  
  End Select
  
  Call SetScrollbars(True)   ' fResetOrigin
  
  RaiseEvent ViewChange(dwViewOld, m_dwView)
  
'  Screen.MousePointer = vbNormal

End Property

' ============================================================================
' private view functions

' called from:
'   Property Let View:
'     vtEmpty, vtPicture (fVisible = False)
'     vtControl (fVisible = True)

Private Sub SetControlView(fVisible As Boolean, _
                                            Optional xView As Long, _
                                            Optional yView As Long)
  Dim nCtrls As Long
  Dim ctrl As Control
  Dim xCtrl As Long
  Dim yCtrl As Long
  Dim hwndChild As Long
  Dim rcChild As RECT
  
#If PROCLOG Then
  Debug.Print "SetControlView", fVisible
#End If
  
  ' ContainedControls may not be supported in the client...
  On Error Resume Next
  nCtrls = ContainedControls.Count
  On Error GoTo 0
  
  If nCtrls Then
    For Each ctrl In ContainedControls
      ctrl.Visible = fVisible
      If fVisible Then
        xCtrl = (ctrl.Left + ctrl.Width) \ m_cxTPP
        If (xCtrl > xView) Then xView = xCtrl
        yCtrl = (ctrl.Top + ctrl.Height) \ m_cyTPP
        If (yCtrl > yView) Then yView = yCtrl
      End If
    Next

  Else
    ' won't find VB intrinsic non-window controls (Label, Shape, etc)
    hwndChild = FindWindowEx(m_hwndUC, 0, vbNullString, vbNullString)
    Do While hwndChild
      Call SetWindowStyle(hwndChild, WS_VISIBLE, False, fVisible)
      If fVisible Then
        Call GetWindowRect(hwndChild, rcChild)
        Call ScreenToClient(m_hwndUC, rcChild.Right)
        If (rcChild.Right > xView) Then xView = rcChild.Right
        If (rcChild.Bottom > yView) Then yView = rcChild.Bottom
      End If
      hwndChild = FindWindowEx(m_hwndUC, hwndChild, vbNullString, vbNullString)
    Loop
    Call InvalidateRect(m_hwndUC, ByVal 0&, CTrue)
    
  End If   ' nCtrls
  
End Sub

' called from:
'   Property Let BackColor (m_dwView = vtPicture)
'   Property Let View (vtPicture, vtControl)
'
Private Sub SetPictureView(pic As StdPicture)
  Dim rcWnd As RECT
  Dim rcClient As RECT
  
#If PROCLOG Then
  Debug.Print "SetPictureView"
#End If

  m_fSizing = True
  
  ' Set the view dimensions and size the UserControl (and it's Picture, it must
  ' be large enough to accept the new graphic or the graphic will be clipped)
  Call GetWindowRect(m_hwndUC, rcWnd)
  Call SendMessage(m_hwndUC, WM_SETREDRAW, 0, 0)
  Call SetWindowPos(m_hwndUC, 0, 0, 0, _
                                  Max(Screen.Width \ m_cxTPP, _
                                          m_cxView + GetSystemMetrics(SM_CXVSCROLL) + (2 * GetSystemMetrics(SM_CXEDGE))), _
                                  Max(Screen.Height \ m_cyTPP, _
                                          m_cyView + GetSystemMetrics(SM_CYHSCROLL) + (2 * GetSystemMetrics(SM_CYEDGE))), _
                                  SWP_NOZORDER Or SWP_NOMOVE)
  
  ' paint the graphic to the UserControl Picture's DC, converting the graphic to a bitmap
  ' (this not only ensures that the UC's update region is as large as the view, but prevents
  ' a metafile from being stretched with the UC Picture's dimensions).
  With UserControl
    .AutoRedraw = True
    ' Not only is FillRect much faster than UserControl.Cls, but when not calling
    ' PaintPicture, it also causes the UserControl to set the size of its persistent
    ' bitmap (update region) to what we set it to above (otherwise it would be set
    ' to the size of the UserControl's client area, and any area outside of the client
    ' area would not be painted when the UserControl is enlarged...)
    Call GetClientRect(m_hwndUC, rcClient)
    Call FillRect(.hdc, rcClient, m_brshUCBack)
    If (pic Is Nothing) = False Then
      If pic Then Call .PaintPicture(pic, 0, 0)
    End If
'    .Picture = .Image   ' <--- not necessary, and requires a lot of processing...
    .AutoRedraw = False
  End With
  
  ' Restore the UserControl to its original window size.
  Call SetWindowPos(m_hwndUC, 0, 0, 0, _
                                  rcWnd.Right - rcWnd.Left, _
                                  rcWnd.Bottom - rcWnd.Top, _
                                  SWP_NOZORDER Or SWP_NOMOVE)
  Call SendMessage(m_hwndUC, WM_SETREDRAW, 1, 0)
'  Call InvalidateRect(m_hwndUC, ByVal 0&, CTrue)
  UserControl.Refresh
  
  m_fSizing = False

End Sub

' ============================================================================
' private scrollbar functions

' called from:
'   UserControl_Resize (fResetOrigin = False, designtime)
'   UCWndProc/WM_SIZE (fResetOrigin = False, runtime)
'   UCWndProc/WM_SHOWWINDOW (fResetOrigin = True)
'   Property Let ShowDisabledScrollbars (fResetOrigin = False)
'   Property Let View (all 4 views, fResetOrigin = True x 4)

Private Static Sub SetScrollbars(fResetOrigin As Boolean)
  Dim rcClient As RECT
  Dim si As SCROLLINFO
  Dim xScroll As Long
  Dim yScroll As Long
  Dim ptOrigin As POINTAPI
  Dim ptCaret As POINTAPI
  
#If PROCLOG Then
  Debug.Print "SetScrollbars start", vbTab, fResetOrigin
#End If

  If (m_hwndUC = 0) Or (m_fShown = False) Then Exit Sub
  
    ' ============================================================
  ' Set the scrollbars according to the current view and UC client window sizes
  
  ' Though SetScrollInfo implicitly sets scrollbar visibility (WS_H/VSCROLL),
  ' explictly set the scrollbars now and re-read GetClientRect so that si.nPage
  ' is set correctly (have to SWP/SWP_DRAWFRAME).
  Call GetClientRect(m_hwndUC, rcClient)
  Call SetWindowStyle(m_hwndUC, WS_HSCROLL, False, _
                                    (m_cxView > rcClient.Right) Or CBool(m_dwDisabledScrollbars And dstHorizontal), True)
  Call SetWindowStyle(m_hwndUC, WS_VSCROLL, False, _
                                    (m_cyView > rcClient.Bottom) Or CBool(m_dwDisabledScrollbars And dstVertical), True)
  Call GetClientRect(m_hwndUC, rcClient)
  
  ' Set the size of the scrollbar thumbs (scrollbars are created with the
  ' following defaults: nMin = 0, nMax = 0, nPage = 0, nPos = 0)
  si.cbSize = Len(si)
  si.nMin = 0
  If fResetOrigin Then si.nPos = 0
  
  si.fMask = SIF_RANGE Or SIF_PAGE Or (fResetOrigin And SIF_POS) Or _
                   (SIF_DISABLENOSCROLL And CBool(m_dwDisabledScrollbars And dstHorizontal))
  si.nMax = m_cxView
  si.nPage = rcClient.Right
  Call SetScrollInfo(m_hwndUC, SB_HORZ, si, 1)
  
  si.fMask = SIF_RANGE Or SIF_PAGE Or (fResetOrigin And SIF_POS) Or _
                   (SIF_DISABLENOSCROLL And CBool(m_dwDisabledScrollbars And dstVertical))
  si.nMax = m_cyView
  si.nPage = rcClient.Bottom
  Call SetScrollInfo(m_hwndUC, SB_VERT, si, 1)
  
    ' ============================================================
  ' Adjust the UserControl's scroll and DC origins with respect to the current thumb
  ' positions (again, si.nPos is 0 when a scrollbar style bit is first set).
  
  ' Get the UserControl's current origin offset
  Call GetWindowOrgEx(UserControl.hdc, ptOrigin)
  
  ' calculate the amount of scroll change (ScrollWindowEx and
  ' OffsetWindowOrgEx accept relative, and not absolute, coords).
  If fResetOrigin Then
    xScroll = ptOrigin.x
    yScroll = ptOrigin.y
  Else
    si.fMask = SIF_POS
    Call GetScrollInfo(m_hwndUC, SB_HORZ, si)
    xScroll = ptOrigin.x - si.nPos
    Call GetScrollInfo(m_hwndUC, SB_VERT, si)
    yScroll = ptOrigin.y - si.nPos
  End If
  
  If (xScroll Or yScroll) Then
    ' changes the caret position with respect to the window origin by the offset amount
    Call ScrollWindowEx(m_hwndUC, xScroll, yScroll, ByVal 0&, ByVal 0&, 0, ByVal 0&, SW_INVALIDATE Or SW_SCROLLCHILDREN)
    Call OffsetWindowOrgEx(UserControl.hdc, -xScroll, -yScroll, ByVal 0&)
  End If   ' (xScroll Or yScroll)
  
  Call UpdateWindow(m_hwndUC)
  
#If PROCLOG Then
  Debug.Print "SetScrollbars end"
#End If

End Sub

' The UserControl doesn't get WM_KEYDOWN and WM_KEYUP window messages
' for accelerators, have to set CanGetFocus and KeyPreview to True, get them here,
' and forward them along to SVWndProc where scrolling is handled.

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  
#If PROCLOG Then
  Debug.Print "UserControl_KeyDown"
#End If

  ' adjust the scroll bars only when Ctrl key is pressed (there is no
  ' Home and End horizontal scrolling functionality, only vertical)
  If (Shift And vbCtrlMask) Then
    Select Case KeyCode
      Case vbKeyPageUp    ' = 33
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_PAGEUP, 0)
      Case vbKeyPageDown  ' 34
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_PAGEDOWN, 0)
      Case vbKeyEnd           ' =35
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_BOTTOM, 0)
      Case vbKeyHome        ' = 36
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_TOP, 0)
      Case vbKeyLeft           ' = 37
        Call UCWndProc(0, m_hwndUC, WM_HSCROLL, SB_LINELEFT, 0)
      Case vbKeyUp            ' = 38
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_LINEUP, 0)
      Case vbKeyRight         ' = 39
        Call UCWndProc(0, m_hwndUC, WM_HSCROLL, SB_LINERIGHT, 0)
      Case vbKeyDown        ' = 40
        Call UCWndProc(0, m_hwndUC, WM_VSCROLL, SB_LINEDOWN, 0)
    End Select
  End If   ' (Shift And vbCtrlMask)
  
End Sub

Friend Function UCWndProc(ByVal lpfnOld As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim pt As POINTAPI
  Dim ptOrigin As POINTAPI
  Dim si As SCROLLINFO
  Dim x As Long
  Dim y As Long
  
#If PROCLOG Then
  Debug.Print "UCWndProc &H" & Hex(uMsg) & ", &H" & Hex(wParam) & ", &H" & Hex(lParam)
#End If

  Select Case uMsg
    
    ' ============================================================
    ' Unsubclass the UserControl in designtime right before is shown.
    ' Called only once on the UC's first WM_SHOWWINDOW, where the
    ' UserControl's size is finally set by the client (for propbag read at init)

    ' fShow = (BOOL) wParam;      // show/hide flag
    ' fnStatus = (int) lParam;    // status flag
    
    Case WM_SHOWWINDOW
        
      If wParam Then
        ' CallWindowProc first and exit, UnSubClass removes lpfnOld
        UCWndProc = CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
        If (m_fShown = False) Then
          m_fShown = True
          If m_fDesigntime Then Call UnSubClass(hWnd)
          Call SetScrollbars(True)   ' fResetOrigin
        End If
        Exit Function
      End If
      
    ' ============================================================
    ' used at runtime instead of UserControl_Resize. m_fSizing is set below, and
    ' in SetPictureView
    
    ' fwSizeType = wParam;      // resizing flag
    ' nWidth = LOWORD(lParam);  // width of client area
    ' nHeight = HIWORD(lParam); // height of client area
    
    Case WM_SIZE
  
      UCWndProc = CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
      
      If m_dwView And (m_fSizing = False) Then
        m_fSizing = True
        ' calls SetWindowStyle which calls SetWindowPos, invoking WM_SIZE
        Call SetScrollbars(False)   ' no fResetOrigin
        m_fSizing = False
      End If

'Debug.Print "WM_SIZE end"
      Exit Function
          
    ' ============================================================
    ' Adjust the view's horizontal scrollbar and view position
    
    ' nScrollCode = (int) LOWORD(wParam);  // scroll bar value
    ' nPos = (short int) HIWORD(wParam);   // scroll box position
    ' hwndScrollBar = (HWND) lParam;       // handle of scroll bar

    Case WM_HSCROLL
      
      si.cbSize = Len(si)
      si.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
      Call GetScrollInfo(hWnd, SB_HORZ, si)
      x = 0
      
      Select Case (wParam And &HFFFF&)    ' LOWORD(wParam)
        Case SB_LINELEFT
          x = Max(si.nMin, si.nPos - m_cxChar)
        Case SB_LINERIGHT
          x = Min((si.nMax - si.nPage) + 1, si.nPos + m_cxChar)
        Case SB_PAGELEFT
          x = Max(si.nMin, si.nPos - si.nPage)
        Case SB_PAGERIGHT
          x = Min((si.nMax - si.nPage) + 1, si.nPos + si.nPage)
        Case SB_LEFT
          x = si.nMin
        Case SB_RIGHT
          x = (si.nMax - si.nPage) + 1
        Case SB_THUMBTRACK
          x = si.nTrackPos    ' HIWORD(wParam)
        Case SB_ENDSCROLL, SB_THUMBPOSITION
          Exit Function
      End Select
      
      If (x <> si.nPos) Then
        Call ScrollWindowEx(hWnd, -(x - si.nPos), 0, ByVal 0&, ByVal 0&, 0, ByVal 0&, SW_INVALIDATE Or SW_SCROLLCHILDREN)
        Call OffsetWindowOrgEx(UserControl.hdc, x - si.nPos, 0, ByVal 0&)
        
        si.fMask = SIF_POS
        si.nPos = x
        Call SetScrollInfo(hWnd, SB_HORZ, si, 1)
      
        Call UpdateWindow(hWnd)
      End If
      
      ' The UserControl doesn't know how to process this message, and will
      ' die on 95/98 if x > 32K (&H8000&)
      Exit Function
      
    ' ============================================================
    ' Adjust the view's vertical scrollbar and view position
    
    Case WM_VSCROLL
      
      si.cbSize = Len(si)
      si.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
      Call GetScrollInfo(hWnd, SB_VERT, si)
      y = 0
      
      Select Case (wParam And &HFFFF&)    ' LOWORD(wParam)
        Case SB_LINEUP
          y = Max(si.nMin, si.nPos - m_cyChar)
        Case SB_LINEDOWN
          y = Min((si.nMax - si.nPage) + 1, si.nPos + m_cyChar)
        Case SB_PAGEUP
          y = Max(si.nMin, si.nPos - si.nPage)
        Case SB_PAGEDOWN
          y = Min((si.nMax - si.nPage) + 1, si.nPos + si.nPage)
        Case SB_TOP
          y = si.nMin
        Case SB_BOTTOM
          y = (si.nMax - si.nPage) + 1
        Case SB_THUMBTRACK
          y = si.nTrackPos   ' HIWORD(wParam)
        Case SB_ENDSCROLL, SB_THUMBPOSITION
          Exit Function
      End Select

      If (y <> si.nPos) Then
        Call ScrollWindowEx(hWnd, 0, -(y - si.nPos), ByVal 0&, ByVal 0&, 0, ByVal 0&, SW_INVALIDATE Or SW_SCROLLCHILDREN)
        Call OffsetWindowOrgEx(UserControl.hdc, 0, y - si.nPos, ByVal 0&)
        
        si.fMask = SIF_POS
        si.nPos = y
        Call SetScrollInfo(hWnd, SB_VERT, si, 1)
        
        Call UpdateWindow(hWnd)
      End If
      
      Exit Function
    
    ' ============================================================
    ' Unsubclass the window.
    
    Case WM_DESTROY
      ' lpfnOld will be gone after UnSubClass is called!
      Call CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  UCWndProc = CallWindowProc(lpfnOld, hWnd, uMsg, wParam, lParam)
  
End Function

' ============================================================================
' private helper functions

' Sets the specified style bits for the specified window.

'   hWnd           - specified window's handle
'   dwNewStyle - style to set
'   fExtended     - flag indicating whether the style is an extended style,
'                          if True, the style is extended, otherwise the style is normal
'   fAdd              - flag indicating whether the style is added or removed,
'                          if True, the style is added, otherwise the style is removed
'   fRedraw        - flag indicating whether to redraw the window so it detects the style change,
'                          if True, the window is redrawn, otherwise the window is not redrawn

' Returns True if the specified style is successfully added or removed,
' or is already set as specified. Returns False otherwise.

Private Function SetWindowStyle(hWnd As Long, _
                                                     dwNewStyle As Long, _
                                                     fExtended As Boolean, _
                                                     fAdd As Boolean, _
                                                     Optional fRedraw As Boolean = False) As Boolean
  Dim dwStyleType As Long
  Dim dwCurStyle As Long
  
  If fExtended Then
    dwStyleType = GWL_EXSTYLE
  Else
    dwStyleType = GWL_STYLE
  End If
  
  dwCurStyle = GetWindowLong(hWnd, dwStyleType)   ' sets GetLastError
  ' Make sure everything went OK...
  If (Err.LastDllError = 0) Then ' dwCurStyle could be zero!!!
    
    ' If adding the new style and is not currently set, set it.
    If fAdd And ((dwCurStyle And dwNewStyle) = 0) Then
      Call SetWindowLong(hWnd, dwStyleType, dwCurStyle Or dwNewStyle)
    
    ' If removing the new style and is currently set, clear it.
    ElseIf (fAdd = False) And (dwCurStyle And dwNewStyle) Then
      Call SetWindowLong(hWnd, dwStyleType, dwCurStyle And (Not dwNewStyle))
    End If
    
    ' Don't test equality, for some reason it returns False...???
    If fAdd Then
      SetWindowStyle = (dwNewStyle And GetWindowLong(hWnd, dwStyleType))
    Else
      SetWindowStyle = (dwNewStyle And (Not GetWindowLong(hWnd, dwStyleType)))
    End If
    
    If fRedraw Then Call SetWindowPos(hWnd, 0, _
                                                              0, 0, 0, 0, _
                                                              SWP_NOZORDER Or _
                                                              SWP_NOMOVE Or _
                                                              SWP_NOSIZE Or _
                                                              SWP_DRAWFRAME)
  End If   ' (Err.LastDllError = 0)
  
End Function

' Returns the larger of the two passed params

Private Function Max(param1 As Long, param2 As Long) As Long
  If (param1 > param2) Then Max = param1 Else Max = param2
End Function

' Returns the smaller of the two passed params

Private Function Min(param1 As Long, param2 As Long) As Long
  If (param1 < param2) Then Min = param1 Else Min = param2
End Function

Private Function OleColorToColorRef(oleclr As Long) As Long
  If (oleclr And &H80000000) Then
    OleColorToColorRef = GetSysColor(oleclr And &HFF&)
  Else
    OleColorToColorRef = oleclr
  End If
End Function
'
'' Returns the low 16-bit integer from a 32-bit long integer
'
'Private Function LOWORD(dwValue As Long) As Integer
'  MoveMemory LOWORD, dwValue, 2
'End Function
'
'' Returns the low 16-bit integer from a 32-bit long integer
'
'Private Function HIWORD(dwValue As Long) As Integer
'  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
'End Function
