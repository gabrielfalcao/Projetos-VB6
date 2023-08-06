VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private BITMAP1 As BITMAP
Private XDC As Long
Public Sub InitBITMAP()
Dim DXH As Long
Dim DXDC As Long
DXH = GetDesktopWindow
DXDC = GetDC(DXH)
XDC = CreateCompatibleDC(DXDC)
Call ReleaseDC(DXH, DXDC)
End Sub
Public Sub KillBITMAP()
Call DeleteDC(XDC)
End Sub
Public Function GetBitmap(ByVal hmodule As Long, ByVal lpName As Long)
Dim tobj As Long
tobj = LoadBitmap(hmodule, lpName)
Call SelectObject(XDC, tobj)
Call GetObj(tobj, Len(BITMAP1), BITMAP1)
Call DeleteObject(tobj)
End Function
Public Property Get BitmapWidth() As Long
BitmapWidth = BITMAP1.bmWidth
End Property
Public Property Get BitmapHeight() As Long
BitmapHeight = BITMAP1.bmHeight
End Property
Public Property Get hdc() As Long
hdc = XDC
End Property

