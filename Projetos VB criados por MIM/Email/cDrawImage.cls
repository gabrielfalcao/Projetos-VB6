VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Enum AxMODO
  [Lado a Lado] = 0
  [Redimensionada] = 1
End Enum

' Declarações de APIs
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Declaração de Constantes
Private Const SRCCOPY = &HCC0020

' Variáveis que serão usadas
Private m_hDC As Long
Private m_hWnd As Long
Private hBmpSrc As Long
Private m_Modo As Integer
Public Property Get hDC() As Long
  hDC = m_hDC
End Property
Public Property Let hDC(ByVal oNew_hDC As Long)
  m_hDC = oNew_hDC
End Property
Public Property Get hWnd() As Long
  hWnd = m_hWnd
End Property
Public Property Let hWnd(ByVal oNew_hWnd As Long)
  m_hWnd = oNew_hWnd
End Property
Public Property Get ImgHandle() As Long
  ImgHandle = hBmpSrc
End Property
Public Property Let ImgHandle(ByVal oNew_ImgHandle As Long)
  hBmpSrc = oNew_ImgHandle
End Property
Public Property Get ModoDeExibição() As AxMODO
  ModoDeExibição = m_Modo
End Property
Public Property Let ModoDeExibição(ByVal oNew_Modo As AxMODO)
  m_Modo = oNew_Modo
End Property
Public Sub TileBlt()
  Dim BMP As BITMAP
  Dim hDCSrc As Long
  Dim hDCDest As Long
  Dim hBmpTmp As Long
  Dim RC As Rect
  Dim Rows As Long
  Dim Cols As Long
  Dim dX As Long
  Dim dY As Long
  Dim i As Long, j As Long

  Call GetClientRect(m_hWnd, RC)
  hDCDest = m_hDC 'GetDC(m_hWnd)

  hDCSrc = CreateCompatibleDC(hDCDest)
  hBmpTmp = SelectObject(hDCSrc, hBmpSrc)

  Call GetObj(hBmpSrc, Len(BMP), BMP)
  Rows = RC.Right \ BMP.bmWidth
  Cols = RC.Bottom \ BMP.bmHeight

  If m_Modo = 0 Then
    For i = 0 To Rows
      dX = i * BMP.bmWidth
      For j = 0 To Cols
        dY = j * BMP.bmHeight
        Call BitBlt(hDCDest, dX, dY, BMP.bmWidth, BMP.bmHeight, hDCSrc, 0, 0, SRCCOPY)
      Next j
    Next i
  ElseIf m_Modo = 1 Then
    Call StretchBlt(hDCDest, 0, 0, RC.Right, RC.Bottom, hDCSrc, 0, 0, BMP.bmWidth, BMP.bmHeight, SRCCOPY)
  End If
  
  Call SelectObject(hDCSrc, hBmpTmp)
  Call DeleteDC(hDCSrc)
  Call ReleaseDC(m_hWnd, hDCDest)
End Sub
