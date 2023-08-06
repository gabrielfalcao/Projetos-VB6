VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBMPRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'»»»»»» API Declares ««««««
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const RGN_DIFF As Long = 4

Private m_hRgn As Long

Public Function CreateFromFile(ByVal lpFile As String) As Boolean
    On Error Resume Next
    Dim lx0 As Long, lxF As Long, Y As Long
    Dim Buffer As String
    
    Dim hFile As Long
    hFile = FreeFile
    Open lpFile For Binary Access Read As #hFile
        Buffer = Space$(LOF(hFile))
        Get #hFile, , Buffer
    Close #hFile
    
    Dim l1 As Long, l2 As Long, l3 As Long, l4 As Long
    l1 = InStr(1, Buffer, vbCrLf, vbBinaryCompare)
    l2 = InStr(1, Buffer, ",", vbBinaryCompare)
    lx0 = Left(Buffer, l2)
    lxF = Mid(Buffer, l2 + 1, l1 - l2 - 1)
    
    m_hRgn = CreateRectRgn(1, 1, lx0, lxF - 1)
    
    Dim hTRgn As Long
    
    l1 = l1 + 1
    Do While (l1 > 0)
        l3 = l1
        l2 = InStr(l1 + 1, Buffer, ",", vbBinaryCompare)
        l4 = InStr(l2 + 1, Buffer, "@", vbBinaryCompare)
        l1 = InStr(l4 + 1, Buffer, "|", vbBinaryCompare)
        lx0 = Mid(Buffer, l3 + 1, l2 - l3 - 1)
        lxF = Mid(Buffer, l2 + 1, l4 - l2 - 1)
        Y = Mid(Buffer, l4 + 1, l1 - l4 - 1)
        hTRgn = CreateRectRgn(lx0, Y, lxF, Y + 1)
        CombineRgn m_hRgn, m_hRgn, hTRgn, RGN_DIFF
        DeleteObject hTRgn
    Loop
    CreateFromFile = True
End Function
Public Function CreateFromPic(ByVal Pic As IPictureDisp, _
        ByVal crColor As OLE_COLOR) As Boolean
    On Error GoTo Err_Create
    Dim BM As BITMAP
    Dim X As Long, Y As Long, lStart As Long
    Dim hTRgn As Long
    Dim hDC As Long
    Dim hOld As Long
    
    Destroy
    
    hDC = CreateCompatibleDC(0)
    GetObject Pic.Handle, Len(BM), BM
    hOld = SelectObject(hDC, Pic.Handle)
    m_hRgn = CreateRectRgn(1, 1, BM.bmWidth, BM.bmHeight)
    
    For Y = 1 To BM.bmHeight
        For X = 1 To BM.bmWidth
            If GetPixel(hDC, X, Y) = crColor Then
                lStart = X
                Do
                    X = X + 1
                Loop Until ((GetPixel(hDC, X, Y) <> crColor) Or _
                        (X = BM.bmWidth))
                hTRgn = CreateRectRgn(lStart, Y, X, Y + 1)
                CombineRgn m_hRgn, m_hRgn, hTRgn, RGN_DIFF
                DeleteObject hTRgn
            End If
        Next
    Next
    
    CreateFromPic = True
    
Err_Create:
    SelectObject hDC, hOld
    DeleteDC hDC
End Function
Public Function Destroy() As Boolean
    If m_hRgn <> 0 Then
        DeleteObject m_hRgn
        m_hRgn = 0
    End If
End Function
Public Property Get Handle() As Long
    Handle = m_hRgn
End Property
Public Function CreateRegionFile(ByVal Pic As IPictureDisp, _
        ByVal crColor As OLE_COLOR, _
        ByVal lpPath As String) As Boolean
    On Error GoTo Err_Create
    Dim BM As BITMAP
    Dim X As Long, Y As Long, lStart As Long
    Dim hDC As Long
    Dim hOld As Long
    Dim Buffer As String
    
    hDC = CreateCompatibleDC(0)
    GetObject Pic.Handle, Len(BM), BM
    hOld = SelectObject(hDC, Pic.Handle)
    Buffer = BM.bmWidth & "," & BM.bmHeight & vbCrLf
    
    For Y = 1 To BM.bmHeight
        For X = 1 To BM.bmWidth
            If GetPixel(hDC, X, Y) = crColor Then
                lStart = X
                Do
                    X = X + 1
                Loop Until ((GetPixel(hDC, X, Y) <> crColor) Or _
                        (X = BM.bmWidth))
                Buffer = Buffer & lStart & "," & X & "@" & Y & "|"
            End If
        Next
    Next
    
    If Len(Buffer) > 0 Then
        Buffer = Left$(Buffer, Len(Buffer) - 1)
    End If
    
    On Error Resume Next
    Dim hFile As Long
    hFile = FreeFile
    Open lpPath For Binary Access Write As #hFile
        Put #hFile, , Buffer
    Close #hFile
    CreateRegionFile = (err.Number = 0)
Err_Create:
    SelectObject hDC, hOld
    DeleteDC hDC
End Function


