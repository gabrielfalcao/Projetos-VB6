Attribute VB_Name = "modImprimir"
Private Type Rect
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type CharRange
cpMin As Long
cpMax As Long
End Type
Private Type FormatRange
hdc As Long
hdcTarget As Long
rc As Rect
rcPage As Rect
chrg As CharRange
End Type
Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpout As Long, ByVal lpInitData As Long) As String

Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
Dim LeftOffset As Long, TopOffset As Long
Dim LeftMargin As Long, TopMargin As Long
Dim RightMargin As Long, BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawTo As Rect
Dim rcPage As Rect
Dim TextLength As Long
Dim NextCharPosition As Long
Dim r As Long
' Start a print job to get a valid Printer.hDC
Printer.Print Space(1)
Printer.ScaleMode = vbTwips
' Get the offsett to the printable area on the page in twips
LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)
' Calculate the Left, Top, Right, and Bottom margins
LeftMargin = LeftMarginWidth - LeftOffset
TopMargin = TopMarginHeight - TopOffset
RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
' Set printable area rect
rcPage.Left = 0
rcPage.Top = 0
rcPage.Right = Printer.ScaleWidth
rcPage.Bottom = Printer.ScaleHeight
' Set rect in which to print (relative to printable area)
rcDrawTo.Left = LeftMargin
rcDrawTo.Top = TopMargin
rcDrawTo.Right = RightMargin
rcDrawTo.Bottom = BottomMargin
' Set up the print instructions
fr.hdc = Printer.hdc ' Use the same DC for measuring and rendering
fr.hdcTarget = Printer.hdc ' Point at printer hDC
fr.rc = rcDrawTo ' Indicate the area on page to draw to
fr.rcPage = rcPage ' Indicate entire size of page
fr.chrg.cpMin = 0 ' Indicate start of text through
fr.chrg.cpMax = -1 ' end of the text
' Get length of text in RTF
TextLength = Len(RTF.Text)
' Loop printing each page until done
Do
' Print the page by sending EM_FORMATRANGE message
NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
If NextCharPosition >= TextLength Then
Exit Do 'If done then exit
End If
fr.chrg.cpMin = NextCharPosition ' Starting position for next page
Printer.NewPage ' Move on to next page
Printer.Print Space(1) ' Re-initialize hDC
fr.hdc = Printer.hdc
fr.hdcTarget = Printer.hdc
Loop
' Commit the print job
Printer.EndDoc ' Allow the RTF to free up memory
r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub

