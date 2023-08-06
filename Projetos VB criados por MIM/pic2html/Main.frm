VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "G@b$ PIC2HTML"
   ClientHeight    =   6510
   ClientLeft      =   1365
   ClientTop       =   2535
   ClientWidth     =   9045
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   Begin VB.CheckBox chkShowPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      Caption         =   "Mostrar no Browser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   195
      Left            =   4867
      TabIndex        =   16
      Top             =   555
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkShowCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      Caption         =   "Mostrar código HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   195
      Left            =   4867
      TabIndex        =   15
      Top             =   345
      Width           =   3015
   End
   Begin VB.PictureBox picPercent2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7D9B0&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   449
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   14
      Top             =   1695
      Width           =   8130
   End
   Begin VB.PictureBox picPercent3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7D9B0&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   449
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   13
      Top             =   1965
      Width           =   8130
   End
   Begin VB.TextBox txtPath2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4237
      TabIndex        =   12
      Top             =   5595
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox picPercent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7D9B0&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   449
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   11
      Top             =   1425
      Width           =   8130
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Carregar Código HTML"
      Height          =   375
      Left            =   6007
      TabIndex        =   10
      Top             =   5550
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      Caption         =   "Centralizar imagem na página"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   195
      Left            =   4867
      TabIndex        =   9
      Top             =   135
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salvar Código HTML"
      Height          =   375
      Left            =   472
      TabIndex        =   7
      Top             =   5550
      Width           =   1785
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Copiar Código HTML"
      Height          =   375
      Left            =   2287
      TabIndex        =   6
      Top             =   5550
      Width           =   1785
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3712
      TabIndex        =   5
      Top             =   495
      Width           =   360
   End
   Begin RichTextLib.RichTextBox txtHTMLSource 
      Height          =   3150
      Left            =   82
      TabIndex        =   4
      Top             =   2295
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5556
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Main.frx":628A
   End
   Begin VB.PictureBox picPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   7500
      Picture         =   "Main.frx":630C
      ScaleHeight     =   0
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   3
      Top             =   6915
      Width           =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00A4E3AC&
      Caption         =   "&Gerar Código!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   930
      Width           =   1935
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   97
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   3570
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   225
      Left            =   2977
      TabIndex        =   8
      Top             =   6090
      Width           =   2895
   End
   Begin VB.Image imgPic 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   10650
      Picture         =   "Main.frx":6F50
      Top             =   7245
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha um arquivo de imagem para transformar em HTML:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   555
      Left            =   112
      TabIndex        =   0
      Top             =   75
      Width           =   3120
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Common Dialog stuff:
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpopenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFiletitle As String
    nMaxfileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'for shelling:
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

'for getting pixel color:
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'For timeout:
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Timeout(dblSeconds As Double)
Dim hCurrent As Long
    
dblSeconds = dblSeconds * 1000
hCurrent = GetTickCount()

Do While GetTickCount() - hCurrent < dblSeconds
    DoEvents
Loop
End Sub

Public Sub SetPercent(Picture As Control, ByVal Percent)
Dim num As String

Picture.Cls
Picture.ScaleHeight = 100
Picture.ScaleWidth = 100
num = Format$(Percent, "###") + "%"
num = Val(num) & "%"

'    Picture.ScaleWidth = 100
num$ = Format$(Percent, "###") + "%"
 Picture.CurrentX = 50 - Picture.TextWidth(num) / 2
 Picture.CurrentY = (Picture.ScaleHeight - Picture.TextHeight(num)) / 2
 Picture.Print num
Picture.Line (0, 0)-(Percent, Picture.ScaleHeight), , BF

Picture.Refresh
Timeout 0.001 'pause to display it
End Sub

Private Function GetHex(intLongValue As Long) As String
Dim intBlue As Long
Dim intGreen As Long
Dim intRed As Long
Dim strBlue As String 'hex
Dim strRed As String 'hex
Dim strGreen As String 'hex

    If intLongValue >= 65536 Then
        intBlue = Int(intLongValue / 65536)
        intLongValue = intLongValue - (65536 * intBlue)
    End If

    If intLongValue >= 256 Then
        intGreen = Int(intLongValue / 256)
        intLongValue = intLongValue - (256 * intGreen)
    End If
    
    intRed = intLongValue

    strBlue = Hex(intBlue)
    strRed = Hex(intRed)
    strGreen = Hex(intGreen)
    
    If Len(strBlue) < 2 Then strBlue = "0" & strBlue
    If Len(strRed) < 2 Then strRed = "0" & strRed
    If Len(strGreen) < 2 Then strGreen = "0" & strGreen

    GetHex = strBlue & strRed & strGreen

End Function


Public Function LaunchSite(strUrl As String) As Long

    Dim lhWnd As Long
    Dim lAns As Long
    ' Execute the url
    lAns = ShellExecute(lhWnd, "open", strUrl, vbNullString, vbNullString, 3) '3=MAXIMIZE WINDOW
   
    OpenLocation = lAns ' return returnval

End Function


Private Sub SaveAndLoadHTML(ByRef strHTMLCode)
'This will save the textbox to a file and launch it in the browser
Dim strText As String
Dim intFreeFile As Byte
Dim intTimes As Currency
Dim i As Long 'allow huge files


Open "C:\TestHTML.html" For Output As #11
Close #11

If Len(strHTMLCode) >= 50000 Then
    intTimes = Len(strHTMLCode) / 50000
    
    For i = 1 To intTimes
        intFreeFile = FreeFile
        Open "C:\TestHTML.html" For Binary As intFreeFile
            strText = Mid$(strHTMLCode, ((i - 1) * 50000) + 1, 50000)
            Put #intFreeFile, LOF(intFreeFile) + 1, strText '((i - 1) * 50000) + 1, strText
        Close #intFreeFile
        Call SetPercent(picPercent3, intMod / i)
    Next 'i
End If

If Len(strHTMLCode) Mod 50000 <> 0 Then 'finish it
    strText = Right$(strHTMLCode, Len(strHTMLCode) Mod 50000)
    intFreeFile = FreeFile
    Open "C:\TestHTML.html" For Binary As intFreeFile
        Put #intFreeFile, LOF(intFreeFile) + 1, strText
    Close #intFreeFile
End If
Call SetPercent(picPercent3, 100)


Call LaunchSite("C:\TestHTML.html")
End Sub

Private Sub Command1_Click()
Dim strHTMLCode As String
Dim strHexColor As String 'Hex code for the current pixel
Dim intPicWidth As Integer
Dim intPicHeight As Integer
Dim intColor As Long
Dim intLastColor As Long
Dim intNextColor As Long
Dim picHDC As Long
Dim intRowNumber As Integer
Dim intStringLength As Currency
Dim intCurrent As Currency
Dim intAdd As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False

picHDC = picPic.hDC
intPicWidth = picPic.Width
intPicHeight = picPic.Height
intRowNumber = 1

'BEGIN Algorithm...


'First dive through the algorithm, get the string length:
For i = 0 To intPicHeight - 1 'vertical rows
    DoEvents
    intStringLength = intStringLength + 13
'    strHTMLCode = strHTMLCode & "<TR HEIGHT=1>" '1 pixel height
    For j = 0 To intPicWidth - 1 'horizontal rows
        DoEvents
        intColor = GetPixel(picHDC, j, i)
        If j < intPicWidth - 1 Then 'Not the last pixel
            intNextColor = GetPixel(picHDC, j + 1, i)
            If intColor = intNextColor Then 'Check total amount of repeating colors
                intRowNumber = 2 'there's definitely 2, check for more:
                For x = j + 2 To intPicWidth - 1 'check remaining row
                    intNextColor = GetPixel(picHDC, x, i)
 '                   MsgBox x & "," & i & "=" & intNextColor, 64, "X LOOP"
                    If intColor = intNextColor Then 'same again
                        If x <> intPicWidth - 1 Then
                            intRowNumber = intRowNumber + 1
                        Else
                            intRowNumber = intRowNumber + 1
                            j = intPicWidth - 1
                        End If
                    Else 'not the same
                        j = x - 1
                        Exit For
                    End If
                Next 'x
            End If
        End If

    'MsgBox intRowNumber, 64, j & ",i:" & i
        If intRowNumber = 1 Then
            strHexColor = Hex(intColor)
            intAdd = Len(strHexColor) + 1
            intStringLength = intStringLength + 27 + intAdd
'            strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """></TD>"
        Else
            strHexColor = Hex(intColor)
            intAdd = Len(Format(intRowNumber)) + Len(strHexColor)
            intStringLength = intStringLength + 38 + intAdd
'            strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """ COLSPAN=" & intRowNumber & "></TD>"
            intRowNumber = 1
        End If
    Next 'j
    intStringLength = intStringLength + 5
'    strHTMLCode = strHTMLCode & "</TR>"
    Call SetPercent(picPercent, (i / (intPicHeight - 1)) * 100)
Next 'i


strHTMLCode = Space(intStringLength) 'build buffer

intCurrent = 1


'Now fill the buffer with delicious HTML... MMM... html.
For i = 0 To intPicHeight - 1 'vertical rows
    DoEvents
    
    Mid$(strHTMLCode, intCurrent, 13) = "<TR HEIGHT=1>"  '1 pixel height
    intCurrent = intCurrent + 13
    
    For j = 0 To intPicWidth - 1 'horizontal rows
        DoEvents
        intColor = GetPixel(picHDC, j, i)
        If j < intPicWidth - 1 Then 'Not the last pixel
            intNextColor = GetPixel(picHDC, j + 1, i)
            If intColor = intNextColor Then 'Check total amount of repeating colors
                intRowNumber = 2 'there's definitely 2, check for more:
                For x = j + 2 To intPicWidth - 1 'check remaining row
                    intNextColor = GetPixel(picHDC, x, i)
 '                   MsgBox x & "," & i & "=" & intNextColor, 64, "X LOOP"
                    If intColor = intNextColor Then 'same again
                        If x <> intPicWidth - 1 Then
                            intRowNumber = intRowNumber + 1
                        Else
                            intRowNumber = intRowNumber + 1
                            j = intPicWidth - 1
                        End If
                    Else 'not the same
                        j = x - 1
                        Exit For
                    End If
                Next 'x
            End If
        End If

    'MsgBox intRowNumber, 64, j & ",i:" & i
        If intRowNumber = 1 Then
            strHexColor = Hex(intColor)
            intAdd = Len(strHexColor) + 1
            Mid$(strHTMLCode, intCurrent, 27 + intAdd) = "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """></TD>"
            intCurrent = intCurrent + 27 + intAdd
'            strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """></TD>"
        Else
            strHexColor = Hex(intColor)
            intAdd = Len(strHexColor) + Len(Format(intRowNumber))
            Mid$(strHTMLCode, intCurrent, 38 + intAdd) = "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """ COLSPAN=" & intRowNumber & "></TD>"
            intCurrent = intCurrent + 38 + intAdd
            'strHTMLCode = strHTMLCode & "<TD WIDTH=" & intRowNumber & " BGCOLOR=""" & strHexColor & """ COLSPAN=" & intRowNumber & "></TD>"
            intRowNumber = 1
        End If
    Next 'j
    
    Mid$(strHTMLCode, intCurrent, 5) = "</TR>"
    intCurrent = intCurrent + 5
    '    strHTMLCode = strHTMLCode & "</TR>"
    Call SetPercent(picPercent2, (i / (intPicHeight - 1)) * 100)
Next 'i




strHTMLCode = "<TABLE WIDTH=" & intPicWidth & " HEIGHT=" & intPicHeight & " CELLSPACING=0 CELLPADDING=0 BORDER=0>" & strHTMLCode & "</TABLE>"
'END Algorithm.

If Check1.Value = 1 Then 'center it
    strHTMLCode = "<HTML><HEAD><TITLE>http://www.gabrielfalcao.i8.com</TITLE></HEAD><BODY BGCOLOR=""#FFFFFF""><CENTER>" & strHTMLCode & "</CENTER></BODY></HTML>"
Else
    strHTMLCode = "<HTML><HEAD><TITLE>http://www.gabrielfalcao.i8.com</TITLE></HEAD><BODY BGCOLOR=""#FFFFFF"">" & strHTMLCode & "</BODY></HTML>"
End If

If chkShowCode.Value = 1 Then
    txtHTMLSource.Text = strHTMLCode
End If


If chkShowPic.Value = 1 Then 'show pic in browser:
    Call SaveAndLoadHTML(strHTMLCode)
End If


Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next

If txtHTMLSource.Text = "" Then Exit Sub

Clipboard.SetText txtHTMLSource.Text

MsgBox "O código fonte foi colocado na área de transferência!", 64
End Sub

Private Sub Command3_Click()
Dim i As Byte
Dim j As Byte
Dim intFreeFile As Byte
Dim strText As String 'The entire file of params
Dim strParam As String 'Current Parameter
'...+
Dim rc As Long 'Return Data Holder
Dim cdFindFile As OPENFILENAME
Dim FileOpen As String
Const MAX_BUFFER_LENGTH = 256
On Error GoTo ErrHandler

'Set dialog's stuff:
cdFindFile.hwndOwner = hwnd
cdFindFile.hInstance = App.hInstance
cdFindFile.lpstrTitle = "Open Image..."
cdFindFile.lpstrInitialDir = CurDir() 'App.Path
cdFindFile.lpstrFilter = "Arquivos de IMAGEM (.bmp, .gif, .jpg)" & Chr$(0) & "*.BMP;*.JPG;*.GIF" & Chr$(0)
cdFindFile.nFilterIndex = 1
cdFindFile.flags = &H4 + &H800 'Don't show read only + Path must exist
cdFindFile.lpstrFile = String(MAX_BUFFER_LENGTH, Chr$(0))
cdFindFile.nMaxFile = MAX_BUFFER_LENGTH - 1
cdFindFile.lpstrFiletitle = cdFindFile.lpstrFile
cdFindFile.nMaxfileTitle = MAX_BUFFER_LENGTH - 1
cdFindFile.lStructSize = Len(cdFindFile)

rc = GetOpenFileName(cdFindFile)

If rc Then
    FileOpen = Left$(cdFindFile.lpstrFile, cdFindFile.nMaxFile)
    txtPath.Text = FileOpen
    FileOpen = txtPath.Text 'Remove nulls

    strSP3Path = FileOpen 'SET THE PROJECT PATH
Else 'Cancel was pressed
    FileOpen = ""
    Exit Sub
End If

imgPic.Picture = LoadPicture(FileOpen)
picPic.Width = imgPic.Width
picPic.Height = imgPic.Height
picPic.Picture = imgPic.Picture

If txtPath.Text <> Empty Then Command1.Enabled = True
Exit Sub
ErrHandler:
MsgBox "Não foi possível abrir a imagem por:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & " " & Err.Number & " - " & Err.Description, 16, "Error!"
End Sub

Private Sub Command4_Click()
Dim i As Byte
Dim j As Byte
Dim intFreeFile As Byte
Dim strText As String 'The entire file of params
Dim strParam As String 'Current Parameter
'...+
Dim rc As Long 'Return Data Holder
Dim cdFindFile As OPENFILENAME
Dim FileOpen As String
Const MAX_BUFFER_LENGTH = 256
On Error GoTo ErrHandler

If txtHTMLSource.Text = "" Then Exit Sub

'Set dialog's stuff:
cdFindFile.hwndOwner = hwnd
cdFindFile.hInstance = App.hInstance
cdFindFile.lpstrTitle = "Save HTML Source..."
cdFindFile.lpstrInitialDir = CurDir() 'App.Path
cdFindFile.lpstrFilter = "Arquivos HTML (*.htm, *.html)" & Chr$(0) & "*.htm;*.html" & Chr$(0)
cdFindFile.nFilterIndex = 1
cdFindFile.flags = &H4 + &H80000 + &H200000 + &H2 'Hide Read Only ;explorer ; OFN_LONGNAMES long names + overwrite prompt
cdFindFile.lpstrFile = String(MAX_BUFFER_LENGTH, Chr$(0))
cdFindFile.nMaxFile = MAX_BUFFER_LENGTH - 1
cdFindFile.lpstrFiletitle = cdFindFile.lpstrFile
cdFindFile.nMaxfileTitle = MAX_BUFFER_LENGTH - 1
cdFindFile.lStructSize = Len(cdFindFile)

rc = GetSaveFileName(cdFindFile)

If rc Then
    FileOpen = Left$(cdFindFile.lpstrFile, cdFindFile.nMaxFile)
    txtPath2.Text = FileOpen
    FileOpen = txtPath2.Text 'Remove nulls

Else 'Cancel was pressed
    FileOpen = ""
    Exit Sub
End If

If Right$(FileOpen, 4) <> ".htm" Or Right$(FileOpen, 5) <> ".html" Then
    FileOpen = FileOpen & ".html"
End If

Open FileOpen For Output As #11
Close #11

Open FileOpen For Binary As #130
    strText = txtHTMLSource.Text
    Put #130, 1, strText
Close #130

MsgBox "Salvo como:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & FileOpen, 64, "SALVO!"

Exit Sub
ErrHandler:
MsgBox "Não foi salvo por:" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & " " & Err.Number & " - " & Err.Description, 16, "Error!"
End Sub

Private Sub Command5_Click()
Dim strText As String
On Error Resume Next

Close

If txtHTMLSource.Text = "" Then Exit Sub

Open "C:\TestHTML.html" For Output As #11
Close #11

Open "C:\TestHTML.html" For Binary As #130
    strText = txtHTMLSource.Text
    Put #130, 1, strText
Close #130

Call LaunchSite("C:\TestHTML.html")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim p As String
If Len(App.Path) = 3 Then
p = App.Path
Else
p = App.Path & "\"
End If
Open p & "teste.html" For Output As #13
Close #13

End
End Sub

Private Sub Label2_Click()
Call LaunchSite("http://www.gabrielfalcao.i8.com")
End Sub
