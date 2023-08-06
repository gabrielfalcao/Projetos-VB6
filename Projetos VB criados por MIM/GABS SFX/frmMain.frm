VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TeraZip Self-Extractor"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pb1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEE7D6&
      FillColor       =   &H00990B0B&
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
      Height          =   270
      Left            =   53
      ScaleHeight     =   240
      ScaleWidth      =   4125
      TabIndex        =   3
      Top             =   825
      Width           =   4155
   End
   Begin VB.TextBox txtDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF9E3&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   405
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Extrair"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3060
      TabIndex        =   0
      Top             =   405
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As String
Dim strSalvaREM As String
Dim Path As String
Dim filelen As String
Dim p_ret As String
Dim dll1 As String
Dim dll2 As String
Dim dll3 As String
Dim dll4 As String
Dim dll5 As String
Dim dll6 As String
Dim dll7 As String
Dim ok As Boolean
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Private Sub Command2_Click()
On Error Resume Next
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  Dim O As New Collection
  szTitle = vbCr & vbCr & "Escolha o destino:"
  
  
  With tBrowseInfo
    .hwndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If sBuffer = Empty Then
    Exit Sub
    End If
    
    If Not Right$(sBuffer, 1) = "\" Then
    strSalvaREM = sBuffer & "\"
    txtDestino = sBuffer & "\"
    Else
    strSalvaREM = sBuffer
    txtDestino = sBuffer
  End If
  End If

pbar pb1, 25, True
Open txtDestino.Text & App.EXEName & ".zip" For Binary As #2
pbar pb1, 55, True
Put #2, , Right(file, filelen)
pbar pb1, 65, True
Close #2
Zipit1.FileName = txtDestino.Text & App.EXEName & ".zip"
 For i = 0 To Zipit1.ZipFiles.Count
  O.Add Zipit1.ZipFiles.item(i), i
  Next i
pbar pb1, 75, True
Zipit1.ExtrAction = zipUpdate
Zipit1.Read
pbar pb1, 85, True
Zipit1.ExtractDir = txtDestino.Text
Zipit1.Extract O
pbar pb1, 95, True
Kill (txtDestino.Text & App.EXEName & ".zip")
pbar pb1, 100, True
MsgBox "Extraído para " & txtDestino.Text & " com sucesso!"
End

End Sub

Private Sub Form_Activate()
If ok = True Then Command2.Enabled = True
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & "(" & App.EXEName & ")"
pbar pb1, 0, True
ok = False
If Len(App.Path) = 3 Then
Path = App.Path
Else
Path = App.Path & "\"
End If
dll3 = Path & "ZIPIT.DLL"
p_ret = StrConv(LoadResData("ZIPIT.DLL", "NEED"), vbUnicode)
Open dll3 For Binary As #1
Put #1, , p_ret
Close #1
dll5 = Path & "ZIPDLL.DLL"
p_ret = StrConv(LoadResData("ZIPDLL.DLL", "NEED"), vbUnicode)
Open dll5 For Binary As #1
Put #1, , p_ret
Close #1
dll6 = Path & "UNZDLL.DLL"
p_ret = StrConv(LoadResData("UNZDLL.DLL", "NEED"), vbUnicode)
Open dll6 For Binary As #1
Put #1, , p_ret
Close #1
dll7 = Path & "ZIP32.DLL"
p_ret = StrConv(LoadResData("ZIP32.DLL", "NEED"), vbUnicode)
Open dll7 For Binary As #1
Put #1, , p_ret
Close #1
ok = True
If Len(App.Path) > 3 Then
Path = App.Path & App.EXEName & ".exe"
Else
Path = App.Path & "\" & App.EXEName & ".exe"
End If
txtDestino = "C:\" & App.EXEName & ".zip"
Open Path For Binary Access Read As #1
file = Input(LOF(1), 1)
Close #1
filelen = Len(file) - Int(InStr(1, file, "|", vbBinaryCompare))
End Sub
Private Function InStrRevVB5(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1) As Long
    Dim lPos        As Long
    Dim lSavePos    As Long
    If Start = -1 Then Start = Len(StringCheck)
    lPos = InStr(1, StringCheck, StringMatch, vbBinaryCompare)
    While lPos > 0 And lPos < Start
        lSavePos = lPos
        lPos = InStr(lPos + 1, StringCheck, StringMatch, vbBinaryCompare)
    Wend
    InStrRevVB5 = lSavePos
End Function
Private Function pbar(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    If ShowPercent = True Then
    Num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
    pb.Print Num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function

