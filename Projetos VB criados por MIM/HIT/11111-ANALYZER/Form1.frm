VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "xTracty - Extrator de Resources"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Salvar como RES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3938
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tocar se Avi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7328
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salvar como Binário"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2243
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin xTracty.ScrollView ScrollView1 
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4260
      BackColor       =   15264511
      ForeColor       =   192
      View            =   1
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E8EAFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   2655
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tocar se Wav"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5633
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Carregar Executável"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFADE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3450
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   840
      Width           =   9135
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Salvar como Bitmap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   548
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3915
      ScaleHeight     =   975
      ScaleWidth      =   3930
      TabIndex        =   0
      Top             =   -435
      Visible         =   0   'False
      Width           =   3930
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   165
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "xTracty - Extrator de Resources"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   690
         TabIndex        =   2
         Top             =   360
         Width           =   3105
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0099D8F7&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nome da Resource"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0069C4FA&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo da Resource"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const RSFilter = "Arquivo de Resource (*.res)" & vbNullChar & "*.res"
Const sFilter = "Bitmap (*.bmp)" & vbNullChar & "*.bmp"
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PlaySound_Res Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszname As Long, ByVal hmodule As Long, ByVal dwFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LB_SETTABSTOPS = &H192
Dim sFile As String
Dim sPath As String
Private BM1 As New BMP

Private Sub Command1_Click()
aa = GetSaveFilePath(hWnd, sFilter, 0, sFilter, "", "", "Salvar BITMAP", sPath)
If aa = False Then Exit Sub
DoEvents
If Dir(sPath) <> "" Then Kill sPath
SavePicture Picture1.Image, sPath
End Sub

Private Sub Command2_Click()

aa = GetOpenFilePath(hWnd, "", 0, sFile, "", "Carregar OCX,DLL,EXE...", sPath)
If aa = False Then Exit Sub
Unload Form2
Unload Form3
Set Picture1.Picture = Nothing
FreeLibrary HMOD
DoEvents
'HMOD = LoadLibraryEx(sPath, 0&, 2)
HMOD = LoadLibrary(sPath)
If Not CBool(HMOD) Then Exit Sub
List1.Clear
ClearCOLLECTION
Call EnumResourceTypes(HMOD, AddressOf EnumRSType, 0)
Dim ccnt As Long
ccnt = RESNAME.Count
If ccnt = 0 Then Exit Sub
For u = 1 To ccnt
List1.AddItem RESTYPENAME.Item(u) & vbTab & RESNAME.Item(u)
Next u

End Sub

Private Sub Command3_Click()
On Error Resume Next
PlaySound_Res ByVal VarPtr(OtherData(0)), 0, &H4 Or &H1
If Err <> 0 Then On Error GoTo 0
End Sub

Private Sub Command4_Click()
aa = GetSaveFilePath(hWnd, "", 0, "", "", "", "Salvar Binário", sPath)
If aa = False Then Exit Sub
DoEvents
If Dir(sPath) <> "" Then Kill sPath
Open sPath For Binary As #1
Put #1, , OtherData
Close #1
End Sub

Private Sub Command5_Click()
On Error GoTo dalje:
Open App.Path & "\temp.temp.avi" For Binary As #1
Put #1, , OtherData
Close #1
Call UpdateWindow(ScrollView1.hWnd)
ScrollView1.View = vtControl

Picture1.Visible = True
Picture1.Height = 2000
Picture1.Width = 2000
DoEvents
PlayAVIPictureBox App.Path & "\temp.temp.avi", Picture1
Kill App.Path & "\temp.temp.avi"
Set Picture1.Picture = Nothing
Exit Sub
dalje:
On Error GoTo 0
Close #1
End Sub



Private Sub Command6_Click()
aa = GetSaveFilePath(hWnd, RSFilter, 0, RSFilter, "", "", "Salvar como Arquivo de Resource", sPath)
If aa = False Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath
DoEvents

Dim ResHedLen As Long 'Resource Header length
ResHedLen = 24

Dim nameQ As Boolean
Dim typeQ As Boolean

If (TrueName < 0) Or (TrueName > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(TrueBuffer(0))) + 1) * 2
nameQ = True
Else
ResHedLen = ResHedLen + 4
End If

If (TypePtr < 0) Or (TypePtr > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(TrueType(0))) + 1) * 2
typeQ = True
Else
ResHedLen = ResHedLen + 4
End If


Open sPath For Binary As #1
'PRE-HEADER
Put #1, , CLng(0)
Put #1, , CLng(&H20)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
'END OF PRE-HEADER
Put #1, , ResTotLen
Put #1, , ResHedLen
If typeQ Then
Dim UNI1 As String
UNI1 = StrConv(TrueType, vbUnicode)
UNI1 = StrConv(UNI1, vbUnicode)
Put #1, , UNI1
Else
Put #1, , CInt(&HFFFF)
Put #1, , CInt(TypePtr)
End If
If nameQ Then
Dim UNI2 As String
UNI2 = StrConv(TrueBuffer, vbUnicode)
UNI2 = StrConv(UNI2, vbUnicode)
Put #1, , UNI2
Else
Put #1, , CInt(&HFFFF)
Put #1, , CInt(TrueName)
End If
Put #1, , CLng(0) 'Data Version
Put #1, , CInt(&H1030) 'Memory Flag
Put #1, , LangID
Put #1, , CLng(0) 'Version
Put #1, , CLng(0) 'Characteristic
Put #1, , OtherData 'Put Memory Data
Close #1
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
ScrollView1.ShowDisabledScrollbars = dstBoth
Top = 0
Left = 0
Height = Screen.Height
Width = Screen.Width
List1.Width = ScaleWidth
Label2(1).Width = ScaleWidth - Label2(1).Left
Dim tabs(1) As Long
tabs(0) = 200
tabs(1) = 150
SendMessage List1.hWnd, LB_SETTABSTOPS, 1, tabs(0)
Erase tabs
InsertCtrl Picture2.hWnd, hWnd, 10, 215
SetWidth = 332
End Sub

Private Sub Form_Paint()
Static RPT As Boolean
If Not RPT Then
RPT = True
ScrollView1.Height = ScaleHeight - ScrollView1.Top
ScrollView1.Width = ScaleWidth - ScrollView1.Left
Call UpdateWindow(ScrollView1.hWnd)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
ClearCOLLECTION
FreeLibrary HMOD
Unload Form2
Unload Form3
Erase TrueType
Erase TrueBuffer
Erase OtherData
End Sub
Private Sub List1_Click()

Unload Form2
Unload Form3
Set Picture1.Picture = Nothing
Erase TrueType
Erase TrueBuffer
Erase OtherData



Dim CASSE As Variant
CASSE = RESTYPE.Item(List1.ListIndex + 1)

Dim TypeNnm As String
TypeNnm = CStr(CASSE) & Chr(CByte(0))
If Not IsNumeric(CASSE) Then
TrueType = StrConv(TypeNnm, vbFromUnicode)
TypePtr = VarPtr(TrueType(0))
Else
TypePtr = CLng(CASSE)
End If

Dim nnm As String
nnm = RESNAME.Item(List1.ListIndex + 1)
If IsNumeric(nnm) Then
TrueName = CLng(nnm)
Else
nnm = nnm & Chr(CByte(0))
TrueBuffer = StrConv(nnm, vbFromUnicode)
TrueName = VarPtr(TrueBuffer(0))
End If

Call EnumResourceLanguages(HMOD, TypePtr, TrueName, AddressOf EnumRSLang, 0)

If CASSE = "2" Then
BM1.InitBITMAP
BM1.GetBitmap HMOD, TrueName
If BM1.BitmapHeight = 0 Or BM1.BitmapWidth = 0 Then Exit Sub
Picture1.Cls
Picture1.Width = BM1.BitmapWidth * 15
Picture1.Height = BM1.BitmapHeight * 15
Call BitBlt(Picture1.hdc, 0, 0, BM1.BitmapWidth, BM1.BitmapHeight, BM1.hdc, 0, 0, &HCC0020)
Picture1.Refresh
Call UpdateWindow(ScrollView1.hWnd)
ScrollView1.View = vtControl
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
LoadIntoMemory TrueName, TypePtr
BM1.KillBITMAP


ElseIf CASSE = "14" Then
Picture1.Cls
Dim rIc As Long
rIc = LoadIcon(HMOD, TrueName)
Picture1.Width = 32 * 16
Picture1.Height = 32 * 16
DrawIcon Picture1.hdc, 0, 0, rIc
Picture1.Refresh
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False



ElseIf CASSE = "12" Then
Picture1.Cls
Dim crsr As Long
crsr = LoadCursor(HMOD, TrueName)
Picture1.Width = 32 * 16
Picture1.Height = 32 * 16
DrawIcon Picture1.hdc, 0, 0, crsr
Picture1.Refresh
Command1.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
LoadIntoMemory TrueName, TypePtr


ElseIf CASSE = "5" Then
SetWinPosByCursor Form2.hWnd, 0
HHDL = CreateDialogParam(HMOD, TrueName, Form2.hWnd, AddressOf dialogProc, 0&)
If Not CBool(HHDL) Then Exit Sub
Form2.Show
Dim RCT1 As RECT
Dim RCT2 As RECT
Dim PT1 As POINTAPI
GetWindowRect HHDL, RCT1
GetWindowRect Form2.hWnd, RCT2
PT1.x = RCT2.Left
PT1.y = RCT2.Top
ScreenToClient Form2.hWnd, PT1
SetParent HHDL, Form2.hWnd
Dim MTY As Long
MTY = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER) + GetSystemMetrics(SM_CYDLGFRAME)
Dim MTX As Long
MTX = GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXDLGFRAME)
SetWindowPos HHDL, 0, PT1.x + MTX, PT1.y + MTY, 0, 0, 1
Form2.Width = (RCT1.Right - RCT1.Left) * 15 + (MTX * 2) * 15
Form2.Height = (RCT1.Bottom - RCT1.Top) * 15 + (MTY + GetSystemMetrics(SM_CYDLGFRAME)) * 15
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
LoadIntoMemory TrueName, TypePtr


ElseIf CASSE = "4" Then
MHDL = LoadMenu(HMOD, TrueName)
If Not CBool(MHDL) Then Exit Sub
Dim USTRX As String
Dim chrllen As Long
Dim MnuCNT As Long
Dim IID As Long
MnuCNT = GetMenuItemCount(MHDL)
For u = 0 To MnuCNT - 1
USTRX = Space(255)
chrllen = GetMenuString(MHDL, u, USTRX, 255, MF_BYPOSITION)
USTRX = Left(USTRX, chrllen)
If USTRX = "" Then
IID = GetMenuItemID(MHDL, u)
Dim modd() As Byte
modd = StrConv("Hiden PopUp" & Chr(CByte(0)), vbFromUnicode)
Call ModifyMenu(MHDL, u, 0 Or MF_BYPOSITION, IID, VarPtr(modd(0)))
End If
Next u
SetWinPosByCursor Form2.hWnd, 0
Form2.Width = 5000
Form2.Height = 900
Form2.Show
Call SetMenu(Form2.hWnd, MHDL)
DrawMenuBar Form2.hWnd
LoadIntoMemory TrueName, TypePtr
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True


ElseIf Not IsNumeric(CASSE) Or RESTYPENAME.Item(List1.ListIndex + 1) = "Custom Defined" Then
LoadIntoMemory TrueName, TypePtr
Form3.Text1 = StrConv(OtherData, vbUnicode)
Form3.Label1 = "Tamanho da Rsource: " & ResTotLen & " Bytes "
Form3.Show
Command1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Else
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False

End If



End Sub





