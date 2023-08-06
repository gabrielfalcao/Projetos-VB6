VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Patch The Palace em Português"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   7020
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5550
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3255
      Width           =   1170
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   615
      ItemData        =   "frmMain.frx":E134C
      Left            =   11205
      List            =   "frmMain.frx":E134E
      TabIndex        =   7
      Top             =   4755
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00D89970&
      Caption         =   "Fazer Backup do Arquivo Original"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2865
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   3795
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D89970&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   2220
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   4380
      Width           =   4695
      Begin VB.PictureBox status 
         BackColor       =   &H00D89970&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   45
         ScaleHeight     =   780
         ScaleWidth      =   4605
         TabIndex        =   3
         Top             =   240
         Width           =   4605
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   615
      ItemData        =   "frmMain.frx":E1350
      Left            =   10725
      List            =   "frmMain.frx":E1352
      TabIndex        =   1
      Top             =   3855
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EFD1AD&
      Caption         =   "&Instalar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4005
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3255
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5850
      Picture         =   "frmMain.frx":E1354
      Top             =   765
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   210
      Left            =   5385
      TabIndex        =   12
      Top             =   5970
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gabrielfalcao@hotmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   210
      Left            =   5385
      TabIndex        =   11
      Top             =   5775
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patch criado por Gabriel Falcão"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   210
      Left            =   5205
      TabIndex        =   10
      Top             =   5550
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   330
      Left            =   5730
      TabIndex        =   9
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "em Português"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   330
      Left            =   5040
      TabIndex        =   8
      Top             =   2565
      Width           =   1785
   End
   Begin VB.Label mov 
      BackStyle       =   0  'Transparent
      Height          =   1080
      Index           =   1
      Left            =   420
      MousePointer    =   5  'Size
      TabIndex        =   6
      Top             =   300
      Width           =   4020
   End
   Begin VB.Label mov 
      BackStyle       =   0  'Transparent
      Height          =   2730
      Index           =   0
      Left            =   225
      MousePointer    =   5  'Size
      TabIndex        =   5
      Top             =   135
      Width           =   2490
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const SW_SHOW = 5
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion
Private mCaptionlessWindowMover As CCaptionlessWindowMover
Dim HDdrive As String
Dim files As Boolean
Dim pastapal As String
Dim nada As Integer
Dim cache As Boolean
Dim contador As Integer
Dim navi As Integer
Dim cnt As Integer
Dim maxcache As Integer
Private Sub GetFiles(Path As String, SubFolder As Boolean)
Screen.MousePointer = 0
        Dim li As ListBox
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    fPath = AddBackslash(Path)
    fName = fPath & "*.*"
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
      If StripNulls(WFD.cFileName) = "Palace32.exe" Then
       List1.AddItem fPath & StripNulls(WFD.cFileName)
       List2.AddItem fPath
             End If
    End If
    While FindNextFile(hFile, WFD)
               If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
         If StripNulls(WFD.cFileName) = "Palace32.exe" Then
          List1.AddItem fPath & StripNulls(WFD.cFileName)
                 List2.AddItem fPath
          Status.Cls
    Status.Print "Encontrado em " & fPath
      End If
        End If
    Wend
    If SubFolder Then
        hFile = FindFirstFile(fName, WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
        StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
            GetFiles fPath & StripNulls(WFD.cFileName), True
        End If
        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
            StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
                GetFiles fPath & StripNulls(WFD.cFileName), True
            End If
        Wend
    End If
    FindClose hFile
    Set li = Nothing
    If List1.ListCount > 0 Then
   If cache = False Then
    Status.Cls
    Status.Print List1.ListCount & " Arquivo(s) encontrado(s)!"
     Screen.MousePointer = 0
     List1.ListIndex = 0
       If cnt = 1 Then instalar
        cnt = cnt + 1
      End If
      Exit Sub
          Else
    Status.Cls
    Status.Print "Procurando em " & fPath
     End If

    
End Sub
Private Function StripNulls(f As String) As String
    StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function
Private Function AddBackslash(S As String) As String
    If Len(S) Then
       If Right$(S, 1) <> "\" Then
          AddBackslash = S & "\"
       Else
          AddBackslash = S
       End If
    Else
       AddBackslash = "\"
    End If
End Function

Private Sub command1_Click()
Status.Cls
Status.Print "Procurando arquivos aguarde..."
GetFiles "C:\", True
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Set m_Rgn = New CBMPRegion
Set mCaptionlessWindowMover = New CCaptionlessWindowMover
  Set mCaptionlessWindowMover.Form = Me
  m_Rgn.CreateFromPic Me.Picture, vbBlack
  SetWindowRgn hwnd, m_Rgn.Handle, True
  cache = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetWindowRgn hwnd, 0, False
  m_Rgn.Destroy
  Set m_Rgn = Nothing
End Sub


Private Sub mov_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseDown X, Y
End Sub

Private Sub mov_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseMove X, Y
End Sub

Private Sub mov_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  mCaptionlessWindowMover.HandleMouseUp
End Sub
Private Sub instalar()
On Error Resume Next
Dim i
MsgBox "Se o " & Chr(34) & "The Palace" & Chr(34) & " estiver aberto, feche-o antes de continuar."
For i = 0 To List1.ListCount And List2.ListCount
List1.ListIndex = i
List2.ListIndex = i
If Check1.Value = 1 Then
Status.Cls
Status.Print "Fazendo Backup do Arquivo Original..." & vbCrLf & "Copiando Arquivo:" & vbCrLf & "DE: " & List1.Text & vbCrLf & "PARA: " & List2.Text & "PalENG.bkp"
FileCopy List1.Text, List2.Text & "Backup Palace32.exe"
Kill List1.Text
End If
Status.Cls
Status.Print "Preparando..."
Dim p_ret As String
Dim nome As String
Status.Cls
Status.Print "Modificando Arquivo..."
nome = List1.Text
p_ret = StrConv(LoadResData("PALACE.EXE", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
Next i
Status.Cls
Status.Print "Arquivo Substituído com Sucesso!" & vbCrLf & "Obrigado por utilizar o Patch!" & vbCrLf & "Visite o site: www.gabrielfalcao.i8.com para releases..."
cache = True
End Sub
