VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox step 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFD1AD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Index           =   1
      Left            =   60
      ScaleHeight     =   3075
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   5685
      Begin VB.ListBox lvfiles 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFD1AD&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         ItemData        =   "frmMain.frx":08CA
         Left            =   1950
         List            =   "frmMain.frx":08CC
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFD1AD&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   150
         TabIndex        =   2
         Top             =   2160
         Width           =   5235
         Begin VB.PictureBox status 
            BackColor       =   &H00EFD1AD&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00990B0B&
            Height          =   285
            Left            =   60
            ScaleHeight     =   285
            ScaleWidth      =   5130
            TabIndex        =   3
            Top             =   150
            Width           =   5130
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "Iniciar Busca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFF9E3&
         BackStyle       =   0  'Transparent
         Caption         =   "Clique em ""Avançar"" para continuar..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00990B0B&
         Height          =   165
         Left            =   3255
         TabIndex        =   5
         Top             =   2835
         Width           =   2370
      End
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Dim HDdrive As String
Dim files As Boolean
Dim nada As Integer
Dim cache As Boolean
Dim contador As Integer
Dim navi As Integer
Dim maxcache As Integer
Private Sub GetFiles(Path As String, SubFolder As Boolean)
    Screen.MousePointer = 11
        Dim li As ListBox
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    fPath = AddBackslash(Path)
    fName = fPath & "*.*"
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
      If Right$(fPath & StripNulls(WFD.cFileName), 3) = "tmp" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "jpg" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "peg" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "gif" Then
      Kill fPath & StripNulls(WFD.cFileName)
      End If
    End If
    While FindNextFile(hFile, WFD)
               If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
          If Right$(fPath & StripNulls(WFD.cFileName), 3) = "tmp" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "jpg" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "peg" Or Right$(fPath & StripNulls(WFD.cFileName), 3) = "gif" Then
     Kill fPath & StripNulls(WFD.cFileName)

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
    If List1.ListCount - 1 > 0 Then

     Screen.MousePointer = 0

    Else

      Screen.MousePointer = 0
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
Private Sub cmdBck_Click()
If navi > 0 Then navi = navi - 1
Passear
End Sub
Private Sub cmdCancel_Click()
If MsgBox("Deseja realmente sair?", vbQuestion + vbYesNo, frmMain.Caption) = vbYes Then End
End Sub
Private Sub cmdFin_Click()
End
End Sub
Private Sub cmdFwd_Click()
navi = navi + 1
Passear
End Sub
Private Sub Passear()
Select Case navi
Case 0
step(0).Visible = True
step(1).Visible = False
step(2).Visible = False
step(3).Visible = False
cmdBck.Enabled = True
cmdFwd.Enabled = True
cmdFin.Enabled = False
cmdCancel.Enabled = True
Case 1
step(0).Visible = False
step(1).Visible = True
step(2).Visible = False
step(3).Visible = False
cmdBck.Enabled = True
cmdFwd.Enabled = True
cmdFin.Enabled = False
cmdCancel.Enabled = True
Case 2
step(0).Visible = False
step(1).Visible = False
step(2).Visible = True
step(3).Visible = False
cmdBck.Enabled = True
cmdFwd.Enabled = True
cmdFin.Enabled = False
cmdCancel.Enabled = True
Case 3
step(0).Visible = False
step(1).Visible = False
step(2).Visible = False
step(3).Visible = True
cmdBck.Enabled = True
cmdFwd.Enabled = False
cmdFin.Enabled = True
filesok.Visible = files
cacheok.Visible = cache
cmdCancel.Enabled = False
If nada > 0 Then
nadaf.Visible = False
Label15.Visible = True
limpeza.Caption = "Limpeza Concluída!"
Else
nadaf.Visible = True
Label15.Visible = False
limpeza.Caption = "Limpeza Incompleta!"
End If
End Select
End Sub
Private Sub Command1_Click()
On Error GoTo err
files = True
nada = nada + 1
Dim i
List1.ListIndex = 0
lvfiles.ListIndex = 0
Command1.Enabled = False
For i = 0 To List1.ListCount - 1
Screen.MousePointer = 11
status.Cls
status.Print "Deletando: " & List1.Text
Kill List1.Text
List1.ListIndex = i
lvfiles.ListIndex = i
If i = List1.ListCount - 1 Then
status.Cls
If List1.ListCount - contador - 2 > 0 Then
MsgBox List1.ListCount - contador - 2 & "  Arquivo(s) removido(s) com sucesso!", vbInformation, Me.Caption
status.Print List1.ListCount - contador - 2 & "  Arquivo(s) removido(s) com sucesso!"
Else
MsgBox "0 Arquivo(s) removido(s) com sucesso!", vbInformation, Me.Caption
status.Print "0 Arquivo(s) removido(s) com sucesso!"
End If
List1.Clear
Command2.Enabled = True
lvfiles.Clear
Screen.MousePointer = 0
Command1.Enabled = False
contador = 0
Exit Sub
End If
Next i

err:
If err.Number = 75 Then
contador = contador + 1
Resume Next
Else
Resume Next
End If
End Sub
Private Sub command2_Click()

End Sub
Private Sub del(arquivo As String)
On Error Resume Next
Kill arquivo
End Sub
Private Sub Command3_Click()
   Dim cachefile As String
   Dim i As Long
   For i = 0 To List2.ListCount - 1
      cachefile = List2.List(i)
      If InStr(cachefile, "Cookie") = 0 Then
         Call DeleteUrlCacheEntry(cachefile)
      End If
   Next
   GetCacheURLList
   Dim tot As Integer
   maxcache = List2.ListCount
   tot = maxcache - Int(List2.ListCount)
     Label8.Caption = "Arquivos de cache removidos!"
nada = nada + 1
cache = True
Command3.Enabled = False
End Sub

Private Sub Drive1_Change()

End Sub

Private Sub Form_Load()
status.Cls
status.Print "Procurando arquivos aguarde..."
HDdrive = Left$(Drive1.Drive, 2) & "\"
GetFiles HDdrive, True
End Sub
Public Sub GetCacheURLList()
   Dim ICEI As INTERNET_CACHE_ENTRY_INFO
   Dim hFile As Long
   Dim cachefile As String
   Dim posUrl As Long
   Dim posEnd As Long
   Dim dwBuffer As Long
   Dim pntrICE As Long
   List2.Clear
   dwBuffer = 0
   hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)
   If (hFile = ERROR_CACHE_FIND_FAIL) And _
      (err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
      pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
      If pntrICE Then
         CopyMemory ByVal pntrICE, dwBuffer, 4
         hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
         If hFile <> ERROR_CACHE_FIND_FAIL Then
            Do
               CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
               If (ICEI.CacheEntryType And _
                   NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
                  cachefile = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                  List2.AddItem cachefile
               End If
               Call LocalFree(pntrICE)
               dwBuffer = 0
               Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
               pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
               CopyMemory ByVal pntrICE, dwBuffer, 4
            Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
         End If
      End If
   End If
   Call LocalFree(pntrICE)
   Call FindCloseUrlCache(hFile)
End Sub
Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = &H800000
Label18.ForeColor = &H800000
End Sub

Private Sub Label18_Click()
ShellExecute GetActiveWindow(), "Open", "http://www.gabrielfalcao.i8.com", "", 0&, 1
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.ForeColor = &HC0&
Label19.ForeColor = &H800000
End Sub

Private Sub Label19_Click()
ShellExecute GetActiveWindow(), "Open", "mailto:gabrielfalcao@hotmail.com", "", 0&, 1
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = &HC0&
Label18.ForeColor = &H800000
End Sub

Private Sub step_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = &H800000
Label18.ForeColor = &H800000
End Sub
