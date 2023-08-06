VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E7F4FA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Hunter"
   ClientHeight    =   5565
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   7035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6FBE6&
      Height          =   2370
      Left            =   3480
      TabIndex        =   16
      Top             =   810
      Width           =   3480
   End
   Begin VB.ListBox lvfiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6FBE6&
      Height          =   2370
      Left            =   3480
      TabIndex        =   15
      Top             =   810
      Width           =   3480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E7F4FA&
      Caption         =   "Ação"
      Height          =   675
      Left            =   75
      TabIndex        =   13
      Top             =   3315
      Width           =   3645
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Mover para..."
         Enabled         =   0   'False
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
         Left            =   2235
         TabIndex        =   18
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H008080FF&
         Caption         =   "Copiar para..."
         Enabled         =   0   'False
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
         Left            =   945
         TabIndex        =   17
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Deletar"
         Enabled         =   0   'False
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
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E7F4FA&
      Caption         =   "Detalhes da Procura"
      Height          =   2475
      Left            =   45
      TabIndex        =   6
      Top             =   720
      Width           =   3375
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DADAFC&
         Enabled         =   0   'False
         Height          =   285
         Left            =   420
         TabIndex        =   12
         Top             =   1920
         Width           =   2700
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DADAFC&
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Top             =   825
         Width           =   2700
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E7F4FA&
         Caption         =   "Por Extensão"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E7F4FA&
         Caption         =   "Por Nome do Arquivo"
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extensão do arquivo:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   1695
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do arquivo:"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   600
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7F4FA&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   52
      TabIndex        =   4
      Top             =   4935
      Width           =   6930
      Begin VB.PictureBox status 
         BackColor       =   &H00E7F4FA&
         BorderStyle     =   0  'None
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
         Left            =   75
         ScaleHeight     =   285
         ScaleWidth      =   6780
         TabIndex        =   5
         Top             =   195
         Width           =   6780
      End
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
      Height          =   270
      Left            =   3030
      TabIndex        =   3
      Top             =   330
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Procurar"
      Default         =   -1  'True
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
      Left            =   5700
      TabIndex        =   0
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Label lblPasta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   2
      Top             =   330
      Width           =   2955
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha a pasta:"
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
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   1185
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
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
Dim HDdrive As String
Dim contador As Integer
Dim achado As String
  Dim dirlen As Integer
  Dim fil As String
Private Sub GetFiles(Path As String, SubFolder As Boolean)
On Error Resume Next
Dim taman As Integer
    Screen.MousePointer = 11
        Dim li As ListBox
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    fPath = AddBackslash(Path)
    fName = fPath & "*.*"
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
    If Option2.Value = True Then
    Dim asygf As String
    asygf = InStrRevVB5(StripNulls(WFD.cFileName), ".")
    taman = Len(StripNulls(WFD.cFileName)) - Len(asygf)
      If Left$(StripNulls(WFD.cFileName), taman) = Text2.Text Then
      lvfiles.AddItem fPath & StripNulls(WFD.cFileName)
      List1.AddItem fPath & StripNulls(WFD.cFileName)
      End If
    End If
  
   
    ElseIf Option1.Value = True Then
    fil = InStr(1, fPath & StripNulls(WFD.cFileName), ".", vbBinaryCompare)
     dirlen = Len(fil) + 1
    If Left$(fPath & StripNulls(WFD.cFileName), dirlen) = Text1.Text Then
     lvfiles.AddItem fPath & StripNulls(WFD.cFileName)
      List1.AddItem fPath & StripNulls(WFD.cFileName)
      End If
      End If
    While FindNextFile(hFile, WFD)
               If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then

        If Option2.Value = True Then
        Dim tfjh As String
        tfjh = InStrRevVB5(StripNulls(WFD.cFileName), ".")
   taman = Len(StripNulls(WFD.cFileName)) - Len(tfjh)
      If Left$(StripNulls(WFD.cFileName), taman) = Text2.Text Then
      lvfiles.AddItem fPath & StripNulls(WFD.cFileName)
      List1.AddItem fPath & StripNulls(WFD.cFileName)
      End If
    End If
    ElseIf Option1.Value = True Then
   fil = InStr(1, fPath & StripNulls(WFD.cFileName), ".", vbBinaryCompare)
     dirlen = Len(fil) + 1
    If Left$(fPath & StripNulls(WFD.cFileName), dirlen) = Text1.Text Then
     lvfiles.AddItem fPath & StripNulls(WFD.cFileName)
      List1.AddItem fPath & StripNulls(WFD.cFileName)
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
    status.Cls
    status.Print List1.ListCount - 1 & " Arquivos encontrados!"
     Screen.MousePointer = 0
    Command1.Enabled = True
      Command4.Enabled = True
        Command5.Enabled = True
    Command2.Enabled = False
    Else
    status.Cls
    status.Print "Nenhum arquivo encontrado!"
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
Private Sub Command1_Click()
On Error GoTo err
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
status.Cls
status.Print "Procurando arquivos aguarde..."
HDdrive = lblPasta.Caption
GetFiles HDdrive, True
End Sub
Private Sub del(arquivo As String)
On Error Resume Next
Kill arquivo
End Sub

Private Sub Command3_Click()
Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Selecione a Pasta Desejada:"
  
  With tBrowseInfo
    .hWndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Len(sBuffer) = 3 Then
lblPasta.Caption = sBuffer
Else
lblPasta.Caption = sBuffer & "\"
End If
  End If
End Sub
Public Function InStrRevVB5(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1) As Long
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

Private Sub Command4_Click()
Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Copiar Arquivos - Selecione o Destino:"
  
  With tBrowseInfo
    .hWndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  Dim pastatocopy As String
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Len(sBuffer) = 3 Then
pastatocopy = sBuffer
Else
pastatocopy = sBuffer & "\"
End If
  End If
  '''''''''''''
  If sBuffer = Empty Then Exit Sub
  Dim i
List1.ListIndex = 0
lvfiles.ListIndex = 0
Command4.Enabled = False
For i = 0 To List1.ListCount - 1
Screen.MousePointer = 11
status.Cls
status.Print "Copiando: " & List1.Text & " para " & pastatocopy
FileCopy List1.Text, InStr(1, List1.Text, "\", vbTextCompare)
List1.ListIndex = i
lvfiles.ListIndex = i
If i = List1.ListCount - 1 Then
status.Cls
If List1.ListCount - contador - 2 > 0 Then
MsgBox List1.ListCount - contador - 2 & "  Arquivo(s) copiado(s) com sucesso!", vbInformation, Me.Caption
status.Print List1.ListCount - contador - 2 & "  Arquivo(s) copiado(s) com sucesso!"
Else
MsgBox "0 Arquivo(s) copiados(s) com sucesso!", vbInformation, Me.Caption
status.Print "0 Arquivo(s) copiado(s) com sucesso!"
End If
List1.Clear
Command2.Enabled = True
lvfiles.Clear
Screen.MousePointer = 0
Command4.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
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

Private Sub Command5_Click()
Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = vbCr & vbCr & "Mover Arquivos - Selecione o Destino:"
  
  With tBrowseInfo
    .hWndOwner = Me.hWnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  Dim pastatocopy As String
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Len(sBuffer) = 3 Then
pastatocopy = sBuffer
Else
pastatocopy = sBuffer & "\"
End If
  End If
  '''''''''''''
  If sBuffer = Empty Then Exit Sub
  Dim i
List1.ListIndex = 0
lvfiles.ListIndex = 0
Command4.Enabled = False
For i = 0 To List1.ListCount - 1
Screen.MousePointer = 11
status.Cls
status.Print "Movendo: " & List1.Text & " para " & pastatocopy
FileCopy List1.Text, InStr(1, List1.Text, "\", vbTextCompare)
Kill List1.Text
List1.ListIndex = i
lvfiles.ListIndex = i
If i = List1.ListCount - 1 Then
status.Cls
If List1.ListCount - contador - 2 > 0 Then
MsgBox List1.ListCount - contador - 2 & "  Arquivo(s) movido(s) com sucesso!", vbInformation, Me.Caption
status.Print List1.ListCount - contador - 2 & "  Arquivo(s) movido(s) com sucesso!"
Else
MsgBox "0 Arquivo(s) copiados(s) com sucesso!", vbInformation, Me.Caption
status.Print "0 Arquivo(s) movido(s) com sucesso!"
End If
List1.Clear
Command2.Enabled = True
lvfiles.Clear
Screen.MousePointer = 0
Command4.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
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

Private Sub Option1_Click()
Label1.Enabled = Option1.Value
Text1.Enabled = Option1.Value
Label2.Enabled = Option2.Value
Text2.Enabled = Option2.Value
End Sub

Private Sub Option2_Click()
Label2.Enabled = Option2.Value
Text2.Enabled = Option2.Value
Label1.Enabled = Option1.Value
Text1.Enabled = Option1.Value
End Sub
