VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gabs Antivirus 1.2"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFind.frx":0E42
   ScaleHeight     =   2850
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remover Encontrados"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      TabIndex        =   4
      Top             =   1950
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFAD64&
      Caption         =   "Iniciar Busca"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      TabIndex        =   3
      Top             =   1635
      Width           =   2700
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   60
      Top             =   105
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillStyle       =   3  'Vertical Line
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00691412&
      Height          =   525
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6885
      TabIndex        =   1
      Top             =   2325
      Width           =   6915
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAEAEA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00691412&
      Height          =   750
      ItemData        =   "frmFind.frx":E804
      Left            =   0
      List            =   "frmFind.frx":E806
      TabIndex        =   0
      Top             =   1590
      Width           =   3705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos Encontrados:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007A4232&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1290
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create a form with a command button (command1), a list box (list1)
'and four text boxes (text1, text2, text3 and text4).
'Type in the first textbox a startingpath like c:\
'and in the second textbox you put a pattern like *.* or *.txt

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Dim t As Integer
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net

    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                List1.AddItem path & FileName
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If
End Function
Sub acharvirus(status As PictureBox, lista As ListBox)
On Error Resume Next
Dim i As Long
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Screen.MousePointer = vbHourglass
    Picture1.Print "Procurando arquivos, aguarde..."
    lista.Clear


    SearchPath = "C:\Windows\"
    FindStr = "ccApp.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)

    FindStr = "POPStart.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    
     FindStr = "AVG7 Update.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    
  FindStr = "POPInicio.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)

    FindStr = "msnUpdate.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    
      SearchPath = "C:\Windows\system\"
    FindStr = "SysApp.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
   FindStr = "rundl32.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
   FindStr = "Garfield.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
   FindStr = "MSN.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
'C:\WINDOWS\system32
      SearchPath = "C:\WINDOWS\system32\"
    FindStr = "WUpdate.exe"
    FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
    status.Cls
If NumFiles > 0 Then
    Picture2.Cls
    Picture2.BackColor = &HC0FF&
    Picture2.ForeColor = &HFF&
Picture2.Print "É Altamente recomendado remover os arquivos encontrados, reiniciar" & vbCrLf & " o computador nesse exato momento e fazer mais uma verficação de vírus!!!"
    Screen.MousePointer = vbDefault
    lista.ListIndex = 0

If lista.ListCount > 0 Then
 Command2.Enabled = True
 Command1.Caption = "PRONTO"
 Else
 Command1.Enabled = True
 Command1.Caption = "Iniciar Busca"
 End If
 Else
 Command1.Enabled = True
 Command1.Caption = "Iniciar Busca"
 Screen.MousePointer = vbDefault
 status.BackColor = &HC000&
 status.Print "Verificação Concluída com Sucesso!" & vbCrLf & "Não foi encontrado nenhum vírus!"
 End If
End Sub


Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Command1.Caption = "Aguarde..."
Command1.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub limpar(list As ListBox)
On Error Resume Next
   For i = 0 To list.ListCount
  
    Kill list.Text
      list.ListIndex = i
    Next i
    
If List1.ListIndex = list.ListCount - 1 Then list.Clear

End Sub

Private Sub Command2_Click()
limpar List1
Command2.Enabled = False
Command1.Enabled = False

Command1.Caption = "!!Reinicie o PC!!"
End Sub

Private Sub Form_Load()
t = 0
End Sub



Private Sub Timer1_Timer()

Select Case t
Case Is = 0
Picture2.Cls
Picture2.Print "Aguarde..." & vbCrLf & "Removendo entradas do registro..."
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "GNUKey .NET", Empty
Case Is = 1
Picture2.Cls
Picture2.Print "Removendo entradas do registro..." & vbCrLf & "Entrada: GNUKey .NET - Removida com sucesso!"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "GNUKey", Empty
Case Is = 2
Picture2.Cls
Picture2.Print "Removendo entradas do registro..." & vbCrLf & "Entrada: GNUKey - Removida com sucesso!"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Microsofts .NET Support", Empty
Case Is = 3
Picture2.Cls
Picture2.Print "Removendo entradas do registro..." & vbCrLf & "Entrada: Microsofts .NET Support - Removida com sucesso!"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "Gnutella .NET", Empty
Case Is = 4
Picture2.Cls
Picture2.Print "Removendo entradas do registro..." & vbCrLf & "Entrada: Gnutella .NET - Removida com sucesso!"
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "POP Discador Init", Empty
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "AVG7 Update", Empty
Case Is = 5
Picture2.Cls
Picture2.Print "Todas as entradas maléficas do registro foram removidas com sucesso!" & vbCrLf & "Aguarde enquanto o programa procura os arquivos de vírus..."
Case Is = 6
Timer1.Enabled = False
acharvirus Picture2, List1
End Select
t = t + 1
End Sub
Private Sub Form_Initialize()
Dim XP As Long
XP = InitCommonControls
End Sub
