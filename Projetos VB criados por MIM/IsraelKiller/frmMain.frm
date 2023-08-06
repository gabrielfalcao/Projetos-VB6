VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00D2D2DC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar arquivos do GKEYLOG"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   2565
      TabIndex        =   6
      Top             =   4905
      Width           =   930
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "frmMain.frx":0E42
      Left            =   60
      List            =   "frmMain.frx":0E44
      TabIndex        =   5
      Top             =   3030
      Width           =   7470
   End
   Begin VB.PictureBox status 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   7560
      TabIndex        =   4
      Top             =   5520
      Width           =   7590
   End
   Begin VB.DriveListBox Drive1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2130
      TabIndex        =   2
      Top             =   450
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procurar!"
      Height          =   585
      Left            =   300
      TabIndex        =   0
      Top             =   405
      Width           =   1500
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "frmMain.frx":0E46
      Left            =   75
      List            =   "frmMain.frx":0E48
      TabIndex        =   3
      Top             =   1080
      Width           =   7470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D2D2DC&
      Caption         =   "Ao achar copiar em:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2205
      TabIndex        =   1
      Top             =   195
      Width           =   1635
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
Dim i
Dim arqui As String
Dim nada As Integer
Dim pastadestino As String
Dim cache As Boolean
Dim contador As Integer
Dim navi As Integer
Dim auto As String
Dim maxcache As Integer
Dim pasta As String

Private Sub GetFiles(Path As String, SubFolder As Boolean)
On Error Resume Next
    Screen.MousePointer = 11
        Dim li As ListBox
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, fName As String
    fPath = AddBackslash(Path)
    fName = fPath & "*.*"
    hFile = FindFirstFile(fName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
      If Right$(fPath & StripNulls(WFD.cFileName), 3) = "txt" Then
      pasta = pasta & "LOG"
           List1.AddItem fPath & StripNulls(WFD.cFileName)
           arqui = StripNulls(WFD.cFileName)
  If Left$(arqui, 4) = "gkey" Then List2.AddItem fPath & StripNulls(WFD.cFileName)
      End If
    End If
    While FindNextFile(hFile, WFD)
               If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
          If Right$(fPath & StripNulls(WFD.cFileName), 3) = "txt" Then
            pasta = pasta & "LOG"
              List1.AddItem fPath & StripNulls(WFD.cFileName)
           arqui = StripNulls(WFD.cFileName)
  If Left$(arqui, 4) = "gkey" Then List2.AddItem fPath & StripNulls(WFD.cFileName)
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
Screen.MousePointer = 0
If Screen.MouseIcon = 0 Then
Command1.Enabled = True
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
Private Sub ativar()
HDdrive = "C:\"
GetFiles HDdrive, True
End Sub

Private Sub Command1_Click()
pasta = "C:\"
HDdrive = "C:\"
GetFiles HDdrive, True

End Sub

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Sub Command2_Click()
For i = 0 To List2.ListCount - 1
Kill List2.Text
List2.ListIndex = i
Next i
End Sub

Private Sub Dir1_Change()
If Len(Dir1.Path) > 3 Then
pastadestino = Dir1.Path & "\"
Else
pastadestino = Dir1.Path
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

