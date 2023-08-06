VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmProcurar 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procurar arquivos"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iml16 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1260
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton CmdAcao 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdAcao 
      Caption         =   "&Adicionar"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdAcao 
      Caption         =   "&Procurar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.ListView LVFiles 
      Height          =   2895
      Left            =   105
      TabIndex        =   4
      Top             =   1440
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "iml16"
      ForeColor       =   -2147483640
      BackColor       =   16775914
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Arquivo"
         Object.Width           =   7408
      EndProperty
   End
   Begin VB.TextBox TxtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.DriveListBox DrvFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackColor       =   &H00EFD1AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do arquivo para pesquisa:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackColor       =   &H00EFD1AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidade de disco para pesquisa:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
End
Attribute VB_Name = "FrmProcurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PATH = 260
Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type


Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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
    cAlternateFileName As String * 14
End Type

Private Enum FILE_ATTRIBUTES
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

Private lCounter As Long
Private lTickCount As Long
Private ShInfo As SHFILEINFO
Private Item As ListItem

Private Function FixPath(sPath As String) As String
 FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
End Function

Private Sub ShowIconsForFiles()
  Set LVFiles.SmallIcons = Nothing
  iml16.ListImages.Clear
  GetAllIcons
  ShowIcons
End Sub

Private Sub GetAllIcons()
  For Each Item In LVFiles.ListItems
    GetIcon Item.Text, Item.Index
  Next
End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
Dim hSIcon As Long, R As Long
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

If hSIcon <> 0 Then
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    R = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Call iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub ShowIcons()
  If iml16.ListImages.Count < LVFiles.ListItems.Count Then Exit Sub
  With LVFiles
    .SmallIcons = iml16
    For Each Item In .ListItems
      Item.SmallIcon = Item.Index
    Next
  End With
End Sub

Private Sub ShowFolderList(sFolder As String, sSearchFor As String)
Dim WFD As WIN32_FIND_DATA
Dim WFD2 As WIN32_FIND_DATA
Dim lReturn As Long
Dim lReturn2 As Long
Dim lNextFile As Long
Dim lNextFile2 As Long
Dim sPath As String
Dim sPath2 As String
Dim sFilename As String
Dim sFileName2 As String
Dim sFullPath As String
Dim ItmX As ListItem
    
    sPath = (sFolder & "*.*") & Chr(0)
    lReturn = FindFirstFile(sPath, WFD)
    
    Do
        If (WFD.dwFileAttributes And vbDirectory) Then
            sFilename = StripNullChar(WFD.cFileName)
            If sFilename <> "." And sFilename <> ".." Then
                Me.Caption = LoadResString(MyLanguage + 166) & sFilename
                GoSub List_Files_In_Folder:
                lCounter = lCounter + 1
                ShowFolderList (sFolder & sFilename & "\"), (sSearchFor)
            End If
        End If
     lNextFile = FindNextFile(lReturn, WFD)
    Loop Until lNextFile = False
    lNextFile = FindClose(lReturn)
Exit Sub


List_Files_In_Folder:
 sPath2 = sFolder & "*.*"
 lReturn2 = FindFirstFile(sPath2, WFD2) & Chr(0)
        
    If lReturn2 > 0 Then
      Do
        If Not (WFD2.dwFileAttributes And vbDirectory) = vbDirectory Then
            sFileName2 = StripNullChar(WFD2.cFileName)
            If IsEmpty(sFileName2) = False Then
              sFullPath = FixPath(sFolder) & sFileName2
              If InStr(1, sFileName2, sSearchFor, vbTextCompare) > 0 And ExistListItem(sFullPath) = False Then
               Set ItmX = LVFiles.ListItems.Add(, sFullPath, sFullPath)
              End If
            End If
        End If
       lNextFile2 = FindNextFile(lReturn2, WFD2)
      Loop Until lNextFile2 <= Val(0)
    End If
 lNextFile2 = FindClose(lReturn2)
Return
End Sub

Private Function ExistListItem(sFile As String) As Boolean
On Error GoTo err1

 ExistListItem = IIf(LVFiles.ListItems(sFile).Text <> "", True, False)
 Exit Function
 
err1:
 ExistListItem = False
 Exit Function
End Function

Private Function StripNullChar(sInput As String)
Dim iSearch As Integer
  iSearch = InStr(1, sInput, Chr(0))
  If iSearch > 0 Then StripNullChar = Left(sInput, iSearch - 1)
End Function

Private Sub CmdAcao_Click(Index As Integer)
Select Case Index
 Case 0
    Screen.MousePointer = 11
    lCounter = 0
    lTickCount = GetTickCount
    LVFiles.ListItems.Clear
    CmdAcao(0).Enabled = False
    CmdAcao(1).Enabled = False
    TxtFile.Text = Replace(TxtFile.Text, "*.", "")
    
    Call ShowFolderList(Left(DrvFile.Drive, 2) & "\", TxtFile.Text)
    
    Screen.MousePointer = 0
    CmdAcao(0).Enabled = True
    
    
    If LVFiles.ListItems.Count = 0 Then
     CmdAcao(1).Enabled = False
     LVFiles.ListItems.Add , , Replace(Replace(LoadResString(MyLanguage + 144), "%1", lCounter), "%2", Format(GetTickCount - lTickCount&, "###,000"))
    Else
     CmdAcao(1).Enabled = True
     ShowIconsForFiles
    End If
    Screen.MousePointer = 0
    Me.Caption = LoadResString(MyLanguage + 145)

 Case 1
    lCounter = 0
    If FrmMenu.Zipit1.FileName = "" Then
     MsgBox LoadResString(MyLanguage + 165)
     Unload Me
     Exit Sub
    End If
    
    Dim COL As New Collection
    For Each Item In LVFiles.ListItems
      If Item.Selected = True Then
       lCounter = lCounter + 1
       COL.Add Item.Text
      End If
    Next
    If lCounter = 0 Then
     LVFiles.SetFocus
     Exit Sub
    Else
     CmdAcao(0).Enabled = False
     CmdAcao(1).Enabled = False
     Screen.MousePointer = 11
     FrmMenu.Zipit1.Add COL
     Screen.MousePointer = 0
     Unload Me
    End If
 
 Case 2
    Unload Me
End Select
End Sub

Private Sub Form_Load()
 MyLanguage = GetSetting(App.EXEName, "last", "language", 0)
 LoadCaption
 
 Me.Icon = FrmMenu.Icon
 TxtFile.Text = GetSetting(App.EXEName, "last", "filesearch", "")

 pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
 pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting App.EXEName, "last", "filesearch", TxtFile.Text
End Sub

Private Sub LVFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With LVFiles
    If .SortKey <> ColumnHeader.Index - 1 Then
     .SortKey = ColumnHeader.Index - 1
     .SortOrder = lvwAscending
    Else
     If .SortOrder = lvwAscending Then
      .SortOrder = lvwDescending
     Else
      .SortOrder = lvwAscending
     End If
    End If
    .Sorted = -1
  End With
End Sub

Private Sub LVFiles_KeyDown(KeyCode As Integer, Shift As Integer)
 If LVFiles.ListItems.Count = 0 Then Exit Sub
 If KeyCode = vbKeyDelete Then
  Dim SOC As New SHFileOPClass
  Dim DCOL As New Collection
  
  For Each Item In LVFiles.ListItems
   If Item.Selected = True Then DCOL.Add Item.Text
  Next
   
  If DCOL.Count > 0 Then
   With SOC
    .ParentWnd = Me.hWnd
    .ClearSourceFiles
     For lCounter = 1 To DCOL.Count
      .AddSourceFile DCOL.Item(lCounter)
     Next lCounter
     .AllowUndo = True
     .ConfirmOperation = True
     If .DeleteFiles = True Then
      For lCounter = 1 To DCOL.Count
       LVFiles.ListItems.Remove DCOL.Item(lCounter)
      Next
     End If
   End With
  End If
 End If
End Sub
Private Sub TxtFile_GotFocus()
 TxtFile.SelStart = 0
 TxtFile.SelLength = Len(TxtFile.Text)
End Sub
Private Sub LoadCaption()
  With Me
   .Caption = LoadResString(MyLanguage + 145)
   .L(0).Caption = LoadResString(MyLanguage + 123)
   .L(1).Caption = LoadResString(MyLanguage + 124)
   .CmdAcao(0).Caption = LoadResString(MyLanguage + 125)
   .CmdAcao(1).Caption = LoadResString(MyLanguage + 126)
   .CmdAcao(2).Caption = LoadResString(MyLanguage + 127)
   .LVFiles.ColumnHeaders(1).Text = LoadResString(MyLanguage + 150)
  End With
End Sub
