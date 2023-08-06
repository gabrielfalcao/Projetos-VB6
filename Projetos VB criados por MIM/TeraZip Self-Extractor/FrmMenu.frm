VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMenu 
   Caption         =   "TeraZip - Compactador e Descompactador "
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9075
   ForeColor       =   &H00754709&
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgTBar 
      Left            =   1320
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":1FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2108
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":22E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":243C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":2596
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":26F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":284A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMenu.frx":29A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1320
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1860
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   4800
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgTBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "options"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "del"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "extract"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "view"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addfolder"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4905
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVFiles 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "iml32"
      SmallIcons      =   "iml16"
      ForeColor       =   7685897
      BackColor       =   16775914
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Arquivo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Tamanho Original"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Tamanho Compactado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "% Compactação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Data de criação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "CRC"
         Object.Width           =   2540
      EndProperty
   End
   Begin ZipDir.Zipit Zipit1 
      Left            =   5400
      Top             =   4320
      _ExtentX        =   979
      _ExtentY        =   979
      ExtractDir      =   "c:\"
      UseDirectoryInfo=   0   'False
      Overwrite       =   -1  'True
      IncludeSystemFiles=   0   'False
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   720
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Novo"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Abrir"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Fechar"
         Index           =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "&Opções"
         Index           =   4
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Recentes"
         Index           =   6
         Begin VB.Menu mnuRecentSub 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRecentSub 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecentSub 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Sai&r"
         Index           =   8
      End
   End
   Begin VB.Menu mnuActionTop 
      Caption         =   "Açõ&es"
      Begin VB.Menu mnuActionSub 
         Caption         =   "&Adicionar..."
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuActionSub 
         Caption         =   "&Excluir"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuActionSub 
         Caption         =   "E&xtrair..."
         Index           =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuActionSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuActionSub 
         Caption         =   "&Visualizar"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuActionSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuViewTop 
         Caption         =   "Exi&bir"
         Begin VB.Menu mnuViewSub 
            Caption         =   "Ícones grandes"
            Index           =   0
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "Ícones pequenos"
            Index           =   1
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "Lista"
            Index           =   2
         End
         Begin VB.Menu mnuViewSub 
            Caption         =   "Detalhes"
            Index           =   3
         End
      End
      Begin VB.Menu mnuActionSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsTop 
         Caption         =   "&Ferramentas"
         Begin VB.Menu mnuToolsSub 
            Caption         =   "&Procurar arquivos..."
            Index           =   0
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuToolsSub 
            Caption         =   "&Adicionar diretório..."
            Index           =   1
            Shortcut        =   ^D
         End
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuHelpSub 
         Caption         =   "&Sobre..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "&Página oficial"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FrmMenu"
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

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private ZipName As String
Private i As Integer
Private J As Integer
Private B As Boolean
Private R As Long

Private O As Collection
Private Item As ListItem
Private ShInfo As SHFILEINFO
Private FSO As FileSystemObject

Private Function FixPath(sPath As String) As String
 FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
End Function

Private Function GetTmpPath() As String
Dim strFolder As String
Dim lngResult As Long
  
  strFolder = String(MAX_PATH, 0)
  lngResult = GetTempPath(MAX_PATH, strFolder)

If lngResult <> 0 Then GetTmpPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1) Else GetTmpPath = CurDir
End Function

Private Function GetFileTitle(FileName As String) As String
If Trim(FileName) = "" Then Exit Function
Dim i As Integer, C As String, Pos As Integer
For i = 1 To Len(FileName) Step 1
  C = Mid(FileName, i, 1)
  If C = "\" Then Pos = i + 1
Next
GetFileTitle = Mid(FileName, Pos, (Len(FileName) + 1 - Pos))
End Function

Private Function FormatFileSize(ByVal lFileSize As Long) As String
Select Case lFileSize
    Case 0 To 1023
      FormatFileSize = Format(lFileSize, "##0") & " Bytes"
    Case 1024 To 1048575
      FormatFileSize = Format(lFileSize / 1024#, "#,##0") & " KB"
    Case 1024# ^ 2 To 1073741823
      FormatFileSize = Format(lFileSize / (1024# ^ 2), "#,##0.00") & " MB"
    Case Is > 1073741823#
      FormatFileSize = Format(lFileSize / (1024# ^ 3), "#,###,##0.00") & " GB"
End Select
End Function

Private Sub Status(Optional sText As String = "", Optional iPanel As Integer = 1)
  SBar.Panels(iPanel).Text = sText
End Sub

Private Sub ShowPBar()
Dim tRC As RECT
SendMessageAny SBar.hWnd, SB_GETRECT, 2, tRC
  
With tRC
 .Top = (.Top * Screen.TwipsPerPixelY)
 .Left = (.Left * Screen.TwipsPerPixelX)
 .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
 .Right = (.Right * Screen.TwipsPerPixelX) - .Left
End With

With PBar
  SetParent .hWnd, SBar.hWnd
 .Move tRC.Left + 10, tRC.Top + 10, tRC.Right - 15, tRC.Bottom - 25
 .Visible = True
 .Value = 0
End With
End Sub

Private Sub Form_Load()
 MyLanguage = GetSetting(App.EXEName, "last", "language", 0)
 LoadCaption
 
 pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
 pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
 pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
 pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY
 Set FSO = New FileSystemObject
 LVFiles.View = GetSetting(App.EXEName, "last", "view", lvwReport)
 
 For R = 1 To LVFiles.ColumnHeaders.Count
  LVFiles.ColumnHeaders(R).Width = GetSetting(App.EXEName, "last", "colw" & R, 1500)
 Next R
 
 Call mnuFileSub_Click(2)
 mnuViewSub(LVFiles.View).Checked = True


 If Command <> "" Then
  If InStr(1, LCase(Command), "-o") > 0 Then GoSub Command_OpenFile
  If InStr(1, LCase(Command), "-d") > 0 Then GoSub Command_ExtractFile
  If InStr(1, LCase(Command), "-a") > 0 Then GoSub Command_ExtractFile
 End If

 J = 0
 For i = 0 To 2
  mnuRecentSub(i).Caption = GetSetting(App.EXEName, "recent", "file" & i + 1, "")
  If Trim(mnuRecentSub(i).Caption) = "" Then
   J = J + 1
   If J < 1 Then mnuRecentSub(i).Visible = False
  End If
 Next i


Exit Sub

Command_OpenFile:
  If InStr(LCase(Command), "zip") > 0 Then
   ZipName = Trim(Mid(Command, 3, Len(Command)))
    If FSO.FileExists(ZipName) = False Then
     MsgBox Replace(LoadResString(MyLanguage + 128), "%1", ZipName), 16
    Else
     LVFiles.ListItems.Clear
     Zipit1.FileName = ZipName
     EMenus
    End If
  End If
Return


Command_ExtractFile:
  Dim LF As String, NF As String, BP As Boolean
  If InStr(1, LCase(Command), "-d") > 0 Then BP = False
  If InStr(1, LCase(Command), "-a") > 0 Then BP = True
  
  If InStr(LCase(Command), "zip") > 0 Then
   ZipName = Trim(Mid(Command, 3, Len(Command)))
    If FSO.FileExists(ZipName) = False Then
     MsgBox Replace(LoadResString(MyLanguage + 128), "%1", ZipName), 16
    Else
     Me.Visible = False
     Zipit1.FileName = ZipName
     If LVFiles.ListItems.Count = 0 Then
      MsgBox Replace(LoadResString(MyLanguage + 129), "%1", ZipName), 16
     End If
     
      Set O = New Collection
      For i = 1 To LVFiles.ListItems.Count
        O.Add LVFiles.ListItems(i).Text
      Next i
      
      If O.Count >= 1 Then
       If BP = False Then
        NF = GetFilePath(ZipName)
       Else
        LF = GetSetting(App.EXEName, "last", "dirextract", App.Path)
        If FSO.FolderExists(LF) = False Then LF = CurDir
        NF = OpenFolder(LF, Me.hWnd)
        If Trim(NF) = "" And FSO.FolderExists(NF) = False Then Encerrar
        NF = FixPath(NF)
        SaveSetting App.EXEName, "last", "dirextract", NF
       End If
        Screen.MousePointer = 11
        Zipit1.ExtractDir = NF
        Zipit1.Extract O
        Screen.MousePointer = 0
      End If
    End If
  End If
 Encerrar
Return
End Sub

Private Sub Encerrar()
On Error Resume Next
  Unload Me
  End
End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
On Local Error Resume Next
Dim hLIcon As Long, hSIcon As Long, imgObj As ListImage

hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

If hLIcon <> 0 Then
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    R = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With

  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    R = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub GetAllIcons()
Dim FileName As String
Dim DirName As String
Dim lNumber As Long

On Local Error Resume Next
  DirName = GetTmpPath
  For Each Item In LVFiles.ListItems
    FileName = FixPath(GetTmpPath) & Item.Text
    lNumber = FreeFile
    If FSO.FileExists(FileName) = False Then
     Open FileName For Output As lNumber: Close lNumber
     GetIcon FileName, Item.Index
     Kill FileName
    Else
     GetIcon FileName, Item.Index
    End If
  Next
End Sub

Private Sub ShowIcons()
On Local Error Resume Next
  With LVFiles
    .Icons = iml32
    .SmallIcons = iml16
    For Each Item In .ListItems
      Item.Icon = Item.Index
      Item.SmallIcon = Item.Index
    Next
  End With
End Sub

Private Sub ShowIconsForFiles()
On Local Error Resume Next
  With LVFiles
   .Icons = Nothing
   .SmallIcons = Nothing
  End With
  
  iml32.ListImages.Clear
  iml16.ListImages.Clear
  
  GetAllIcons
  ShowIcons
End Sub


Private Sub Form_Resize()
On Error Resume Next
 LVFiles.Move 0, TBar.Height, Me.ScaleWidth, Me.ScaleHeight - (TBar.Height + SBar.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSetting App.EXEName, "last", "view", LVFiles.View
 For R = 1 To LVFiles.ColumnHeaders.Count
  SaveSetting App.EXEName, "last", "colw" & R, LVFiles.ColumnHeaders(R).Width
 Next R
 End
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

Private Sub LVFiles_DblClick()
 If Not (LVFiles.SelectedItem Is Nothing) Then Call mnuActionSub_Click(4)
End Sub

Private Sub LVFiles_KeyDown(KeyCode As Integer, Shift As Integer)
 If LVFiles.ListItems.Count > 0 Then
  If KeyCode = vbKeyDelete Then Call mnuActionSub_Click(1)
 End If
End Sub

Private Sub LVFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then PopupMenu mnuActionTop
End Sub

Private Sub mnuActionSub_Click(Index As Integer)
Select Case Index
 Case 0
    Dim vFiles As Variant
    Dim lFile As Long
    Dim cFiles As New Collection
    With CDial
        On Error Resume Next
        .FileName = ""
        .CancelError = True
        .DialogTitle = LoadResString(MyLanguage + 130)
        .flags = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNHideReadOnly
        .Filter = LoadResString(MyLanguage + 148)
        .ShowOpen
        If Trim(.FileName) <> "" Then
         vFiles = Split(.FileName, Chr(0))
          If UBound(vFiles) = 0 Then
            cFiles.Add .FileName
          Else
            For lFile = 1 To UBound(vFiles)
              cFiles.Add vFiles(0) + "\" & vFiles(lFile)
            Next
          End If
        End If
        Screen.MousePointer = 11
        If cFiles.Count >= 1 Then Zipit1.Add cFiles
        Screen.MousePointer = 0
    End With
    
  Case 1
   If LVFiles.ListItems.Count = 0 Then Exit Sub
   Set O = New Collection
   For i = 1 To LVFiles.ListItems.Count
    If LVFiles.ListItems(i).Selected = True Then
     O.Add LVFiles.ListItems(i).Text
    End If
   Next i
   
   If O.Count >= 1 Then
    If MsgBox(IIf(O.Count = 1, LoadResString(MyLanguage + 131), Replace(LoadResString(MyLanguage + 132), "%1", O.Count)) & " ?", 292) = 6 Then
     Screen.MousePointer = 11
     Zipit1.Delete O
     Screen.MousePointer = 0
    End If
   End If
   
  Case 2
   Dim LF As String, NF As String
   If LVFiles.ListItems.Count = 0 Then Exit Sub
   
   Set O = New Collection
   For i = 1 To LVFiles.ListItems.Count
    If LVFiles.ListItems(i).Selected = True Then
     O.Add LVFiles.ListItems(i).Text
    End If
   Next i
   
   If O.Count >= 1 Then
      LF = GetSetting(App.EXEName, "last", "dir", App.Path)
      If FSO.FolderExists(LF) = False Then LF = CurDir
      NF = OpenFolder(LF, Me.hWnd)
      If Trim(NF) = "" And FSO.FolderExists(NF) = False Then Exit Sub
      NF = FixPath(NF)
      SaveSetting App.EXEName, "last", "dir", NF
      Zipit1.ExtractDir = NF
      Screen.MousePointer = 11
      Zipit1.Extract O
      Screen.MousePointer = 0
   End If
  
  Case 4
    Dim sFile As String, sExtract As String
    sFile = FixPath(GetTmpPath) & LVFiles.SelectedItem.Text
    If FSO.FileExists(sFile) = True Then FSO.DeleteFile sFile, True
    sExtract = Zipit1.ExtractDir
    Set O = New Collection
    O.Add LVFiles.SelectedItem.Text
    Zipit1.ExtractDir = GetTmpPath
    Zipit1.Extract O
    Zipit1.ExtractDir = sExtract
    Call ShellExecute(Me.hWnd, "Open", sFile, "", App.Path, 1)
   
End Select
End Sub

Private Sub mnuHelpSub_Click(Index As Integer)
Select Case Index
 Case 0
  Unload FrmAboutSimple
  FrmAboutSimple.Show 1
  Set FrmAboutSimple = Nothing
 Case 2
  Call ShellExecute(Me.hWnd, "Open", "http://www.gabrielfalcao.i8.com", "", App.Path, 0)
End Select
End Sub

Private Sub mnuRecentSub_Click(Index As Integer)
 If Trim(mnuRecentSub(Index).Caption) = "" Then Exit Sub
 If Dir(mnuRecentSub(Index).Caption) = "" Then Exit Sub
 
 ZipName = mnuRecentSub(Index).Caption
 Zipit1.FileName = ZipName
 EMenus
End Sub

Private Sub mnuViewSub_Click(Index As Integer)
 For i = 0 To mnuViewSub.Count - 1
  mnuViewSub(i).Checked = False
 Next i
 
 LVFiles.View = Index
 mnuViewSub(Index).Checked = True
End Sub

Private Sub mnuToolsSub_Click(Index As Integer)
Select Case Index
Case 0
    FrmProcurar.Show 1

Case 1
   Dim LF As String, NF As String, FI As File
   Set O = New Collection
   
      LF = GetSetting(App.EXEName, "last", "dir_add", App.Path)
      If FSO.FolderExists(LF) = False Then LF = CurDir
      NF = OpenFolder(LF, Me.hWnd)
      If Trim(NF) = "" Then Exit Sub
      NF = NF & IIf(Right(NF, 1) <> "\", "\", "")
      SaveSetting App.EXEName, "last", "dir_add", NF
      If FSO.FolderExists(NF) = True Then
        Status Replace(LoadResString(MyLanguage + 133), "%1", FSO.GetFolder(NF).Files.Count)
        If FSO.GetFolder(NF).Files.Count = 0 Then Exit Sub
         For Each FI In FSO.GetFolder(NF).Files
          O.Add FI.Path
         Next
         If Zipit1.FileName <> "" And FSO.FileExists(Zipit1.FileName) Then Zipit1.Add O
       End If
End Select
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
On Error Resume Next

Select Case Index
 Case 0
  With CDial
   .CancelError = True
   .flags = cdlOFNHideReadOnly
   .FileName = ""
   .DialogTitle = LoadResString(MyLanguage + 134)
   .Filter = LoadResString(MyLanguage + 135)
   .ShowSave
   If Trim(.FileName) <> "" Then
    LVFiles.ListItems.Clear
    ZipName = .FileName
    Zipit1.FileName = ZipName
    EMenus
    Call mnuActionSub_Click(0)
   End If
  End With
  
 Case 1
  With CDial
   .CancelError = True
   .flags = cdlOFNHideReadOnly
   .FileName = ""
   .DialogTitle = LoadResString(MyLanguage + 136)
   .Filter = LoadResString(MyLanguage + 135)
   .ShowOpen
   If Trim(.FileName) <> "" Then
    ZipName = .FileName
    Zipit1.FileName = ZipName
    EMenus
   End If
  End With
  
 Case 2
  LVFiles.ListItems.Clear
  Zipit1.FileName = ""
  EMenus

  
 Case 4
  FrmOpcoes.Show 1
  
End Select
End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case LCase(Button.Key)
 Case "new"
  Call mnuFileSub_Click(0)
 Case "open"
  Call mnuFileSub_Click(1)
 Case "close"
  Call mnuFileSub_Click(2)
 Case "options"
  Call mnuFileSub_Click(4)
 Case "add"
  Call mnuActionSub_Click(0)
 Case "del"
  Call mnuActionSub_Click(1)
 Case "extract"
  Call mnuActionSub_Click(2)
 Case "view"
  Call mnuActionSub_Click(4)
 Case "find"
  Call mnuToolsSub_Click(0)
 Case "addfolder"
  Call mnuToolsSub_Click(1)
End Select
End Sub

Private Sub Zipit1_Change()
    Set O = New Collection
    Dim Num As Long
    Dim FileName As String
    Dim Path As String
    Dim ret As Long
    Dim Files As New ZipFileEntry
    Dim Location As Long
    Dim ItmX As ListItem
    
    LVFiles.ListItems.Clear
    R = Zipit1.ZipFiles.Count
    Num = R
    For i = 1 To R
        Set Files = Zipit1.ZipFiles.Item(i)
        With Files
            Status Replace(LoadResString(MyLanguage + 160), "%1", .FileName)
            If Right(.FileName, 1) <> "/" Then
                Location = InStrRev(.FileName, "/", -1)
                Path = Left(.FileName, Location)
                FileName = Mid(.FileName, Location + 1)
                Set ItmX = LVFiles.ListItems.Add(, , FileName)
                    ItmX.SubItems(1) = FormatFileSize(.UncompressedSize)
                    ItmX.SubItems(2) = FormatFileSize(.CompressedSize)
                    ItmX.SubItems(3) = Format(100 - (.CompressedSize / .UncompressedSize * 100), "#,##0.00")
                    ItmX.SubItems(4) = .FileDateTime
                    ItmX.SubItems(5) = .CRC32
                    If ItmX.SubItems(3) < 30 Then
                     ItmX.ForeColor = &H808080
                     For J = 1 To 5
                      ItmX.ListSubItems(J).ForeColor = ItmX.ForeColor
                     Next J
                    End If
                O.Add FileName
            End If
        End With
    Next i
    Status Replace(LoadResString(MyLanguage + 137), "%1", LVFiles.ListItems.Count), 2
    If LVFiles.ListItems.Count > 0 Then ShowIconsForFiles
    If Trim(Zipit1.FileName) <> "" Then SaveRecent Zipit1.FileName
    Status ""
End Sub

Private Sub Zipit1_DeleteComplete(Successful As Long)
 PBar.Visible = False
 Status
End Sub

Private Sub Zipit1_DeleteProgress(Percentage As Integer, FileName As String)
 ShowPBar
 PBar.Value = Percentage
 Status Replace(LoadResString(MyLanguage + 138), "%1", FileName)
End Sub

Private Sub Zipit1_UnzipComplete(Successful As Long)
 PBar.Visible = False
 Status
End Sub

Private Sub Zipit1_UnzipProgress(Percentage As Integer, FileName As String)
 ShowPBar
 PBar.Value = Percentage
  Status Replace(LoadResString(MyLanguage + 139), "%1", FileName)
End Sub

Private Sub Zipit1_ZipComplete(Successful As Long)
 PBar.Visible = False
 Status
End Sub

Private Sub Zipit1_ZipProgress(Percentage As Integer, FileName As String)
 ShowPBar
 PBar.Value = Percentage
 Status Replace(LoadResString(MyLanguage + 140), "%1", FileName)
End Sub



Public Sub LoadCaption()
 With FrmMenu
  .mnuFileTop.Caption = LoadResString(MyLanguage + 100)
  .mnuFileSub(0).Caption = LoadResString(MyLanguage + 101)
  .mnuFileSub(1).Caption = LoadResString(MyLanguage + 102)
  .mnuFileSub(2).Caption = LoadResString(MyLanguage + 103)
  .mnuFileSub(4).Caption = LoadResString(MyLanguage + 104)
  .mnuFileSub(6).Caption = LoadResString(MyLanguage + 164)
  .mnuFileSub(8).Caption = LoadResString(MyLanguage + 105)
  
  
  .mnuActionTop.Caption = LoadResString(MyLanguage + 106)
  .mnuActionSub(0).Caption = LoadResString(MyLanguage + 107)
  .mnuActionSub(1).Caption = LoadResString(MyLanguage + 108)
  .mnuActionSub(2).Caption = LoadResString(MyLanguage + 109)
  .mnuActionSub(4).Caption = LoadResString(MyLanguage + 110)
  
  .mnuViewTop.Caption = LoadResString(MyLanguage + 111)
  .mnuViewSub(0).Caption = LoadResString(MyLanguage + 112)
  .mnuViewSub(1).Caption = LoadResString(MyLanguage + 113)
  .mnuViewSub(2).Caption = LoadResString(MyLanguage + 114)
  .mnuViewSub(3).Caption = LoadResString(MyLanguage + 115)
  
  .mnuToolsTop.Caption = LoadResString(MyLanguage + 116)
  .mnuToolsSub(0).Caption = LoadResString(MyLanguage + 117)
  .mnuToolsSub(1).Caption = LoadResString(MyLanguage + 118)
  
  .mnuHelpTop.Caption = LoadResString(MyLanguage + 156)
  .mnuHelpSub(0).Caption = LoadResString(MyLanguage + 157)
  .mnuHelpSub(2).Caption = LoadResString(MyLanguage + 158)
  
  .LVFiles.ColumnHeaders(1).Text = LoadResString(MyLanguage + 150)
  .LVFiles.ColumnHeaders(2).Text = LoadResString(MyLanguage + 151)
  .LVFiles.ColumnHeaders(3).Text = LoadResString(MyLanguage + 152)
  .LVFiles.ColumnHeaders(4).Text = LoadResString(MyLanguage + 153)
  .LVFiles.ColumnHeaders(5).Text = LoadResString(MyLanguage + 154)
  .LVFiles.ColumnHeaders(6).Text = LoadResString(MyLanguage + 155)
  
  .TBar.Buttons("new").ToolTipText = Replace(LoadResString(MyLanguage + 101), "&", "")
  .TBar.Buttons("open").ToolTipText = Replace(LoadResString(MyLanguage + 102), "&", "")
  .TBar.Buttons("close").ToolTipText = Replace(LoadResString(MyLanguage + 103), "&", "")
  .TBar.Buttons("options").ToolTipText = Replace(LoadResString(MyLanguage + 104), "&", "")
  .TBar.Buttons("add").ToolTipText = Replace(LoadResString(MyLanguage + 107), "&", "")
  .TBar.Buttons("del").ToolTipText = Replace(LoadResString(MyLanguage + 108), "&", "")
  .TBar.Buttons("extract").ToolTipText = Replace(LoadResString(MyLanguage + 109), "&", "")
  .TBar.Buttons("view").ToolTipText = Replace(LoadResString(MyLanguage + 110), "&", "")
  .TBar.Buttons("find").ToolTipText = Replace(LoadResString(MyLanguage + 117), "&", "")
  .TBar.Buttons("addfolder").ToolTipText = Replace(LoadResString(MyLanguage + 118), "&", "")
 
 End With
End Sub

Public Sub EMenus()
 Me.Caption = LoadResString(MyLanguage + 161) & IIf(Zipit1.FileName = "", "", " [" & GetFileTitle(Zipit1.FileName) & "]")
 
 mnuFileSub.Item(2).Enabled = Not (Zipit1.FileName = "")
 mnuActionSub(0).Enabled = Not (Zipit1.FileName = "")
 mnuActionSub(1).Enabled = Not (Zipit1.FileName = "")
 mnuActionSub(2).Enabled = Not (Zipit1.FileName = "")
 mnuActionSub(4).Enabled = Not (Zipit1.FileName = "")
 
 mnuToolsSub(0).Enabled = Not (Zipit1.FileName = "")
 mnuToolsSub(1).Enabled = Not (Zipit1.FileName = "")
 
 TBar.Buttons("close").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("add").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("del").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("extract").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("view").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("find").Enabled = Not (Zipit1.FileName = "")
 TBar.Buttons("addfolder").Enabled = Not (Zipit1.FileName = "")
End Sub


Private Sub SaveRecent(sFilename As String)

 For i = 1 To 3
  If LCase(sFilename) = LCase(GetSetting(App.EXEName, "recent", "file" & i, "")) Then Exit Sub
 Next i
 
 SaveSetting App.EXEName, "recent", "file3", GetSetting(App.EXEName, "recent", "file2", "")
 SaveSetting App.EXEName, "recent", "file2", GetSetting(App.EXEName, "recent", "file1", "")
 SaveSetting App.EXEName, "recent", "file1", sFilename
 
 For i = 0 To 2
  mnuRecentSub(i).Caption = GetSetting(App.EXEName, "recent", "file" & i + 1, "")
  If Trim(mnuRecentSub(i).Caption) = "" Then
   mnuRecentSub(i).Visible = False
  End If
 Next i
End Sub
