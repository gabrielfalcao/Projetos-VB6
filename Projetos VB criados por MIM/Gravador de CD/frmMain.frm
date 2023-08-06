VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BurnIT!"
   ClientHeight    =   5835
   ClientLeft      =   255
   ClientTop       =   720
   ClientWidth     =   9225
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9225
   Begin MSComctlLib.ListView lvnew 
      Height          =   4800
      Left            =   4620
      TabIndex        =   9
      Top             =   705
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   8467
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15264511
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome do Arquivo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tamanho"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data de Modificação"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvfiles 
      Height          =   4800
      Left            =   0
      TabIndex        =   3
      Top             =   690
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   8467
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16710366
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome do Arquivo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tamanho"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data de Modificação"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   2190
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   2055
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1050
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":174E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   2070
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3945
      TabIndex        =   4
      Top             =   1500
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9225
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   9225
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   4575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CD de Destino:"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4635
         TabIndex        =   8
         Tag             =   " ListView:"
         Top             =   15
         Width           =   4545
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2835
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5565
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11086
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "20/06/04"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "00:40"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1755
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F058
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F59A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2001E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20130
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20A0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "burn"
            Object.ToolTipText     =   "Gravar CD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Deletar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Gnome Icon 39"
            Object.ToolTipText     =   "Opções"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":212E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23666
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28E58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivos"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Proprie&dades"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Fechar"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Desfazer"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Recor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "C&olar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Selecionar &Tudo"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Visualizar"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Barra de Ferramentas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barra de Sta&tus"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Ícones G&randes"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Ícones P&equenos"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Lista"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Detalhes"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Organizar Íco&nes"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Atualizar"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Opções..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Apagar CD Regravável..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
  Dim NameOfFile As String
Dim clmX As ColumnHeader
Dim itmX As ListItem
Dim Counter As Long
Dim dname As String
Dim TempDname As String
Dim counter2 As Integer
Dim CurrentDir As String
Dim mbMoving As Boolean
Dim Item As ListItem
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

Private ShInfo As SHFILEINFO
Private FSO As FileSystemObject
Dim xpos As Long, ypos As Long
Const sglSplitLimit = 500

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
refreshLV
End Sub

Private Sub Form_Load()


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2




' Create an object variable for the ColumnHeader object.
' Add ColumnHeaders.  The width of the columns is the width
' of the control divided by the number of ColumnHeader objects.


    
lvfiles.BorderStyle = ccFixedSingle ' Set BorderStyle property.
    
' To use ImageList controls with the ListView control, you must
' associate a particular ImageList control with the Icons and
' Icons were previously Added to list

' SmallIcons properties.
refreshLV
End Sub

Private Sub refreshLV()
lvfiles.ListItems.Clear
lvfiles.Icons = ImageList1
lvfiles.SmallIcons = ImageList2

'Start Off With Current Drive and directory
ChDrive Drive1.Drive
Dir1.Path = CurDir

'Add BackSlash if Necessary
'If Right(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir & "\"
'Dir1.Path = CurrentDir
'Dir1.Path = Drive1.Drive

'NameOfFile = Dir$(CurrentDir & "*.*", vbDirectory)
Dim Fname As String

'If we are in a subdirectory then do the following
If Right(Dir1.Path, 1) <> "\" Then
    CurrentDir = Dir1.Path & "\"
    dname = ".."
    Set itmX = lvfiles.ListItems.Add(, , dname)
    itmX.SubItems(1) = ""
    itmX.Icon = 3           ' Set an icon from ImageList1.
    itmX.SmallIcon = 3      ' Set an icon from ImageList2.
    itmX.SubItems(2) = ""
Else
    'If not in a subdirectory then do the following
    CurrentDir = Dir1.Path
End If

'Get the FileNames
For Counter = 0 To File1.ListCount - 1
    Fname = File1.List(Counter)
    Set itmX = lvfiles.ListItems.Add(, , Fname)
    itmX.SubItems(1) = CStr(FileLen(CurrentDir & Fname)) & "KB"
    itmX.Icon = 2           ' Set an icon from ImageList1.
    itmX.SmallIcon = 2      ' Set an icon from ImageList2.
    itmX.SubItems(2) = FileDateTime(CurrentDir & Fname)
Next Counter

'Get the Directory Names
For Counter = 0 To Dir1.ListCount - 1
    dname = Dir1.List(Counter)
    For counter2 = Len(dname) To 1 Step -1
        If Mid$(dname, counter2, 1) = "\" Then
            TempDname = Right(dname, Len(dname) - counter2)
            Exit For
        End If
    Next counter2
    Set itmX = lvfiles.ListItems.Add(, , TempDname)
    itmX.SubItems(1) = ""
    itmX.Icon = 1           ' Set an icon from ImageList1.
    itmX.SmallIcon = 1      ' Set an icon from ImageList2.
    itmX.SubItems(2) = FileDateTime(dname)
Next Counter



lvfiles.Arrange = 0 'lvwNoArrange
lvfiles.LabelWrap = False
lvfiles.Sorted = True
sbStatusBar.Panels(1).Text = CurrentDir


End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
  
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 1500 Then X = 1500
    If X > (Me.Width - 1500) Then X = Me.Width - 1500
    tvTreeView.Width = X
    imgSplitter.Left = X
    lvListView.Left = X + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40


    'set the top
  

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If

  lvListView.Top = tvTreeView.Top
    

    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    

    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub



Private Sub lvfiles_DblClick()

If lvfiles.HitTest(xpos, ypos) Is Nothing Then
    Exit Sub
Else
    Set Item = lvfiles.HitTest(xpos, ypos)
End If
Set Item = lvfiles.SelectedItem



If Item.SmallIcon = "2" Then
Call ShellExecute(Me.hWnd, "Open", Item, "", CurDir, 1)
End If
'If you Click on a filename just exit this subroutine
If Right(Dir1.Path, 1) <> "\" Then
    CurrentDir = Dir1.Path & "\"
Else
    CurrentDir = Dir1.Path
End If



If (GetAttr(CurrentDir & Item) And vbDirectory) <= 0 Then Exit Sub
lvfiles.ListItems.Clear 'Clear Out Old Items

'Change to selected Directory - Let Visual Basic do the work
ChDir Item

'Change the Directory List Box to equal the new Current Directory
Dir1.Path = CurDir

Dim Fname As String

'If we are in a subdirectory then add the backup ".." directory name
If Right(Dir1.Path, 1) <> "\" Then
    CurrentDir = Dir1.Path & "\"
    dname = ".."
    Set itmX = lvfiles.ListItems.Add(, , dname)
    itmX.SubItems(1) = ""
    itmX.Icon = 3           ' Set an icon from ImageList1.
    itmX.SmallIcon = 3      ' Set an icon from ImageList2.
    itmX.SubItems(2) = ""
Else
    'If we are not in a subdirectory then just set our temporary Directory variable
    CurrentDir = Dir1.Path
End If

'Start adding Filenames to ListView
For Counter = 0 To File1.ListCount - 1
    Fname = File1.List(Counter)
    Set itmX = lvfiles.ListItems.Add(, , Fname)
    itmX.SubItems(1) = CStr(FileLen(CurrentDir & Fname)) & "KB"
    itmX.Icon = 2           ' Set an icon from ImageList1.
    itmX.SmallIcon = 2      ' Set an icon from ImageList2.
    itmX.SubItems(2) = FileDateTime(CurrentDir & Fname)
Next Counter

'Add directory names to ListView
For Counter = 0 To Dir1.ListCount - 1
    dname = Dir1.List(Counter)
    'Get the actual directory name, not the full path and directory
    For counter2 = Len(dname) To 1 Step -1
        If Mid$(dname, counter2, 1) = "\" Then
            TempDname = Right(dname, Len(dname) - counter2)
            Exit For
        End If
    Next counter2
    Set itmX = lvfiles.ListItems.Add(, , TempDname)
    itmX.SubItems(1) = ""
    itmX.Icon = 1           ' Set an icon from ImageList1.
    itmX.SmallIcon = 1      ' Set an icon from ImageList2.
    itmX.SubItems(2) = FileDateTime(dname)
Next Counter
sbStatusBar.Panels(1).Text = CurrentDir

End Sub

Private Sub lvfiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y
End Sub

Private Sub mnuListViewMode_Click(Index As Integer)
Select Case Index
Case 0
lvfiles.View = lvwIcon
Case 1
lvfiles.View = lvwSmallIcon

Case 2
lvfiles.View = lvwList

Case 3
lvfiles.View = lvwReport
End Select
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            'ToDo: Add 'New' button code.
            MsgBox "Add 'New' button code."
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            'ToDo: Add 'Save' button code.
            MsgBox "Add 'Save' button code."
        Case "burn"
            'ToDo: Add 'burn' button code.
            MsgBox "Add 'burn' button code."
        Case "Delete"
            mnuFileDelete_Click
        Case "Gnome Icon 39"
            'ToDo: Add 'Gnome Icon 39' button code.
            MsgBox "Add 'Gnome Icon 39' button code."
    End Select
End Sub

Private Sub mnuToolsOptions_Click()
    'ToDo: Add 'mnuToolsOptions_Click' code.
    MsgBox "Add 'mnuToolsOptions_Click' code."
End Sub

Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
    'ToDo: Add 'mnuViewOptions_Click' code.
    MsgBox "Add 'mnuViewOptions_Click' code."
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub



Private Sub mnuVAIByDate_Click()
    'ToDo: Add 'mnuVAIByDate_Click' code.
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'ToDo: Add 'mnuVAIByName_Click' code.
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'ToDo: Add 'mnuVAIBySize_Click' code.
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'ToDo: Add 'mnuVAIByType_Click' code.
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
  
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked

End Sub

Private Sub mnuEditInvertSelection_Click()
    'ToDo: Add 'mnuEditInvertSelection_Click' code.
    MsgBox "Add 'mnuEditInvertSelection_Click' code."
End Sub

Private Sub mnuEditSelectAll_Click()
    'ToDo: Add 'mnuEditSelectAll_Click' code.
    MsgBox "Add 'mnuEditSelectAll_Click' code."
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    'ToDo: Add 'mnuEditPaste_Click' code.
    MsgBox "Add 'mnuEditPaste_Click' code."
End Sub

Private Sub mnuEditCopy_Click()
    'ToDo: Add 'mnuEditCopy_Click' code.
    MsgBox "Add 'mnuEditCopy_Click' code."
End Sub

Private Sub mnuEditCut_Click()
    'ToDo: Add 'mnuEditCut_Click' code.
    MsgBox "Add 'mnuEditCut_Click' code."
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileRename_Click()
    'ToDo: Add 'mnuFileRename_Click' code.
    MsgBox "Add 'mnuFileRename_Click' code."
End Sub

Private Sub mnuFileDelete_Click()
    'ToDo: Add 'mnuFileDelete_Click' code.
    MsgBox "Add 'mnuFileDelete_Click' code."
End Sub

Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    MsgBox "Add 'mnuFileNew_Click' code."
End Sub

Private Sub mnuFileSendTo_Click()
    'ToDo: Add 'mnuFileSendTo_Click' code.
    MsgBox "Add 'mnuFileSendTo_Click' code."
End Sub

Private Sub mnuFileFind_Click()
    'ToDo: Add 'mnuFileFind_Click' code.
    MsgBox "Add 'mnuFileFind_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

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
  For Each Item In lvfiles.ListItems
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

