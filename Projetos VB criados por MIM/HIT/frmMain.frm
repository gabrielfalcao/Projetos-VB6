VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00000000&
   Caption         =   "H1T - H4ck3R & 1Nt3rn3t T00lz"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   -465
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1515
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2625
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir MP3"
      Filter          =   "Arquivos MP3|*.mp3"
      InitDir         =   "C:\"
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   11850
      TabIndex        =   1
      Top             =   7635
      Visible         =   0   'False
      Width           =   11880
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abrir MP3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   915
         TabIndex        =   4
         Top             =   75
         Width           =   1365
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4155
         TabIndex        =   3
         Top             =   90
         Width           =   4350
      End
      Begin VB.Image imgPause 
         Height          =   300
         Index           =   0
         Left            =   3450
         Picture         =   "frmMain.frx":24090C
         Top             =   45
         Width           =   330
      End
      Begin VB.Image imgStop 
         Height          =   300
         Index           =   0
         Left            =   3780
         Picture         =   "frmMain.frx":240E9E
         Top             =   45
         Width           =   330
      End
      Begin VB.Image imgPlay 
         Height          =   300
         Index           =   0
         Left            =   2400
         Picture         =   "frmMain.frx":241430
         Top             =   45
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP3 Player"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   45
         TabIndex        =   2
         Top             =   90
         Width           =   795
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   8025
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10160
            Text            =   "Pronto"
            TextSave        =   "Pronto"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "30/11/04"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:56"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Fechar"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Salvar &Como..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimir..."
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
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "S&air"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Desfazer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Re&cortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Colar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Ferramentas"
      Begin VB.Menu mnuInvader 
         Caption         =   "SOCKET"
         Shortcut        =   {F2}
      End
      Begin VB.Menu xtrac 
         Caption         =   "Extrator de Resources"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPlaya 
         Caption         =   "MP3 Player"
         Shortcut        =   {F4}
      End
      Begin VB.Menu enab 
         Caption         =   "Super Habilitador"
         Shortcut        =   {F5}
      End
      Begin VB.Menu pass 
         Caption         =   "Password Cracker"
         Shortcut        =   {F6}
      End
      Begin VB.Menu criphex 
         Caption         =   "Encriptador/Decriptador Hexadecimal"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuLang 
      Caption         =   "Linguagem"
      Begin VB.Menu mnuC 
         Caption         =   "C/C++"
      End
      Begin VB.Menu mnuBasic 
         Caption         =   "Basic"
      End
      Begin VB.Menu mnuPascal 
         Caption         =   "Pascal"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "HTML"
      End
      Begin VB.Menu mnuXML 
         Caption         =   "XML"
      End
      Begin VB.Menu mnuJava 
         Caption         =   "Java"
      End
      Begin VB.Menu mnuSQL 
         Caption         =   "SQL"
      End
   End
   Begin VB.Menu mnuViewWebBrowser 
      Caption         =   "&Web Browser"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Organizar"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Em &Cascata"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Alinhar na Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Alinhar na Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Organizar Ícones"
      End
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "&Sobre..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim Percent As Integer
Dim i As Integer
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private mmOpen As String, sec As Integer, mins As Integer
Dim nFileName As String
Public Function MP3Play(wndHandle As Long, sFileName As String)
  Dim cmdToDo As String * 255
  Dim dwReturn As Long
  Dim ret As String * 128
  Dim tmp As String * 255
  Dim lenShort As Long
  Dim ShortPathAndFie As String, glo_HWND As Long
    If Dir(sFileName) = "" Then
         mmOpen = "Error with input file"
         Exit Function
    End If
  lenShort = GetShortPathName(sFileName, tmp, 255)
  ShortPathAndFie = Left$(tmp, lenShort)
  glo_HWND = wndHandle
  cmdToDo = "open " & ShortPathAndFie & " type MPEGVideo Alias MP3Play"
  dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
  If dwReturn <> 0 Then  'not success
     mciGetErrorString dwReturn, ret, 128
     mmOpen = ret
     MsgBox ret, vbCritical
     Exit Function
  End If
  mmOpen = "Success"
  mciSendString "play MP3Play", 0, 0, 0
End Function

Private Sub Open_file()
Dim cderr As Long
cd.ShowOpen
'************
DoEvents
If cd.Filename <> Empty Then
MP3Play Me.hwnd, cd.Filename
Label1.Caption = cd.Filename
End If


End Sub

Public Function IsPlaying() As Boolean
Static s As String * 30
mciSendString "status MP3Play mode", s, Len(s), 0
IsPlaying = (Mid$(s, 1, 7) = "playing")
End Function

Public Function MP3Pause()
  mciSendString "pause MP3Play", 0, 0, 0
End Function

Public Function MP3UnPause()
  mciSendString "play MP3Play", 0, 0, 0
End Function

Public Function MP3Stop() As String
  mciSendString "stop MP3Play", 0, 0, 0
  mciSendString "close MP3Play", 0, 0, 0
End Function
Public Function Ups(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    'Dim Num$
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

Private Sub criphex_Click()
frmCrip.Show
End Sub

Private Sub enab_Click()
frmEnab.Show
End Sub

Private Sub imgPause_Click(Index As Integer)
 MP3Pause
End Sub

Private Sub imgPlay_Click(Index As Integer)
On Error Resume Next

MP3UnPause
End Sub

Private Sub imgStop_Click(Index As Integer)
MP3Stop
End Sub

Private Sub Label3_Click()
Open_file
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HFFFFFF
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HC0C0C0
End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)

End Sub


Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Código Fonte " & lDocumentCount
    frmD.Show
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MP3Stop
End
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
Label1.Width = Me.Width - 4500
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
    End If
    End
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            ActiveForm.rtftext.SelBold = Not ActiveForm.rtftext.SelBold
            Button.Value = IIf(ActiveForm.rtftext.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtftext.SelItalic = Not ActiveForm.rtftext.SelItalic
            Button.Value = IIf(ActiveForm.rtftext.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtftext.SelUnderline = Not ActiveForm.rtftext.SelUnderline
            Button.Value = IIf(ActiveForm.rtftext.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtftext.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtftext.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtftext.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuFind_Click()

End Sub

Private Sub mnuBasic_Click()
ActiveForm.rtftext.Language = mnuBasic.Caption
End Sub

Private Sub mnuC_Click()
On Error Resume Next
ActiveForm.rtftext.Language = mnuC.Caption
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuHTML_Click()
ActiveForm.rtftext.Language = mnuHTML.Caption
End Sub

Private Sub mnuInvader_Click()
frmInvasor.Show
End Sub

Private Sub mnuJava_Click()
ActiveForm.rtftext.Language = mnuJava.Caption
End Sub

Private Sub mnuPascal_Click()
ActiveForm.rtftext.Language = mnuPascal.Caption
End Sub

Private Sub mnuPlaya_Click()
If mnuPlaya.Checked = True Then
mnuPlaya.Checked = False
MP3Stop
Picture2.Visible = False
Else
mnuPlaya.Checked = True
Picture2.Visible = True
End If
End Sub

Private Sub mnuSQL_Click()
ActiveForm.rtftext.Language = mnuSQL.Caption
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewWebBrowser_Click()
  
    frmBrowser.StartingAddress = "http://www.tebugho.i8.com"
    frmBrowser.Show
End Sub

Private Sub mnuViewOptions_Click()
frmOptions.Show
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtftext.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF
    ActiveForm.rtftext.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
On Error Resume Next
ActiveForm.rtftext.Undo

End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtftext.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtftext.SelPrint .hdc
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Todos os Arquivos (*.*)|*.*"
        .ShowSave
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtftext.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
On Error GoTo 10
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Código F" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "All Files (*.*)|*.*"
            .ShowSave
            If Len(.Filename) = 0 Then
                Exit Sub
            End If
            sFile = .Filename
        End With
        ActiveForm.rtftext.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtftext.SaveFile sFile, True
    End If
10 Exit Sub
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
    Dim sFile As String


    If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Abrir arquivo"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Todos os Códigos Fonte|*.asm;*.c;*.h;*.xml;*.java;*.html;*.htm;*.php;*.txt;*.ini|Códigos Fonte Assembly (*.asm)|*.asm|Códigos Fonte C++ (*.c,*.h)|*.c;*.h|Códigos Fonte XML (*.xml)|*.xml|Códigos Fonte HTML (*.htm,*.html)|*.htm;*.html|Códigos Fonte Java (*.java)|*.java|Códigos Fonte PHP (*.php)|*.php|Arquivos de Texto (*.txt)|*.txt|Arquivos de Inicialização (*.ini)|*.ini|Todos os Arquivos (*.*)|*.*"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    End With
    ActiveForm.rtftext.OpenFile sFile
    ActiveForm.Caption = sFile
If Left$(sFile, 3) = "xml" Then ActiveForm.rtftext.Language = "XML"
If Left$(sFile, 1) = "c" Or Left$(sFile, 1) = "h" Then ActiveForm.rtftext.Language = "C/C++"
If Left$(sFile, 3) = "tml" Or Left$(sFile, 3) = "htm" Then ActiveForm.rtftext.Language = "HTML"
If Left$(sFile, 3) = "bas" Then ActiveForm.rtftext.Language = "Basic"
If Left$(sFile, 3) = "txt" Or Left$(sFile, 3) = "ini" Then ActiveForm.rtftext.Language = "HTML"
If Left$(sFile, 3) = "ava" Or Left$(sFile, 3) = "ini" Then ActiveForm.rtftext.Language = "Java"
If Left$(sFile, 3) = "sql" Or Left$(sFile, 3) = "ini" Then ActiveForm.rtftext.Language = "SQL"
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Private Sub mnuXML_Click()
ActiveForm.rtftext.Language = mnuXML.Caption
End Sub

Private Sub pass_Click()
frmCrack.Show
End Sub


Private Sub xtrac_Click()
Form1.Show
End Sub
