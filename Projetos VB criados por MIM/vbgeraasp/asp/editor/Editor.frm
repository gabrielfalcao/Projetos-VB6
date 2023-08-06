VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmEditor 
   BackColor       =   &H8000000C&
   Caption         =   "Editor"
   ClientHeight    =   5640
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7515
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar BarraDeFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "Ícones"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   25
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Find"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Scratched out"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Left"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Center"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Right"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.ComboBox cboTamanhoDaFonte 
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Text            =   "cboTamanhoDaFonte"
         ToolTipText     =   "Font"
         Top             =   30
         Width           =   780
      End
   End
   Begin MSComDlg.CommonDialog cmmArquivo 
      Left            =   5760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timBarraDeStatus 
      Interval        =   1000
      Left            =   6360
      Top             =   600
   End
   Begin ComctlLib.StatusBar BarraDeStatus 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Text            =   "Status"
            TextSave        =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Information"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status of Caps Lock: On or Off"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            TextSave        =   "NUM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status of Num Lock: On or Of"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Object.ToolTipText     =   "Hour"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Ícones 
      Left            =   6840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0448
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":055A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":066C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":09A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":1120
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&File"
      Begin VB.Menu mnuNovo 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSalvar 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuEncr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "&View"
      Begin VB.Menu mnuBarraDeFerramentas 
         Caption         =   "&Tool Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBarraDeStatus 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Edit"
      Begin VB.Menu mnuRecortar 
         Caption         =   "&Cut    Ctrl + X"
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copy        Ctrl + C"
      End
      Begin VB.Menu mnuColar 
         Caption         =   "Paste          Ctrl + V"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocalizar 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuInserir 
      Caption         =   "&Insert"
      Begin VB.Menu mnuFigura 
         Caption         =   "Picture"
      End
   End
   Begin VB.Menu mnuFormatar 
      Caption         =   "&Format"
      Begin VB.Menu mnuFonte 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuParágrafo 
         Caption         =   "&Paragraph..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMudarData 
         Caption         =   "Change Date System..."
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascata 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuLadoALadoHorizontalmente 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuLadoALadoVerticalmente 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuOrganizarÍcones 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuSobreOEditor 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuVoltar 
      Caption         =   "&Return"
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BarraDeFerramentas_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1
mnuNovo_Click
Case 2
mnuAbrir_Click
Case 3
mnuSalvar_Click
Case 5
mnuImprimir_Click
Case 6
mnuLocalizar_Click
Case 8
mnuRecortar_Click
Case 9
mnuCopiar_Click
Case 10
mnuColar_Click
Case 12
If frmEditor.ActiveForm.rtbDocumento.SelBold = False Then
frmEditor.ActiveForm.rtbDocumento.SelBold = True
Else
frmEditor.ActiveForm.rtbDocumento.SelBold = False
End If
Case 13
If frmEditor.ActiveForm.rtbDocumento.SelItalic = False Then
frmEditor.ActiveForm.rtbDocumento.SelItalic = True
Else
frmEditor.ActiveForm.rtbDocumento.SelItalic = False
End If
Case 14
If frmEditor.ActiveForm.rtbDocumento.SelUnderline = False Then
frmEditor.ActiveForm.rtbDocumento.SelUnderline = True
Else
frmEditor.ActiveForm.rtbDocumento.SelUnderline = False
End If
Case 15
If frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = False Then
frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = True
Else
frmEditor.ActiveForm.rtbDocumento.SelStrikeThru = False
End If
Case 23
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfLeft
Case 24
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfCenter
Case 25
frmEditor.ActiveForm.rtbDocumento.SelAlignment = rtfRight
End Select
End Sub

Private Sub cboTamanhoDaFonte_Click()
On Error Resume Next
frmEditor.ActiveForm.rtbDocumento.SelFontSize = cboTamanhoDaFonte.Text
Me.ActiveForm.SetFocus
End Sub

Private Sub MDIForm_Load()
Dim Contador
For Contador = 6 To 120
cboTamanhoDaFonte.AddItem (Contador)
Next
Me.Show
mnuNovo_Click
cboTamanhoDaFonte.Text = frmEditor.ActiveForm.rtbDocumento.SelFontSize
End Sub

Private Sub mnuAbrir_Click()
BarraDeStatus.Panels.Item(1).Text = "Wait to open file..."
Dim Documento As New frmModelo
cmmArquivo.DialogTitle = "Open"
cmmArquivo.InitDir = strPathEdit
cmmArquivo.Filter = "Active Server Pages (*.asp)|*.asp|Formato Rich Text (*.rtf)|*.rtf|Arquivos de texto (*.txt)|*.txt|Arquivos de lote (*.bat)|*.bat|Arquivos de Inicialização (*.ini)|*.ini|Todos os arquivos (*.*)|*.*|"
cmmArquivo.ShowOpen
If cmmArquivo.FileName <> "" Then
Documento.rtbDocumento.FileName = cmmArquivo.FileName
Documento.Caption = Documento.rtbDocumento.FileName
BarraDeStatus.Panels.Item(1).Text = "File is open"
cmmArquivo.FileName = ""
Exit Sub
Else
BarraDeStatus.Panels.Item(1).Text = "Unable to Open the File"
End If
End Sub



Private Sub mnuBarraDeFerramentas_Click()
If mnuBarraDeFerramentas.Checked = True Then
mnuBarraDeFerramentas.Checked = False
BarraDeFerramentas.Visible = False
Else
mnuBarraDeFerramentas.Checked = True
BarraDeFerramentas.Visible = True
End If
End Sub

Private Sub mnuBarraDeStatus_Click()
If mnuBarraDeStatus.Checked = True Then
mnuBarraDeStatus.Checked = False
BarraDeStatus.Visible = False
Else
mnuBarraDeStatus.Checked = True
BarraDeStatus.Visible = True
End If
End Sub

Private Sub mnuCascata_Click()
Me.Arrange vbCascade
End Sub

Private Sub mnuColar_Click()
frmEditor.ActiveForm.rtbDocumento.SelText = Clipboard.GetText
End Sub




Private Sub mnuCopiar_Click()
Clipboard.Clear
Clipboard.SetText frmEditor.ActiveForm.rtbDocumento.SelText
End Sub

Private Sub mnuEncr_Click()
    Load frmEncrypt
    frmEncrypt.Show
End Sub

Private Sub mnuFechar_Click()
Unload frmEditor.ActiveForm
End Sub

Private Sub mnuFigura_Click()
Load frmInserirFigura
frmInserirFigura.Show
Me.Enabled = False
End Sub

Private Sub mnuFonte_Click()
Load frmFormatarFontes
frmFormatarFontes.Show
Me.Enabled = False
End Sub

Private Sub mnuImprimir_Click()
On Error GoTo Erro
PrintRTF frmEditor.ActiveForm.rtbDocumento, 1440, 1440, 1440, 1440
Exit Sub
Erro:
MsgBox "Verify if Printer is off; Or without paper; Don't exists printer installed; Don't exists opened document", vbOKOnly, "Error in printing"
End Sub

Private Sub mnuLadoALadoHorizontalmente_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuLadoALadoVerticalmente_Click()
Me.Arrange vbTileVertical
End Sub

Private Sub mnuLocalizar_Click()
Load frmLocalizar
frmLocalizar.Show
Me.Enabled = False
End Sub

Private Sub mnuMudarData_Click()
Load frmMudarDataDoSistema
frmMudarDataDoSistema.Show
Me.Enabled = False
End Sub

Private Sub mnuNovo_Click()
BarraDeStatus.Panels.Item(1).Text = "Making a template of the document..."
Static NúmeroDoDocumento As Integer
NúmeroDoDocumento = NúmeroDoDocumento + 1
Dim Documento As New frmModelo
Documento.Caption = "Document " & NúmeroDoDocumento
Load Documento
Documento.Show
BarraDeStatus.Panels.Item(1).Text = "Document created"
End Sub

Private Sub mnuOrganizarÍcones_Click()
Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuParágrafo_Click()
Load frmFormatarParágrafo
frmFormatarParágrafo.Show
Me.Enabled = False
End Sub

Private Sub mnuRecortar_Click()
Clipboard.Clear
Clipboard.SetText frmEditor.ActiveForm.rtbDocumento.SelText
frmEditor.ActiveForm.rtbDocumento.SelText = ""
End Sub


Private Sub mnuSair_Click()
End
End Sub

Private Sub mnuSalvar_Click()
BarraDeStatus.Panels.Item(1).Text = "Wait, saving file..."
cmmArquivo.DialogTitle = "Save"
cmmArquivo.Filter = "Active Server Pages (*.asp)|*.asp|Rich Text Format (*.rtf)|*.rtf|Arquivos de texto (*.txt)|*.txt|Arquivos de lote (*.bat)|*.bat|Ini File (*.ini)|*.ini|"
cmmArquivo.ShowSave
If cmmArquivo.FileName <> "" Then
If cmmArquivo.FilterIndex = 1 Then
Salvar = frmEditor.ActiveForm.rtbDocumento.SaveFile(cmmArquivo.FileName, 0)
frmEditor.ActiveForm.Caption = cmmArquivo.FileName
BarraDeStatus.Panels.Item(1).Text = "File Save in Rich Text Format"
cmmArquivo.FileName = ""
Else
Salvar = frmEditor.ActiveForm.rtbDocumento.SaveFile(cmmArquivo.FileName, 1)
frmEditor.ActiveForm.Caption = cmmArquivo.FileName
BarraDeStatus.Panels.Item(1).Text = "File Save in Text Format"
cmmArquivo.FileName = ""
End If
End If
Exit Sub
Erro:
BarraDeStatus.Panels.Item(1).Text = "Error in saving file"
MsgBox "Unable to save file." & Chr(13) & "Verify free space in disk and if exists file to save."
End Sub

Private Sub mnuSegurança_Click()
Me.Enabled = False
Load frmEncriptaDesencripta
frmEncriptaDesencripta.Show
End Sub

Private Sub mnuSobreOEditor_Click()
Me.Enabled = False
Load frmSobre
frmSobre.Show
End Sub

Private Sub mnuVoltar_Click()
    Unload frmEditor
End Sub

Private Sub timBarraDeStatus_Timer()
BarraDeStatus.Panels.Item(4).Text = Time
If BarraDeStatus.Panels.Item(5).Text <> Date Then
BarraDeStatus.Panels.Item(5).Text = Date
End If
End Sub
