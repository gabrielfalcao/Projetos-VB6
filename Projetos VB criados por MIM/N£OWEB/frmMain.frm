VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "GNavigator :: powered by Internet Explorer"
   ClientHeight    =   4920
   ClientLeft      =   270
   ClientTop       =   450
   ClientWidth     =   9060
   FillColor       =   &H00849B99&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":2CFA
   ScaleHeight     =   4920
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   " IR"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6855
      TabIndex        =   4
      Top             =   1605
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "PARAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7470
      TabIndex        =   3
      Top             =   1605
      Width           =   870
   End
   Begin VB.ComboBox txtInput 
      Height          =   315
      Left            =   930
      TabIndex        =   2
      Top             =   1575
      Width           =   5925
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   1140
      Picture         =   "frmMain.frx":BE53C
      ScaleHeight     =   1530
      ScaleWidth      =   6750
      TabIndex        =   1
      Top             =   0
      Width           =   6750
      Begin VB.Image btnHome 
         Height          =   1530
         Left            =   5220
         ToolTipText     =   "Vai para a página inicial padrão"
         Top             =   0
         Width           =   1530
      End
      Begin VB.Image btnRefresh 
         Height          =   1530
         Left            =   4095
         ToolTipText     =   "Atualiza(recarrega) a página atual"
         Top             =   0
         Width           =   1125
      End
      Begin VB.Image btnFoward 
         Height          =   1530
         Left            =   2625
         ToolTipText     =   "avança uma página no histórico"
         Top             =   0
         Width           =   1470
      End
      Begin VB.Image btnStop 
         Height          =   1530
         Left            =   1515
         ToolTipText     =   "Para o carregamento da página atual"
         Top             =   0
         Width           =   1110
      End
      Begin VB.Image btnBack 
         Height          =   1530
         Left            =   0
         ToolTipText     =   "Volta uma página no histórico"
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11085
      Top             =   1005
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   9705
      Top             =   6450
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmdOpen 
      Left            =   9705
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3000
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   9060
      ExtentX         =   15981
      ExtentY         =   5292
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8340
      TabIndex        =   5
      Top             =   1605
      Width           =   645
   End
   Begin VB.Image imgHome 
      Height          =   1530
      Left            =   2220
      Picture         =   "frmMain.frx":E002E
      Top             =   135
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image imgRefresh 
      Height          =   1530
      Left            =   2220
      Picture         =   "frmMain.frx":E7B28
      Top             =   135
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgFoward 
      Height          =   1530
      Left            =   2220
      Picture         =   "frmMain.frx":ED642
      Top             =   135
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image imgStop 
      Height          =   1530
      Left            =   2220
      Picture         =   "frmMain.frx":F4C74
      Top             =   135
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image imgBack 
      Height          =   1530
      Left            =   2220
      Picture         =   "frmMain.frx":FA5F6
      Top             =   135
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu Hyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Sair"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMillBack 
         Caption         =   "&Voltar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuMillForward 
         Caption         =   "&Avançar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuMillHome 
         Caption         =   "&Página Inicial"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuMillRefresh 
         Caption         =   "&Atualizar"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuMillStop 
         Caption         =   "&Parar"
         Shortcut        =   ^S
      End
      Begin VB.Menu v 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Opções"
      End
      Begin VB.Menu mnuAboutBrowser 
         Caption         =   "Sobre o &Browser"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was programmed by: Josh Walters
'If have any questions or problems email me at:
'matrixgod@crosswinds.net
'I don't comment this code much cause it's pretty much
'self explanatory.

Option Explicit
Public strURL As String

Private Sub btnBack_Click()
  On Error GoTo errHandle
WebBrowser1.GoBack
errHandle:
    Exit Sub
End Sub

Private Sub btnBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnBack.Picture = imgBack.Picture
End Sub

Private Sub btnBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnBack.Picture = LoadPicture("")
End Sub

Private Sub btnFoward_Click()
  On Error GoTo errHandle
WebBrowser1.GoForward
errHandle:
    Exit Sub
End Sub

Private Sub btnFoward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnFoward.Picture = imgFoward.Picture
End Sub

Private Sub btnFoward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnFoward.Picture = LoadPicture("")
End Sub

Private Sub btnHome_Click()
  On Error GoTo errHandle
WebBrowser1.Navigate "http://www.megaaccesshp.hpg.com.br"
errHandle:
    Exit Sub
End Sub

Private Sub btnHome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnHome.Picture = imgHome.Picture
End Sub

Private Sub btnHome_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnHome.Picture = LoadPicture("")
End Sub

Private Sub btnRefresh_Click()
  On Error GoTo errHandle
    WebBrowser1.Refresh
errHandle:
    Exit Sub
End Sub

Private Sub btnRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnRefresh.Picture = imgRefresh.Picture
End Sub

Private Sub btnRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnRefresh.Picture = LoadPicture("")
End Sub

Private Sub btnStop_Click()
  On Error GoTo errHandle
WebBrowser1.Stop
errHandle:
    Exit Sub
End Sub

Private Sub btnStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStop.Picture = imgStop.Picture
End Sub

Private Sub btnStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnStop.Picture = LoadPicture("")
End Sub

Private Sub Command1_Click()
    If txtInput.Text = "" Then
        'do nothing
    Else
        WebBrowser1.Navigate (txtInput.Text)
    End If
    strURL = txtInput.Text
    If Left(LCase(strURL), 7) = "http://" Or Left(LCase(strURL), 6) = "ftp://" Then
        txtInput.Text = strURL
    Else
        If Left(strURL, 7) <> "http://" Then
            txtInput.Text = "http://" & strURL
        Else
            If Left(strURL, 6) = "ftp://" Then
                txtInput.Text = "ftp://" & strURL
            End If
        End If
    End If
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim nome As String
Dim p_ret As String
If Len(App.Path) = 3 Then
nome = App.Path & "gnavi.swf"
Else
nome = App.Path & "\" & "gnavi.swf"
End If
p_ret = StrConv(LoadResData("GNAVI", "FLASH"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
    Dim lngTemp As Long
    txtInput.Text = ""
            WebBrowser1.Navigate (nome)
           txtInput.Text = ""
Me.Width = 11925
Me.Height = 9000
End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload frmAbout
End Sub

Private Sub form_resize()
On Error GoTo pau
WebBrowser1.Width = frmMain.Width - 105
WebBrowser1.Height = frmMain.Height - 2325
pau:
Exit Sub
End Sub



Private Sub Label2_Click()
PopupMenu menu
Label2.BackStyle = 0
Label2.BorderStyle = 0
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackStyle = 1
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackStyle = 0
Label2.BorderStyle = 0
End Sub

Private Sub mnuAboutBrowser_Click()
    frmAbout.Show
    
End Sub





Private Sub mnuFileExit_Click()
  End
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo errHandle
    cmdOpen.CancelError = True
    cmdOpen.Filter = "Arquivos da Web |*.html;*.htm;*.jpg;*.gif;*.bmp;*.swf|"
    cmdOpen.ShowOpen
    WebBrowser1.Navigate (cmdOpen.FileName)
    
errHandle:
    Exit Sub
    
End Sub

Private Sub mnuMillBack_Click()
    On Error GoTo errStop
    WebBrowser1.GoBack
    
errStop:
    Exit Sub
    
End Sub

Private Sub mnuMillForward_Click()
    On Error GoTo errNstop
    WebBrowser1.GoForward
    
errNstop:
    Exit Sub
    
End Sub

Private Sub mnuMillHome_Click()
    WebBrowser1.Navigate "home"
    
End Sub

Private Sub mnuMillRefresh_Click()
    WebBrowser1.Refresh
    
End Sub

Private Sub mnuMillStop_Click()
    WebBrowser1.Stop
    
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
     txtInput.Text = WebBrowser1.LocationURL
     If WebBrowser1.LocationName <> "gnavi.swf" Then
     frmMain.Caption = "GNavigator :: " & WebBrowser1.LocationName & " :: powered by Internet Explorer"
     Else
     Me.Caption = "GNavigator :: powered by Internet Explorer"
     txtInput.Text = "http://www.google.com.br"
     End If
     txtInput.SelStart = 0
     txtInput.SelLength = Len(txtInput.Text)
     
End Sub
