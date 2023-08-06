VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00596F6D&
   Caption         =   "N£OWEB"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11805
   FillColor       =   &H00849B99&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txtInput 
      Height          =   315
      Left            =   1110
      TabIndex        =   4
      Top             =   975
      Width           =   5925
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   4980
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6510
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":304A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":349E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":419A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51A2
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   7725
      TabIndex        =   2
      Top             =   1005
      Width           =   870
   End
   Begin MSComDlg.CommonDialog cmdOpen 
      Left            =   5790
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   7095
      TabIndex        =   1
      Top             =   1005
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6675
      Left            =   120
      TabIndex        =   0
      Top             =   1335
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   11765
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
   Begin VB.Image Image4 
      Height          =   9450
      Left            =   0
      Picture         =   "frmMain.frx":55F6
      Stretch         =   -1  'True
      Top             =   1335
      Width           =   15195
   End
   Begin VB.Image Image3 
      Height          =   1170
      Left            =   13770
      Picture         =   "frmMain.frx":695D
      Top             =   75
      Width           =   1425
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   4020
      Picture         =   "frmMain.frx":77EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   150
      TabIndex        =   3
      Top             =   1005
      Width           =   960
   End
   Begin VB.Image Command6 
      Height          =   720
      Left            =   3060
      Picture         =   "frmMain.frx":803D
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command5 
      Height          =   720
      Left            =   2100
      Picture         =   "frmMain.frx":88A0
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command4 
      Height          =   720
      Left            =   1140
      Picture         =   "frmMain.frx":9421
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command3 
      Height          =   720
      Left            =   180
      Picture         =   "frmMain.frx":9B5A
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   0
      Picture         =   "frmMain.frx":A292
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
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
   End
   Begin VB.Menu mnuMill 
      Caption         =   "&Millenium"
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
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Sobre"
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

Private Sub Command2_Click()
    WebBrowser1.Stop
    
End Sub

Private Sub Command3_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BorderStyle = 1

End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BorderStyle = 0
End Sub

Private Sub Command4_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BorderStyle = 1
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BorderStyle = 1
End Sub

Private Sub Command5_Click()
WebBrowser1.Stop
End Sub

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BorderStyle = 1
End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BorderStyle = 0
End Sub

Private Sub Command6_Click()
    WebBrowser1.Refresh
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BorderStyle = 1
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BorderStyle = 0
End Sub

Private Sub Form_Load()
    Dim lngTemp As Long
    frmMain.BackColor = &H849B99
    WebBrowser1.Navigate "www.msn.com.br"
End Sub


Private Sub form_resize()
WebBrowser1.Width = frmMain.Width - 350
WebBrowser1.Height = frmMain.Height - 2200
Image1.Width = frmMain.Width - 150

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
    
End Sub

Private Sub Image2_Click()
WebBrowser1.Navigate "http://www.megaaccesshp.hpg.com.br"
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 0
End Sub

Private Sub mnuAboutBrowser_Click()
    frmAbout.Show
    
End Sub





Private Sub mnuFileExit_Click()
    Unload frmMain
    
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo errHandle
    cmdOpen.CancelError = True
    cmdOpen.Filter = "Html Files (*.html)|*.html;*.htm|"
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
    WebBrowser1.Navigate ("http://www.semo.net")
    
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
     frmMain.Caption = "N£OWEB - " & WebBrowser1.LocationName
     txtInput.SelStart = 0
     txtInput.SelLength = Len(txtInput.Text)
     
End Sub
