VERSION 5.00
Begin VB.Form FrmAboutSimple 
   BackColor       =   &H00EFD1AD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre o TeraZip"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5190
      TabIndex        =   1
      Top             =   1470
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por gabriel Falcão usando os Controles RichSoft ZipIt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   195
      Left            =   630
      TabIndex        =   3
      Top             =   1050
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "FrmAboutSimple.frx":0000
      Top             =   60
      Width           =   720
   End
   Begin VB.Label LSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.gabrielfalcao.i8.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1980
      TabIndex        =   2
      Top             =   1410
      Width           =   1860
   End
   Begin VB.Label LApp 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAboutSimple.frx":11F8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00990B0B&
      Height          =   960
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   5085
   End
End
Attribute VB_Name = "FrmAboutSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const CorUrl = vbBlue

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub Form_Load()
 MyLanguage = GetSetting(App.EXEName, "last", "language", 0)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub LApp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub

Private Sub LSite_Click()
 Call ShellExecute(Me.hWnd, "Open", LSite.Caption, "", CurDir, 0)
End Sub

Private Sub LSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = CorUrl
End Sub

Private Sub PicLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 LSite.ForeColor = vbButtonText
End Sub
