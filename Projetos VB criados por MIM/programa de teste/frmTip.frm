VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definir senha:"
   ClientHeight    =   4290
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTip.frx":0442
   ScaleHeight     =   4290
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      BackColor       =   &H0030B3C4&
      Caption         =   "&Não mostrar esta tela ao iniciar"
      Height          =   315
      Left            =   1005
      MaskColor       =   &H00FF8080&
      TabIndex        =   3
      Top             =   2250
      UseMaskColor    =   -1  'True
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   765
      TabIndex        =   2
      Top             =   1650
      Width           =   3330
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0030B3C4&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ä"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   30
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   4125
      TabIndex        =   4
      Top             =   -90
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   3975
      Picture         =   "frmTip.frx":438C4
      Stretch         =   -1  'True
      Top             =   15
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00BEA2A2&
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1620
      TabIndex        =   0
      Top             =   990
      Width           =   1620
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

Private Sub chkLoadTipsAtStartup_Click()
SaveSetting App.EXEName, "Options", "Mostrar esta tela ao iniciar", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdOK_Click()
SaveSetting App.EXEName, "Options", "Mostrar esta tela ao iniciar", Text1.Text
SaveSetting App.EXEName, "Options", "Mostrar esta tela ao iniciar", chkLoadTipsAtStartup.Value
Form3.Show
Me.Hide
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ForeColor = &HFFFF00
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ForeColor = &H0&
End Sub

Private Sub Form_Load()
If chkLoadTipsAtStartup.Value = 1 Then
Form3.Show
Me.Hide
Else
End If
End Sub
