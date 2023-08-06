VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeG Chat v1.00 - Opções Iniciais"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3435
      TabIndex        =   24
      Top             =   2040
      Width           =   1380
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1380
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   4845
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   750
      Width           =   390
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1425
      Width           =   390
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   765
      Width           =   390
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1410
      Width           =   390
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1110
      Width           =   390
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   765
      Width           =   390
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   465
      Width           =   390
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1860
      Width           =   390
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000000&
      Height          =   270
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1860
      Width           =   390
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1410
      Width           =   390
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1110
      Width           =   390
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   765
      Width           =   390
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   390
   End
   Begin MSComDlg.CommonDialog cor 
      Left            =   1650
      Top             =   3810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image7 
      Height          =   180
      Left            =   3525
      Top             =   2055
      Width           =   270
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Papo:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2895
      TabIndex        =   18
      Top             =   1410
      Width           =   1530
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2895
      TabIndex        =   16
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   780
      Width           =   1530
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2880
      TabIndex        =   12
      Top             =   495
      Width           =   720
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Papo:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1860
      Width           =   1530
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seu IP"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1425
      Width           =   1530
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seu IP:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP e Apelido"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   1530
   End
   Begin VB.Label lblIp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP e Apelido"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cores do Sistema:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   1755
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command10_Click()
cor.ShowColor
Command10.BackColor = cor.Color
Label11.BackColor = cor.Color
frmChatCliente.lblStatus.ForeColor = cor.Color
End Sub

Private Sub Command11_Click()
cor.ShowColor
Command11.BackColor = cor.Color
Label12.BackColor = cor.Color
frmChatCliente.lblInfo.BackColor = cor.Color
End Sub

Private Sub Command12_Click()
cor.ShowColor
Command12.BackColor = cor.Color
Label4.ForeColor = cor.Color
frmChatCliente.txtIp.ForeColor = cor.Color
frmChatCliente.txtApelido.ForeColor = cor.Color
End Sub

Private Sub Command13_Click()
cor.ShowColor
Command13.BackColor = cor.Color
Label6.ForeColor = cor.Color
frmChatCliente.Text1.ForeColor = cor.Color
End Sub

Private Sub Command14_Click()

frmChatCliente.Show
Unload Me
End Sub

Private Sub Command15_Click()
cor.ShowColor
Command15.BackColor = cor.Color
Label10.ForeColor = cor.Color
frmChatCliente.txtEnviar.ForeColor = cor.Color
End Sub

Private Sub Command16_Click()
cor.ShowColor
Command16.BackColor = cor.Color
Label12.ForeColor = cor.Color
frmChatCliente.lblInfo.ForeColor = cor.Color
End Sub

Private Sub Command2_Click()
cor.ShowColor
Command2.BackColor = cor.Color
lblIp.BackColor = cor.Color
frmChatCliente.lblIp.ForeColor = cor.Color
frmChatCliente.lblApelido.ForeColor = cor.Color
End Sub

Private Sub Command3_Click()
cor.ShowColor
Command3.BackColor = cor.Color
Label4.BackColor = cor.Color
frmChatCliente.txtIp.BackColor = cor.Color
frmChatCliente.txtApelido.BackColor = cor.Color
End Sub

Private Sub Command4_Click()
cor.ShowColor
Command4.BackColor = cor.Color
Label5.BackColor = cor.Color
frmChatCliente.Label1.ForeColor = cor.Color
End Sub

Private Sub Command5_Click()
cor.ShowColor
Command5.BackColor = cor.Color
Label6.BackColor = cor.Color
frmChatCliente.Text1.BackColor = cor.Color
End Sub

Private Sub Command6_Click()
cor.ShowColor
Command6.BackColor = cor.Color
Label8.ForeColor = cor.Color
frmChatCliente.txtPapo.ForeColor = cor.Color
End Sub

Private Sub Command7_Click()
cor.ShowColor
Command7.BackColor = cor.Color
Label8.BackColor = cor.Color
frmChatCliente.txtPapo.BackColor = cor.Color
End Sub

Private Sub Command8_Click()
cor.ShowColor
Command8.BackColor = cor.Color
Label9.BackColor = cor.Color
frmChatCliente.lblTexto.ForeColor = cor.Color
End Sub

Private Sub Command9_Click()
cor.ShowColor
Command9.BackColor = cor.Color
Label10.BackColor = cor.Color
frmChatCliente.txtEnviar.BackColor = cor.Color
End Sub

