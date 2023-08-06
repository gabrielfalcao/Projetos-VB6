VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00F4C686&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Sobre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   833
      TabIndex        =   4
      Top             =   1245
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   833
      TabIndex        =   1
      Top             =   1650
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convênios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   833
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.CheckBox eregistrado 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1230
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BSC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   1988
      TabIndex        =   3
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Conveniados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   593
      TabIndex        =   2
      Top             =   30
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmConvenio.Show

End Sub

Private Sub Command2_Click()
frmDRT.Show
End Sub

Private Sub Command3_Click()
frmOPT.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
frmABT.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
eregistrado.Value = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "isregs")
If eregistrado.Value <> "1" Then
MsgBox "Cópia do programa não registrada, contate o distribuidor!", vbInformation, Me.Caption
Unload Me
End If
End Sub
