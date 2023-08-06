VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSobre 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Editor"
   ClientHeight    =   5265
   ClientLeft      =   3195
   ClientTop       =   1440
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "&Return"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbLicensiado 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   8421376
      BorderStyle     =   0
      Enabled         =   0   'False
      MultiLine       =   0   'False
      TextRTF         =   $"Sobre.frx":0000
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808000&
      Caption         =   "More programs in my home page:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label lblEndereço2 
      BackColor       =   &H00808000&
      Caption         =   "http://www.terravista.pt/ilhadomel/4128"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "Sobre.frx":00AE
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Editor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2025
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   4275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "version 3.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   390
      Left            =   3960
      TabIndex        =   4
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Now with more resources"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Developer Leandro Carísio Fernandes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   5055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Editor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2025
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   4275
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To open the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdVoltar_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
rtbLicensiado.SelColor = Label6.ForeColor
rtbLicensiado.LoadFile App.Path & "\User.txt", 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmEditor.Enabled = True
frmEditor.ActiveForm.rtbDocumento.SetFocus
End Sub

Private Sub lblEndereço2_Click()
    'Open browser
    Call ShellExecute(0, vbNullString, "http://www.lprogramacao.hpg.com.br/", vbNullString, vbNullString, 0)
End Sub
