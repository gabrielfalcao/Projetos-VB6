VERSION 5.00
Begin VB.Form frmCalcula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora de IMC - por Gabriel Falcão"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1635
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1905
      Width           =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   345
      Left            =   1470
      TabIndex        =   6
      Top             =   1275
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3105
      TabIndex        =   4
      Top             =   720
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1215
      TabIndex        =   3
      Top             =   720
      Width           =   1200
   End
   Begin VB.Line Line2 
      X1              =   270
      X2              =   270
      Y1              =   210
      Y2              =   4740
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   195
      Y2              =   4725
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   615
      Picture         =   "Form1.frx":0442
      Top             =   2790
      Width           =   3765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela"
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
      Left            =   2182
      TabIndex        =   7
      Top             =   2370
      Width           =   630
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculadora de Índice de Massa Corporal"
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
      Left            =   532
      TabIndex        =   5
      Top             =   195
      Width           =   3930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMC:"
      Height          =   195
      Left            =   1185
      TabIndex        =   2
      Top             =   1980
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altura:"
      Height          =   195
      Left            =   2535
      TabIndex        =   1
      Top             =   765
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      Height          =   195
      Left            =   690
      TabIndex        =   0
      Top             =   765
      Width           =   405
   End
End
Attribute VB_Name = "frmCalcula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo err
Dim peso As String
Dim altura As String
Dim imc As String
peso = Text1.Text
altura = Text2.Text * Text2.Text
imc = peso / altura
label5.Text = imc
err:
If err.Number <> 0 Then
MsgBox "Ocorreu o seguinte erro: " & err.Description & ", Causado no objeto " & err.Source & ", verifique se os valores inseridos são números e se todas as informações estão corretas e tente novamente!", vbCritical, Me.Caption
End If
Exit Sub
End Sub

