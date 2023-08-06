VERSION 5.00
Begin VB.Form frmMedidor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medidor de desempenho físico"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULAR"
      Height          =   420
      Left            =   1740
      TabIndex        =   7
      Top             =   810
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   885
      Left            =   60
      TabIndex        =   6
      Top             =   1920
      Width           =   3330
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   3180
         TabIndex        =   8
         Top             =   225
         Width           =   3180
      End
   End
   Begin VB.TextBox tFCT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   885
      TabIndex        =   5
      Top             =   1260
      Width           =   675
   End
   Begin VB.TextBox tFCR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   885
      TabIndex        =   3
      Top             =   855
      Width           =   675
   End
   Begin VB.TextBox tId 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   885
      TabIndex        =   1
      Top             =   450
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FCT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   345
      TabIndex        =   4
      Top             =   1305
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FCR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   345
      TabIndex        =   2
      Top             =   885
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idade:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   285
      TabIndex        =   0
      Top             =   495
      Width           =   480
   End
End
Attribute VB_Name = "frmMedidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fct As Integer
Dim ff As Integer
Dim fcm As Integer
Dim fcr As Integer
Dim itr As Integer
Dim itrei As Integer
Dim res As Integer
Private Sub Command1_Click()
If tId.Text <> Empty And tFCR.Text <> Empty And tFCT.Text <> Empty Then
fcm = 220 - Int(tId.Text)
fcr = Int(tFCR.Text)
fct = Int(tFCT.Text)
ff = fcm - fcr
itr = ff + fcr
itrei = fct - itr / 100
res = Int(itrei) / Int(itr)
Picture1.Cls
Picture1.Print "A sua intensidade de treinamento é de " & itrei & "% !"
End If
End Sub

'FCT=[(FCM-FCR).ITREI/100]+FCR
Private Sub Form_Load()

End Sub
