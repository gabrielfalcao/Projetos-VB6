VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   2625
   ClientTop       =   1980
   ClientWidth     =   6000
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtbLicensiado 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   8421376
      BorderStyle     =   0
      Enabled         =   0   'False
      MultiLine       =   0   'False
      TextRTF         =   $"Splash Screen.frx":0000
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5400
      Top             =   1680
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
      TabIndex        =   5
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Developer: Leandro Carísio Fernandes"
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
      TabIndex        =   4
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "Agora com mais recursos"
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
      Top             =   120
      Width           =   975
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1380
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
      TabIndex        =   0
      Top             =   240
      Width           =   4275
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
      TabIndex        =   1
      Top             =   360
      Width           =   4275
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
rtbLicensiado.SelColor = Label6.ForeColor
rtbLicensiado.LoadFile App.Path & "\User.txt", 1
End Sub

Private Sub Timer1_Timer()
frmEditor.Show
Unload Me
End Sub
