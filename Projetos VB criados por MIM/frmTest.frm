VERSION 5.00
Object = "{3D186C2C-C1E1-11D4-A4F7-B23873A93835}#2.0#0"; "My3DLabel.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Form for Label 3D"
   ClientHeight    =   4470
   ClientLeft      =   3195
   ClientTop       =   2790
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdShadow 
      Caption         =   "Shadow ON/OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1620
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   465
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   90
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   820
      Caption         =   "3D Label ver 1.0 By Khalid Pervaz."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   65535
      ShadowColor     =   4210752
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   405
      Index           =   1
      Left            =   810
      TabIndex        =   1
      Top             =   675
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   714
      Caption         =   "3D Label ver 1.0 By Khalid Pervaz."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   255
      ShadowColor     =   4210752
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   345
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1170
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   609
      Caption         =   "3D Label ver 1.0 By Khalid Pervaz."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   65280
      ShadowColor     =   4210752
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   1005
      Index           =   3
      Left            =   1350
      TabIndex        =   3
      Top             =   1575
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   1773
      Caption         =   "3D Label"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   255
      ShadowColor     =   0
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   330
      Index           =   4
      Left            =   1215
      TabIndex        =   5
      Top             =   2700
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      Caption         =   "3D Label ver 1.0 By Khalid Pervaz."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   0
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   330
      Index           =   5
      Left            =   1215
      TabIndex        =   6
      Top             =   3150
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      Caption         =   "3D Label ver 1.0 By Khalid Pervaz."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   16777215
      ShadowColor     =   0
   End
   Begin My3DLabel.Label3D Label3D1 
      Height          =   330
      Index           =   6
      Left            =   1485
      TabIndex        =   7
      Top             =   3555
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   582
      Caption         =   "khalid_pervaz@yahoo.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FaceColor       =   65280
      ShadowColor     =   4210752
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdShadow_Click()
    If Label3D1(3).Shadow Then
        Label3D1(3).Shadow = False
    Else
        Label3D1(3).Shadow = True
    End If
End Sub


