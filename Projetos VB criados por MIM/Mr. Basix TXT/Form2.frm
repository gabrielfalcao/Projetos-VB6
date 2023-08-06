VERSION 5.00
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Sobre o Mr. Basic TXT"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   390
      Left            =   3735
      TabIndex        =   0
      Top             =   2985
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   688
      BTYPE           =   2
      TX              =   "Sair"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   485035
      BCOLO           =   485035
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form2.frx":113C9
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Unload Me
End Sub
