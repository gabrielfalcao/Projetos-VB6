VERSION 5.00
Begin VB.Form frm3DLAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Label3D"
   ClientHeight    =   1575
   ClientLeft      =   3405
   ClientTop       =   4110
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   420
      Left            =   3645
      TabIndex        =   0
      Top             =   990
      Width           =   915
   End
   Begin VB.Label lblEmailAdd 
      AutoSize        =   -1  'True
      Caption         =   "khalid_pervaz@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1215
      Width           =   1935
   End
   Begin VB.Label lblicon2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Khalid Pervaz."
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   585
      Index           =   0
      Left            =   540
      TabIndex        =   3
      Top             =   585
      Width           =   2685
   End
   Begin VB.Label lblicon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Label ver 1.0 By "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label lblicon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3D Label ver 1.0 By "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label lblicon2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Khalid Pervaz."
      BeginProperty Font 
         Name            =   "Brush Script MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   1
      Left            =   540
      TabIndex        =   4
      Top             =   540
      Width           =   2685
   End
End
Attribute VB_Name = "frm3DLAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    Unload Me
End Sub


