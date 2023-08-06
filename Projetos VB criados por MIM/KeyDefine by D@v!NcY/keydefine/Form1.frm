VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Max261 Examples"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "dmf@wildsoftware85.freeserve.co.uk"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Little Example From MAX261"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Apply the changes to the set strings
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", (Text1.Text)
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", (Text2.Text)
End Sub

Private Sub Form_Load()
' Get users name from the registry
Text1.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
' Get the company from the registry
Text2.Text = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
End Sub

