VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o Kit de MP3"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "www.gsite2003.kit.net"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "gabrielfalcao@hotmail.com"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   1920
      Picture         =   "frmAbout.frx":08CA
      Top             =   2280
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Programado e Desenhado por:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   360
      Picture         =   "frmAbout.frx":1095
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4245
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
