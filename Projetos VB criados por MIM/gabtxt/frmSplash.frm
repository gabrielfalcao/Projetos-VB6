VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   97
      Width           =   7080
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "gabrielfalcao@hotmail.com"
         Height          =   255
         Left            =   2250
         TabIndex        =   2
         Top             =   2310
         Width           =   2310
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Programador: Gabriel Falcão"
         Height          =   255
         Left            =   2250
         TabIndex        =   1
         Top             =   1950
         Width           =   2310
      End
      Begin VB.Image imgLogo 
         Height          =   480
         Left            =   5790
         Picture         =   "frmSplash.frx":000C
         Top             =   420
         Width           =   480
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   945
         Left            =   570
         Picture         =   "frmSplash.frx":044E
         Stretch         =   -1  'True
         Top             =   675
         Width           =   5460
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
Private Sub Frame1_Click()
    Unload Me
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
