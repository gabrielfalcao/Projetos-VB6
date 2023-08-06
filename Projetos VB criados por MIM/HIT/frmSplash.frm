VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4335
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   4335
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   615
      Top             =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2618
      TabIndex        =   1
      Top             =   2550
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2633
      TabIndex        =   0
      Top             =   2565
      Width           =   2220
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If Command$ <> "" Then
        If InStr(Command$, "-enabler") > 0 Then
        frmEnab.Show
        Unload Me

        End If
        If InStr(Command$, "-invader") > 0 Then
        frmInvasor.Show
        Unload Me

        End If
        If InStr(Command$, "-cript") > 0 Then
        frmCrip.Show
        Unload Me

        End If
        If InStr(Command$, "-crack") > 0 Then
        frmCrack.Show
        Unload Me
        End If
        If InStr(Command$, "-res") > 0 Then
        Form1.Show
        Unload Me
       
        End If
        Else
        Timer1.Enabled = True

    End If
  
End Sub

Private Sub Timer1_Timer()
  Unload Me
  frmMain.Show
  frmAjuda.Show
End Sub
