VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instalar ScreenSaver Triplo X"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "setup xXx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "<- Voltar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   930
      TabIndex        =   5
      Top             =   1575
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Avançar ->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2400
      TabIndex        =   4
      Top             =   1575
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Caption         =   "1º Passo"
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   4845
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   195
         TabIndex        =   12
         Text            =   "Digite aqui o diretório!!!!"
         Top             =   945
         Width           =   4455
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4815
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C:\Windows\"
         Height          =   300
         Left            =   1260
         TabIndex        =   3
         Top             =   465
         Width           =   2610
      End
      Begin VB.Label Label2 
         Caption         =   "Exemplo:"
         Height          =   285
         Left            =   540
         TabIndex        =   2
         Top             =   435
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digite abaixo a localização da sua pasta ""Windows"":"
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   210
         Width           =   3750
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Final"
      Height          =   1365
      Left            =   120
      TabIndex        =   10
      Top             =   150
      Visible         =   0   'False
      Width           =   4845
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1890
         Top             =   735
      End
      Begin VB.PictureBox Picture1 
         Height          =   390
         Left            =   300
         ScaleHeight     =   330
         ScaleWidth      =   4200
         TabIndex        =   11
         Top             =   225
         Width           =   4260
         Begin VB.Shape Shape1 
            BackColor       =   &H000000C0&
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   30
            Top             =   30
            Width           =   4140
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2º Passo"
      Height          =   1365
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   4845
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   105
         TabIndex        =   9
         Top             =   405
         Width           =   4440
      End
      Begin VB.Label Label4 
         Caption         =   "Diretório escolhido:"
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pressione avançar para continuar com a instalação..."
         Height          =   195
         Left            =   465
         TabIndex        =   7
         Top             =   1050
         Width           =   3765
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Concluir" Then
Unload Me
End If
If Frame2.Visible = True Then
Frame3.Visible = True
Frame2.Visible = False
Command2.Enabled = False
Command1.Enabled = False
Command1.Caption = "Concluir"
Command2.Enabled = False
Timer1.Enabled = True
End If
If Frame1.Visible = True Then
Frame2.Visible = True
Frame1.Visible = False
Command2.Enabled = True
Label5.Caption = Text1.Text
End If


End Sub

Private Sub Command2_Click()
If Frame2.Visible = True Then
Frame1.Visible = True
Frame2.Visible = False
Command2.Enabled = False
End If
If Frame3.Visible = True Then
Frame2.Visible = True
Frame3.Visible = False
End If
End Sub

Private Sub Form_Load()
Dim barrawidth As String
barrawidth = Shape1.Width
Shape1.Width = 1
End Sub

Private Sub Text1_Click()
Text1.ForeColor = &H80000012
Text1.BackColor = &H8000000E
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
On Error GoTo msg

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "Triplo X.scr", Text1.Text & "Triplo X.scr"
Shape1.Width = Shape1.Width + 10

FileCopy App.Path & "ASYCFILT.DLL", Text1.Text & "ASYCFILT.DLL"
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "COMCAT.DLL", Text1.Text & "COMCAT.DLL"
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "COMDLG32.OCX", Text1.Text & "COMDLG32.OCX"
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "MSCOMCTL.OCX", Text1.Text & "MSCOMCTL.OCX"
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "MSINET.OCX", Text1.Text & "MSINET.OCX"
Shape1.Width = Shape1.Width + 10

FileCopy App.Path & "MSVBVM60.DLL", Text1.Text & "MSVBVM60.DLL"
Shape1.Width = Shape1.Width + 10

FileCopy App.Path & "OLEAUT32.DLL", Text1.Text & "OLEAUT32.DLL"
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
FileCopy App.Path & "VB6STKIT.DLL", Text1.Text & VB6STKIT.DLL
Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10

Shape1.Width = Shape1.Width + 10
Shape1.Width = 4140
Timer1.Enabled = False
MsgBox "Instalação Concluída"
msg:
MsgBox Err.Number & " " & Err.Description & " " & Err.Source
Resume Next

End Sub
