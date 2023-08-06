VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   7320
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   9750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   7320
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4665
      Top             =   4057
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4665
      Top             =   4057
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4665
      Top             =   4057
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4665
      Top             =   4057
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4665
      Top             =   4057
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   4665
      Top             =   4057
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1230
      Left            =   3375
      ScaleHeight     =   1170
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   5895
      Width           =   3015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Index           =   1
         Left            =   -30
         TabIndex        =   5
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label des 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carregando Banco de Dados..."
         Height          =   255
         Left            =   -30
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programador: Gabriel Falcão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   0
         Width           =   2325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.megaaccesshp.hpg.com.br"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   150
         MouseIcon       =   "frmSplash.frx":E8950
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   450
         Width           =   2745
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gabrielfalcao@hotmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   405
         MouseIcon       =   "frmSplash.frx":E8AA2
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   225
         Width           =   2235
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4665
      Top             =   4057
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   1095
      Index           =   0
      Left            =   4748
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1931
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   3600
      TabIndex        =   7
      Top             =   2730
      Width           =   405
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
ProgressBar1.Item(1).Value = 4
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = False
Label3.FontUnderline = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = False
Label3.FontUnderline = False
End Sub

Private Sub Label2_Click()
  ShellExecute GetActiveWindow(), "Open", "http://" & Label2.Caption, "", 0&, 1
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = True
Label3.FontUnderline = False
End Sub

Private Sub Label3_Click()
ShellExecute GetActiveWindow(), "Open", "mailto:" & Label3.Caption, "", 0&, 1
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = True
Label2.FontUnderline = False

End Sub

Private Sub Label4_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False

Load Principal
Principal.Show
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = False
Label3.FontUnderline = False
End Sub

Private Sub Timer2_Timer()
ProgressBar1.Item(0).Value = 8
ProgressBar1.Item(1).Value = 6
des.Caption = "Carregando Clientes..."

Timer2.Enabled = False
Timer3.Enabled = True

End Sub

Private Sub Timer1_Timer()


Timer1.Enabled = False

Load Principal
Principal.Show
Unload Me
End Sub

Private Sub Timer3_Timer()
des.Caption = "Carregando Fornecedores..."
ProgressBar1.Item(0).Value = 20
ProgressBar1.Item(1).Value = 20
Timer3.Enabled = False
Timer4.Enabled = True


End Sub

Private Sub Timer4_Timer()
des.Caption = "Carregando Funcionários..."
ProgressBar1.Item(1).Value = 30
ProgressBar1.Item(0).Value = 30
Timer4.Enabled = False
Timer5.Enabled = True

End Sub

Private Sub Timer5_Timer()
des.Caption = "Carregando Pedidos..."
ProgressBar1.Item(0).Value = 40
ProgressBar1.Item(1).Value = 40
Timer5.Enabled = False
Timer6.Enabled = True

End Sub

Private Sub Timer6_Timer()
des.Caption = "Finalizando ... "
ProgressBar1.Item(1).Value = 75
ProgressBar1.Item(0).Value = 75
Timer6.Enabled = False
Timer7.Enabled = True

End Sub

Private Sub Timer7_Timer()
ProgressBar1.Item(0).Value = 100
ProgressBar1.Item(1).Value = 100
Timer7.Enabled = False
Timer1.Enabled = True
End Sub
