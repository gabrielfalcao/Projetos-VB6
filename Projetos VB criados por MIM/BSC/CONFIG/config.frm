VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form config 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração do SiCon BSC"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton p1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3645
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin MSComctlLib.ProgressBar pb 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   1095
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   953
      _Version        =   393216
      Appearance      =   0
      Max             =   40
      Scrolling       =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1590
      TabIndex        =   2
      Top             =   1650
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1432
      Top             =   1665
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1432
      Top             =   1725
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   255
      ItemData        =   "config.frx":0442
      Left            =   450
      List            =   "config.frx":0444
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox data 
      Height          =   240
      Left            =   855
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "config.frx":0446
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2250
      TabIndex        =   3
      Top             =   300
      Width           =   1635
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Command1.Enabled = False
End Sub

Private Sub Form_Load()
Dim datafile As String
If Len(App.Path) = 3 Then
datafile = App.Path & "drt.txt"
Else
datafile = App.Path & "\drt.txt"
End If
pb.Max = 4
Dim maximo As String
End Sub



Private Sub p1_Click()
MsgBox "Instalação e configuração finalizada com sucesso!", vbInformation, Me.Caption
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim datafile As String
If Len(App.Path) = 3 Then
datafile = App.Path & "drt.txt"
Else
datafile = App.Path & "\drt.txt"
End If
On Error GoTo err
Select Case List1.ListIndex + 1
Case 1
Me.Caption = "Preparando configuração..."
Label1.Caption = "Preparando instalação..."
pb.Value = 1
List1.ListIndex = List1.ListIndex + 1
Case 2
Me.Caption = "Compilando Banco de Dados..."
Label1.Caption = "Compilando Banco de Dados..."
Open datafile For Output As #1
Print #1, data
Close #1
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l"
pb.Value = 2
List1.ListIndex = List1.ListIndex + 1
Case Is = 3
Me.Caption = "Configurando registro do sistema..."
Label1.Caption = "Configurando registro do sistema..."

SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE1", "0800323"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "CODE2", "0.00N"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "isregs", "1"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD1", "1"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\8$C3l", "savCOD2", "1"
pb.Value = 3
List1.ListIndex = List1.ListIndex + 1
Case Is = 4
Me.Caption = "INSTALAÇÃO BEM SUCEDIDA!"
Label1.Caption = "INSTALAÇÃO BEM SUCEDIDA!"
pb.Value = 4
List1.ListIndex = List1.ListIndex + 1
p1.Visible = True
Case Else
End Select
err:
Exit Sub
End Sub

Private Sub Timer2_Timer()
If File1.ListIndex < 4 Then
 File1.ListIndex = File1.ListIndex + 1
List1.AddItem (File1.Path & "\" & File1.FileName)
Else
List1.AddItem File1.Path & File1.FileName
Timer2.Enabled = False
List1.Enabled = True
List1.ListIndex = 0
maximo = List1.ListIndex
Timer1.Enabled = True
End If
'Else

'End If
End Sub
