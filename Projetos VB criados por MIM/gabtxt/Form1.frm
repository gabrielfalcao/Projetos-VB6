VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Comm Text's"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Height          =   405
      Left            =   9045
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   645
      UseMaskColor    =   -1  'True
      Width           =   720
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1290
      TabIndex        =   22
      Top             =   727
      Width           =   435
      Begin VB.CheckBox bullet 
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   285
      End
   End
   Begin VB.ComboBox txtsize 
      Height          =   315
      Left            =   8205
      TabIndex        =   19
      Text            =   "8"
      Top             =   735
      Width           =   615
   End
   Begin VB.ComboBox fonte 
      Height          =   315
      Left            =   4395
      TabIndex        =   18
      Text            =   "Arial"
      Top             =   735
      Width           =   2850
   End
   Begin VB.OptionButton Option1 
      Caption         =   "E"
      Height          =   315
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   735
      Value           =   -1  'True
      Width           =   450
   End
   Begin VB.OptionButton Option2 
      Caption         =   "C"
      Height          =   315
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   735
      Width           =   450
   End
   Begin VB.OptionButton Option3 
      Caption         =   "D"
      Height          =   315
      Left            =   2895
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   735
      Width           =   450
   End
   Begin VB.CommandButton Command11 
      Height          =   570
      Left            =   2925
      Picture         =   "Form1.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command10 
      Height          =   570
      Left            =   9180
      Picture         =   "Form1.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Inserir figura"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Height          =   570
      Left            =   4275
      Picture         =   "Form1.frx":1218
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sobre"
      Height          =   570
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command4 
      Height          =   570
      Left            =   900
      Picture         =   "Form1.frx":1882
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Height          =   570
      Left            =   1575
      Picture         =   "Form1.frx":1EEC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CheckBox Check3 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CheckBox Check2 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cor do Fundo"
      Height          =   570
      Left            =   6075
      Picture         =   "Form1.frx":2556
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cor da Letra"
      Height          =   570
      Left            =   4950
      Picture         =   "Form1.frx":2BC0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sair"
      Height          =   570
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Height          =   570
      Left            =   3600
      Picture         =   "Form1.frx":322A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Height          =   570
      Left            =   2250
      Picture         =   "Form1.frx":3894
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin RichTextLib.RichTextBox txt 
      Height          =   7140
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   12594
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":3EFE
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   4965
      Top             =   4290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cor 
      Left            =   3735
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   49152
   End
   Begin MSComDlg.CommonDialog print 
      Left            =   8670
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fonte:"
      Height          =   195
      Left            =   3825
      TabIndex        =   21
      Top             =   795
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tamanho:"
      Height          =   195
      Left            =   7365
      TabIndex        =   20
      Top             =   795
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bullet_Click()
txt.SelBullet = bullet.Value
End Sub

Private Sub Check1_Click()
txt.SelBold = Check1.Value
End Sub

Private Sub Check2_Click()
txt.SelItalic = Check2.Value
End Sub

Private Sub Check3_Click()
txt.SelUnderline = Check3.Value
End Sub



Private Sub Command10_Click()
frmInserirFigura.Show
Me.Enabled = False
End Sub



Private Sub Command11_Click()
Clipboard.SetText (txt.SelText)
txt.SelText = ""
End Sub



Private Sub Command12_Click()
Printer.Font = txt.Font
Printer.FontSize = txtsize.Text
Printer.ForeColor = cor.Color
Printer.Print Spc(25)
Printer.Print Tab(3)
Printer.Print txt.Text

Printer.EndDoc
End Sub

Private Sub Command4_Click()
txt.Text = ""
Me.Caption = "Comm text's"
End Sub

Private Sub Command8_Click()
frmSplash.Show
End Sub

Private Sub Command9_Click()
Me.txt.SelText = Clipboard.GetText
End Sub

Private Sub fonte_Change()
txt.SelFontName = fonte.Text
End Sub

Private Sub fonte_Click()
txt.SelFontName = fonte.Text
End Sub

Private Sub fonte_KeyPress(KeyAscii As Integer)
txt.SelFontName = fonte.Text
End Sub

Private Sub Form_Load()
txtsize.AddItem "8"
txtsize.AddItem "10"
txtsize.AddItem "12"
txtsize.AddItem "14"
txtsize.AddItem "16"
txtsize.AddItem "18"
txtsize.AddItem "24"
txtsize.AddItem "28"
txtsize.AddItem "40"
txtsize.AddItem "48"
txtsize.AddItem "72"
Dim Contador
For Contador = 1 To Screen.FontCount - 1
fonte.AddItem Screen.Fonts(Contador)
Next
End Sub


Private Sub Option1_Click()
txt.SelAlignment = 0
End Sub

Private Sub Option2_Click()
txt.SelAlignment = 2
End Sub

Private Sub Option3_Click()
txt.SelAlignment = 1
End Sub

Private Sub Option4_Click()
txt.SelBullet = Option4.Value
End Sub

Private Sub txtsize_Change()
txt.SelFontSize = txtsize.Text
End Sub

Private Sub Command1_Click()
Com.DialogTitle = "Abrir"
Com.Filter = "Arquivos de texto (*.txt)|*.txt|Arquivos de lote (*.bat)|*.bat|Todos os Arquivos (*.*)|*.*"
Com.ShowOpen
txt.LoadFile (Com.FileName)
Me.Caption = Me.Caption & " (" & Com.FileName & ") "
End Sub


Private Sub Command2_Click()
Com.DialogTitle = "Salvar Como"
Com.FileName = "Com2Arquivodetextos.txt"
Com.Filter = "Arquivos de texto (*.txt)|*.txt|Arquivos de lote (*.bat)|*.bat|Todos os Arquivos (*.*)|*.*"
Com.ShowSave
txt.SaveFile (Com.FileName)
Me.Caption = Me.Caption & " (" & Com.FileName & ") "
End Sub

Private Sub Command3_Click()
Clipboard.SetText (txt.SelText)
End Sub

Private Sub Command5_Click()
Unload Me
End
End Sub

Private Sub Command6_Click()
cor.ShowColor
txt.SelColor = cor.Color
End Sub

Private Sub Command7_Click()
cor.ShowColor
txt.BackColor = cor.Color
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
txt.Height = Me.Height - 1800
txt.Width = Me.Width - 225
End If
If Me.WindowState = 0 Then
txt.Height = Me.Height - 1800
txt.Width = Me.Width - 225
End If
End Sub

Private Sub txtsize_Click()
txt.SelFontSize = txtsize.Text
End Sub

Private Sub txtsize_KeyPress(KeyAscii As Integer)
txt.SelFontSize = txtsize.Text
End Sub
