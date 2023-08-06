VERSION 5.00
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mr. Basic TXT"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton chameleonButton4 
      Height          =   660
      Left            =   6885
      TabIndex        =   5
      Top             =   0
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1164
      BTYPE           =   2
      TX              =   "SOBRE"
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
      BCOL            =   14678527
      BCOLO           =   14678527
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEF3D8&
      Height          =   4890
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   690
      Width           =   8715
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3135
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":101A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1694
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   660
      Left            =   7830
      TabIndex        =   2
      Top             =   0
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1164
      BTYPE           =   2
      TX              =   "SAIR"
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
      BCOL            =   12640511
      BCOLO           =   12640511
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":1D0E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1191
      ButtonWidth     =   1349
      ButtonHeight    =   1138
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "new"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Imprimir "
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5580
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "15:25"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "16/12/2003"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5159
            Text            =   "gabrielfalcao@hotmail.com"
            TextSave        =   "gabrielfalcao@hotmail.com"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir arquivo de texto..."
      Filter          =   "Arquivo de Texto|*.txt|Arquivo de Lote|*.bat"
   End
   Begin RichTextLib.RichTextBox carregador 
      Height          =   615
      Left            =   3060
      TabIndex        =   4
      Top             =   1980
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":1D2A
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1785
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar texto como..."
      Filter          =   "Arquivo de Texto|*.txt|Arquivo de Lote|*.bat"
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuNovo 
         Caption         =   "Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuColar 
         Caption         =   "Colar"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub

Private Sub Slider1_Change()
ProgressBar1.Value = Slider1.Value
End Sub

Private Sub chameleonButton1_Click()
Unload Me
End Sub

Private Sub chameleonButton2_Click()
mnuCopiar_Click
End Sub

Private Sub chameleonButton3_Click()
mnuColar_Click
End Sub

Private Sub chameleonButton4_Click()
Form2.Show
End Sub

Private Sub mnuAbrir_Click()
cd2.ShowOpen
carregador.LoadFile cd2.FileName
Text1.Text = Empty
Text1.Text = carregador.Text
End Sub

Private Sub mnuColar_Click()
Text1.Text = Text1.Text & Clipboard.GetText
End Sub

Private Sub mnuCopiar_Click()
Clipboard.SetText (Text1.SelText)
End Sub

Private Sub mnuImprimir_Click()
 
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 8
    Printer.Font.Bold = False
    Printer.Print Text1.Text & vbCrLf & vbCrLf & "Documento criado com Mr. Basic TXT" & vbCrLf & "Programador: Gabriel Falcão"
    Printer.EndDoc
End Sub

Private Sub mnuNovo_Click()
Text1.Text = Empty
End Sub

Private Sub mnuSalvar_Click()
On Error GoTo err
cd.ShowSave
   Open cd.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
err:
If err.Number <> 0 Then
MsgBox "Não foi possível salvar", vbExclamation, Me.Caption
Exit Sub
ElseIf err.Number = 76 Then
Exit Sub
End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button
Case Is = "Novo"
mnuNovo_Click
Case Is = "Salvar"
mnuSalvar_Click
Case Is = "Abrir"
  mnuAbrir_Click
    Case Is = "Imprimir"
   mnuImprimir_Click
    Case Else
    End Select
err:
If err.Number <> 0 Then
Exit Sub
End If
End Sub
