VERSION 5.00
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form Find 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   1470
   ClientTop       =   4275
   ClientWidth     =   3180
   ControlBox      =   0   'False
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   405
      Left            =   225
      TabIndex        =   4
      Top             =   1410
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   714
      BTYPE           =   9
      TX              =   "Localizar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmfind.frx":08CA
      PICN            =   "frmfind.frx":08E6
      PICH            =   "frmfind.frx":09F8
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   2760
   End
   Begin VB.TextBox crit 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   390
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No item:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   750
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   930
   End
End
Attribute VB_Name = "Find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Dim criterio As String
Dim cont
criterio = Combo1.Text & "=" & "'" & crit.Text & "'"
If cont = 0 Then
frmAgenda.Data1.Recordset.FindFirst criterio
Else
frmAgenda.Data1.Recordset.FindNext criterio
If cont = frmAgenda.Data1.Recordset.RecordCount - 1 Then
frmAgenda.Data1.Recordset.FindFirst criterio
End If
End If
If frmAgenda.Data1.Recordset.NoMatch = True Then
'MsgBox Combo1.Text & " não localizado!"
cont = 0
Else
cont = cont + 1
End If
End Sub

Private Sub Form_Load()

Combo1.AddItem "Nome"
Combo1.AddItem "Sobrenome"
Combo1.AddItem "Endereço"
Combo1.AddItem "Bairro"
Combo1.AddItem "CEP"
Combo1.AddItem "Cidade"
Combo1.AddItem "UF"
Combo1.AddItem "Telefone"
Combo1.AddItem "E-Mail"
Combo1.ListIndex = 0

End Sub
