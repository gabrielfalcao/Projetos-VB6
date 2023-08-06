VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EDE6871F-B292-4B86-B602-523B7F4DC820}#1.0#0"; "ChameleonButton.ocx"
Begin VB.Form frmAgenda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerenciamento de Contatos"
   ClientHeight    =   6330
   ClientLeft      =   885
   ClientTop       =   165
   ClientWidth     =   11505
   ControlBox      =   0   'False
   Icon            =   "frmAgendaMDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Bairro"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   50
      Top             =   1095
      Width           =   1470
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "CEP"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   3180
      MaxLength       =   50
      TabIndex        =   49
      Top             =   1095
      Width           =   1530
   End
   Begin Chameleon.chameleonButton cmdClose 
      Height          =   435
      Left            =   75
      TabIndex        =   39
      Top             =   45
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   767
      BTYPE           =   1
      TX              =   "SAIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   16761024
      FCOL            =   65535
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgendaMDB.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command6 
      Height          =   270
      Left            =   9360
      TabIndex        =   47
      Top             =   3450
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   476
      BTYPE           =   9
      TX              =   "Restaurar Padrão"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   12582912
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmAgendaMDB.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command7 
      Height          =   435
      Left            =   8760
      TabIndex        =   48
      Top             =   5895
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   767
      BTYPE           =   9
      TX              =   "Salvar Texto Atual"
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
      MICON           =   "frmAgendaMDB.frx":0902
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command4 
      Height          =   885
      Left            =   10050
      TabIndex        =   46
      Top             =   5010
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1561
      BTYPE           =   9
      TX              =   "Imprimir"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmAgendaMDB.frx":091E
      PICN            =   "frmAgendaMDB.frx":093A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command1 
      Height          =   255
      Left            =   4650
      TabIndex        =   41
      Top             =   3135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "SOBRE"
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
      MICON           =   "frmAgendaMDB.frx":158C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command2 
      Height          =   240
      Left            =   4650
      TabIndex        =   42
      Top             =   2880
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   423
      BTYPE           =   8
      TX              =   "Sem Foto"
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
      MICON           =   "frmAgendaMDB.frx":15A8
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      Connect         =   "Access"
      DatabaseName    =   "Agenda.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Agenda"
      Top             =   3045
      Width           =   4455
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   8745
      ScaleHeight     =   360
      ScaleWidth      =   2640
      TabIndex        =   33
      Top             =   4605
      Width           =   2640
      Begin Chameleon.chameleonButton option1 
         Height          =   285
         Left            =   15
         TabIndex        =   43
         Top             =   15
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Negrito"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAgendaMDB.frx":15C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1290
         ScaleHeight     =   285
         ScaleWidth      =   1290
         TabIndex        =   34
         Top             =   15
         Width           =   1290
         Begin Chameleon.chameleonButton option2 
            Height          =   285
            Left            =   15
            TabIndex        =   44
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Itálico"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            BCOLO           =   13160660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAgendaMDB.frx":15E0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
   End
   Begin VB.ListBox cmbFonte 
      Appearance      =   0  'Flat
      BackColor       =   &H00ECEDFF&
      Height          =   810
      Left            =   8745
      TabIndex        =   32
      Top             =   3750
      Width           =   2640
   End
   Begin VB.TextBox crit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      Height          =   300
      Left            =   5925
      TabIndex        =   29
      Top             =   1740
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5925
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2400
      Width           =   2760
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF4EA&
      ForeColor       =   &H00FF8080&
      Height          =   3000
      Left            =   8745
      ScaleHeight     =   2970
      ScaleWidth      =   2610
      TabIndex        =   14
      Top             =   330
      Width           =   2640
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF4FF&
         DataField       =   "Foto"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   -15
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2400
         Width           =   2640
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   2220
         Left            =   390
         ScaleHeight     =   2190
         ScaleWidth      =   1830
         TabIndex        =   21
         Top             =   195
         Width           =   1860
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2220
            Left            =   -15
            Stretch         =   -1  'True
            Top             =   -15
            Width           =   1860
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sem Foto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   540
            TabIndex        =   22
            Top             =   1005
            Width           =   795
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1110
         TabIndex        =   19
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FOTO "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   9
         Left            =   13000
         TabIndex        =   15
         Top             =   -15
         Width           =   525
      End
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "E-Mail"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   3180
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1635
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Telefone"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1635
      Width           =   1440
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "UF"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   3675
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1365
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Cidade"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1365
      Width           =   2175
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Endereço"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   5
      Top             =   825
      Width           =   4575
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Sobrenome"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3930
      MaxLength       =   50
      TabIndex        =   3
      Top             =   555
      Width           =   1665
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFF7FD&
      DataField       =   "Nome"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   1
      Top             =   555
      Width           =   1665
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   9315
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Escolha uma foto para o contato:"
      Filter          =   "Arquivos de Imagem|*.bmp;*.jpg;*.gif"
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "Código"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   10
      Left            =   9015
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "1"
      Top             =   990
      Visible         =   0   'False
      Width           =   135
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   405
      Left            =   5925
      TabIndex        =   27
      Top             =   2760
      Width           =   2700
      _ExtentX        =   4763
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
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAgendaMDB.frx":15FC
      PICN            =   "frmAgendaMDB.frx":1618
      PICH            =   "frmAgendaMDB.frx":172A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog savetxt 
      Left            =   9405
      Top             =   1875
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Salvar texto como:"
      Filter          =   "Arquivos de Texto|*.txt"
   End
   Begin Chameleon.chameleonButton cmdAdd 
      Height          =   870
      Left            =   525
      TabIndex        =   35
      Top             =   2145
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1535
      BTYPE           =   9
      TX              =   "Novo"
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
      MICON           =   "frmAgendaMDB.frx":183C
      PICN            =   "frmAgendaMDB.frx":1858
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdDelete 
      Height          =   870
      Left            =   1425
      TabIndex        =   36
      Top             =   2145
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1535
      BTYPE           =   9
      TX              =   "Excluir"
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
      MICON           =   "frmAgendaMDB.frx":1ED2
      PICN            =   "frmAgendaMDB.frx":1EEE
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdUpdate 
      Height          =   870
      Left            =   3225
      TabIndex        =   37
      Top             =   2145
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1535
      BTYPE           =   9
      TX              =   "Salvar"
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
      MICON           =   "frmAgendaMDB.frx":2088
      PICN            =   "frmAgendaMDB.frx":20A4
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdRefresh 
      Height          =   870
      Left            =   2325
      TabIndex        =   38
      Top             =   2145
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1535
      BTYPE           =   9
      TX              =   "Atualizar"
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
      MICON           =   "frmAgendaMDB.frx":271E
      PICN            =   "frmAgendaMDB.frx":273A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdChangeFoto 
      Height          =   810
      Left            =   4650
      TabIndex        =   40
      Top             =   2055
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1429
      BTYPE           =   8
      TX              =   "Escolher Foto"
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
      MICON           =   "frmAgendaMDB.frx":2DB4
      PICN            =   "frmAgendaMDB.frx":2DD0
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton Command5 
      Height          =   885
      Left            =   8760
      TabIndex        =   45
      Top             =   5010
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1561
      BTYPE           =   9
      TX              =   "Atualizar Texto"
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
      MICON           =   "frmAgendaMDB.frx":349A
      PICN            =   "frmAgendaMDB.frx":34B6
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox impresso 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   3750
      Width           =   8565
   End
   Begin VB.Shape Shape5 
      Height          =   285
      Left            =   45
      Top             =   1635
      Width           =   1200
   End
   Begin VB.Shape Shape4 
      Height          =   285
      Left            =   45
      Top             =   1365
      Width           =   1200
   End
   Begin VB.Shape Shape3 
      Height          =   285
      Left            =   45
      Top             =   1095
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      Height          =   285
      Left            =   45
      Top             =   825
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      Height          =   285
      Left            =   45
      Top             =   555
      Width           =   1200
   End
   Begin VB.Line Line4 
      X1              =   5580
      X2              =   5580
      Y1              =   675
      Y2              =   1875
   End
   Begin VB.Line Line3 
      X1              =   2295
      X2              =   5130
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   3495
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   52
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2670
      TabIndex        =   51
      Top             =   1125
      Width           =   345
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6690
      TabIndex        =   31
      Top             =   1275
      Width           =   1170
   End
   Begin VB.Label Label10 
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
      Left            =   5925
      TabIndex        =   30
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fonte:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8745
      TabIndex        =   26
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visualizar Impressão:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   25
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visualizar Impressão:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   3525
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   90
      X2              =   11370
      Y1              =   3405
      Y2              =   3420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Código"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6555
      TabIndex        =   20
      Top             =   495
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   10
      Left            =   6690
      TabIndex        =   17
      Top             =   15
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   2565
      TabIndex        =   12
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   165
      TabIndex        =   10
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3300
      TabIndex        =   8
      Top             =   1410
      Width           =   255
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   165
      TabIndex        =   6
      Top             =   1410
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   4
      Top             =   870
      Width           =   825
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sobrenome:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2790
      TabIndex        =   2
      Top             =   585
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VisualizarImpressão()
impresso.Text = "Código: " & Label4.Caption _
& vbCrLf & "Nome: " & txtFields(0).Text _
& vbCrLf & "Sobrenome: " & txtFields(1).Text _
& vbCrLf & "Endereço: " & txtFields(2).Text _
& vbCrLf & "Bairro: " & txtFields(3).Text _
& vbCrLf & "CEP: " & txtFields(4).Text _
& vbCrLf & "Cidade: " & txtFields(5).Text _
& vbCrLf & "UF: " & txtFields(6).Text _
& vbCrLf & "Telefone: " & txtFields(7).Text _
& vbCrLf & "E-Mail: " & txtFields(8).Text _
& vbCrLf & "Arquivo de Foto: " & txtFields(9).Text
End Sub
Function FileExists(filename As String) As Integer
    On Error Resume Next
        x% = Len(Dir$(filename))
    If err Or x% = 0 Then FileExists = False Else FileExists = True
End Function

Private Sub Check1_Click()
impresso.FontBold = Check1.Value
End Sub

Private Sub chameleonButton1_Click()
On Error Resume Next
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
crit.Text = Empty
End Sub

Private Sub close_Click()
cmdClose_Click
End Sub

Private Sub cmbFonte_Change()
impresso.Font = cmbFonte.Text
End Sub

Private Sub cmbFonte_Click()
impresso.Font = cmbFonte.Text
End Sub

Private Sub cmbTamanho_Change()
impresso.Font.Size = cmbTamanho.Text
End Sub

Private Sub cmbTamanho_Click()
impresso.Font.Size = cmbTamanho.Text
End Sub

Private Sub cmdAdd_Click()
Dim codigo As String
codigo = txtFields(10).Text + 1

On Error Resume Next


   If txtFields(0).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(0).DataField & "'", vbCritical, Me.Caption
  Else
    If txtFields(1).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(1).DataField & "'", vbCritical, Me.Caption
  Data1.UpdateRecord
    Else
        If txtFields(2).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(2).DataField & "'", vbCritical, Me.Caption
  Else
        If txtFields(3).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(3).DataField & "'", vbCritical, Me.Caption
  Else
  If txtFields(4).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(4).DataField & "'", vbCritical, Me.Caption
  Else
  If txtFields(5).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(5).DataField & "'", vbCritical, Me.Caption
  Else
  If txtFields(6).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(6).DataField & "'", vbCritical, Me.Caption
  Else
  If txtFields(7).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(7).DataField & "'", vbCritical, Me.Caption
  Else
  If txtFields(8).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(8).DataField & "'", vbCritical, Me.Caption
  Else
If txtFields.Item(0).Text = "" Then
   Data1.Recordset.AddNew
     Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  cmdDelete.Enabled = True
   Data1.Recordset.AddNew
Else
 
    If txtFields(10).Text = Empty Then
  MsgBox "Preencha o campo '" & txtFields(10).DataField & "'", vbCritical, Me.Caption
  Else
  'Data1.Recordset.Bookmark = Data1.Recordset.LastModified
  Data1.Recordset.AddNew
  txtFields(10).Text = codigo
  Label4.Caption = codigo
  cmdDelete.Enabled = False
  End If
  End If
End If
  End If
  End If
  End If
    End If
  End If
  End If
End If
End If
txtFields(10).Locked = False

End Sub

Private Sub cmdChangeFoto_Click()
On Error Resume Next
cd.filename = Empty
cd.ShowOpen
txtFields(9).Text = cd.filename
  

   Image1.Picture = LoadPicture(txtFields(9).Text)
   Data1.UpdateRecord
   Data1.Recordset.MoveLast

End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim mensagem As Integer
Dim a As String

If Not Data1.Recordset.RecordCount = 1 Then
If txtFields(0).Text = "" Then
Exit Sub
Else
a = txtFields(0).Text
End If
mensagem = MsgBox("Tem certeza que deseja excluir o registro: " & a & " ?", vbYesNo, "!ATENÇÃO OPERADOR!")
Select Case mensagem
Case vbYes
If txtFields.Item(0).Text = "" Then
MsgBox "Não há nenhum registro à apagar!"", vbCritical, Me.Caption"

Else

   Data1.Recordset.Delete
Data1.Recordset.Bookmark = Data1.Recordset.LastModified

If Not txtFields(9).Text = Empty Then
     Image1.Picture = LoadPicture(txtFields(9).Text)
  End If
  End If
  'this may produce an error if you delete the last
  'record or the only record in the recordset
Case Else
End Select
Else
MsgBox "Não é possível excluir o último registro!", vbCritical + vbOKOnly, Me.Caption
End If
Data1.Refresh
Data1.Recordset.MoveLast
End Sub


Private Sub cmdRefresh_Click()
On Error Resume Next
  'this is really only needed for multi user apps
  Data1.Refresh

     Image1.Picture = LoadPicture(txtFields(9).Text)
     VisualizarImpressão
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified

     Image1.Picture = LoadPicture(txtFields(9).Text)
VisualizarImpressão
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
If MsgBox("Deseja realmente sair do programa?", vbYesNo, Me.Caption) = vbYes Then Unload Me
  
End Sub



Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
txtFields(9).Text = "Sem foto"
Image1.Picture = LoadPicture("")
Image1.Refresh
End Sub


Private Sub Command4_Click()
On Error Resume Next
Printer.Font.Bold = option1.Value
Printer.Font.Italic = option2.Value
Printer.Font.Name = cmbFonte.Text
Printer.Font.Size = cmbTamanho.Text
On Error GoTo err
Dim intResp As Integer
intResp = MsgBox("Deseja realmente imprimir os dados?", vbYesNo, App.Title)
Select Case intResp
Case vbYes
Printer.Print "Gerenciamento de Contatos, criado por Gabriel Falcão - gabrielfalcao@hotmail.com" & vbCrLf & impresso.Text
Printer.EndDoc
Case Else
End Select
err:
Me.Caption = err.Description
Exit Sub
End Sub

Private Sub Command5_Click()
VisualizarImpressão
End Sub

Private Sub Command7_Click()
On Error GoTo err
savetxt.ShowSave
Open savetxt.filename For Output As #1
Print #1, "Gerenciamento de Contatos, criado por Gabriel Falcão - gabrielfalcao@hotmail.com" & vbCrLf & impresso.Text
Close #1
err:
Exit Sub
End Sub

Private Sub Command6_Click()
impresso.Font = "Tahoma"
impresso.FontSize = 8
End Sub

Private Sub crit_GotFocus()
chameleonButton1.Default = True
End Sub

Private Sub crit_KeyPress(KeyAscii As Integer)
chameleonButton1.Default = True
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
On Error Resume Next
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
End Sub

Private Sub Data1_Reposition()
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots

  Data1.Caption = "Total: " & (Data1.Recordset.AbsolutePosition + 1)
    Image1.Picture = LoadPicture(txtFields(9).Text)
If Image1.Picture = "Sem Foto" Then
Image1.Picture = Empty
End If
VisualizarImpressão
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, save As Integer)
  'This is where you put validation code
  On Error Resume Next


  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Data1.Caption = "Total: " & (Data1.Recordset.AbsolutePosition + 1)
   Image1.Picture = LoadPicture(txtFields(9).Text)
VisualizarImpressão
End Sub

Private Sub del_Click()
cmdDelete_Click
End Sub

Private Sub Form_Activate()
On Error Resume Next
VisualizarImpressão
    Image1.Picture = LoadPicture(txtFields(9).Text)
'txtFields(10).Text = Data1.Recordset.Index + 1
'Data1.Caption = "Agenda - Total cadastrado: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub

Private Sub Form_Load()
On Error Resume Next
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

If FileExists("C:\Agenda\Agenda.mdb") = True Then
Data1.DatabaseName = "C:\Agenda\Agenda.mdb"
Else
Data1.DatabaseName = "Agenda.mdb"
End If

If Not txtFields(9).Text = Empty Then
   Image1.Picture = LoadPicture(txtFields(9).Text)
   End If
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
txtFields(10).Locked = False
If txtFields(0).Text = Empty Then
cmdDelete.Enabled = False
Else
cmdDelete.Enabled = True
End If

Dim Contador
For Contador = 1 To Screen.FontCount - 1
cmbFonte.AddItem Screen.Fonts(Contador)
Next
Dim Tamanho
For Tamanho = 6 To 72
cmbTamanho.AddItem (Tamanho)
Next
End Sub

Private Sub new_Click()
cmdAdd_Click
End Sub

Private Sub Option1_Click()
If impresso.FontBold = True Then
impresso.FontBold = False
Else
impresso.FontBold = True
End If
End Sub

Private Sub Option2_Click()
If impresso.FontItalic = True Then
impresso.FontItalic = False
Else
impresso.FontItalic = True
End If

End Sub



Private Sub txtFields_Click(Index As Integer)
'Dim resultado As Integer
'Select Case Index
'Case Is = 10
'resultado = MsgBox("Deseja mudar o código?", vbYesNo + vbInformation, Me.Caption)
'Select Case resultado
'Case vbYes
'txtFields(10).Locked = True
'Case vbNo
'txtFields(10).Locked = False
'Case Else
'End Select
'Case Else
'End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case Index

Case Is = 0
If Len(txtFields(0).Text) > 2 Then
cmdDelete.Enabled = True
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
Else
cmdDelete.Enabled = False
End If
Case Else
End Select

End Sub
