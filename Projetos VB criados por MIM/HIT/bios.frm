VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0000C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre o H1T"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   Icon            =   "bios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   195
      Picture         =   "bios.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   4
      Top             =   75
      Width           =   510
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1185
      ItemData        =   "bios.frx":1194
      Left            =   135
      List            =   "bios.frx":1196
      TabIndex        =   1
      Top             =   4545
      Width           =   5445
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1185
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2970
      Width           =   5430
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   600
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   5564
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "H1T - H4ck3r 1nt3rn3t T00lz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   960
      TabIndex        =   9
      Top             =   60
      Width           =   4545
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Criado para satisfazer as necessidades dos melhores programadores, geeks, hackers, etc..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   960
      TabIndex        =   8
      Top             =   930
      Width           =   4080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "www.tebugho.i8.com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1725
      TabIndex        =   7
      Top             =   1710
      Width           =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "tebugho@hotmail.com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1665
      TabIndex        =   6
      Top             =   1980
      Width           =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5564
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000054FF&
      Caption         =   "&OK"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4410
      TabIndex        =   5
      Top             =   2475
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivos suportados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   135
      TabIndex        =   3
      Top             =   4275
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalhes da BIOS do sistema:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   150
      Left            =   180
      TabIndex        =   2
      Top             =   2790
      Width           =   1890
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Characteristics(40) As String
Dim bc() As String
Dim ObjWmi As Object
Dim ObjBiosCollection As Object
Dim ObjBios As Object
Dim Str1 As String

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
lblVersion = "v" & App.Major & "." & App.Minor & App.Revision
Characteristics(0) = "Reservado"
Characteristics(1) = "Reservado"
Characteristics(2) = "Desconhecido"
Characteristics(3) = "Características da BIOS não suportadas"
Characteristics(4) = "Suporte à ISA"
Characteristics(5) = "Suporte à MCA"
Characteristics(6) = "Suporte à EISA"
Characteristics(7) = "Suporte à PCI"
Characteristics(8) = "Suporte à PC Card (PCMCIA)"
Characteristics(9) = "Suporte à Plug and Play"
Characteristics(10) = "Suporte à APM"
Characteristics(11) = "A BIOS é atualizável (Flash)"
Characteristics(12) = "Sombreamento de BIOS Habilitado"
Characteristics(13) = "Suporte à VL-VESA"
Characteristics(14) = "Suporte à ESCD está disponível"
Characteristics(15) = "Suporte à Boot pelo CD-ROM"
Characteristics(16) = "Suporte à seleção de BOOT"
Characteristics(17) = "A ROM da BIOS ROM é SOCKET"
Characteristics(18) = "Suporte à Boot pelo PC Card (PCMCIA)"
Characteristics(19) = "Suporte à especificação EDD (Enhanced Disk Drive)"
Characteristics(20) = "Int 13h - Suporte à Floppy Japonês NEC 9800 1.2mb (3.5, 1k Bytes/Setor, 360 RPM)"
Characteristics(21) = "Int 13h - Suporte à Floppy Japonês Toshiba 1.2mb (3.5, 360 RPM)"
Characteristics(22) = "Int 13h - Suporte aos Serviços de Floppy 5.25 / 360 KB"
Characteristics(23) = "Int 13h - Suporte aos Serviços de Floppy 5.25 /1.2MB"
Characteristics(24) = "Int 13h - Suporte aos Serviços de Floppy 3.5 / 720 KB"
Characteristics(25) = "Int 13h - Suporte aos Serviços de Floppy 3.5 / 2.88 MB"
Characteristics(26) = "Int 5h, Suporte à Print Screen"
Characteristics(27) = "Int 9h, 8042 Suporte à Teclado"
Characteristics(28) = "Int 14h, Suporte à Serviços do Serial"
Characteristics(29) = "Int 17h, Suporte à Impressora"
Characteristics(30) = "Int 10h, Suporte à Vídeo CGA/Mono"
Characteristics(31) = "NEC PC-98"
Characteristics(32) = "Suporte à ACPI"
Characteristics(33) = "Suporte à USB Legacy"
Characteristics(34) = "Suporte à AGP"
Characteristics(35) = "Suporte à Boot I2O"
Characteristics(36) = "Suporte à Boot LS-120"
Characteristics(37) = "Suporte à Boot através de ATAPI ZIP Drive"
Characteristics(38) = "Suporte à Boot1394"
Characteristics(39) = "Suporte à Bateria Pequena"

Set ObjWmi = GetObject("Winmgmts:")

Set ObjBiosCollection = ObjWmi.ExecQuery("SELECT Name,BiosCharacteristics,SMBIOSPresent,SMBIOSMajorVersion,SMBIOSMinorVersion,Version,SMBIOSBIOSVersion,ListOfLanguages,Description,CurrentLanguage,Manufacturer,Status FROM Win32_BIOS")


For Each ObjBios In ObjBiosCollection

    For i = 0 To UBound(ObjBios.BiosCharacteristics) - 1
        List1.AddItem Characteristics(ObjBios.BiosCharacteristics(i))
    Next
    
    Str1 = Str1 & "Fabricante" & vbTab & ":" & vbTab & ObjBios.Manufacturer & vbCrLf
    Str1 = Str1 & "Descrição" & vbTab & vbTab & ":" & vbTab & ObjBios.Description & vbCrLf
    Str1 = Str1 & "Data de Liberação" & vbTab & ":" & vbTab & Format(Left(ObjBios.Version, 8), "DD - mmmm - YYYY") & vbCrLf
    Str1 = Str1 & "Status" & vbTab & vbTab & ":" & vbTab & ObjBios.status & vbCrLf
    Str1 = Str1 & "Versão da BIOS" & vbTab & ":" & vbTab & ObjBios.SMBIOSBIOSVersion & vbCrLf
    Str1 = Str1 & "Linguagem atual" & vbTab & ":" & vbTab & ObjBios.CurrentLanguage & vbCrLf
    
    Str1 = Str1 & "----------------------------------------------------------" & vbCrLf
    Str1 = Str1 & "SM BIOS Presente" & vbTab & vbTab & ":" & vbTab & ObjBios.SMBIOSPresent & vbCrLf
    Str1 = Str1 & "SM BIOS Versão(Maior)" & vbTab & ":" & vbTab & ObjBios.SMBIOSMinorVersion & vbCrLf
    Str1 = Str1 & "SM BIOS Versão(Menor)" & vbTab & ":" & vbTab & ObjBios.SMBIOSMajorVersion & vbCrLf
    Str1 = Str1 & "----------------------------------------------------------" & vbCrLf
    
    For i = 1 To UBound(ObjBios.ListOfLanguages) - 1
        If Trim(ObjBios.ListOfLanguages(i)) <> vbNullString Then
            Str1 = Str1 & "Linguagens instaláveis da BIOS" & vbTab & " : " & vbTab & ObjBios.ListOfLanguages(i) & vbCrLf
        End If
    Next
    
    Str1 = Str1 & "----------------------------------------------------------" & vbCrLf
Next

Text1 = Str1


End Sub
Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HFF0000
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HC0C0C0
End Sub

