VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00A4E3AC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informações sobre a BIOS do sistema"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "bios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F0E7&
      ForeColor       =   &H00004000&
      Height          =   3150
      ItemData        =   "bios.frx":08CA
      Left            =   52
      List            =   "bios.frx":08CC
      TabIndex        =   1
      Top             =   3360
      Width           =   7680
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F0E7&
      ForeColor       =   &H00004000&
      Height          =   2565
      Left            =   52
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   415
      Width           =   7680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispositivos suportados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   52
      TabIndex        =   3
      Top             =   3050
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalhes da BIOS do sistema:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   52
      TabIndex        =   2
      Top             =   105
      Width           =   2550
   End
End
Attribute VB_Name = "Form1"
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
    Str1 = Str1 & "Descrição" & vbTab & ":" & vbTab & ObjBios.Description & vbCrLf
    Str1 = Str1 & "Data de Liberação" & vbTab & ":" & vbTab & Format(Left(ObjBios.Version, 8), "DD - mmmm - YYYY") & vbCrLf
    Str1 = Str1 & "Status" & vbTab & vbTab & ":" & vbTab & ObjBios.Status & vbCrLf
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

