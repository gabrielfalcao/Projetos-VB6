VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lembrete 1.0"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDia 
      Height          =   1035
      Left            =   1725
      TabIndex        =   1
      Top             =   420
      Width           =   1575
   End
   Begin VB.ListBox lstMes 
      Height          =   2400
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2235
      TabIndex        =   3
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mêses:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   450
      TabIndex        =   2
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim agenda As New clsIniFiles
Dim mepath As String
   Dim Cnt As Long
Private Sub Form_Load()
If Right$(App.Path, 1) = "\" Then
mepath = App.Path
Else
mepath = App.Path & "\"
End If
agenda.IniFile = mepath & "agenda.ini"
 Dim dia As New Collection
    Set dia = agenda.ReadSection("Dia")
 
    For Cnt = 1 To dia.Count
        lstDia.AddItem dia.Item(Cnt)  ' & vbCrLf
    Next Cnt
    
     Dim mes As New Collection
    Set mes = agenda.ReadSection("Mes")
 
    For Cnt = 1 To mes.Count
        lstMes.AddItem mes.Item(Cnt)  ' & vbCrLf
    Next Cnt
End Sub
