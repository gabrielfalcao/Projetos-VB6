VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discar para UAI"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Comando1 
      Appearance      =   0  'Flat
      Caption         =   "Conectar"
      Height          =   570
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function conectar(conexão As String)
Dim x
x = Shell("rundll32.exe rnaui.dll,RnaDial " & conexão, 1)
DoEvents
SendKeys "{6}{0}{0}{5}"
SendKeys "{TAB}"
SendKeys "{3}{2}{6}{3}{7}{3}{5}{6}"
'32637356
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{j}{m}{p}{m}"
End Function
Private Sub Comando1_Click()
conectar "Ubbi"
Unload Me
End Sub
