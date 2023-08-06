VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   375
      Left            =   5175
      TabIndex        =   1
      Top             =   3375
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Text            =   "filecopy"
      Top             =   3420
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ori As String
Dim des As String
Dim tem As String
Dim l As Integer
If Left$(Text1.Text, 9) = "filecopy " Then
tem = Right$(Text1.Text, Len(Text1.Text) - 9)
Me.Print tem
l = InStr(1, tem, "?", vbTextCompare) - 1
ori = Left$(tem, l)
des = Right$(tem, Len(tem) - l - 1)
Me.Print "ORIGEM: " & ori & vbCrLf & "DESTINO:" & des
End If
End Sub
