VERSION 5.00
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Begin VB.Form frmDocument 
   BackColor       =   &H00E0E0E0&
   Caption         =   "frmDocument"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   FillColor       =   &H00E9C7B8&
   ForeColor       =   &H00404040&
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   5595
   Begin CodeMaxCtl.CodeMax rtftext 
      Height          =   3705
      Left            =   60
      OleObjectBlob   =   "frmDocument.frx":0A02
      TabIndex        =   0
      Top             =   45
      Width           =   5490
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Form_Resize
Me.WindowState = 2
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtftext.Move 5, 5, Me.ScaleWidth - 10, Me.ScaleHeight - 10
 
  rtftext.SetColor cmClrLeftMargin, &HE0E0E0
  
End Sub


