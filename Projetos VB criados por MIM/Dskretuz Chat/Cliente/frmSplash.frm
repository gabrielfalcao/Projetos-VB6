VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4170
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4170
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2730
      Top             =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde Carregando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1905
      TabIndex        =   0
      Top             =   3705
      Width           =   1965
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim m_Rgn As CBMPRegion

Private Sub ctr_Click(Index As Integer)
Select Case Index
Case Is = 0
End
Case Is = 1
Me.BorderStyle = 3
Me.WindowState = 1
End Select
End Sub


Private Sub Form_Load()
Set m_Rgn = New CBMPRegion

  m_Rgn.CreateFromPic Me.Picture, vbWhite
  SetWindowRgn hwnd, m_Rgn.Handle, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetWindowRgn hwnd, 0, False
  m_Rgn.Destroy
  Set m_Rgn = Nothing
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ctr.ForeColor = &HFFFF&
End Sub

Private Sub Timer1_Timer()
frmChatCliente.Show
Unload Me
End Sub
