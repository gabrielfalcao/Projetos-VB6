VERSION 5.00
Begin VB.Form frmImag 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D6D6FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POP Anonimail v0.1 by 7£r@t0$"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMail.frx":37B2
   ScaleHeight     =   3975
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSMTPServer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1035
      TabIndex        =   11
      Text            =   "smtp.pop.com.br"
      Top             =   1110
      Width           =   3315
   End
   Begin VB.TextBox TXTmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCE1B6&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1185
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmMail.frx":1952C
      Top             =   1905
      Width           =   4215
   End
   Begin Project1.Winsock Winsock1 
      Left            =   735
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox statusb 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FEFADE&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4740
      TabIndex        =   9
      Top             =   3720
      Width           =   4800
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H80000004&
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3210
      Width           =   1140
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FCE1B6&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6030
      TabIndex        =   2
      Text            =   "pop@pop.com.br"
      Top             =   300
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCE1B6&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1035
      TabIndex        =   1
      Text            =   "alguem@pop.com.br"
      Top             =   150
      Width           =   3315
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCE1B6&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1035
      TabIndex        =   0
      Text            =   "Mensagem do Servidor"
      Top             =   630
      Width           =   3315
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.pop.com.br"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmMail.frx":19547
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   3375
      Width           =   2955
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "www.tebugho.i8.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmMail.frx":19851
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3105
      Width           =   2955
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   8
      Top             =   150
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5925
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Assunto:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   630
      Width           =   870
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servidor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1110
      Width           =   870
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1635
      Width           =   4215
   End
End
Attribute VB_Name = "frmImag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NewImg As String
Dim cDraw As cDrawImage
Dim server As String
Dim from As String
Private Declare Function ShellExecute Lib "Shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
  Set cDraw = New cDrawImage
 cDraw.ModoDeExibição = [Lado a Lado]
      statusb.Cls
      from = "campolinadiniz@terra.com.br"
      server = "smtp.poa.terra.com.br"
           statusb.Print "Pronto"
Me.Caption = "POP Anonimail v0.1." & App.Revision & " by 7£r@t0$"
End Sub
Private Sub Form_Paint()
  Me.Cls
  cDraw.hDC = Me.hDC
  cDraw.hwnd = Me.hwnd
  cDraw.ImgHandle = Me.Picture.Handle
  cDraw.TileBlt
  Me.Refresh
End Sub
Private Sub Form_Resize()
  Call Form_Paint
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set cDraw = Nothing
  Set frmImag = Nothing
End Sub

Private Sub Label6_Click()
    Dim lRet As Long
    lRet = ShellExecute(0, "open", Label6.Caption, vbNullString, vbNullString, 1)
End Sub

Private Sub Label7_Click()
    Dim lRet As Long
    lRet = ShellExecute(0, "open", Label7.Caption, vbNullString, vbNullString, 1)
End Sub

Private Sub Winsock1_Connect()

  Winsock1.Tag = "conectado"

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim strdata As String
Dim MsgTexto As String
Dim status As String
Dim Erro As Boolean

If Trim(Winsock1.Tag) <> "" Then
  Winsock1.GetData strdata
  status = Left(strdata, 3)
  
  'Verifica de o servidor retornou alguma msg de erro
  Select Case status
     Case "250", "220", "354", "221": Erro = False
     Case Else:
       Erro = True
       Winsock1.Tag = "fechar"
       status = Mid(strdata, 4)
  End Select
  
  Select Case Winsock1.Tag
    Case "conectado":
      Winsock1.SendData "helo " & Winsock1.LocalIP & vbCrLf
      Winsock1.Tag = "conectou"
     statusb.Cls
     statusb.Print "Conectado."
    
    Case "conectou":
         statusb.Cls
     statusb.Print "Enviando..."
      Winsock1.SendData "mail from:<" & txtFrom.Text & ">" & vbCrLf
      Winsock1.Tag = "from"
    
    Case "from":
      Winsock1.SendData "rcpt to:<" & txtTo.Text & ">" & vbCrLf
      Winsock1.Tag = "to"
    
    Case "to":
      Winsock1.SendData "data" & vbCrLf
      Winsock1.Tag = "data"
      
    Case "data":
      'A sequencia "." e quebra de linha deve ser substituida por ".." e quebra de linha
      'para evitar que o servidor entenda fim de email antes do fim do texto
      MsgTexto = TXTmsg.Text & vbCrLf
      While InStr(MsgTexto, vbCrLf & "." & vbCrLf) <> 0
        MsgTexto = Replace(MsgTexto, vbCrLf & "." & vbCrLf, vbCrLf & ".." & vbCrLf)
      Wend
      
      Winsock1.SendData "subject: " & txtSubject & vbCrLf & MsgTexto & vbCrLf & "." & vbCrLf
      Winsock1.Tag = "fim"
      
    Case "fim":
         statusb.Cls
     statusb.Print "Desconectando..."
      Winsock1.SendData "quit" & vbCrLf
      Winsock1.Tag = "fechar"
      
    Case "fechar":
      If Not Erro Then
           statusb.Cls
       statusb.Print "Enviado com sucesso!"
      Else
           statusb.Cls
       statusb.Print "Erro ao enviar email!"
        MsgBox status, vbCritical, "Erro"
      End If
      
      Winsock1.CloseSck
      Winsock1.Tag = ""
  
  End Select
  
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  MsgBox "Erro ao conectar" & vbNewLine & "Verifique sua conexão ou o endereço do servidor", vbCritical, "Erro"
End Sub
Private Sub cmdEnviar_Click()
Dim a1 As Integer
Dim a2 As String
Dim l As String
a1 = InStr(1, txtTo.Text, "@", vbTextCompare)
a1 = a1 - 1
l = Len(txtTo.Text) - a1

a2 = Right$(txtTo.Text, l)
If a2 = "@pop.com.br" Then
  'Verificar se nenhuma conexão está em andamento
    
  If Winsock1.Tag = "" Then
    If Winsock1.State <> sckClosed Then Winsock1.CloseSck
    Winsock1.Connect txtSMTPServer.Text, 25
  End If
  Else
  If MsgBox("O destinatário não é um email POP, deseja tentar mesmo assim?", vbYesNo, Me.Caption) = vbYes Then
    If Winsock1.Tag = "" Then
    If Winsock1.State <> sckClosed Then Winsock1.CloseSck
    Winsock1.Connect txtSMTPServer.Text, 25
  End If
  End If
  End If
End Sub
