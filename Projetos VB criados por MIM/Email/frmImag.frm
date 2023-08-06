VERSION 5.00
Begin VB.Form frmImag 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0069C4FA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terra & Yahoo Anonimail v0.1.22"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmImag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmImag.frx":1042
   ScaleHeight     =   4485
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTmsg 
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
      Left            =   225
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2295
      Width           =   4215
   End
   Begin Project1.Winsock Winsock1 
      Left            =   720
      Top             =   2820
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ComboBox txtSMTPServer 
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
      Height          =   315
      Left            =   1095
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1545
      Width           =   3315
   End
   Begin VB.PictureBox statusb 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H0099D8F7&
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4950
      TabIndex        =   9
      Top             =   4230
      Width           =   5010
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H000080FF&
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3645
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
      Left            =   1095
      TabIndex        =   2
      Text            =   "terra@terra.com.br"
      Top             =   135
      Width           =   3315
   End
   Begin VB.TextBox txtTo 
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
      Left            =   1095
      TabIndex        =   1
      Top             =   615
      Width           =   3315
   End
   Begin VB.TextBox txtSubject 
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
      Left            =   1095
      TabIndex        =   0
      Top             =   1095
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   225
      TabIndex        =   8
      Top             =   615
      Width           =   465
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
      Left            =   225
      TabIndex        =   7
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1095
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   1575
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   2055
      Width           =   990
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
Private Sub Form_Load()
  Set cDraw = New cDrawImage
 cDraw.ModoDeExibição = [Lado a Lado]
      statusb.Cls
      from = "campolinadiniz@terra.com.br"
      server = "smtp.poa.terra.com.br"
           statusb.Print "Pronto"
           txtSMTPServer.AddItem server
           txtSMTPServer.AddItem "smtp.mail.yahoo.com.br"
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
If a2 = "@terra.com.br" Or "@yahoo.com.br" Or "@yahoo.com" Then
  'Verificar se nenhuma conexão está em andamento
  If a2 = "@yahoo.com.br" Or "@yahoo.com" Then
  txtFrom.Text = "yahoo@yahoo.com.br"
  If Winsock1.Tag = "" Then
    If Winsock1.State <> sckClosed Then Winsock1.CloseSck
    Winsock1.Connect txtSMTPServer.Text, 25
  End If
  Else
  If MsgBox("O destinatário não é um email Terra, deseja tentar mesmo assim?", vbYesNo, Me.Caption) = vbYes Then
    If Winsock1.Tag = "" Then
    If Winsock1.State <> sckClosed Then Winsock1.CloseSck
    Winsock1.Connect txtSMTPServer.Text, 25
  End If
  End If
  End If
End Sub
