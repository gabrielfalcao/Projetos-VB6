VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form statustxt 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmailSubject 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   7
      Text            =   "Assunto aqui"
      Top             =   1380
      Width           =   3210
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1605
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "statustxt.frx":0000
      Top             =   1695
      Width           =   3210
   End
   Begin VB.CommandButton CmdSendMail 
      Caption         =   "Enviar"
      Height          =   315
      Left            =   3135
      TabIndex        =   5
      Top             =   3075
      Width           =   1080
   End
   Begin VB.TextBox txtFromEmailAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   4
      Text            =   "gabrielfalcaogm@bol.com.br"
      Top             =   750
      Width           =   3210
   End
   Begin VB.TextBox txtToEmailAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   3
      Text            =   "gabrielfalcao@hotmail.com"
      Top             =   1065
      Width           =   3210
   End
   Begin VB.TextBox txtEmailServer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   2
      Text            =   "smtp.bol.com.br"
      Top             =   435
      Width           =   3210
   End
   Begin VB.TextBox txtFromName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   1
      Text            =   "Prego"
      Top             =   120
      Width           =   3210
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   3015
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label statustxt 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3630
      Width           =   6795
   End
End
Attribute VB_Name = "statustxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String
Dim Start As Single, Tmr As Single


Sub SendEmail(MailServerName As String, FromName As String, fromemailaddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
Dim DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String
With Winsock1
    If .State = sckClosed Then ' Check to see if socket is closed
        DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
        first = "mail from: " & fromemailaddress & vbCrLf ' Get who's sending E-Mail address
        Second = "rcpt to: " & ToEmailAddress & vbCrLf ' Get who mail is going to
        Third = "Date: " & DateNow & vbCrLf ' Date when being sent
        Fourth = "From: """ & FromName & """ <" & fromemailaddress & ">" + vbCrLf ' Who's Sending
        Fifth = "To: " & ToNametxt & vbCrLf ' Who it going to
        Sixth = "Subject: " & EmailSubject & vbCrLf ' Subject of E-Mail
        Seventh = EmailBodyOfMessage & vbCrLf ' E-mail message body
        Ninth = "X-Mailer: STMP Sender" & vbCrLf ' What program sent the e-mail, customize this
        .LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail per program start
        .Protocol = sckTCPProtocol ' Set protocol for sending
        .RemoteHost = MailServerName ' Set the server address
        .RemotePort = 25 ' Set the SMTP Port
        .Connect ' Start connection
        WaitFor ("220")
        statustxt.Caption = "Connecting...."
        .SendData ("HELO EnterComputerNameHere" & vbCrLf)
        WaitFor ("250")
        statustxt.Caption = "Connected"

        .SendData (first)
        statustxt.Caption = "Sending Message"

        WaitFor ("250")
        .SendData (Second)
        WaitFor ("250")
        .SendData ("data" & vbCrLf)
        WaitFor ("354")
        .SendData (Fourth & Third & Ninth & Fifth & Sixth & vbCrLf)
        .SendData (Seventh & vbCrLf)
        .SendData ("." & vbCrLf)
        WaitFor ("250")
        .SendData ("quit" & vbCrLf)
        statustxt.Caption = "Disconnecting"

        WaitFor ("221")
        .Close
    Else
        MsgBox (Str(.State))
    End If
End With
End Sub


Sub WaitFor(ResponseCode As String)
    Start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
           MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
           Exit Sub
        End If
    Wend
    Response = "" ' Sent response code to blank **IMPORTANT**
End Sub


Private Sub CmdSendMail_Click()
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtMessage.Text
    statustxt.Caption = "Enviando"
    Beep
    Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*
End Sub


