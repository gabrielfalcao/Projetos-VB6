VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl SMTP 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   ScaleHeight     =   375
   ScaleWidth      =   735
   Begin MSWinsockLib.Winsock Sock 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "SMTP.ctx":0000
      ScaleHeight     =   375
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------
'
'   This is the WinsockVB.com SMTP ActiveX UserControl v1.0
'   You are free to redistribute this control as long as this
'   section remains intact and unchanged - give credit where it
'   is due. Thankyou
'   Other than that - Enjoy using it!!
'
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
'
'   EVENTS
'
' ------------------------------------------------------------------------------
Public Event Connected(ByVal Host As String, ByVal Port As Long)
Public Event ReceivedData(ByVal Data As String)
Public Event SentData(ByVal Data As String)
Public Event MailCompleted()
Public Event Error(ByVal Error As String)

' ------------------------------------------------------------------------------
'
'   PROPERTY VARIABLES
'
' ------------------------------------------------------------------------------
Dim m_Server As String      ' mail server host
Dim m_Port As String        ' mail server port
Dim m_MailFrom As String    ' from address
Dim m_MailTo As String      ' to address
Dim m_BCC As String         ' blind carbon copy addresses
Dim m_CCC As String         ' carbon copy addresses
Dim m_Subject As String     ' email subject
Dim m_NameFrom As String    ' from name
Dim m_NameTo As String      ' to name
Dim m_Body As String        ' email body
Dim m_Log As String         ' log of transaction

' private state variables
Dim LastResponse As String

' ------------------------------------------------------------------------------
'
'   PUBLIC PROPERTIES
'
' ------------------------------------------------------------------------------
    Public Property Get Server() As String
        Server = m_Server
    End Property
    
    Public Property Let Server(ByVal Data As String)
        m_Server = Data
    End Property
    
    Public Property Get Port() As String
        Port = m_Port
    End Property
    
    Public Property Let Port(ByVal Data As String)
        m_Port = Data
    End Property
    
    Public Property Get MailFrom() As String
        MailFrom = m_MailFrom
    End Property
    
    Public Property Let MailFrom(ByVal Data As String)
        m_MailFrom = Data
    End Property
    
    Public Property Get MailTo() As String
        MailTo = m_MailTo
    End Property
    
    Public Property Let MailTo(ByVal Data As String)
        m_MailTo = Data
    End Property
    
    Public Property Get BCC() As String
        BCC = m_BCC
    End Property
    
    Public Property Let BCC(ByVal Data As String)
        m_BCC = Data
    End Property
    
    Public Property Get CCC() As String
        CCC = m_CCC
    End Property
    
    Public Property Let CCC(ByVal Data As String)
        m_CCC = Data
    End Property
    
    Public Property Get Subject() As String
        Subject = m_Subject
    End Property
    
    Public Property Let Subject(ByVal Data As String)
        m_Subject = Data
    End Property
    
    Public Property Get NameTo() As String
        NameTo = m_NameTo
    End Property
    
    Public Property Let NameTo(ByVal Data As String)
        m_NameTo = Data
    End Property
    
    Public Property Get NameFrom() As String
        NameFrom = m_NameFrom
    End Property
    
    Public Property Let NameFrom(ByVal Data As String)
        m_NameFrom = Data
    End Property
    
    Public Property Get Body() As String
        Body = m_Body
    End Property
    
    Public Property Let Body(ByVal Data As String)
        m_Body = Data
    End Property
    
    Public Property Get Log() As String
        Log = m_Log
    End Property
    
    Public Property Let Log(ByVal Data As String)
        m_Log = Data
    End Property
    
' ------------------------------------------------------------------------------
'
'   PUBLIC SUBS
'
' ------------------------------------------------------------------------------
    Public Function SendMail() As Boolean
    Dim SMTPCommands(0 To 10) As String
    Dim SMTPResponses(0 To 10) As String
    Dim Success As Boolean
    Dim i As Integer
    
        ' construct an array of commands
        SMTPCommands(0) = "HELO " & Me.Server
        SMTPCommands(1) = "MAIL FROM:" & Me.MailFrom
        SMTPCommands(2) = "RCPT TO:" & Me.MailTo
        SMTPCommands(3) = "DATA"
        SMTPCommands(4) = "BCC:" & Me.BCC
        SMTPCommands(5) = "CCC:" & Me.CCC
        SMTPCommands(6) = "SUBJECT:" & Me.Subject
        SMTPCommands(7) = "TO:" & Me.NameTo
        SMTPCommands(8) = "FROM:" & Me.NameFrom & vbCrLf ' extra vbCrLf
        SMTPCommands(9) = Me.Body & vbCrLf & "."
        SMTPCommands(10) = "QUIT"
        
        SMTPResponses(0) = "250"
        SMTPResponses(1) = "250"
        SMTPResponses(2) = "250"
        SMTPResponses(3) = "354"
        SMTPResponses(4) = ""
        SMTPResponses(5) = ""
        SMTPResponses(6) = ""
        SMTPResponses(7) = ""
        SMTPResponses(8) = ""
        SMTPResponses(9) = "250"
        SMTPResponses(10) = "221"
        
        ' connect to the server
        If ConnectToServer = False Then
            RaiseEvent Error("Couldn't connect to server")
            Exit Function
        Else
            ' wait for the welcome message to be received
            WaitForResponse "220"
   End If
        
        ' send each command, waiting for a response
        For i = 0 To 10
        
            ' send the command
            SMTPSend SMTPCommands(i)
            
            ' wait for the response
            Success = WaitForResponse(SMTPResponses(i))
            
            ' check if we were successful
            If Success = False Then
                RaiseEvent Error("Failed at server side. Check the SMTP.Log property for more details")
                Exit Function
   End If
        
        Next i
        
        ' finished
        RaiseEvent MailCompleted
    End Function
    
    Private Function ConnectToServer() As Boolean

        ' connect to the host
        Sock.RemoteHost = Me.Server
         Sock.RemotePort = Me.Port
        Sock.Connect
        
        ' wait for connection
        Do While Sock.State <> sckConnected
            DoEvents
            If Sock.State = sckError Then
                Exit Function
   End If
        Loop
        
        ' return true
        ConnectToServer = True
    End Function
    
    Private Function WaitForResponse(ByVal Response As String) As Boolean
    
        ' if we're not waiting for a response then exit
        If Response = "" Then
            WaitForResponse = True
            Exit Function
        Else

            ' wait for the response property to change
            Do While LastResponse = ""
                DoEvents
            Loop
    
            ' if it matches, return true, else return false
            If Response = LastResponse Then
                ' return true
                WaitForResponse = True
                
            ' check for errors
            Else
                WaitForResponse = False
   End If
    
   End If
        
        ' clear the static variable for next time
        LastResponse = ""
    End Function

    Private Sub SMTPSend(ByVal Data As String)

        ' send the passed string with a vbCrLf
        If Sock.State = sckConnected Then
            Sock.SendData Data & vbCrLf
            DoEvents
   End If
        
        ' raise event
        RaiseEvent SentData(Data)
        
        ' log
        AppendLog Data & vbCrLf
    End Sub

    Public Sub AppendLog(ByVal Data As String)
        ' append the passed data to the log
        Me.Log = Me.Log & Data
    End Sub

' ------------------------------------------------------------------------------
'
'   INTERNAL SUBS
'
' ------------------------------------------------------------------------------
    Private Sub Sock_Connect()

        ' raise event
        RaiseEvent Connected(Sock.RemoteHost, Sock.RemotePort)
    End Sub
    
    Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    
        ' grab the data and set the last response
        Sock.GetData Data
        LastResponse = Mid$(Data, 1, 3)
            
        ' raise event
        RaiseEvent ReceivedData(Data)
        
        ' log
        AppendLog Data
    End Sub

    Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

        ' raise event
        RaiseEvent Error(Description)
    End Sub

    Private Sub UserControl_Resize()
    
        ' fit the control to the image
        With UserControl
            .Height = picImage.Height
            .Width = picImage.Width
        End With
    End Sub
Public Sub ccl()
 Sock.Close
 End Sub
