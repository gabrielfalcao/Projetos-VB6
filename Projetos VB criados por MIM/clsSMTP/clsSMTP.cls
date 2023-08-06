VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Design and create by Jeff, 2002
Option Explicit

' Winscok that will be use for connect
Private WithEvents sckMailSock As Winsock
Attribute sckMailSock.VB_VarHelpID = -1

Private blnRaiseError As Boolean
Private lngErrorNumber As Long
Private strErrDescription As String

Private strSMTPserver As String
Private lngPort As Long

Private strMailTo As String
Private strMailFrom As String
Private strSubject As String
Private strMessage As String

Private strDataCol As String

Private Const B_ErrTimeOut = 7780
Private Const S_ErrTimeOut = "Time Out or error response from server."
Private Const B_ErrUnableConnect = 7781
Private Const S_ErrUnableConnect = "Unable to connect to server."
Private Const B_ErrServerEnd = 7782
Private Const S_ErrServerEnd = "Connection close on server side."

'Private Const B_ErrNoWinsock = 7797
'Private Const S_ErrNoWinsock = "Parameter Sock do not contain any winsock."

Private Const TmOut = 3


' SMTP Status
Private Const smtpStatus = 211
Private Const smtpHelp = 214
Private Const smtpReady = 220
Private Const smtpClosing = 221
Private Const smtpDone = 250
Private Const smtpWillForward = 251
Private Const smtpStartMail = 354
Private Const smtpShuttingDown = 421
Private Const smtpMailboxUnavailable = 450
Private Const smtpLocalError = 451
Private Const smtpNoSpace = 452
Private Const smtpSyntaxError = 500
Private Const smtpArgError = 501
Private Const smtpNoCommand = 502
Private Const smtpBadSequence = 503
Private Const smtpNoParamater = 504
Private Const smtpMailboxUnavailable2 = 550
Private Const smtpUserRejected = 551
Private Const smtpTooBig = 552
Private Const smtpInvalidMailboxName = 553
Private Const smtpFailed = 554

Public Property Get RaiseError() As Boolean
    RaiseError = blnRaiseError
End Property

Public Property Let RaiseError(blnError As Boolean)
    blnRaiseError = blnError
End Property

Public Property Get ErrNumber() As Integer
    ErrNumber = lngErrorNumber
    ' Clear the error
    lngErrorNumber = 0
End Property

Public Property Get ErrDescription() As String
    ErrDescription = strErrDescription
End Property

Public Property Let Sock(WSock As Winsock)
    Set sckMailSock = WSock
End Property

Public Property Get SMTPserver() As String
    SMTPserver = strSMTPserver
End Property

Public Property Let SMTPserver(strSMTP As String)
    strSMTPserver = strSMTP
End Property

Public Property Get Port() As Long
    Port = lngPort
End Property

Public Property Let Port(lngP As Long)
    lngPort = lngP
End Property

Public Property Get MailFrom() As String
    MailFrom = strMailFrom
End Property

Public Property Let MailFrom(strMF As String)
    strMailFrom = strMF
End Property

Public Property Get MailTo() As String
    MailTo = strMailTo
End Property

Public Property Let MailTo(strMT As String)
    strMailTo = strMT
End Property

Public Property Get Subject() As String
    Subject = strSubject
End Property

Public Property Let Subject(strSub As String)
    strSubject = strSub
End Property

Public Property Get Message() As String
    Message = strMessage
End Property

Public Property Let Message(strMsg As String)
    strMessage = strMsg
End Property

Public Property Get Connected() As Boolean
    Connected = sckMailSock.State = sckConnected
End Property

Private Sub SendData(strData As String)
    ' Only send message when have connection.
    ' Igone error raise, let WaitResponse do it.
    If sckMailSock.State = sckConnected Then
        sckMailSock.SendData strData
    End If
End Sub

Private Function WaitResponse(Code1 As Integer, Optional Code2 As Integer) As Boolean
' A lot of error checking, hope it work
    Dim sglTime As Single
    Dim intCode As Integer
    Dim intPos As Integer
    Dim strServResponse As String
    
    sglTime = Timer
    
    ' Clear buffer
    strServResponse = ""
    
    ' Within time response
    While Timer < sglTime + TmOut
        DoEvents
        
        ' If server close connection
        If sckMailSock.State <> sckConnected Then
            WaitResponse = False
            sckMailSock.Close
            mRaiseError B_ErrServerEnd, S_ErrServerEnd
            Exit Function
        End If
        
        ' If a clear buffer
        If Len(strServResponse) = 0 Then
            ' Search for nearless response
            intPos = InStr(1, strDataCol, vbCrLf)
            
            ' Found nearless response
            If intPos Then
                ' Get nearless response
                strServResponse = Left$(strDataCol, intPos - 1)
                ' Clear current response
                strDataCol = Right$(strDataCol, Len(strDataCol) - intPos - 1)
            End If
            
        Else
            ' Verify code
            intCode = Val(Left$(strServResponse, 3))
            If intCode = Code1 Or intCode = Code2 Then
                ' Success
                WaitResponse = True
                Exit Function
            End If
            ' Clear buffer, wait for another message
            strServResponse = ""
        End If
    Wend
    
    ' Time Out or error response from server
    WaitResponse = False
    mRaiseError B_ErrTimeOut, S_ErrTimeOut
    
    ' Close the sock, terminal connection
    If sckMailSock.State <> sckClosed Then sckMailSock.Close
End Function

Public Sub Connect()
    Dim sglTime As Single
    
    If sckMailSock.State <> sckClosed Then sckMailSock.Close
    sckMailSock.RemoteHost = strSMTPserver
    sckMailSock.RemotePort = lngPort
    sckMailSock.Connect
    
    sglTime = Timer
    
    ' Wait for connect
    While (Timer < sglTime + TmOut) And sckMailSock.State <> sckConnected
        DoEvents: DoEvents: DoEvents
    Wend
    
    ' Time out, unable to connect
    If sckMailSock.State <> sckConnected Then
        ' Close
        sckMailSock.Close
        mRaiseError B_ErrUnableConnect, S_ErrUnableConnect
        Exit Sub
    End If
              
    ' Wait for ready response
    If Not WaitResponse(smtpReady) Then Exit Sub
    
    ' Said Helo to server
    SendData "helo " & sckMailSock.LocalIP & vbCrLf
    If Not WaitResponse(smtpDone) Then Exit Sub
    
End Sub

Public Sub Send()
    Dim colRecp
    Dim x As Integer
    
    ' Verify sender
    SendData "mail from:" & strMailFrom & vbCrLf
    If Not WaitResponse(smtpDone) Then Exit Sub
    
    ' Send Recepient
    colRecp = Split(strMailTo, ",")
    
    For x = 0 To UBound(colRecp)
        SendData "rcpt to:" & colRecp(x) & vbCrLf
        If Not WaitResponse(smtpDone, smtpWillForward) Then Exit Sub
    Next
    
    SendData "DATA" & vbCrLf
    If Not WaitResponse(smtpStartMail) Then Exit Sub
          
    ' Send hearder
    SendData "From: " & strMailFrom & vbCrLf
    SendData "Subject: " & strSubject & vbCrLf
    SendData "To: " & strMailTo & vbCrLf
    ' Send end hearder
    SendData vbCrLf
    
    ' Send message
    SendData strMessage & vbCrLf
    SendData vbCrLf & "." & vbCrLf
    
    If Not WaitResponse(smtpDone) Then Exit Sub
End Sub

Public Sub Quit()
    SendData "quit" & vbCrLf
    If sckMailSock.State = sckConnected Then sckMailSock.Close
End Sub

Private Sub mRaiseError(lngErr As Long, strDesc As String)
    ' Point error
    lngErrorNumber = lngErr
    strErrDescription = strDesc
            
    ' Raise error, use in dll
    If blnRaiseError = True Then Err.Raise lngErrorNumber, "clsBarcode39.Barcode", strErrDescription
End Sub

Private Sub Class_Initialize()
    lngPort = 25
End Sub

Private Sub sckMailSock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    ' Get the data
    sckMailSock.GetData strData
    
    ' Link to the mail collection
    strDataCol = strDataCol & strData
End Sub

