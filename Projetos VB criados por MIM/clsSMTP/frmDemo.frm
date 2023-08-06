VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDemo 
   Caption         =   "Demo for clsSMTP"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtSMTPserver 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Write down your SMTP server. "
         Top             =   960
         Width           =   4695
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtMailFrom 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "Type in your POP3 email account. Must be an exist account"
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox txtMailTo 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "Type in receiver email address. Can be any email address."
         Top             =   3240
         Width           =   4695
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "Your subject."
         Top             =   3600
         Width           =   4695
      End
      Begin VB.TextBox txtMessage 
         Height          =   1095
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmDemo.frx":0000
         Top             =   3960
         Width           =   4695
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   5880
         Width           =   1575
      End
      Begin MSWinsockLib.Winsock wsSMTP 
         Left            =   480
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   6600
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6480
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "SMTP server:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "1) Fill in you SMTP server and then click [Start] button. It will connect to your SMTP server and say halo. >o< .."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label3 
         Caption         =   $"frmDemo.frx":0018
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Mail From:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Mail To:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Message:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "3) Finish sending, click quit to disconnect from server."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   5760
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       For people who have been use Outlook Express, you
'       may skip this.
'
' Note: This demo need a SMTP and POP3 account. SMTP stand
'       for Simple Mail Transaction Protocol and POP3 stand
'       for Post Office Protocol. (Don't know
'       spell correct, I am still a student and will go to
'       work for another two month)
'
'       POP3 is for incoming mail and SMTP is for outgoing
'       mail. For this demo, it just use for sending mail,
'       not for receivering mail, so we will only use SMTP
'       protocol.
'
'       To success use this demo. First we must have a POP3
'       mail account. POP3 mail account usually is not a
'       free mail account or web mail(It is not like a free
'       web mail account, like Hotmail, Yahoo mail account.
'       Some ISP may provide a POP3 account for their
'       customer, so you may check with your ISP for
'       further informatiom).
'
'       If you confuse or never use POP3 mail, then start
'       search Internet for research, you will know what
'       is all about it. (Type in POP3 mail and search for
'       it in any search engine, then it will come a lot
'       of information, try to read some of it. It may seen
'       not related to what we want, but it will give you
'       a good understanding.) >,<  .. Yes, if you find
'       a free PO3 mail, please inform me. Thank you.
'
'       NOTE, I am not an advance programer so some may be
'       mistake, if you find some mistake, please tell me.
'       ahsooon@hotmail.com with title "Hi"
'
'       YES, the arrange of code is follow sequence, so
'       you can read it from top to down.

Option Explicit
Dim smtp As clsSMTP

Private Sub Form_Load()
    ' Create a new class
    Set smtp = New clsSMTP
End Sub

Private Sub cmdStart_Click()
    ' Point the winsock to the class, we need a winsock
    ' to send and receive data
    smtp.Sock = wsSMTP
    
    ' Set the SMTP server address
    smtp.SMTPserver = txtSMTPserver.Text
    
    ' Now we connect to SMTP server
    smtp.Connect
    
    ' What, we got an error?
    If smtp.ErrNumber <> 0 Then
        ' Tell user the error we get
        MsgBox smtp.ErrDescription, vbOKOnly + vbCritical, "Error"
        
        ' No thing to do, just end it.
        Exit Sub
    End If
    
    ' No error, connect successful
    ' Enable another two button for send email
    cmdSend.Enabled = True
    cmdQuit.Enabled = True
    
    cmdStart.Enabled = False
    
End Sub

Private Sub cmdSend_Click()
    ' Set all the field, however, I have not include
    ' error handing for valid email format. Please do
    ' it on your own.
    smtp.MailFrom = txtMailFrom.Text
    smtp.MailTo = txtMailTo.Text
    
    ' Set the subject and message.
    smtp.Subject = txtSubject.Text
    smtp.Message = txtMessage.Text
    
    ' Send the mail.
    smtp.Send
    
    ' What, cause a problem?
    ' I also don't know how to fix it!!!!!!
    ' Most problem is invalid email address. That is,
    ' not exist POP3 email account in use for mail from.
    
    If smtp.ErrNumber <> 0 Then
        ' Tell user the error we get
        MsgBox smtp.ErrDescription, vbOKOnly + vbCritical, "Error"
        
        ' No thing to do, just end it.
        Exit Sub
    End If
    
    ' The email has send successful
    MsgBox "The email has been send successfully.", vbOKOnly + vbInformation, "Gong Xi Fa Chai"
End Sub

Private Sub cmdQuit_Click()
    ' Disconnect it from server
    smtp.Quit
    
    ' Reset the button
    cmdSend.Enabled = False
    cmdQuit.Enabled = False
    
    cmdStart.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Remove memory
    Set smtp = Nothing
End Sub
