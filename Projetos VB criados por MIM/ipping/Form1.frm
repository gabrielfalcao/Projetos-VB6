VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstip 
      Height          =   1425
      Left            =   885
      TabIndex        =   2
      Top             =   1095
      Width           =   3615
   End
   Begin MSMask.MaskEdBox txtip 
      Height          =   840
      Index           =   0
      Left            =   885
      TabIndex        =   1
      Top             =   2895
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   1482
      _Version        =   393216
      MaxLength       =   11
      Mask            =   "999.999.999"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdping 
      Caption         =   "PING!"
      Height          =   405
      Left            =   1965
      TabIndex        =   0
      Top             =   465
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' Título: Ping de uma faixa de IP de 1 a 255
' Por: Joao Ricardo Ziglio do Amaral
' URL: http://www.coders.com.br/vb/codigo.asp?codigo=73&id=1
'**************************************************************
 
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Private Const SOCKET_ERROR = 0

Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type



Private Sub cmdPing_Click()
Dim hFile As Long, lpWSAdata As WSAdata
Dim hHostent As Hostent, AddrList As Long
Dim Address As Long, rIP As String
Dim OptInfo As IP_OPTION_INFORMATION
Dim EchoReply As IP_ECHO_REPLY
Dim numero As Integer
Dim parte As String
Dim total As String
Dim final As String
Dim espaco As Integer
For numero = 1 To 255
    final = ""
    total = Str(numero)
    For espaco = 1 To Len(total)
    parte = Mid(total, espaco, 1)
    If parte <> " " Then
     final = final + parte
    End If
    Next
    HostName = txtip(0).Text + final
    Call WSAStartup(&H101, lpWSAdata)
    
    If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    
    hFile = IcmpCreateFile()
    
    If hFile = 0 Then
        'MsgBox "Não foi possível criar o arquivo ICMP"
        Exit Sub
    End If
    
    OptInfo.TTL = 255
    
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
      ' MsgBox "Timeout"
    End If
    
    If EchoReply.Status = 0 Then
        'MsgBox "A Resposta de " + HostName + " (" + rIP + ") foi recebida após " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
        lstip.AddItem HostName + " (" + rIP + ") foi recebida após " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
        DoEvents
    Else
       ' MsgBox "Falhou ..."
    End If
    
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
      Next
       txtip(0).SetFocus
End Sub

