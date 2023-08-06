VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDL 
   Caption         =   "Downloader"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar P 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Download to..."
   End
   Begin VB.CommandButton cmdDL 
      Caption         =   "Download"
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtAd 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblDLFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "Download From:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblSave 
      BackStyle       =   0  'Transparent
      Caption         =   "Saving to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specify a file"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "frmDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Here is my HTTP File Downloader

Dim File As String
Dim Ad As String
Dim X As Integer
Dim I As Integer
Dim File2 As String
Dim Rec As String
Dim Num As Integer
Dim FSize As String
Dim Temp2 As String
Dim Str As String
Dim Range As String
Dim F As Integer

Private Sub cmdBrowse_Click()
    CD.Filter = "All Files (*.*)|*.*"   'sets the file types
    CD.ShowSave 'shows the save dialog
    txtS.Text = CD.FileName
End Sub

Private Sub cmdDL_Click()
    Ad = txtAd.Text 'here is the long messy and confusing process of
    'seperating the address, file path, and file name
    'make sure ur only using /'s or else this wont work
    If InStr(1, Ad, "http://") Then 'remover "http://" if it is present
        Ad = Right(Ad, Len(Ad) - 7)
        MsgBox Ad
    End If
    
    Do Until X = Len(Ad)    'scans for the first /
        DoEvents
        X = X + 1
        If Mid(Ad, X, 1) = "/" Then
            File = Mid(Ad, X, Len(Ad))  'gets the end...the file path
            Ad = Mid(Ad, 1, X - 1)  'gets the address
            MsgBox File 'shows message boxes to ensure it is correct, there are a few of these
            MsgBox Ad
            Exit Do 'stops loop
        End If
    Loop
    File2 = File
    If InStr(2, File2, "/") Then    'this will go to the final / to get just the file name
        Do Until InStr(2, File2, "/") = False
            Do
                DoEvents
                I = I + 1
                If Mid(File2, I, 1) = "/" Then  'when it finds a /
                    File2 = Mid(File2, I, Len(File2))   'file2 will contain from the current / until the end..
                    MsgBox File2
                    Exit Do
                End If
            Loop
        Loop
    End If
    
    CD.Filter = "All Files (*.*)|*.*"
    CD.FileName = Right(File2, Len(File2) - 1)
    CD.ShowSave
    File2 = CD.FileName
    
    F = FreeFile()  'this makes file access easier
    Open File2 For Binary As #F 'it will create if it doesnt exist
    Close #F

    lblS.Caption = File2
    
    WS.Close    'closes winsock
    WS.Connect Ad, 80   'connects to the address on port 80
    lblStat.Caption = "Connecting to: " & Ad
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WS.Close    'closes winsock before u exit the app
End Sub

Private Sub WS_Connect()
Dim Header As String
    lblStat.Caption = "Connected, requesting " & File
    
    If FileLen(File2) > 0 Then  'if the length of the file is greater than 0 bytes
        Range = FileLen(File2)  'it will remember the length of the file for later
        MsgBox "Resuming..."    'lets the user kno its resuming a download
        Else
        Range = 0   'if there is nothing in the file
        Rec = 0
    End If
    
    'prepares the header
    Header = Header & "GET " & File & " HTTP/1.1" & vbCrLf
    Header = Header & "Host: " & WS.RemoteHostIP & vbCrLf
    Header = Header & "Range: bytes=" & Range & "-" & vbCrLf    'the range value comes in use here
    Header = Header & "User-Agent: Nullific Downloader\1.0" & vbCrLf
    Header = Header & "Accept: */*" & vbCrLf
    
    WS.SendData Header & vbCrLf 'sends the header
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    WS.GetData Data 'gets the data that is sent

    Num = 0
    
If InStr(1, Data, "HTTP") Then  'processes the servers response for the file size
    Str = Data
    Do
        'Num = Num + 1
        If Data = "" Then
            MsgBox "Error, Invalid Request" 'sumthin went wrong
        End If
        Data = Right(Data, Len(Data) - 1)   'front of header
        If Mid(Data, 1, 15) = "Content-Length:" Then    'when the front is...
            MsgBox Data
                Do  'for the file size...
                    Num = Num + 1
                    If Mid(Data, Num, 2) = vbCrLf Then  'finds the vbcrlf, telling us that the line with the file size has ended
                        Temp2 = Mid(Data, 1, Num)   'isolates the line with the size
                        MsgBox Temp2
                        FSize = Mid(Temp2, 16, Len(Temp2))  'removes "Content-Length: " and leaves only the file size
                        MsgBox FSize

                        P.Max = FSize
                        Exit Do
                    End If
                Loop
            Exit Do
        End If
    Loop
    
    Num = 0
    
    Do
        Num = Num + 1
            If Mid(Str, Num, 4) = (vbCrLf & vbCrLf) Then    'at the end of the header may be the beginning of the file, seperated by two vbcrlfs
                Str = Mid(Str, Num + 4, Len(Str))   'when they are found
                MsgBox "VBCRLF X2"
                P.Value = P.Value + Len(Str)
                Rec = Len(Str)
                lblDL.Caption = Rec & "/" & FSize
                
                F = FreeFile()
                Open File2 For Binary As #F 'writes to the file
                    If Range = 0 Then   'if there is nothing in the file
                        Put #F, , Str
                        Else
                        Put #F, Range, Str
                    End If
                Close #F
             Exit Do
             End If
        Loop
    Else
    
    F = FreeFile()
    Open File2 For Binary As #F
        P.Value = Rec
        Rec = Int(Rec) + Len(Data)  'adds to how many bytes have been recieved
        lblDL.Caption = Rec & "/" & FSize
        Temp = (LOF(F) + 1)
        If Temp = 0 Then
            Put #F, , Data
            Else
            Put #F, Temp, Data
        End If
    Close #F
    
    If P.Value = P.Max Then
        MsgBox "Completed"
    End If
End If

End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Error: " & Number & ", " & Description & "."    'if there is an error
End Sub

'I'll clean this up and update it later
'please give me feed back and suggestions
'i'm already in the process of resumable downloads
'too lazy to devise a way to show the kbps
