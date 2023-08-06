VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHexCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrPassword As String

Public Property Get Password() As String
   Password = mstrPassword
End Property
Public Property Let Password(ByVal strPassword As String)
   mstrPassword = strPassword
End Property

Public Function Decrypt(InputBuffer As String) As String
    Dim X As Long
    Dim sHexedAscii As String
    Dim iAscii As Integer
    Dim iLoopCount As Long
    Dim sDeHexed As String
    Form1.Picture1.Cls
    Form1.Picture1.Print "Desencriptando..."
    sDeHexed = Space(Len(InputBuffer) / 2)
    For X = 1 To Len(InputBuffer) Step 2
        iLoopCount = iLoopCount + 1
        sHexedAscii = Mid$(InputBuffer, X, 2)
        iAscii = Val("&H" & sHexedAscii)
        Mid$(sDeHexed, iLoopCount, 1) = Chr$(iAscii)
    Next
    Decrypt = DecryptText(sDeHexed)
    Form1.Picture1.Cls
End Function

Private Function EncryptText(strText As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

    If Len(mstrPassword) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(mstrPassword, (i Mod Len(mstrPassword)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

Private Function DecryptText(strText As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

    If Len(mstrPassword) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(mstrPassword, (i Mod Len(mstrPassword)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function


Public Function Encrypt(InputBuffer As String) As String
    Dim sEncrypted As String
    Dim sHexed As String
    Dim sHexEntry As String
    Dim X As Long
    Dim iAscii As Integer
    Form1.Picture1.Cls
    Form1.Picture1.Print "Encriptando..."
    sEncrypted = EncryptText(InputBuffer)
    For X = 1 To Len(sEncrypted)
        iAscii = Asc(Mid$(sEncrypted, X, 1))
        sHexEntry = Hex$(iAscii)
        If Len(sHexEntry) = 1 Then sHexEntry = "0" + sHexEntry
        sHexed = sHexed & sHexEntry
    Next
    Encrypt = sHexed
    Form1.Picture1.Cls
End Function

Private Sub Class_Initialize()
   Password = "pwd001"
End Sub
