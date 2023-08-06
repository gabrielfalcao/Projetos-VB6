Attribute VB_Name = "modSelfExtract"
Public Sub NewExe(FileName, Data)
On Error Resume Next
 Dim ff As Integer
 Dim i As Long
 Dim ReadStream() As Byte
 Dim WriteStream() As Byte
 ff = FreeFile
 
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #ff
        ReDim ReadStream(LOF(ff) - 1)
        Get #ff, , ReadStream
    Close #ff
    
    ReDim WriteStream(UBound(ReadStream) + UBound(Data) + 12)
    
    For i = 0 To UBound(ReadStream)
        WriteStream(i) = ReadStream(i)
    Next
    
    WriteStream(UBound(ReadStream) + 1) = CByte(Asc("%"))
    WriteStream(UBound(ReadStream) + 2) = CByte(Asc("A"))
    WriteStream(UBound(ReadStream) + 3) = CByte(Asc("L"))
    WriteStream(UBound(ReadStream) + 4) = CByte(Asc("O"))
    WriteStream(UBound(ReadStream) + 5) = CByte(Asc("N"))
    WriteStream(UBound(ReadStream) + 6) = CByte(Asc("R"))
    WriteStream(UBound(ReadStream) + 7) = CByte(Asc("A"))
    WriteStream(UBound(ReadStream) + 8) = CByte(Asc("S"))
    WriteStream(UBound(ReadStream) + 9) = CByte(Asc("H"))
    WriteStream(UBound(ReadStream) + 10) = CByte(Asc("E"))
    WriteStream(UBound(ReadStream) + 11) = CByte(Asc("%"))
    
    For i = 0 To UBound(Data)
        WriteStream(i + UBound(ReadStream) + 12) = Data(i)
    Next
    
    Open FileName For Binary Access Write As #ff
        Put #ff, , WriteStream
    Close #ff
End Sub

Public Function LoadData(FileName)
On Error Resume Next
    Dim ff As Integer
    Dim ByteArray() As Byte
    
    ff = FreeFile
    Open FileName For Binary Access Read As #ff
        ReDim ByteArray(LOF(ff) - 1)
        Get #ff, , ByteArray
    Close #ff
    LoadData = ByteArray
End Function

Public Function SaveData(FileName) As Boolean
On Error Resume Next
 Dim ff As Integer
 Dim i As Long
 Dim start As Long
 Dim ReadStream() As Byte
 Dim WriteStream() As Byte
 ff = FreeFile
 
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #ff
        ReDim ReadStream(LOF(ff) - 1)
        Get #ff, , ReadStream
    Close #ff

    For i = 0 To UBound(ReadStream)
        If ReadStream(i) = CByte(Asc("%")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("A")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("L")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("O")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("N")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("R")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("A")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("S")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("H")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("E")) Then
        i = i + 1
        If ReadStream(i) = CByte(Asc("%")) Then
            start = i + 1
            SaveData = True
            Exit For
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
        End If
    Next
    If Not (SaveData) Then Exit Function
    SaveData = False
    
    ReDim WriteStream(UBound(ReadStream) - start)
    
    For i = 0 To UBound(WriteStream)
        WriteStream(i) = ReadStream(i + start)
    Next
    
    Open FileName For Binary Access Write As #ff
        Put #ff, , WriteStream
    Close #ff
    
    SaveData = True
End Function
