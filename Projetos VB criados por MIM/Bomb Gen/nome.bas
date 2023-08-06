Attribute VB_Name = "nome"
Dim nome As String
Public Function iniciar_programa()
If Len(App.Path) = 3 Then
nome = App.Path & "som.wav"
Else
nome = App.Path & "\som.wav"
End If
End Function
Public Function Fechar_Programa()
Kill nome
End Function
