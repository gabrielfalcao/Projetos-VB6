VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      Height          =   330
      Left            =   3495
      TabIndex        =   1
      Top             =   2835
      Width           =   1140
   End
   Begin MSMask.MaskEdBox maskcpfcgc 
      Height          =   450
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   794
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function calculaCGC(numero As String) As String
Dim I As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(numero) Then
calculaCGC = ""
Exit Function
End If
mult = 2
For I = Len(numero) To 1 Step -1
prod = prod + Val(Mid(numero, I, 1)) * mult
mult = IIf(digito = 10 Or digito = 11, 0, digito)

calculaCGC = Trim(Str(digito))
Next
End Function


Public Function validaCGC(CGC As String) As Boolean
If calculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
validaCGC = False
Exit Function
End If
validaCGC = True
End Function
Function calculacpf(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
On Error GoTo err_CPF
Dim I As Integer
Dim strcampo As String
Dim strcaracter As String
Dim intnumero As Integer
Dim intmais As Integer
Dim ingsoma As Long
Dim dbldivisao As Double
Dim inginstiro As Long
Dim intresto As Integer
Dim intdig1 As Integer
Dim intdig2 As Integer
Dim strconf As String

ingsoma = 0
intnumero = 0
intmais = 0
strcampo = Left(CPF, 9)
'Inicia cálculos do 1º dígito
For I = 2 To 10
strcaracter = Right(strcampo, I - 1)
intnumero = Left(strcaracter, 1)
intmais = intnumero * I
ingsoma = ingsoma + intmais
Next I
dbldivisão = ingsoma - inginteiro
If intresto = 0 Or intresto = 1 Then
intdig1 = 0
Else
intdig1 = 11 - intresto
End If

strcampo = strcampo & intdig1 'concatena o cpf com o primeiro digito verificador
ingsoma = 0
intnumero = 0
intmais = 0
'inicia cálculos do 2º dígito
For I = 2 To 11
strcaracter = Right(strcampo, I - 1)
intnumero = Left(strcaracter, 1)
intmais = intnumero * I
ingsoma = ingsoma + intmais
Next I
dbldivisao = ingsoma / 11
inginteiro = Int(dbldivisao) * 11
intresto = ingsoma - inginteiro
If intresto = 0 Or intresto = 1 Then
intdig2 = 0
Else
intdig2 = 11 - intresto
End If
strconf = intdig1 & intdig2
'caso o cpf esteja errado dispara a mensagem
If strconf <> Right(CPF, 2) Then
calculacpf = False
Else
calculacpf = True
End If
Exit Function
exit_CPF:
Exit Function
err_CPF:
MsgBox Error$
Resume exit_CPF
End Function

Private Sub Command1_Click()
End
End Sub

Private Sub maskcpfcgc_GotFocus()
maskcpfcgc.Mask = "##############"
End Sub

Private Sub maskcpfcgc_KeyPress(KeyAscii As Integer)
'se teclar ENTER envia um TAB
If KeyAscii = 13 Then
SendKeys "{tab}"
KeyAscii = 0
End If
End Sub

Private Sub maskcpfcgc_LostFocus()
If Len(maskcpfcgc.Text) > 0 Then
Select Case Len(maskcpfcgc.Text)
Case Is = 11
maskcpfcgc.Mask = "###.###.###-##"
If Not calculacpf(maskcpfcgc.Text) Then
MsgBox "CPF com DV incorreto !!!"
maskcpfcgc = ""
maskcpfcgc.Mask = "##############"
maskcpfcgc.SetFocus
End If
Case Is = 14
maskcpfcgc.Mask = "##.###.###/####-##"
If Not validaCGC(maskcpfcgc.Text) Then
MsgBox "CGC com DV incorreto !!!"
maskcpfcgc = ""
maskcpfcgc.Mask = "##############"
maskcpfcgc.SetFocus
End If
End Select
End If
End Sub
