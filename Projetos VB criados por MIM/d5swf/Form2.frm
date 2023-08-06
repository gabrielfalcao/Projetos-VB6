VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Incúria Screensaver"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form2"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
If App.PrevInstance = True Then End
Dim p_ret As String
Dim ass As String
If Len(App.Path) = 3 Then
ass = App.Path
Else
ass = App.Path & "\"
End If
p_ret = StrConv(LoadResData("ocx", "swf"), vbUnicode)
Open ass & "flash.ocx" For Binary As #1
Put #1, , p_ret
Close #1
p_ret = StrConv(LoadResData("SP", "SWF"), vbUnicode)
Open "c:\sp.swf" For Binary As #2
Put #2, , p_ret
Close #2
Form1.Show
Unload Me
End Sub

