VERSION 5.00
Begin VB.Form frmProperties 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   1920
   ClientTop       =   2505
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Menu mnuProperties 
      Caption         =   "Properties"
      Begin VB.Menu mnuProperties1 
         Caption         =   "Drive Properties"
      End
      Begin VB.Menu mnuProperties2 
         Caption         =   "Folder Properties"
      End
      Begin VB.Menu mnuProperties3 
         Caption         =   "File Properties"
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuProperties1_Click()
Dim R As Long
   R = ShowFileProperties(frmFileSystem.Text4.Text, Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If R <= 32 Then MsgBox "Error"
  
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub mnuProperties2_Click()
Dim R As Long
   R = ShowFileProperties(frmFileSystem.Text1.Text, Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If R <= 32 Then MsgBox "Error"
  
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub mnuProperties3_Click()
Dim R As Long
   R = ShowFileProperties(frmFileSystem.Text5.Text, Me.hwnd) 'To show the properties dialog pass the filename and the owner of the dialog
    If R <= 32 Then MsgBox "Error"
  
    MousePointer = vbDefault
    Exit Sub
End Sub
