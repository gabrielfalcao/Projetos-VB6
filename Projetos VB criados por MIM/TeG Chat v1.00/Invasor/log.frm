VERSION 5.00
Begin VB.Form log 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log"
   ClientHeight    =   2370
   ClientLeft      =   165
   ClientTop       =   7500
   ClientWidth     =   7950
   ControlBox      =   0   'False
   Icon            =   "log.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7950
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2295
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   7875
   End
End
Attribute VB_Name = "log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub log_Change()
log.SelStart = Len(log.Text)
End Sub
' Example code for the CCaptionlessWindowMover class
'
' To try this example, do the following:
' 1. Create a new form
' 2. Paste all the code from this example to the new form's module
' 4. Run the form, and try moving the form around by
'    clicking anywhere on the body of the form and dragging

' In the Declarations section of the form declare the variable


