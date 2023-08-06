VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmPreview 
   Caption         =   "Visualizar Impressão"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CME\os.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   585
      Visible         =   0   'False
      Width           =   1140
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6120
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   8670
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: An introduction to using the Report Designer.  Some
'          of the basic methods used in manipulating and creating
'          reports are discussed in the code comments

Option Explicit

Dim m_Report As New ORDM    ' Create a new instance of the Crystal Report
'Dim m_RS As New ADOR.Recordset        ' Create and ADO record set

Private Sub Command1_Click()
Unload Me
OS.Show
End Sub

' *************************************************************
' You can also capture all kinds of events from the Smart Viewer
'
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'    MsgBox "The Smart Viewer fires 27 different events!"
End Sub

Private Sub Form_Load()
'Dim pasta As String
'If Len(App.Path) = 3 Then
'pasta = App.Path
'Else
'pasta = App.Path & "\"
'End If
    ' Open the recordset
  '  m_RS.Open "Select * from customer where country = 'USA'", "xtreme sample database"
CRViewer1.Zoom (100)
Me.WindowState = 2
    ' Pass this recordset to the report engine to use as the datasource
    ' If you comment out the following line the report will show all customer
    ' in the world because the report was created to show all customers and
    ' this information is stored in the DSR.
    m_Report.Database.SetDataSource Data1

    ' The Smart Viewer (CRVIEWER) is the preview window for the report
    ' If you are don't wish to show the report to the user, ie. only want
    ' to print the report, then you don't need to use the viewer at all.
    CRViewer1.ReportSource = m_Report

    ' You have full access to all the report objects, so you may want
    ' to change the text in the title of the report
    ' Comment out the line below to see the text that is in the report
    ' definition.
   ' m_Report.Text6.SetText "You can change text objects through code!"

    '  The Smart viewer has a object model that allows you to modify the
    '  look and feel of the preview window at runtime.
    '  The ViewReport method will start the process of the report.
    CRViewer1.ViewReport
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err


Dim mgs As Integer
mgs = MsgBox("Deseja salvar voltar ao programa?", vbYesNo + vbQuestion, Me.Caption)
Select Case mgs
Case vbYes
OS.Show
Unload Me
  Case Else
End Select
Err:
If Err.Number <> 0 Then
MsgBox "Não foi possível salvar, verifique a integridade do programa, do banco de dados, ou contate o distribuidor do programa!", vbCritical, "Erro de Gravação"
End If
Exit Sub
End Sub

' *************************************************************
' This code resizes the Smart Viewer when the form is resized
'
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
