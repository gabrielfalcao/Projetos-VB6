VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} ORDM 
   ClientHeight    =   6885
   ClientLeft      =   1365
   ClientTop       =   540
   ClientWidth     =   11295
   OleObjectBlob   =   "OS.dsx":0000
End
Attribute VB_Name = "ORDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
' Purpose: An introduction to using the Report Designer.  Some
'          of the basic methods used in manipulating reports
'          are discussed in the code comments
'
'          This report uses the 'xtreme sample data' as it's data source, ie.
'          the CrystalReport1.dsr remembers that information.  You can also
'          use the SetDataSource method on the Database object to change the
'          data source at runtime.  For example you can create a DAO/RDO/ADO
'          record set at runtime and pass it to the report engine at runtime
'          and the data in the report will reflect the data from the record set.
'
'          Obviously, the report field names and types must match the names
'          and types in the record set.
'
'          Take a look at the Form Load event in the Preview form to see how
'          to use this method.
'

Option Explicit

' *************************************************************
' You can do condition formating on report objects using VB code!  Crystal
' Reports has allowed this in the past with our own formula language but
' now you can do it in VB code.
' This event is fired each time this section is formatted, so you can get the
' values of the field object and do all kinds of cool formatting things.
' Here is a small example that simply changes the format of the field.
'
Private Sub Section6_Format(ByVal pFormattingInfo As Object)
    ' Change Field9's color to red if the value is less than 50000
    'If Field9.Value < 50000 Then
    '    Field9.TextColor = vbRed
    'Else
    '    Field9.TextColor = vbGreen
    'End If
End Sub

