VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "ASP Generator"
   ClientHeight    =   3120
   ClientLeft      =   2070
   ClientTop       =   1950
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5865
   Begin VB.ListBox lstTables 
      Height          =   1620
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtProjectPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\ASPGen"
      Top             =   2160
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Database"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblTables 
      Caption         =   "Select a Table"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Project Folder Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Menu mnuTipo 
      Caption         =   "&Project Type"
      Begin VB.Menu mnuFull 
         Caption         =   "&Full"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabAdd 
         Caption         =   "&Add Register"
      End
      Begin VB.Menu mnuTabView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuTabUpd 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuTabDel 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Edit Files"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAutor 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code created by joel konecny as open source project 05/29/01
'and adapted by Ribamar FS ribafs@yahoo.com

'Original code in Planet-source-code:
'http://www.planet-source-code.com/xq/ASP/txtCodeId.6688/lngWId.4/qx/vb/scripts/ShowCode.htm

'     the engine i am providing is an asp code generator that i am currently using on my
'website (www.intratelligent.com) to allow asp developers to upload access databases
'online and build simple asp interfaces to interact with their data. (insert / update /
'delete and reporting functionality). although it is coded in vb, it can be compiled
'into a dll to be used as an object referenced by asp. otherwise it can be run as a
'stand-alone vb app. i have included a screen shot and added commenting to my code to
'make it more understandable. if you do not have a lot of experience in vb and asp this
' will be difficult for you to modify although it can be implemented by anyone.

'------------------------------------
'Thank your for your feedback of bugs and implementations.
'Ribamar FS ribafs@yahoo.com - http://ribafs.hp10.com.br


Option Explicit

Dim arrProperties(14) As String
Dim TableBGColor As String
Dim FirstHalfSQL As String
Dim SecondHalfSQL As String
Dim EndSQL As String
Dim strType As String
Dim TableName As String
Dim iCount As Integer

Public Function CreateProject(idbPath, iUserRootPath, iProjectName) As Boolean

    'this is the core engine. this loops through the database gathering required
    'information, builds initial directories and makes sure the database is valid.
    'there are three main for...each...next loops in this method that are of utmost
    'importance. the outer loops is the table loop. the code loops through each table
    'gathering required information for the table. the next loops is the index loop.
    'the code loops through the data searching for all indexes in a given table. the
    'final loop is the field loop. this loop reads each field in each table and from
    'there we have the base information to build our asp pages.
    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim objTable As DAO.TableDef
    Dim objField As DAO.Field
    Dim objProperty As DAO.Property
    Dim intProjectID As Integer
    Dim intTableID As Integer
    Dim intColumnID As Integer
    Dim UserName As String
    Dim iProjectRootPath As String
    Dim iProjectTablePath As String
    Dim fsoObject As New FileSystemObject
    Dim bDoesFolderExist As Boolean
    Dim xCount As Integer
    Dim iProjectID As String
    Dim CurrentPrimaryKey As String
    Dim objIndex As DAO.Index
    Dim objIndexField As DAO.Field
    Dim tfolders As Variant
    Dim driveRoot As String
    Dim zCount As Integer
    Dim CurrentPath As String
        
    'Test path database
    If CheckGoodDatabase(idbPath) = False Then
        CreateProject = False
        Exit Function
    End If
    
    'Make path of the new database: database.mdb
    tfolders = Split(iUserRootPath, "\", , vbTextCompare)
    driveRoot = tfolders(0)
    For xCount = 1 To UBound(tfolders)
        CurrentPath = driveRoot
        For zCount = 1 To xCount
            CurrentPath = CurrentPath & "\" & tfolders(zCount)
        Next zCount
        If Not fsoObject.FolderExists(CurrentPath) Then
            Call fsoObject.CreateFolder(CurrentPath)
        End If
    Next
    
    xCount = 1
    iProjectRootPath = iUserRootPath & "\Project" & xCount
'strPathEdit = iProjectRootPath
    If fsoObject.FolderExists(iProjectRootPath) Then
        bDoesFolderExist = True
        While bDoesFolderExist = True
            xCount = xCount + 1
            iProjectRootPath = iUserRootPath & "\Project" & xCount
            bDoesFolderExist = fsoObject.FolderExists(iProjectRootPath)
        Wend
        iProjectRootPath = iUserRootPath & "\Project" & xCount
    End If
    iProjectID = xCount
    
    Call AddProjectFolder(idbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath)
    
    Set objDatabase = objDAO.OpenDatabase(idbPath)
    
    If strType = "Full" Then
    
    For Each objTable In objDatabase.TableDefs
            If Mid(objTable.Name, 1, 4) <> "MSys" Then
            iProjectTablePath = iProjectRootPath & "\" & objTable.Name

            Call AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, objTable.Name)
            
            For Each objIndex In objTable.Indexes
                For Each objIndexField In objIndex.Fields
                    If objIndex.Primary = True Then
                        'primary key located for table
                        CurrentPrimaryKey = objIndexField.Name
                    End If
                Next
            Next
            
            For Each objField In objTable.Fields
                
                Erase arrProperties
                For Each objProperty In objField.Properties
                    
                    Select Case objProperty.Name
                        Case "Attributes"
                            arrProperties(1) = objProperty.Value
                        Case "Type"
                            arrProperties(2) = objProperty.Value
                            If objProperty.Value = 11 Then
                                objField.Properties("Required").Value = False
                                arrProperties(8) = False
                            End If
                        Case "OrdinalPosition"
                            arrProperties(3) = objProperty.Value
                        Case "Size"
                            arrProperties(4) = objProperty.Value
                        Case "DefaultValue"
                            arrProperties(5) = objProperty.Value
                        Case "ValidationRule"
                            arrProperties(6) = objProperty.Value
                        Case "ValidationText"
                            arrProperties(7) = objProperty.Value
                        Case "Required"
                            arrProperties(8) = objProperty.Value
                        Case "AllowZeroLength"
                            arrProperties(9) = objProperty.Value
                        Case "DecimalPlaces"
                            arrProperties(10) = objProperty.Value
                        Case "Format"
                            arrProperties(11) = objProperty.Value
                        Case "Name"
                            arrProperties(12) = objProperty.Value
                            If CurrentPrimaryKey = objProperty.Value Then
                                arrProperties(13) = True
                            Else
                                arrProperties(13) = False
                            End If
                        Case Else
                            
                    End Select
                Next
                
                Call CreateAddTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateAddTableASP(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateViewTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateGetTableData(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateDeleteTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateUpdateTableASP(iProjectRootPath, iProjectID, objTable.Name)
                
            Next
        End If
    Next
    End If

    If strType <> "Full" Then
    
        For Each objTable In objDatabase.TableDefs

        If objTable.Name = TableName Then

            iProjectTablePath = iProjectRootPath & "\" & objTable.Name
            
            Call AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, objTable.Name)
            
            For Each objIndex In objTable.Indexes
                For Each objIndexField In objIndex.Fields
                    If objIndex.Primary = True Then
                        'primary key located for table
                        CurrentPrimaryKey = objIndexField.Name
                    End If
                Next
            Next
            
            For Each objField In objTable.Fields
                
                Erase arrProperties
                For Each objProperty In objField.Properties
                    
                    Select Case objProperty.Name
                        Case "Attributes"
                            arrProperties(1) = objProperty.Value
                        Case "Type"
                            arrProperties(2) = objProperty.Value
                            If objProperty.Value = 11 Then
                                objField.Properties("Required").Value = False
                                arrProperties(8) = False
                            End If
                        Case "OrdinalPosition"
                            arrProperties(3) = objProperty.Value
                        Case "Size"
                            arrProperties(4) = objProperty.Value
                        Case "DefaultValue"
                            arrProperties(5) = objProperty.Value
                        Case "ValidationRule"
                            arrProperties(6) = objProperty.Value
                        Case "ValidationText"
                            arrProperties(7) = objProperty.Value
                        Case "Required"
                            arrProperties(8) = objProperty.Value
                        Case "AllowZeroLength"
                            arrProperties(9) = objProperty.Value
                        Case "DecimalPlaces"
                            arrProperties(10) = objProperty.Value
                        Case "Format"
                            arrProperties(11) = objProperty.Value
                        Case "Name"
                            arrProperties(12) = objProperty.Value
                            If CurrentPrimaryKey = objProperty.Value Then
                                arrProperties(13) = True
                            Else
                                arrProperties(13) = False
                            End If
                        Case Else
                            
                    End Select
                Next
                
If strType = "Full" Or strType = "TableAdd" Then
                Call CreateAddTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateAddTableASP(iProjectRootPath, iProjectID, objTable.Name)
ElseIf strType = "Full" Or strType = "TableView" Then
                Call CreateViewTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateGetTableData(iProjectRootPath, iProjectID, objTable.Name)
ElseIf strType = "Full" Or strType = "TableDel" Then
                Call CreateDeleteTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateViewTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateGetTableData(iProjectRootPath, iProjectID, objTable.Name)
ElseIf strType = "Full" Or strType = "TableUpd" Then
                Call CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateUpdateTableASP(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateViewTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateGetTableData(iProjectRootPath, iProjectID, objTable.Name)
End If

            Next
        End If
        Next
        
        End If
    
    CreateProject = True
    
    Set objDAO = Nothing
    Set objDatabase = Nothing
    Set objTable = Nothing
    Set objField = Nothing
    Set objProperty = Nothing
    
End Function

Private Function AddProjectFolder(ByRef idbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath) As Boolean

    'this method builds the project folder and initial default.asp page that will
    'show a listing of the tables in the project.
    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim CurrentLine As String
    Dim fsoTextStreamTable As TextStream
    Dim BGColor As String
    
    Call fsoObject.CreateFolder(iProjectRootPath)
    Call fsoObject.CopyFile(idbPath, iProjectRootPath & "\database.mdb")
    
    idbPath = iProjectRootPath & "\database.mdb"

    Set fsoTextStreamTable = fsoObject.CreateTextFile(iProjectRootPath & "\default.asp")
    
    Call fsoTextStreamTable.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamTable.WriteLine("<HTML>")
    Call fsoTextStreamTable.WriteLine("<HEAD>")
        Call fsoTextStreamTable.WriteLine("<TITLE>" & iProjectName & " table directory</TITLE>")
    Call fsoTextStreamTable.WriteLine("</HEAD>")
    Call fsoTextStreamTable.WriteLine("<BODY>")
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamTable.WriteLine("<tr>")
    Call fsoTextStreamTable.WriteLine("<td><font color=""#FFFFFF"">Table Directory</font></td>")
    Call fsoTextStreamTable.WriteLine("</tr>")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.WriteLine("<br>")
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""450"">")
    Call fsoTextStreamTable.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.WriteLine("<p><a href=""../"">Return to project list</a></p>")
    Call fsoTextStreamTable.WriteLine("</BODY>")
    Call fsoTextStreamTable.WriteLine("</HTML>")
    Call fsoTextStreamTable.Close
    
    Set fsoTextStreamTable = Nothing
    
End Function

Private Function AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, iTableName) As Boolean
        
    'for each table that exists i create it's own subdirectory folder. this folder
    'contains its own custom pages for insert / update / delete and select functionality.
    'some of these pages will contain tags like this... <!-- LinkStarter. Do not modyify -->.
    'all of the code you see in this method is standard code for each table, no matter what
    'the design for the given table. what will happen in later functions is that we will modify
    'the asp code to customize it for the table that requests it. as i read each field in the
    'CreateProject method i will modify these pages specifically for that field...
    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoTextStreamAddPage As TextStream
    Dim CurrentLine As String
    
    Call fsoObject.CreateFolder(iProjectRootPath & "\" & iTableName)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\default.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\defaultTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
        
            If TableBGColor = "#C0C0C0" Then
                TableBGColor = "#F3F3DC"
            Else
                TableBGColor = "#C0C0C0"
            End If
        
            Call fsoTextStreamTemp.WriteLine("<tr>")
            Call fsoTextStreamTemp.WriteLine("<td width=""200"" bgcolor=""" & TableBGColor & """>" & iTableName & "</td>")
            
            If strType = "Full" Or strType = "TableAdd" Then
                Call fsoTextStreamTemp.WriteLine("<td width=""200"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/Add.asp"">Add Register</a></center></td>")
            End If
            
            If strType = "Full" Or strType = "TableView" Or strType = "TableUpd" Or strType = "TableDel" Then
                Call fsoTextStreamTemp.WriteLine("<td width=""200"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/View.asp"">View, Edit or delete</a></center></td>")
            End If
            Call fsoTextStreamTemp.WriteLine("</tr>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\default.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\defaultTemp.asp", iProjectRootPath & "\default.asp")

If strType = "Full" Or strType = "TableAdd" Then
    'builds add.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\add.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<% response.buffer = true %>")     'I Add (Ribamar FS)
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Insert Data Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Add Information To Table " & iTableName & "</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<% if request.querystring(""fieldempty"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Please fill all required fields. Field "" & request.querystring(""fieldempty"") & "" was left empty."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""duplicatedata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Duplicate data entered in primary key field."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""invaliddata"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Invalid data entered in field "" & request.querystring(""invaliddata"") & ""."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""successful"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Record added successfully."")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""nodata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Unable to add data. No data submitted."")")
    Call fsoTextStreamAddPage.WriteLine("end if %>")
    Call fsoTextStreamAddPage.WriteLine("<p><font color=""#008000"">*</font>Primary Key/ ")
    Call fsoTextStreamAddPage.WriteLine("<font color=""#FF0000"">*</font>Required Field</p>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""POST"" action=""ASPAdd.asp"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Submit"" name=""B1"">   <input type=""reset"" value=""Limpar"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to table list</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    'builds ASPadd.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\ASPAdd.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<% response.buffer = true %>")     'I Add (Ribamar FS)
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Insert Data ASP Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<% Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 10")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iFieldCount = 0")
    Call fsoTextStreamAddPage.WriteLine("FirstHalfSQL = ""insert into [" & iTableName & "] (""")
    Call fsoTextStreamAddPage.WriteLine("SecondHalfSQL = "") Values (""")
    Call fsoTextStreamAddPage.WriteLine("EndSQL = "")""%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<% adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SQLInsert = FirstHalfSQL & SecondHalfSQL & EndSQL")
    Call fsoTextStreamAddPage.WriteLine("if SQLInsert <> ""insert into [" & iTableName & "] () Values ()"" then")
    Call fsoTextStreamAddPage.WriteLine("on error resume next")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLInsert)")
    Call fsoTextStreamAddPage.WriteLine("if err.number = -2147467259 then")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?duplicatedata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("on error goto 0")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?nodata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?successful=true"") %>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
End If
    
If strType = "Full" Or strType = "TableView" Or strType = "TableDel" Or strType = "TableUpd" Then
    
    'builds view.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\view.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<title>Select Data</title>")
    Call fsoTextStreamAddPage.WriteLine("</head>")
    Call fsoTextStreamAddPage.WriteLine("<body>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Select fields to retrieve from " & iTableName & ":</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""GET"" action=""getdata.asp"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Submit"" name=""B1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Limpar"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to table list</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</body>")
    Call fsoTextStreamAddPage.WriteLine("</html>")
    Call fsoTextStreamAddPage.Close
End If
    
    'builds getdata.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\getdata.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<%response.buffer = false%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Data</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<%Dim arrField()")
    Call fsoTextStreamAddPage.WriteLine("Dim totFieldCount")
    Call fsoTextStreamAddPage.WriteLine("If Request.QueryString(""NAV"") = """" Then")
    Call fsoTextStreamAddPage.WriteLine("intPage = 1")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("intPage = Request.QueryString(""NAV"")")
    Call fsoTextStreamAddPage.WriteLine("End If")
    Call fsoTextStreamAddPage.WriteLine("totFieldCount = 0%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<%if totfieldcount <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 11")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SqlSelect = ""select * from [" & iTableName & "]""")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CursorLocation = 3")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CursorType = 3")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLSelect)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.PageSize = 10")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CacheSize = adoRecordset.PageSize")
    Call fsoTextStreamAddPage.WriteLine("intPageCount = adoRecordset.PageCount")
    Call fsoTextStreamAddPage.WriteLine("intRecordCount = adoRecordset.RecordCount")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) <= 0 Then intPage = 1")
    Call fsoTextStreamAddPage.WriteLine("If intRecordCount > 0 Then")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.AbsolutePage = intPage")
    Call fsoTextStreamAddPage.WriteLine("intStart = adoRecordset.AbsolutePosition")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) = CInt(intPageCount) Then")
    Call fsoTextStreamAddPage.WriteLine("intFinish = intRecordCount")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("intFinish = intStart + (adoRecordset.PageSize - 1)")
    Call fsoTextStreamAddPage.WriteLine("End If%>")
    Call fsoTextStreamAddPage.WriteLine("<h4>Registros")
    Call fsoTextStreamAddPage.WriteLine("<%=intStart%> até <%=intFinish%> de <%=intRecordCount%>.</h4>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<%fieldcount = 0")
    Call fsoTextStreamAddPage.WriteLine("for each tempField in arrField %>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""#000080""><font color=""#FFFFFF""><%=tempField%></font>&nbsp;&nbsp;</td>")
    Call fsoTextStreamAddPage.WriteLine("<%fieldcount = fieldcount + 1")
    Call fsoTextStreamAddPage.WriteLine("next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<%for xcount = 1 to fieldcount%>")
    Call fsoTextStreamAddPage.WriteLine("<td>&nbsp;</td>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<%bcolor = ""#COCOCO""%>")
    Call fsoTextStreamAddPage.WriteLine("<%For intRecord = 1 To adoRecordset.PageSize%>")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<% qString = """" %>")
    Call fsoTextStreamAddPage.WriteLine("<% for each temparrfield in arrfield %>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<%=adorecordset(temparrfield)%></td>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    Call fsoTextStreamAddPage.WriteLine("<% for each tempField in adoRecordset.fields %>")
    Call fsoTextStreamAddPage.WriteLine("<% if not isnull(adorecordset(tempfield.name)) then")
    Call fsoTextStreamAddPage.WriteLine(" encodeField = server.urlencode(adorecordset(tempfield.name))")
    Call fsoTextStreamAddPage.WriteLine(" else")
    Call fsoTextStreamAddPage.WriteLine(" encodeField = """"")
    Call fsoTextStreamAddPage.WriteLine(" end if")
    Call fsoTextStreamAddPage.WriteLine(" tempFieldName = Replace(tempfield.name, "" "", """")")
    Call fsoTextStreamAddPage.WriteLine(" qString = qString & ""&"" & ""dat"" & tempFieldName & ""="" & encodeField%>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    
If strType = "Full" Or strType = "TableDel" Then
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<a href=""delete.asp?<%=request.querystring & qString%>"">delete</a></td>")
End If

If strType = "Full" Or strType = "TableUpd" Then
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<a href=""update.asp?<%=request.querystring & qString%>"">update</a></td>")
End If

    Call fsoTextStreamAddPage.WriteLine("<%adorecordset.MoveNext")
    Call fsoTextStreamAddPage.WriteLine("If bcolor = ""#COCOCO"" Then")
    Call fsoTextStreamAddPage.WriteLine("bcolor = ""#F3F3DC""")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("bcolor = ""#COCOCO""")
    Call fsoTextStreamAddPage.WriteLine("End If")
    Call fsoTextStreamAddPage.WriteLine("If adorecordset.EOF Then Exit For")
    Call fsoTextStreamAddPage.WriteLine("Next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<%else")
    Call fsoTextStreamAddPage.WriteLine("response.write(""<i>No Data Available</i>"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""view.asp?nodata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<%if intRecordCount > 0 then%>")
    Call fsoTextStreamAddPage.WriteLine("<%tempstring = request.querystring")
    Call fsoTextStreamAddPage.WriteLine("foundvalue = InStrRev(tempstring, ""&NAV="", Len(tempstring), vbTextCompare)")
    Call fsoTextStreamAddPage.WriteLine("if foundvalue <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("tempstring = Mid(tempstring, 1, foundvalue - 1)")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<%If CInt(intPage) > 1 Then%>")
    Call fsoTextStreamAddPage.WriteLine("<a href=""getdata.asp?<%=tempstring%>&NAV=<%=intPage - 1%>""><< Anterior</a>")
    Call fsoTextStreamAddPage.WriteLine("<%else%>")
    Call fsoTextStreamAddPage.WriteLine("<< Anterior")
    Call fsoTextStreamAddPage.WriteLine("<%End IF")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) < CInt(intPageCount) Then%>")
    Call fsoTextStreamAddPage.WriteLine("<a href=""getdata.asp?<%=tempstring%>&NAV=<%=intPage + 1%>"">Próximo >></a>")
    Call fsoTextStreamAddPage.WriteLine("<%else%>")
    Call fsoTextStreamAddPage.WriteLine("Próximo >>")
    Call fsoTextStreamAddPage.WriteLine("<%End If%>")
    Call fsoTextStreamAddPage.WriteLine("<%End If%>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""view.asp"">Return to selection page</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
If strType = "Full" Or strType = "TableDel" Then
    'builds delete.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\delete.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<title>Delete Data</title>")
    Call fsoTextStreamAddPage.WriteLine("</head>")
    Call fsoTextStreamAddPage.WriteLine("<body>")
    Call fsoTextStreamAddPage.WriteLine("<%totfieldcount = 0")
    Call fsoTextStreamAddPage.WriteLine("qString = """"%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<%if totfieldcount <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 10")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("if qString <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(qstring)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.open (""delete from [" & iTableName & "] where "" & qString)")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""getdata.asp?"" & mid(QueryCheckString,1,len(QueryCheckString)-1)) & ""&NAV="" & request.querystring(""NAV"")")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.write(""no data specified"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("<br><br>" & "Delete OK!" & "<br><br>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to Table list</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</body>")
    Call fsoTextStreamAddPage.WriteLine("</html>")
    Call fsoTextStreamAddPage.Close
End If

If strType = "Full" Or strType = "TableUpd" Then
    'builds update.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\update.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Update Data Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Update Information In " & iTableName & "</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<% if request.querystring(""fieldempty"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Please fill all required fields. Field "" & request.querystring(""fieldempty"") & "" was left empty."")")
    Call fsoTextStreamAddPage.WriteLine("request.querystring(""datHyperlink"")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""duplicatedata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Duplicate data entered in primary key field."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""invaliddata"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Invalid data entered in field "" & request.querystring(""invaliddata"") & ""."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""successful"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Record added successfully."")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""nodata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Unable to add data. No data submitted."")")
    Call fsoTextStreamAddPage.WriteLine("end if %>")
    Call fsoTextStreamAddPage.WriteLine("<p><font color=""#008000"">*</font>Primary Key / ")
    Call fsoTextStreamAddPage.WriteLine("<font color=""#FF0000"">*</font>Required Field</p>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""POST"" action=""ASPupdate.asp?<%=request.querystring%>"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Update"" name=""B1"">   <input type=""reset"" value=""Reset"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=request.querystring%>"" name=""UpdateQueryString"">")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=qstring%>"" name=""qString"">")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=mid(QueryCheckString,1,len(QueryCheckString)-1) & ""&NAV="" & request.querystring(""NAV"")%>"" name=""QueryCheckString"">")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""getdata.asp?<%=mid(QueryCheckString,1,len(QueryCheckString)-1) & ""&NAV="" & request.querystring(""NAV"")%>"">Voltar para a visualização de dados</p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    'builds ASPupdate.asp page for table
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\ASPupdate.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Update Data ASP Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<% Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 13")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("QueryCheckString = request.form(""QueryCheckString"")")
    Call fsoTextStreamAddPage.WriteLine("SecondHalfSQL = request.form(""qString"")")
    Call fsoTextStreamAddPage.WriteLine("iFieldCount = 0")
    Call fsoTextStreamAddPage.WriteLine("FirstHalfSQL = ""update [" & iTableName & "] set ""%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<% adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SQLInsert = FirstHalfSQL & "" where "" & SecondHalfSQL")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLInsert)")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""getdata.asp?"" & QueryCheckString)")
    Call fsoTextStreamAddPage.WriteLine("%>")
    Call fsoTextStreamAddPage.WriteLine("Update OK!" & "<br><br>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to Table lst</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
End If
    Set fsoTextStreamAddPage = Nothing
    
End Function

Private Function CreateAddTablePostHTML(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this is where i modify the add.asp page to customize a given table. this method
    'will be called once for each field in the table.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)

    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\add.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            If Required = "True" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
            If PrimaryKey = True Then
                StarPrimary = "<font color=""#008000"">*</font>"
            Else
                StarPrimary = ""
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""></textarea></p>")
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":  ")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No""></p>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\add.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", iProjectRootPath & "\" & iTableName & "\add.asp")
    
    
End Function

Private Function CreateAddTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method is where i customize the ASPadd.asp page to handle and insert into the
    'database anything that has been posted by the add.asp page.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\ASPadd.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" And Attrib <> 17 And DataType <> 11 Then
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            If Required = "True" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Response.Redirect (""add.asp?fieldempty=" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                        
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("On Error GoTo 0")

                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\ASPadd.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", iProjectRootPath & "\" & iTableName & "\ASPadd.asp")
    
End Function

Private Function CreateViewTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method builds a table with checkboxes for each field in a table. this page
    'will be used to allow the user to select which fields he wants displayed in
    'the report of the data. in the future this page will also display query builder
    'functionality.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\view.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            If DataType <> 11 Then
                fsoTextStreamTemp.WriteLine ("<p><input type=""checkbox"" name=""" & FieldName & """ value=""ON""> " & FieldName & "</p>")
            End If
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\view.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", iProjectRootPath & "\" & iTableName & "\view.asp")
    
End Function

Private Function CreateGetTableData(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method creates the pages that will dynamically build sql queries and display
    'the given data that has been posted to the getdata.asp page.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\getdata.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            If DataType <> 11 Then
                FieldNameVariable = Replace(FieldName, " ", "")
                FieldNameVariable = "var" & FieldNameVariable
                
                fsoTextStreamTemp.WriteLine ("<%")
                fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.querystring(""" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""ON"" then")
                fsoTextStreamTemp.WriteLine ("ReDim Preserve arrField(totFieldCount)")
                fsoTextStreamTemp.WriteLine ("arrField(totFieldCount) = """ & FieldName & """")
                fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                fsoTextStreamTemp.WriteLine ("End If")
                fsoTextStreamTemp.WriteLine ("%>")
            End If
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\getdata.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", iProjectRootPath & "\" & iTableName & "\getdata.asp")
    
End Function

Private Function CreateDeleteTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method creates a page that will delete a record from a table.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim QueryCheck As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\delete.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
                    
            QueryCheck = "Q" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            fsoTextStreamTemp.WriteLine (QueryCheck & " = request.querystring(""" & FieldName & """)")
            fsoTextStreamTemp.WriteLine ("If " & QueryCheck & " = ""ON"" then")
            fsoTextStreamTemp.WriteLine ("QueryCheckString = QueryCheckString & """ & FieldName & """ & ""=ON&""")
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.querystring(""dat" & FieldName & """)")
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
            fsoTextStreamTemp.WriteLine ("ReDim Preserve arrField(totFieldCount)")
            fsoTextStreamTemp.WriteLine ("arrField(totFieldCount) = """ & FieldName & """")
            fsoTextStreamTemp.WriteLine ("if " & FieldNameVariable & " <> """" then")
            Select Case DataType
                Case 10, 12        'Text, Memo, Hyperlink
                    
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 3, 4, 2, 6, 7, 15, 20, 5, 1       'Number, Currency, YesNo
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = "" & " & FieldNameVariable & "")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 8             'Date
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")

                Case 11            'OLEOBject
                                   'is not supported
            End Select
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine ("%>")

        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\delete.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", iProjectRootPath & "\" & iTableName & "\delete.asp")
    
End Function

Private Function CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method builds the asp page that will disply text boxes / textareas / checkboxes
    'that will display current data that is in a given record. this data can be changed
    'by the end user and the new data will be posted to the ASPupdate.asp page where the
    'specified record will be updated in the database.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    Dim valueFieldName As String
    Dim QueryCheck As String
    Dim FieldNameVariable As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\update.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            valueFieldName = Replace(FieldName, " ", "")
            valueFieldName = "dat" & valueFieldName
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            QueryCheck = "Q" & valueFieldName
            
            If Required = "True" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
            If PrimaryKey = True Then
                StarPrimary = "<font color=""#008000"">*</font>"
            Else
                StarPrimary = ""
            End If
             
            fsoTextStreamTemp.WriteLine ("<%")
            fsoTextStreamTemp.WriteLine (QueryCheck & " = request.querystring(""" & FieldName & """)")
            fsoTextStreamTemp.WriteLine ("If " & QueryCheck & " = ""ON"" then")
            fsoTextStreamTemp.WriteLine ("QueryCheckString = QueryCheckString & """ & FieldName & """ & ""=ON&""")
            fsoTextStreamTemp.WriteLine ("end if")
            fsoTextStreamTemp.WriteLine ("%>")
            
            Select Case DataType
                Case 10, 12        'Text, Memo, Hyperlink
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = '"" & request.querystring(""" & valueFieldName & """) & ""'""")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 3, 4, 2, 6, 7, 15, 20, 5, 1       'Number, Currency, YesNo
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = "" & request.querystring(""" & valueFieldName & """)")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 8             'Date
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = cdate('"" & request.querystring(""" & valueFieldName & """) & ""')""")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            Select Case DataType
                Case 10            '     Text
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""><%=request.querystring(""" & valueFieldName & """)%></textarea></p>")
                        Case 32770 'hyperlink
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<%if request.querystring(""" & valueFieldName & """) <> """" then%>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=mid(request.querystring(""" & valueFieldName & """),1,(len(request.querystring(""" & valueFieldName & """))/2)-1)%>""></p>")
                            Call fsoTextStreamTemp.WriteLine("<%else%>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                            Call fsoTextStreamTemp.WriteLine("<%end if%>")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                    End Select
                Case 8             '  DateTime
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 5             '  Currency
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 1             '     YesNo
                    
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":  ")
                    Call fsoTextStreamTemp.WriteLine("<%if request.querystring(""" & valueFieldName & """) <> ""False"" then%>")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No""></p>")
                    Call fsoTextStreamTemp.WriteLine("<%else%>")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No"" checked></p>")
                    Call fsoTextStreamTemp.WriteLine("<%end if%>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\update.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", iProjectRootPath & "\" & iTableName & "\update.asp")
    
    
End Function

Private Function CreateUpdateTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method builds the ASPupdate.asp page which is used to process the data that
    'has been sent to it by the update.asp page and update a given record in the database.
    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\ASPupdate.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" And Attrib <> 17 And DataType <> 11 Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            If Required = "True" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Response.Redirect (""update.asp?"" & request.querystring & ""&fieldempty=" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                     
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\ASPupdate.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", iProjectRootPath & "\" & iTableName & "\ASPupdate.asp")
    
End Function

Private Function CheckGoodDatabase(idbPath) As Boolean

    'function checks to see if database is valid by
    'making a connection to it and then disconnecting
    
    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim DBPassword As String
    Dim wrkJet As DAO.Workspace
    
    On Error GoTo errorhandler:
    
    Set objDatabase = objDAO.OpenDatabase(idbPath)
    
    CheckGoodDatabase = True
    
    Exit Function
    
errorhandler:
    CheckGoodDatabase = False
    
End Function

Private Sub Command1_Click()
    On Error GoTo errorhandler

    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim objTable As DAO.TableDef
    Dim matTable(99) As String
            
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    Command1.Enabled = False
    DoEvents
            
    Set objDatabase = objDAO.OpenDatabase(CommonDialog1.FileName)
    
    If strType <> "Full" Then
        lstTables.Visible = True: lblTables.Visible = True
        lstTables.Clear
        For Each objTable In objDatabase.TableDefs
            If Mid(objTable.Name, 1, 4) <> "MSys" Then
                iCount = iCount + 1
                matTable(iCount) = objTable.Name
                lstTables.AddItem (matTable(iCount))
            End If
        Next
        Exit Sub
    End If
    If CreateProject(CommonDialog1.FileName, txtProjectPath, "Test") = True Then
        MsgBox ("Project created successfully")
    Else
        MsgBox ("Unable to create project")
    End If
    Command1.Enabled = True

errorhandler:

End Sub

Private Sub lstTables_Click()
    TableName = lstTables.Text
    If CreateProject(CommonDialog1.FileName, txtProjectPath, "Test") = True Then
        MsgBox ("Project created successfully")
    Else
        MsgBox ("Unable to create project")
    End If
    Command1.Enabled = True
End Sub

Private Sub mnuAutor_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuFull_Click()
    strType = "Full"
    Command1.Enabled = True
    Call Command1_Click
End Sub

Private Sub mnuEditar_Click()
    Load frmEditor
    frmEditor.Show
End Sub

Private Sub mnuHelp_Click()
    Load Help
    Help.Show
End Sub

Private Sub mnuTabAdd_Click()
    strType = "TableAdd"
    Command1.Enabled = True
    Call Command1_Click
End Sub

Private Sub mnuTabUpd_Click()
    strType = "TableUpd"
    Command1.Enabled = True
    Call Command1_Click
End Sub

Private Sub mnuTabDel_Click()
    strType = "TableDel"
    Command1.Enabled = True
    Call Command1_Click
End Sub

Private Sub mnuTabView_Click()
    strType = "TableView"
    Command1.Enabled = True
    Call Command1_Click
End Sub

