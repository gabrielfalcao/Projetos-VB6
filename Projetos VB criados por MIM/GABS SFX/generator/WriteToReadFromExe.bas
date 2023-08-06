Attribute VB_Name = "WriteToReadFromExe"
Option Explicit

'Declarations

Type VAREXE                                     'Variabletype used to store a variablename
                                                'and the corresponding Data.
    VarName As String
    VarData As String

End Type

Public Vars(200) As VAREXE                      'Array used to store Variablenames and Values for our new exe file
                                                'If you need more than 200 Variables you'll have to modify this line
Public Vars2(200) As VAREXE                     'Stores variablenames and data read from an exe file

Public NumberOfVarsStored As Integer            'how many variables are stored in our exe file ? (needed when reading out data from an exe file)

                                      
'#############################################################################
'
' Function: AddVariable
'
' Usage:    Add new Variables + correponding Data
'           / Modify Data of existing Variables
'
' Syntax:   AddVariable "VariableName", "VariableData"
'
' Example:  AddVariable "WindowTitle", "Example by Over. overkillpage@gmx.net"
'
'#############################################################################

Public Sub AddVariable(VarName As String, VarData As String)
    
    'Declarations
    Dim Looop As Integer                        'Variable used for any kind of loops
    Dim VarDoesAlreadyExist As Boolean          'Do we only have to change the VarData or do we have to add
                                                'a completly new Variable to the Vars Array ???
                                        
    Do                                          'In this Do-Loop we check if the Var does already exist
    
        If Vars(Looop).VarName = VarName Then   'if true it does exist !
            VarDoesAlreadyExist = True
            Exit Do
        End If
        
        Looop = Looop + 1                       'Increasing our loop variable
        
    Loop Until Vars(Looop).VarName = ""
    
    If VarDoesAlreadyExist = True Then
        
        Vars(Looop).VarData = VarData           'We only have to modify an existing Variable
    
    Else
        
        Vars(Looop).VarName = VarName           'We add a new Var + Data
        Vars(Looop).VarData = VarData
        
    End If
    
End Sub


'#############################################################################
'
' Function: WriteExeFile
'
' Usage:    Stores Variables + Data collected by AddVariable into a new exe
'           file DURING RUNTIME !
'
' Syntax:   WriteExeFile "FileNameOfAnTemplateExeFile", "NameOfNewExeFile"
'
' Example:  WriteExeFile "c:\temp.exe", "c:\new.exe"
'
'#############################################################################

Public Function WriteExeFile(TemplateExeFile As String, NameOfNewExeFile As String) As Boolean

    'Declarations
    Dim Looop As Integer                        'Variable used for any kind of loops
    
    Dim Sep1 As String * 12                     'We're gonna insert all data (vars + data) which was "collected"
    Dim Sep2 As String * 12                     'using AddVar at the end of our template file. To do so we will
    Dim Sep3 As String * 12                     'Generate one single string. Special seperators help us to split
                                                'this string into the original data when the new exe file wants
                                                'to read out the stored data.
                                                'ok guys ;) i think nobody understand this part. My English is
                                                'too bad and it's much too late.
    
    Sep1 = "|-|-sep1-|-|"                       'Ok. If you want to you can of course change these seperator strings
    Sep2 = "_-|-sep2_-|-"                       'But be carefull. Chosing strings like "a" is quite dangerous.
    Sep3 = "=-=_sep3-_=-"                       'The hole new created Exe file is searched for this seperatorstrings
                                                'and splitted accordingly. So you have to chose strings that do
                                                'NOT exist in the Temp Exe File !!
                                                

    If Dir(TemplateExeFile) = "" Then           'Checking if Template File exists, if not exiting the function
                                                'and returning an error msg
        MsgBox "Template file doesn't exist !", vbExclamation, "ERROR"
        WriteExeFile = False
        Exit Function
    
    End If

    Open TemplateExeFile For Binary Access Read As #1     'Opening the Template File.
    Open NameOfNewExeFile For Binary Access Write As #2   'opening a file to "create" a new exe file
        
                                                'Now we're gonna create the string which will be inserted at the
                                                'end of the Template Exe File
        Dim OutString As String
                                            
        OutString = Sep1                        'Inserting opening seperator
                                            
        Do                                      'Looping through all stored Variables + Data
                                
            If Vars(Looop).VarName <> "" Then   'If there is an var + data it will be added to outstring
                   
                   OutString = OutString & Vars(Looop).VarName & Sep3 & Vars(Looop).VarData & Sep2
                   
            End If
            
            Looop = Looop + 1                   'Increasing our loop variable
            
        Loop Until Vars(Looop).VarName = ""
            
        OutString = OutString & Sep1 & LOF(1)   'There are two parts in our string. one containing the vars +data
                                                'and one containing the original filelen of the tempfile. This helps
                                                'us later to read out the created string again. (When it was inserted
                                                'into the template file). Ok. The data + vars have been added so there
                                                'is another seperator followed by the filesize.
        
                                                'Finally we'll have to "create" the new exe file....
        Dim TempFile As String
        TempFile = Space$(LOF(1))
        Get #1, , TempFile                      '1. Reading out the content of the Templatefile
        
        Put #2, , TempFile & OutString          '2. Writing the Templatefile + Outstring into a new exe file !
    
    Close #1                                    'tiddy up ;)
    Close #2
    
    WriteExeFile = True                         'Everything went fine so we return True
    
End Function


'#############################################################################
'
' Function: ReadVariable
'
' Usage:    Read DURING RUNTIME a variable stored in the project's very own
'           exe file
'
' Syntax:   ReadVariable "VariableName"
'
' Example:  ReadVariable "WindowTitle"
'
'#############################################################################

Public Function ReadVariable(VarName As String) As String

    'Declarations
    Dim Looop As Integer                        'Variable used for any kind of loops
    Dim SizeOfTemplateFile                      'Will store the size of the Template File
    Dim VarString As String                     'Will hold the string of variable names and data
    
    Dim Sep1 As String * 12                     'We are using the same seperator strings we used in WriteExeFile
    Dim Sep2 As String * 12                     'Now we using to: 1. Get the OutString we added to our exe file
    Dim Sep3 As String * 12                     '2. split it into the single vars + data
                                                                                                
    Sep1 = "|-|-sep1-|-|"
    Sep2 = "_-|-sep2_-|-"
    Sep3 = "=-=_sep3-_=-"
    
    If Vars2(0).VarName = "" Then              'If NO variables are stored in the Vars2 Array we check THIS exe file
                                                'for stored Vars + Data ! Else we only return the stored Data.
    
        
        Open App.EXEName & ".exe" For Binary Access Read As #3     'Funny ;) isn't it ? But this is THE trick of this sourcecode ;)
                                                            'we open the exe file of this program !! and read out data from
                                                            'the end of the file !
                                                            'For German users ;): "Der Moment wo der Elefant das Wasser l‰ﬂt ;)"
        Dim EXEString As String
        EXEString = Space$(LOF(3))
        Get #3, , EXEString                                 'Reading out the content of the exe file of THIS Project
        
        Looop = LOF(3) - Len(Sep3)                          'We start reading out the content beginning at the end of the file
        
        Do                                                  'We stored the len of the template file at the end, before this there
                                                            'is a sep1. So we search it :)
           Looop = Looop - 1
           Dim egal
        
        Loop Until Mid(EXEString, Looop, Len(Sep3)) = Sep1
            
        SizeOfTemplateFile = Mid(EXEString, Looop + Len(Sep1))      'Read out the filelen of the templatefile
        VarString = Mid(EXEString, SizeOfTemplateFile + 1)          'Now we read out the Var + Data String
        VarString = Left(VarString, Looop - SizeOfTemplateFile - 1) 'Cutting of end seperator
        VarString = Mid(VarString, Len(Sep1) + 1)                   'We cut of the initialition seperator
        VarString = Left(VarString, Len(VarString) - Len(Sep2))     'cutting of the sep2 at the end
    
        For Looop = 1 To Len(VarString)                             'Count number of Variables stored in VarString
            
            If Mid(VarString, Looop, Len(Sep2)) = Sep2 Then NumberOfVarsStored = NumberOfVarsStored + 1
        
        Next Looop
        
        For Looop = 0 To NumberOfVarsStored                         'storing vars + data in the vars2 array
                
            Vars2(Looop).VarName = Split((Split(VarString, Sep2)(Looop)), Sep3)(0)
            Vars2(Looop).VarData = Split((Split(VarString, Sep2)(Looop)), Sep3)(1)
        
        Next Looop
            
        Close #3
        
    
    
    End If
    
    'Reading out data for the requested variablename
    For Looop = 0 To NumberOfVarsStored                         'storing vars + data in the vars2 array
            
        If UCase(Vars2(Looop).VarName) = UCase(VarName) Then    'We use ucase to make the check not case sensitive
            ReadVariable = Vars2(Looop).VarData                 'Returning requested data
            Exit For
        End If
    
    Next Looop
    
End Function

