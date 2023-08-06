Option Explicit

'***Abra o 'References dialog box', e então selecione o 'Microsoft DAO 3.51 Object Library'***
'***Execute a Procedure MakeDataBase***

Dim P_W As Workspace
Dim P_DB As Database
Dim P_T As TableDef
Dim P_F As Field
Dim P_I As Index
Dim P_RS As Recordset
Dim P_R As Relation
Dim P_Q As QueryDef
Dim P_P As Parameter

Public Sub MakeDataBase()
  Screen.MousePointer = vbHourglass
  If Dir(App.Path & "Cópia de os.mdb\Cópia de os.mdb") = "" Then MakeDB_Cópia_de_os
  Screen.MousePointer = vbDefault
End Sub

Private Sub MakeDB_Cópia_de_os_Tables()

  Rem ---Tabela---

  Set P_T = P_DB.CreateTableDef("os")

    Rem ---Campos = 13---

    Set P_F = P_T.CreateField("Cliente")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 0
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Endereço")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 1
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Nº")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 2
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Complemento")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 3
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Bairro")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 4
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Telefone")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 5
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Horário")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 6
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Data")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 7
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Contato")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 8
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Áreas de aplicação")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 9
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Serviço(s)")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = False
      .DefaultValue = ""
      .OrdinalPosition = 10
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Nº da Ordem")

    With P_F
      .Type = dbLong
      .Required = False
      .DefaultValue = ""
      .OrdinalPosition = 11
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbFixedField or dbAutoIncrField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

    Set P_F = P_T.CreateField("Preço")

    With P_F
      .Type = dbText
      .Size = 50
      .Required = False
      .AllowZeroLength = True
      .DefaultValue = ""
      .OrdinalPosition = 12
      .ValidationText = ""
      .ValidationRule = ""
      .Attributes = dbVariableField or dbUpdatableField
    End With

    P_T.Fields.Append P_F

  P_DB.TableDefs.Append P_T

    Rem ---Indices = 1---

    Set P_I = P_T.CreateIndex("PrimaryKey")

    Set P_F = P_I.CreateField("Nº da Ordem")
    P_I.Fields.Append P_F

    P_I.Primary = True      'Foreign = False (Read-Only)
    P_I.Clustered = False
    P_I.Required = True
    P_I.Unique = True
    P_I.IgnoreNulls = False

    P_T.Indexes.Append P_I
End Sub

Private Sub MakeDB_Cópia_de_os_Relations()
  Rem ---Relations do Arquivo = 0---

End Sub

Private Sub MakeDB_Cópia_de_os_Querys()
  Dim P_S as String

  Rem ---QueryDefs = 1---

  P_S = ""
  P_S = P_S & "SELECT DISTINCTROW * FROM os;"
  Set P_Q = P_DB.CreateQueryDef("~sq_fOrdem de Serviço", P_S)
  P_Q.ReturnsRecords = True
  'P_Q.Type = 3
  P_Q.Close

  Set P_Q = Nothing
End Sub

Public Sub MakeDB_Cópia_de_os()
  Rem ---Criar Arquivo 'Cópia de os.mdb'---

  Set P_W = CreateWorkspace("", "admin", "", dbUseJet)
  Set P_DB = P_W.CreateDatabase(App.Path & "Cópia de os.mdb\Cópia de os.mdb", dbLangGeneral)

  MakeDB_Cópia_de_os_Tables
  MakeDB_Cópia_de_os_Relations
  MakeDB_Cópia_de_os_Querys

  Rem ---Fechar Arquivo 'Cópia de os.mdb'---
  P_DB.Close
  P_W.Close
  Set P_DB = Nothing
  Set P_W = Nothing
End Sub 

