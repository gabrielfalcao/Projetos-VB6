Type=Exe
Form=Form1.frm
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\System32\stdole2.tlb#OLE Automation
Form=Form2.frm
IconForm="Form1"
Startup="Form1"
HelpFile=""
Title="Assistente de Duelo 1.0"
ExeName32="Assistente de Duelo.exe"
Path32="C:\"
Command32=""
Name="Assistente_de_Duelo"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionComments="Criado por Gabriel Falc�o"
VersionCompanyName="na"
VersionFileDescription="Criado por Gabriel Falc�o"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1

[MS Transaction Server]
AutoRefresh=1
                                                                                                                                                                        =   3
         Top             =   300
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duelista 1"
      Height          =   825
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   1770
      Begin VB.TextBox d1 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1485
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.d1.Caption = d1.Text
Form2.d2.Caption = d2.Text
Me.Hide
Unload Me
End Sub
