VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   1035
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir imagem para wallpaper"
      Filter          =   "Arquivos de Imagem|*.gif;*.jpg;*.bmp;*.jpeg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   945
      Left            =   840
      TabIndex        =   0
      Top             =   1815
      Width           =   2610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lngSuccess As Long
Dim strBitmapImage As String
cd.ShowOpen
strBitmapImage = cd.FileName
lngSuccess = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, strBitmapImage, 0)

End Sub
