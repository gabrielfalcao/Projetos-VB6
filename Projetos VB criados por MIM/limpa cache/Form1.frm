VERSION 5.00
Begin VB.Form frmCleanCache 
   BackColor       =   &H00F4C686&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tera - Cache Cleaner"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFAEA&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00754709&
      Height          =   2955
      ItemData        =   "Form1.frx":030A
      Left            =   75
      List            =   "Form1.frx":0311
      TabIndex        =   2
      Top             =   300
      Width           =   2745
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2925
      TabIndex        =   1
      Top             =   1395
      Width           =   1440
   End
   Begin VB.CommandButton command1 
      Caption         =   "Listar Arquivos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2925
      TabIndex        =   0
      Top             =   1005
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00A4E3AC&
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos de Cache Existentes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00754709&
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   60
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3435
      Picture         =   "Form1.frx":031C
      Top             =   300
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4C686&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00754709&
      Height          =   945
      Left            =   2850
      TabIndex        =   3
      Top             =   1695
      Width           =   1800
   End
End
Attribute VB_Name = "frmCleanCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

  'load the cached file list
   GetCacheURLList
   Label2.Caption = List2.ListCount & " Arquivos Listados."

End Sub

Private Sub Command3_Click()

   Dim cachefile As String
   Dim i As Long
     
  'delete all files except..
   For i = 0 To List2.ListCount - 1
   
      cachefile = List2.List(i)
      
     '..if the file is a cookie, don't screw
     'up saved passwords, so skip it
      If InStr(cachefile, "Cookie") = 0 Then

         Call DeleteUrlCacheEntry(cachefile)

      End If
   
   Next
   
  'reload the list
   GetCacheURLList
   
End Sub


Private Sub frmCleanCache_Click()

End Sub

Private Sub Form_Load()

  'load the cached file list
   GetCacheURLList
   Label2.Caption = List2.ListCount & " Arquivos Listados."

End Sub

Public Sub GetCacheURLList()
    
   Dim ICEI As INTERNET_CACHE_ENTRY_INFO
   Dim hFile As Long
   Dim cachefile As String
   Dim posUrl As Long
   Dim posEnd As Long
   Dim dwBuffer As Long
   Dim pntrICE As Long
   
   List2.Clear
   
  'Like other APIs, calling FindFirstUrlCacheEntry or
  'FindNextUrlCacheEntry with an insufficient buffer will
  'cause the API to fail, and the buffer pointing to the
  'correct size required for a successful call.
   dwBuffer = 0

  'Call to determine the required buffer size
   hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)
   
  'both conditions hould be met by the first call
   If (hFile = ERROR_CACHE_FIND_FAIL) And _
      (err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
   
     'The INTERNET_CACHE_ENTRY_INFO data type is a
     'variable-length type. It is neccessary to allocate
     'memnory for the result of the call and pass the
     'pointer to this memory location to the API.
      pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
        
     'allocation successful
      If pntrICE Then
         
        'set a Long pointer to the memory location
         CopyMemory ByVal pntrICE, dwBuffer, 4
         
        'and call the first find API again passing the
        'pointer to the allocated memory
         hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
       
        'hfile should = 1 (success)
         If hFile <> ERROR_CACHE_FIND_FAIL Then
         
           'loop through the cache
            Do
            
              'the pointer has ben filled, so move the
              'data back into a ICEI structure
               CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
            
              'CacheEntryType is a long representing
              'the type of entry returned
               If (ICEI.CacheEntryType And _
                   NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
               
                 'extract the string from the memory location
                 'pointed to by the lpszSourceUrlName member
                 'and add to a list
                  cachefile = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                  List2.AddItem cachefile

               End If
               
              'free the pointer and memory associated
              'with the last-retrieved file
               Call LocalFree(pntrICE)
               
              'and again repeat the procedure, this time calling
              'FindNextUrlCacheEntry with a buffer size set to 0.
              'This will cause the call to once again fail,
              'returning the required size as dwBuffer
               dwBuffer = 0
               Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
               
              'allocate and assign the memory to the pointer
               pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
               CopyMemory ByVal pntrICE, dwBuffer, 4
               
           'and call again with the valid parameters.
           'If the call fails (no more data), the loop exits.
           'If the call is successful, the Do portion of the
           'loop is executed again, extracting the data from
           'the returned type
            Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
  
         End If 'hFile
         
      End If 'pntrICE
   
   End If 'hFile
   
  'clean up by closing the find handle, as
  'well as calling LocalFree again to be safe
   Call LocalFree(pntrICE)
   Call FindCloseUrlCache(hFile)
   
End Sub


Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function
