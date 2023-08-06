Attribute VB_Name = "basSoundRec_2"
Option Explicit

Public Const CALLBACK_FUNCTION = &H30000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
Public Const MMIO_CREATE = &H1000      '  create new file (or truncate file)
Public Const MMIO_CREATERIFF = &H20    '  mmioCreateChunk(): make a LIST chunk
Public Const MMIO_WRITE = &H1         '  open file for writing only

Public Const MMIOERR_BASE = 256
Public Const MMIOERR_CANNOTEXPAND = (MMIOERR_BASE + 8)  '  cannot expand file
Public Const MMIOERR_CANNOTREAD = (MMIOERR_BASE + 5)  '  cannot read
Public Const MMIOERR_CANNOTWRITE = (MMIOERR_BASE + 6) '  cannot write
Public Const MMIOERR_OUTOFMEMORY = (MMIOERR_BASE + 2)  '  out of memory
Public Const MMIOERR_UNBUFFERED = (MMIOERR_BASE + 10) '  file is unbuffered

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type


Public Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Public Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Public Declare Function mmioWrite Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Public Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Public Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Public Declare Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long

' variables for managing wave file
Public mmckinfoParentIn As MMCKINFO
Public mmckinfoSubchunkIn As MMCKINFO
Public mmioinf As mmioinfo
Public hmmioIn As Long
