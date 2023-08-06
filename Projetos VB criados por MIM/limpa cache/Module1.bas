Attribute VB_Name = "modCleanCache"
Option Explicit
Public Const ERROR_CACHE_FIND_FAIL As Long = 0
Public Const ERROR_CACHE_FIND_SUCCESS As Long = 1
Public Const ERROR_FILE_NOT_FOUND As Long = 2
Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Public Const MAX_PATH As Long = 260
Public Const MAX_CACHE_ENTRY_INFO_SIZE As Long = 4096

Public Const LMEM_FIXED As Long = &H0
Public Const LMEM_ZEROINIT As Long = &H40
Public Const LPTR As Long = (LMEM_FIXED Or LMEM_ZEROINIT)

Public Const NORMAL_CACHE_ENTRY As Long = &H1
Public Const EDITED_CACHE_ENTRY As Long = &H8
Public Const TRACK_OFFLINE_CACHE_ENTRY As Long = &H10
Public Const TRACK_ONLINE_CACHE_ENTRY As Long = &H20
Public Const STICKY_CACHE_ENTRY As Long = &H40
Public Const SPARSE_CACHE_ENTRY As Long = &H10000
Public Const COOKIE_CACHE_ENTRY As Long = &H100000
Public Const URLHISTORY_CACHE_ENTRY As Long = &H200000
Public Const URLCACHE_FIND_DEFAULT_FILTER As Long = NORMAL_CACHE_ENTRY Or _
                                                    COOKIE_CACHE_ENTRY Or _
                                                    URLHISTORY_CACHE_ENTRY Or _
                                                    TRACK_OFFLINE_CACHE_ENTRY Or _
                                                    TRACK_ONLINE_CACHE_ENTRY Or _
                                                    STICKY_CACHE_ENTRY
Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type INTERNET_CACHE_ENTRY_INFO
   dwStructSize As Long
   lpszSourceUrlName As Long
   lpszLocalFileName As Long
   CacheEntryType  As Long
   dwUseCount As Long
   dwHitRate As Long
   dwSizeLow As Long
   dwSizeHigh As Long
   LastModifiedTime As FILETIME
   ExpireTime As FILETIME
   LastAccessTime As FILETIME
   LastSyncTime As FILETIME
   lpHeaderInfo As Long
   dwHeaderInfoSize As Long
   lpszFileExtension As Long
   dwExemptDelta  As Long
End Type

Public Declare Function FindFirstUrlCacheEntry Lib "wininet" _
   Alias "FindFirstUrlCacheEntryA" _
  (ByVal lpszUrlSearchPattern As String, _
   lpFirstCacheEntryInfo As Any, _
   lpdwFirstCacheEntryInfoBufferSize As Long) As Long

Public Declare Function FindNextUrlCacheEntry Lib "wininet" _
   Alias "FindNextUrlCacheEntryA" _
  (ByVal hEnumHandle As Long, _
   lpNextCacheEntryInfo As Any, _
   lpdwNextCacheEntryInfoBufferSize As Long) As Long

Public Declare Function FindCloseUrlCache Lib "wininet" _
   (ByVal hEnumHandle As Long) As Long

Public Declare Function DeleteUrlCacheEntry Lib "wininet" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

Public Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
  
Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long




