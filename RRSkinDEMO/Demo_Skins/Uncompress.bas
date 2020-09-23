Attribute VB_Name = "Module2"


' Uncompress byte array
'-------------------------------
'

'Public Declare Function compress Lib "zlib.dll" _
'   (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Declare Function uncompress Lib "zlib.dll" _
  (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

' zlib responses
Public Const Z_OK = 0
Public Const Z_DATA_ERROR = -3
Public Const Z_MEM_ERROR = -4
Public Const Z_BUF_ERROR = -5


