VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Basic init/release functions.

'InitializeEngine needs to create *both* a compressor and decompressor.  (Most 3rd-party libraries don't
' differentiate between these - but the built-in Windows compression libraries do.)
Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

End Function

Public Sub ReleaseEngine()

End Sub

'Actual compression/decompression functions.  Only arrays and pointers are standardized.  It's assumed
' that users can write simple wrappers for other data types, as necessary.
Public Function CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False) As Boolean

End Function

Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean

End Function

Public Function DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean

End Function

Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean

End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Public Function GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long

End Function

Public Function GetDefaultCompressionLevel() As Long

End Function

Public Function GetMinCompressionLevel() As Long

End Function

Public Function GetMaxCompressionLevel() As Long

End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Public Function GetCompressorName() As String

End Function

Public Function IsCompressorReady() As Boolean

End Function
