VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressLz4HC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Lz4_HC Compression Library Interface
'Copyright 2016-2017 by Tanner Helland
'Created: 04/December/16
'Last updated: 07/September/17
'Last update: rewrite as a ICompress implementation
'
'Per its documentation (available at https://github.com/lz4/lz4), lz4 is...
'
' "...a lossless compression algorithm, providing compression speed at 400 MB/s per core, scalable with
'  multi-cores CPU. It features an extremely fast decoder, with speed in multiple GB/s per core, typically
'  reaching RAM speed limits on multi-core systems."
'
'lz4 is BSD-licensed and written by Yann Collet, the same genius behind the zstd compression library.  As of
' Dec 2016, development is very active and performance numbers rank among the best available for open-source
' compression libraries.  As PD writes a ton of huge files, improved compression performance is a big win
' for us, particularly on old systems with 5400 RPM HDDs.
'
'lz4-hc support is also provided.  lz4-hc is a high-compression variant of lz4.  It is much slower
' (6-10x depending on workload), but provides compression levels close to zlib.  Decompression speed is
' identical to regular lz4, so it is a good fit for things like run-time resources, where you have ample
' time available during compression stages, but you still want decompression to be as fast as possible
' (e.g. "compress once, decompress many").
'
'As of v7.0, most internal PD temp files and caches are written using Lz4, so this library sees heavy usage
' during a typical session.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Implements ICompress


'This constant was originally declared in lz4.c.  Note that lz4 does *not* support variable compression levels.
' Instead, it supports variable *acceleration* levels.  The difference is that bigger values = worse compression.
Private Const LZ4_MIN_ALEVEL As Long = 1
Private Const LZ4_DEFAULT_ALEVEL As Long = 1

'This value is not declared by the lz4 library, and technically, there is no maximum value.  Compression just
' approaches 0% as you increase the level.  I provide a "magic number" cap simply so it supports the same
' default/min/max functions as the other libraries
Private Const LZ4_MAX_ALEVEL As Long = 100

'These constants were originally declared in lz4_hc.h
Private Const LZ4HC_MIN_CLEVEL As Long = 3
Private Const LZ4HC_DEFAULT_CLEVEL As Long = 9
Private Const LZ4HC_MAX_CLEVEL As Long = 12

Private Declare Function LZ4_versionNumber Lib "liblz4" Alias "_LZ4_versionNumber@0" () As Long
Private Declare Function LZ4_compress_fast Lib "liblz4" Alias "_LZ4_compress_fast@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cAccelerationLevel As Long) As Long
Private Declare Function LZ4_compress_HC Lib "liblz4" Alias "_LZ4_compress_HC@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cCompressionLevel As Long) As Long
Private Declare Function LZ4_decompress_safe Lib "liblz4" Alias "_LZ4_decompress_safe@16" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long) As Long
Private Declare Function LZ4_compressBound Lib "liblz4" Alias "_LZ4_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array

'A single lz4 handle is maintained for the life of a PD instance; see InitializeLz4 and ReleaseLz4, below.
Private m_Lz4Handle As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Manually load the DLL
    Dim lz4Path As String
    lz4Path = pathToDLLFolder & "liblz4.dll"
    m_Lz4Handle = LoadLibraryW(StrPtr(lz4Path))
    ICompress_InitializeEngine = (m_Lz4Handle <> 0)
    
    'If we initialized the library successfully, cache some lz4-specific data
    If (Not ICompress_InitializeEngine) Then
        Debug.Print "WARNING!  LoadLibraryW failed to load lz4.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & lz4Path & ")"
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    If (m_Lz4Handle <> 0) Then
        FreeLibrary m_Lz4Handle
        m_Lz4Handle = 0
    End If
End Sub

'Actual compression/decompression functions.  Only arrays and pointers are standardized.  It's assumed
' that users can write simple wrappers for other data types, as necessary.
Private Function ICompress_CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False) As Boolean

    ValidateCompressionLevel compressionLevel
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = ICompress_GetWorstCaseSize(constSrcSizeInBytes)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
    'Perform the compression
    Dim finalSize As Long
    finalSize = LZ4_compress_HC(constSrcPtr, VarPtr(dstArray(0)), constSrcSizeInBytes, dstCompressedSizeInBytes, compressionLevel)
    
    dstCompressedSizeInBytes = finalSize
    ICompress_CompressPtrToDstArray = (finalSize <> 0)
    If (Not ICompress_CompressPtrToDstArray) Then InternalError "lz4_compress failed", finalSize
    
    'Trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    ValidateCompressionLevel compressionLevel
    
    Dim finalSize As Long
    finalSize = LZ4_compress_HC(constSrcPtr, constDstPtr, constSrcSizeInBytes, dstSizeInBytes, compressionLevel)
    
    'Check for error returns
    ICompress_CompressPtrToPtr = (finalSize <> 0)
    If ICompress_CompressPtrToPtr Then
        dstSizeInBytes = finalSize
    Else
        InternalError "lz4_compress failed", finalSize
        finalSize = 0
    End If
    
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = LZ4_decompress_safe(constSrcPtr, VarPtr(dstArray(0)), constSrcSizeInBytes, constDstSizeInBytes)
    
    'Check for error returns
    If (finalSize <= 0) Then
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    ICompress_DecompressPtrToDstArray = (finalSize > 0)
    
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = LZ4_decompress_safe(constSrcPtr, constDstPtr, constSrcSizeInBytes, constDstSizeInBytes)
    
    'Check for error returns
    If (finalSize <= 0) Then
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    ICompress_DecompressPtrToPtr = (finalSize > 0)
    
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = LZ4_compressBound(srcBufferSizeInBytes)
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = LZ4HC_DEFAULT_CLEVEL
End Function

'PAY ATTENTION - lz4 uses an "acceleration factor" instead of a "compression factor"; because of this,
' "max acceleration factor" actually correlates to "lowest compression ratio"
Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = LZ4HC_MIN_CLEVEL
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = LZ4HC_MAX_CLEVEL
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "lz4_hc"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_Lz4Handle <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow
Public Function GetCompressorVersion() As Long
    Dim ptrVersion As Long
    ptrVersion = LZ4_versionNumber()
    GetCompressorVersion = ptrVersion
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = LZ4HC_DEFAULT_CLEVEL
    ElseIf (inputLevel < LZ4HC_MIN_CLEVEL) Then
        inputLevel = LZ4HC_MIN_CLEVEL
    ElseIf (inputLevel > LZ4HC_MAX_CLEVEL) Then
        inputLevel = LZ4HC_MAX_CLEVEL
    End If
End Sub

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    
    If (faultyReturnCode <> 256) Then
        Debug.Print "lz4 returned an error code: " & faultyReturnCode
    Else
        Debug.Print "lz4 experienced an error; additional explanation may be: " & errString
    End If
    
End Sub


