VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressZstd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Zstd Compression Library Interface
'Copyright 2016-2017 by Tanner Helland
'Created: 01/December/16
'Last updated: 07/September/17
'Last update: rewrite as a ICompress implementation
'
'Per its documentation (available at https://github.com/facebook/zstd), zstd is...
'
' "...a fast lossless compression algorithm, targeting real-time compression scenarios
'  at zlib-level and better compression ratios."
'
'zstd is BSD-licensed and sponsored by Facebook.  As of Dec 2016, development is very active and performance
' numbers are very favorable compared to zLib.  (3-4x faster at compressing, ~1.5x faster at decompressing,
' with better compression ratios across the board.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Implements ICompress

'These constants were originally declared in zstd.h
Private Const ZSTD_MIN_CLEVEL As Long = 1
Private Const ZSTD_DEFAULT_CLEVEL As Long = 3
Private Const ZSTD_MAX_CLEVEL As Long = 22

Private Declare Function ZSTD_versionNumber Lib "libzstd" Alias "_ZSTD_versionNumber@0" () As Long
Private Declare Function ZSTD_compress Lib "libzstd" Alias "_ZSTD_compress@20" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long
Private Declare Function ZSTD_decompress Lib "libzstd" Alias "_ZSTD_decompress@16" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long

'These functions are not as self-explanatory as the ones above:
Private Declare Function ZSTD_maxCLevel Lib "libzstd" Alias "_ZSTD_maxCLevel@0" () As Long  'Maximum compression level available
Private Declare Function ZSTD_compressBound Lib "libzstd" Alias "_ZSTD_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array
Private Declare Function ZSTD_isError Lib "libzstd" Alias "_ZSTD_isError@4" (ByVal returnCode As Long) As Long 'Tells you if a function result is an error code or a valid size return
Private Declare Function ZSTD_getErrorName Lib "libzstd" Alias "_ZSTD_getErrorName@4" (ByVal returnCode As Long) As Long 'Returns a pointer to a const char string, with a human-readable string describing the given error code

'If you want, you can ask zstd to tell you how much size is required to decompress a given compression array.  PD doesn't need this
' (as we track compression sizes manually), but it's here if you need it.  Note that automatic calculations like this are generally
' discouraged, as a malicious user can send malformed streams with faulty compression sizes embedded, leading to buffer overflow
' exploits.  Be good, and always manually supply known buffer sizes to external libraries!
'unsigned long long ZSTD_getDecompressedSize(const void* src, size_t srcSize);

'A single zstd handle is maintained for the life of a class instance; see Initialize and Release functions, below.
Private m_ZstdHandle As Long

'Maximum compression level that the library currently supports.  This is cached at initialization time.
Private m_ZstdCompressLevelMax As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim zstdPath As String
    zstdPath = pathToDLLFolder & "libzstd.dll"
    m_ZstdHandle = LoadLibraryW(StrPtr(zstdPath))
    ICompress_InitializeEngine = (m_ZstdHandle <> 0)
    
    'If we initialized the library successfully, cache some zstd-specific data
    If ICompress_InitializeEngine Then
        m_ZstdCompressLevelMax = ZSTD_maxCLevel()
    Else
        Debug.Print "WARNING!  LoadLibraryW failed to load zstd.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & zstdPath & ")"
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    If (m_ZstdHandle <> 0) Then
        FreeLibrary m_ZstdHandle
        m_ZstdHandle = 0
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
    
    Dim finalSize As Long
    finalSize = ZSTD_compress(VarPtr(dstArray(0)), dstCompressedSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        InternalError "ZSTD_compress failed", finalSize
        dstCompressedSizeInBytes = 0
        ICompress_CompressPtrToDstArray = False
    Else
        ICompress_CompressPtrToDstArray = True
        dstCompressedSizeInBytes = finalSize
    End If
    
    'Trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    ValidateCompressionLevel compressionLevel
    
    Dim finalSize As Long
    finalSize = ZSTD_compress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    ICompress_CompressPtrToPtr = (ZSTD_isError(finalSize) = 0)
    
    If ICompress_CompressPtrToPtr Then
        dstSizeInBytes = finalSize
    Else
        InternalError "ZSTD_compress failed", finalSize
        dstSizeInBytes = 0
    End If
    
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = ZSTD_decompress(VarPtr(dstArray(0)), constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ICompress_DecompressPtrToDstArray = (finalSize <> 0)
    
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = ZSTD_decompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ICompress_DecompressPtrToPtr = (finalSize <> 0)
    
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = ZSTD_compressBound(srcBufferSizeInBytes)
    If (ZSTD_isError(ICompress_GetWorstCaseSize) <> 0) Then
        InternalError "ZstdGetMaxCompressedSize failed", ICompress_GetWorstCaseSize
        ICompress_GetWorstCaseSize = 0
    End If
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = ZSTD_DEFAULT_CLEVEL
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = ZSTD_MIN_CLEVEL
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = ZSTD_MAX_CLEVEL
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "zstd"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_ZstdHandle <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow
Public Function GetCompressorVersion() As Long

    If ICompress_IsCompressorReady() Then
        GetCompressorVersion = ZSTD_versionNumber()
    Else
        GetCompressorVersion = 0
    End If
    
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = ZSTD_DEFAULT_CLEVEL
    ElseIf (inputLevel < ZSTD_MIN_CLEVEL) Then
        inputLevel = ZSTD_MIN_CLEVEL
    ElseIf (inputLevel > ZSTD_MAX_CLEVEL) Then
        inputLevel = ZSTD_MAX_CLEVEL
    End If
End Sub

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    
    If (faultyReturnCode <> 256) Then
        
        'Get a char pointer that describes this error
        Dim ptrChar As Long
        ptrChar = ZSTD_getErrorName(faultyReturnCode)
        
        'Convert the char * to a VB string
        Dim errDescription As String
        errDescription = VBHacks.ConvertCharPointerToVBString(ptrChar, False, 255)

        Debug.Print "zstd returned an error code (" & faultyReturnCode & "): " & errDescription
    Else
        Debug.Print "zstd experienced an error: " & errString
    End If
    
End Sub

