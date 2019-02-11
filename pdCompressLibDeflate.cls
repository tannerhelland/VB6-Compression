VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressLibDeflate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'File Compression Interface (via libdeflate)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 11/February/19
'Last update: intiial build
'
'LibDeflate: https://github.com/ebiggers/libdeflate
' - "libdeflate is a library for fast, whole-buffer DEFLATE-based compression and decompression."
' - "libdeflate is heavily optimized. It is significantly faster than the zlib library, both for
'    compression and decompression, and especially on x86 processors."
'
'This wrapper class uses a shorthand implementation of DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Implements ICompress

Private Enum LibDeflate_Result
    ld_Success = 0
    ld_BadData = 1
    ld_ShortOutput = 2
    ld_InsufficientSpace = 3
End Enum

#If False Then
    Private Const ld_Success = 0, ld_BadData = 1, ld_ShortOutput = 2, ld_InsufficientSpace = 3
#End If

'LibDeflate is zlib-compatible, but it exposes even higher compression levels (12 vs zlib's 9) for
' better-but-slower compression.  The default value remains 6; these are all declared in libdeflate.h
Private Const LIBDEFLATE_MIN_CLEVEL = 1
Private Const LIBDEFLATE_MAX_CLEVEL = 12
Private Const LIBDEFLATE_DEFAULT_CLEVEL = 6

'libdeflate has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile myself, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'A single libdeflate handle is maintained for the life of a class instance; see Initialize and Release functions, below.
Private m_libDeflateHandle As Long

'To simplify interactions, we declare persistent libdeflate compressor and decompressor instances
' "as we go".  Compressors are unique in-that they are specific to a given compression level
' (e.g. "compress level 1" compressor != "compress level 6" compressor) which makes interactions
' a little wonky as far as this class is concerned; for best results, you should only
Private m_hCompressor As Long, m_hDecompressor As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim libDeflatePath As String
    libDeflatePath = pathToDLLFolder & "libdeflate.dll"
    m_libDeflateHandle = LoadLibraryW(StrPtr(libDeflatePath))
    ICompress_InitializeEngine = (m_libDeflateHandle <> 0)
    
    If (Not ICompress_InitializeEngine) Then
        Debug.Print "WARNING!  LoadLibraryW failed to load libdeflate.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & libDeflatePath & ")"
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    If (m_libDeflateHandle <> 0) Then
        FreeLibrary m_libDeflateHandle
        m_libDeflateHandle = 0
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
    
    'Compress the data
    ICompress_CompressPtrToDstArray = LibDeflateCompress(VarPtr(dstArray(0)), dstCompressedSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
        
    'If compression was successful, trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    ICompress_CompressPtrToPtr = LibDeflateCompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    ICompress_DecompressPtrToDstArray = LibDeflateDecompress(VarPtr(dstArray(0)), constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    ICompress_DecompressPtrToPtr = LibDeflateDecompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function LibDeflateCompress(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    ValidateCompressionLevel compressionLevel
    
    'Allocate a compressor
    ' LIBDEFLATEAPI struct libdeflate_compressor * libdeflate_alloc_compressor(int compression_level)
    Dim hCompress As Long
    hCompress = CallCDeclW("libdeflate_alloc_compressor", vbLong, compressionLevel)
    
    If (hCompress <> 0) Then
        
        'Perform compression
        ' LIBDEFLATEAPI size_t libdeflate_zlib_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        Dim lReturn As Long
        lReturn = CallCDeclW("libdeflate_zlib_compress", vbLong, hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        LibDeflateCompress = (lReturn <> 0)
        If LibDeflateCompress Then
            dstSizeInBytes = lReturn
        Else
            Debug.Print "libdeflate_zlib_compress() failed"
        End If
        
        'Free the compressor before exiting
        ' LIBDEFLATEAPI void libdeflate_free_compressor(struct libdeflate_compressor *compressor)
        CallCDeclW "libdeflate_free_compressor", vbNull, hCompress
        
    Else
        Debug.Print "WARNING!  Failed to initialize a libdeflate compressor."
    End If
    
    
End Function

Private Function LibDeflateDecompress(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = CallCDeclW("libdeflate_alloc_decompressor", vbLong)
    If (hDecompress <> 0) Then
        
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        Dim lReturn As Long
        lReturn = CallCDeclW("libdeflate_zlib_decompress", vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        LibDeflateDecompress = (lReturn = 0)
        If (Not LibDeflateDecompress) Then Debug.Print "libdeflate_zlib_decompress() failed; return was " & lReturn
        
        ' Make sure we free the compressor before exiting
        'LIBDEFLATEAPI void libdeflate_free_decompressor(struct libdeflate_decompressor *decompressor);
        CallCDeclW "libdeflate_free_decompressor", vbEmpty, hDecompress
        
    Else
        Debug.Print "WARNING!  Failed to initialize a libdeflate decompressor."
    End If
    
    
End Function

'Note that libdeflate exports its own "get worst-case dst size" function.  However, it requires you to
' pass a compressor handle that has been initialized to the target compression level... which creates
' problems for the way our ICompress interface works.  Because there's no good way to mimic this,
' we simply use the standard zlib "worst case" calculation, but add extra bytes for the gzip case
' (as gzip headers/trailers are larger than zlib ones).
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = srcBufferSizeInBytes + Int(CDbl(srcBufferSizeInBytes) * 0.01) + 40&
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = LIBDEFLATE_DEFAULT_CLEVEL
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = LIBDEFLATE_MIN_CLEVEL
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = LIBDEFLATE_MAX_CLEVEL
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "libdeflate"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_libDeflateHandle <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow
Public Function GetCompressorVersion() As String
    'libdeflate doesn't export a version function, but this class was designed against the v1.2 release.
    ' No promises are made about compatibility with other releases.
    GetCompressorVersion = "1.2"
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = LIBDEFLATE_DEFAULT_CLEVEL
    ElseIf (inputLevel < LIBDEFLATE_MIN_CLEVEL) Then
        inputLevel = LIBDEFLATE_MIN_CLEVEL
    ElseIf (inputLevel > LIBDEFLATE_MAX_CLEVEL) Then
        inputLevel = LIBDEFLATE_MAX_CLEVEL
    End If
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Public Function CallCDeclW(ByRef sFunc As String, ByVal fRetType As VbVarType, ParamArray pA() As Variant) As Variant

    Dim i As Long, pFunc As Long, vTemp() As Variant, hResult As Long
    Dim vType() As Integer, vPtr() As Long
    
    Dim numParams As Long
    
    If (UBound(pA) < LBound(pA)) Then numParams = 0 Else numParams = UBound(pA) + 1
    ReDim vType(0 To numParams) As Integer
    ReDim vPtr(0 To numParams) As Long
    
    vTemp = pA 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    For i = 0 To numParams - 1
        If VarType(pA(i)) = vbString Then vTemp(i) = StrPtr(pA(i))
        vType(i) = VarType(vTemp(i))
        vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL = 1
    hResult = DispCallFunc(0, GetProcAddress(m_libDeflateHandle, sFunc), CC_CDECL, fRetType, i, vType(0), vPtr(0), CallCDeclW)
    If hResult Then Err.Raise hResult
    
End Function
