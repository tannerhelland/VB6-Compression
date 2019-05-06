VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressBrotli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Brotli Compression Library Interface
'Copyright 2019-2019 by Tanner Helland
'Created: 06/May/19
'Last updated: 06/May/19
'Last update: initial build
'
'Per its documentation (available at https://brotli.org/), brotli is...
'
' "...a generic-purpose lossless compression algorithm that compresses data using a combination of a modern
' variant of the LZ77 algorithm, Huffman coding and 2nd order context modeling, with a compression ratio
' comparable to the best currently available general-purpose compression methods."
'
'Brotli is MIT-licensed and sponsored by Google.
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

'These constants were originally declared in encode.h
Private Const BROTLI_MIN_QUALITY As Long = 0
Private Const BROTLI_DEFAULT_QUALITY As Long = 11
Private Const BROTLI_MAX_QUALITY As Long = 11
Private Const BROTLI_MODE_GENERIC As Long = 0
Private Const BROTLI_DEFAULT_WINDOW As Long = 22
Private Const BROTLI_DEFAULT_MODE As Long = BROTLI_MODE_GENERIC

'The following functions are used in this module, but instead of being called directly, calls are routed
' through DispCallFunc (which allows us to use the prebuilt release DLLs provided by the library authors):
'BROTLI_BOOL BrotliEncoderCompress(int quality, int lgwin, BrotliEncoderMode mode, size_t input_size, const uint8_t input_buffer[input_size], size_t * encoded_size, uint8_t encoded_buffer[*encoded_size])
'size_t BrotliEncoderMaxCompressedSize(size_t input_size)
'uint32_t BrotliEncoderVersion(void)
'BrotliDecoderResult BrotliDecoderDecompress (size_t encoded_size, const uint8_t encoded_buffer[encoded_size], size_t *decoded_size, uint8_t decoded_buffer[*decoded_size])

Private Enum BrotliDecoderResult
    '/** Decoding error, e.g. corrupted input or memory allocation problem. */
    BROTLI_DECODER_RESULT_ERROR = 0
    '/** Decoding successfully completed. */
    BROTLI_DECODER_RESULT_SUCCESS = 1
    '/** Partially done; should be called again with more input. */
    BROTLI_DECODER_RESULT_NEEDS_MORE_INPUT = 2
    '/** Partially done; should be called again with more output. */
    BROTLI_DECODER_RESULT_NEEDS_MORE_OUTPUT = 3
End Enum

#If False Then
    Private Const BROTLI_DECODER_RESULT_ERROR = 0, BROTLI_DECODER_RESULT_SUCCESS = 1, BROTLI_DECODER_RESULT_NEEDS_MORE_INPUT = 2, BROTLI_DECODER_RESULT_NEEDS_MORE_OUTPUT = 3
#End If

'Single brotlie enc/dec handles are maintained for the life of a class instance; see Initialize and Release functions, below.
Private m_BrotliCommonHandle As Long, m_BrotliEncHandle As Long, m_BrotliDecHandle As Long

'Required for calling non-stdcall flat dlls in VB6
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum Brotli_ProcAddress
    BrotliEncoderCompress
    BrotliEncoderMaxCompressedSize
    BrotliEncoderVersion
    BrotliDecoderDecompress
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim brotliCommonPath As String, brotliEncPath As String, brotliDecPath As String
    brotliCommonPath = pathToDLLFolder & "brotlicommon.dll"
    brotliEncPath = pathToDLLFolder & "brotlienc.dll"
    brotliDecPath = pathToDLLFolder & "brotlidec.dll"
    m_BrotliCommonHandle = LoadLibraryW(StrPtr(brotliCommonPath))
    m_BrotliEncHandle = LoadLibraryW(StrPtr(brotliEncPath))
    m_BrotliDecHandle = LoadLibraryW(StrPtr(brotliDecPath))
    ICompress_InitializeEngine = (m_BrotliCommonHandle <> 0) And (m_BrotliEncHandle <> 0) And (m_BrotliDecHandle <> 0)
    
    'If we initialized the library successfully, cache some brotli-specific data
    If ICompress_InitializeEngine Then
        
        'Pre-load all relevant proc addresses, and note that they do *not* all come from the same lib
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(BrotliEncoderCompress) = GetProcAddress(m_BrotliEncHandle, "BrotliEncoderCompress")
        m_ProcAddresses(BrotliEncoderMaxCompressedSize) = GetProcAddress(m_BrotliEncHandle, "BrotliEncoderMaxCompressedSize")
        m_ProcAddresses(BrotliEncoderVersion) = GetProcAddress(m_BrotliEncHandle, "BrotliEncoderVersion")
        m_ProcAddresses(BrotliDecoderDecompress) = GetProcAddress(m_BrotliDecHandle, "BrotliDecoderDecompress")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
    Else
        If (m_BrotliEncHandle = 0) Then
            Debug.Print "WARNING!  LoadLibraryW failed to load brotlienc.  Last DLL error: " & Err.LastDllError
            Debug.Print "(FYI, the attempted path was: " & brotliEncPath & ")"
        End If
        If (m_BrotliDecHandle = 0) Then
            Debug.Print "WARNING!  LoadLibraryW failed to load brotlidec.  Last DLL error: " & Err.LastDllError
            Debug.Print "(FYI, the attempted path was: " & brotliDecPath & ")"
        End If
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    
    If (m_BrotliEncHandle <> 0) Then
        FreeLibrary m_BrotliEncHandle
        m_BrotliEncHandle = 0
    End If
    
    If (m_BrotliDecHandle <> 0) Then
        FreeLibrary m_BrotliDecHandle
        m_BrotliDecHandle = 0
    End If
    
    If (m_BrotliCommonHandle <> 0) Then
        FreeLibrary m_BrotliCommonHandle
        m_BrotliCommonHandle = 0
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
    
    'Perform compression
    Dim finalSize As Long
    finalSize = dstCompressedSizeInBytes
    ICompress_CompressPtrToDstArray = (CallCDeclW(BrotliEncoderCompress, vbLong, compressionLevel, BROTLI_DEFAULT_WINDOW, BROTLI_DEFAULT_MODE, constSrcSizeInBytes, constSrcPtr, VarPtr(finalSize), VarPtr(dstArray(0))) <> 0)
    If ICompress_CompressPtrToDstArray Then
        dstCompressedSizeInBytes = finalSize
    Else
        InternalError "Brotli compression failed"
        dstCompressedSizeInBytes = 0
    End If
    
    'Trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    ValidateCompressionLevel compressionLevel
    
    Dim finalSize As Long
    finalSize = dstSizeInBytes
    ICompress_CompressPtrToPtr = (CallCDeclW(BrotliEncoderCompress, vbLong, compressionLevel, BROTLI_DEFAULT_WINDOW, BROTLI_DEFAULT_MODE, constSrcSizeInBytes, constSrcPtr, VarPtr(finalSize), constDstPtr) <> 0)
    
    If ICompress_CompressPtrToPtr Then
        dstSizeInBytes = finalSize
    Else
        InternalError "Brotli compression failed"
        dstSizeInBytes = 0
    End If
    
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = constDstSizeInBytes
    ICompress_DecompressPtrToDstArray = (CallCDeclW(BrotliDecoderDecompress, vbLong, constSrcSizeInBytes, constSrcPtr, VarPtr(finalSize), VarPtr(dstArray(0))) = BROTLI_DECODER_RESULT_SUCCESS)
    If (Not ICompress_DecompressPtrToDstArray) Then InternalError "Brotli decompression failed"
    
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    Dim finalSize As Long
    finalSize = constDstSizeInBytes
    ICompress_DecompressPtrToPtr = (CallCDeclW(BrotliDecoderDecompress, vbLong, constSrcSizeInBytes, constSrcPtr, VarPtr(finalSize), constDstPtr) = BROTLI_DECODER_RESULT_SUCCESS)
    If (Not ICompress_DecompressPtrToPtr) Then InternalError "Brotli decompression failed"
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = CallCDeclW(BrotliEncoderMaxCompressedSize, vbLong, srcBufferSizeInBytes)
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = BROTLI_DEFAULT_QUALITY
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = BROTLI_MIN_QUALITY
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = BROTLI_MAX_QUALITY
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "brotli"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_BrotliEncHandle <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow
Public Function GetCompressorVersion() As Long

    If ICompress_IsCompressorReady() Then
        GetCompressorVersion = CallCDeclW(BrotliEncoderVersion, vbLong)
    Else
        GetCompressorVersion = 0
    End If
    
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = BROTLI_DEFAULT_QUALITY
    ElseIf (inputLevel < BROTLI_MIN_QUALITY) Then
        inputLevel = BROTLI_MIN_QUALITY
    ElseIf (inputLevel > BROTLI_MAX_QUALITY) Then
        inputLevel = BROTLI_MAX_QUALITY
    End If
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As Brotli_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, pFunc As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    vTemp = pa 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    For i = 0 To numParams - 1
        If VarType(pa(i)) = vbString Then vTemp(i) = StrPtr(pa(i))
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    If hResult Then Err.Raise hResult
    
End Function

Private Sub InternalError(ByVal errString As String)
    Debug.Print "brotli experienced an error: " & errString & ", LastDLLError: " & Err.LastDllError
End Sub
