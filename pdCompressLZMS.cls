VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressLZMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Microsoft Windows Compression Library Interface
'Copyright 2016-2017 by Tanner Helland
'Created: 04/December/16
'Last updated: 07/September/17
'Last update: rewrite as a ICompress implementation
'
'This class provides support for one of the built-in Windows compression algorithms.  This class will only
' function on Win 8 or later, making it a poor choice for portability -- but if you're only targeting new PCs,
' this gives you compression access without any external dependencies.  (Note that - like most things MS -
' their internal algorithms do not tend to outperform 3rd-party solutions, so adjust expectations accordingly.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Implements ICompress

Private Enum MS_CompressAlgorithm
    ca_MSZIP = 2
    ca_XPRESS = 3
    ca_XPRESS_HUFF = 4
    ca_LZMS = 5
End Enum

#If False Then
    Private Const ca_MSZIP = 2, ca_XPRESS = 3, ca_XPRESS_HUFF = 4, ca_LZMS = 5
#End If

'**All of these functions require Windows 8 or later!**
Private Declare Function CloseCompressor Lib "cabinet" (ByVal hCompressor As Long) As Long
Private Declare Function CloseDecompressor Lib "cabinet" (ByVal hDecompressor As Long) As Long

'We use an aliased name for this function so that it doesn't cause IDE case changes of matching 3rd-party functions
Private Declare Function MS_Compress Lib "cabinet" Alias "Compress" (ByVal hCompressor As Long, ByVal ptrToUncompressedData As Long, ByVal sizeOfUncompressedData As Long, ByVal ptrToCompressedData As Long, ByVal sizeOfCompressedBuffer As Long, ByRef finalCompressedSize As Long) As Long
Private Declare Function Decompress Lib "cabinet" (ByVal hCompressor As Long, ByVal ptrToCompressedData As Long, ByVal sizeOfCompressedData As Long, ByVal ptrToUncompressedData As Long, ByVal sizeOfUncompressedBuffer As Long, ByRef finalUncompressedSize As Long) As Long
Private Declare Function CreateCompressor Lib "cabinet" (ByVal whichAlgorithm As MS_CompressAlgorithm, ByVal ptrToAllocationRoutines As Long, ByRef hCompressor As Long) As Long
Private Declare Function CreateDecompressor Lib "cabinet" (ByVal whichAlgorithm As MS_CompressAlgorithm, ByVal ptrToAllocationRoutines As Long, ByRef hDecompressor As Long) As Long

'OS-level compression APIs are only available on Win 8 or later; we check for this automatically
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByVal lpVersionInformation As Long) As Long

Private Type OS_VersionInfoEx
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(0 To 255) As Byte
    wServicePackMajor  As Integer
    wServicePackMinor  As Integer
    wSuiteMask         As Integer
    wProductType       As Byte
    wReserved          As Byte
End Type

'To improve performance, the first call to GetVersionEx is cached, and subsequent calls just use the cached value.
Private m_OSVI As OS_VersionInfoEx, m_VersionInfoCached As Boolean

'Persistent compressor and decompressor handles are cached by InitializeEngine;
' these handles are valid for the lifetime of the class, unless ReleaseEngine is called.
Private m_hCompressor As Long, m_hDecompressor As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Manually validate the current windows version
    If IsWin8OrLater Then
    
        ICompress_InitializeEngine = (CreateCompressor(ca_LZMS, 0&, m_hCompressor) <> 0)
        
        'Double-check the returned handle
        If ICompress_InitializeEngine Then
            ICompress_InitializeEngine = (m_hCompressor <> 0)
        Else
            InternalError
        End If
        
        'Repeat the above steps for a decompressor
        If ICompress_InitializeEngine Then
            ICompress_InitializeEngine = (CreateDecompressor(ca_LZMS, 0&, m_hDecompressor) <> 0)
            If ICompress_InitializeEngine Then
                ICompress_InitializeEngine = (m_hDecompressor <> 0)
            Else
                InternalError
            End If
        End If
        
    Else
        ICompress_InitializeEngine = False
        InternalError "WARNING!  This compressor requires Win 8 or later."
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    If (m_hCompressor <> 0) Then
        CloseCompressor m_hCompressor
        m_hCompressor = 0
    End If
    If (m_hDecompressor <> 0) Then
        CloseDecompressor m_hDecompressor
        m_hDecompressor = 0
    End If
End Sub

'Actual compression/decompression functions.  Only arrays and pointers are standardized.  It's assumed
' that users can write simple wrappers for other data types, as necessary.
Private Function ICompress_CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False) As Boolean

    'Prep the destination array, as necessary
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = ICompress_GetWorstCaseSize(constSrcSizeInBytes)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
    'Perform the compression
    Dim outputSizeUsed As Long
    ICompress_CompressPtrToDstArray = (MS_Compress(m_hCompressor, constSrcPtr, constSrcSizeInBytes, VarPtr(dstArray(0)), dstCompressedSizeInBytes, outputSizeUsed) <> 0)
    If (Not ICompress_CompressPtrToDstArray) Then InternalError
    
    'Return the number of bytes used
    dstCompressedSizeInBytes = outputSizeUsed
    
    'Trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    'Perform the compression
    Dim outputSizeUsed As Long
    ICompress_CompressPtrToPtr = (MS_Compress(m_hCompressor, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes, outputSizeUsed) <> 0)
    If (Not ICompress_CompressPtrToPtr) Then InternalError
    
    'Return the number of bytes used
    dstSizeInBytes = outputSizeUsed
    
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    
    'Perform decompression
    Dim outputSizeUsed As Long
    ICompress_DecompressPtrToDstArray = (Decompress(m_hDecompressor, constSrcPtr, constSrcSizeInBytes, VarPtr(dstArray(0)), constDstSizeInBytes, outputSizeUsed) <> 0)
            
    'Check for error returns
    If (Not ICompress_DecompressPtrToDstArray) Then InternalError
    
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    
    'Perform decompression
    Dim outputSizeUsed As Long
    ICompress_DecompressPtrToPtr = (Decompress(m_hDecompressor, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, outputSizeUsed) <> 0)
            
    'Check for error returns
    If (Not ICompress_DecompressPtrToPtr) Then InternalError
    
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    
    'Microsoft compressor limits must be acquired at run-time
    Dim outputSizeRequired As Long
    If (MS_Compress(m_hCompressor, 0&, srcBufferSizeInBytes, 0&, 0&, outputSizeRequired) <> 0) Then
        ICompress_GetWorstCaseSize = outputSizeRequired
    
    'I'm not sure why this function would fail, but if it does, request an extraordinarily large buffer,
    ' "just in case"
    Else
        ICompress_GetWorstCaseSize = srcBufferSizeInBytes * 2
    End If
    
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = 0
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = 0
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = 0
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "LZMS"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_hCompressor <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow

'Private methods follow
Private Sub InternalError(Optional ByVal errString As String = vbNullString)
        
    If (Len(errString) = 0) Then
        Debug.Print "A Microsoft compressor returned an error code: " & Err.LastDllError()
    Else
        Debug.Print "A Microsoft compressor failed; additional explanation may be: " & errString
    End If
    
End Sub

Private Sub CacheOSVersion()
    If (Not m_VersionInfoCached) Then
        m_OSVI.dwOSVersionInfoSize = Len(m_OSVI)
        GetVersionEx VarPtr(m_OSVI)
        m_VersionInfoCached = True
    End If
End Sub

Private Function IsWin8OrLater() As Boolean
    CacheOSVersion
    IsWin8OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 2))
End Function

