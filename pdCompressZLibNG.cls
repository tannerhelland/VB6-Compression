VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressZLibNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'File Compression Interface (via zLib-ng)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 07/September/17
'Last update: rewrite as a ICompress implementation
'
'Module to handle file compression and decompression to a custom file format via the zLib-ng compression library.
' The "ng" in "zLib-ng" stands for "next-generation", and you can read more about various attempts to modernize
' the zLib codebase here: https://github.com/Dead2/zlib-ng
'
'The current zLib-ng build I'm used is taken from this commit, dated 25 August 2017:
' https://github.com/Dead2/zlib-ng/commit/d5a9b75872f15f885cbfa35f08e42faf0cdef5a5
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Implements ICompress

Private Const ZLIB_OK = 0

'These constants were originally declared in zlib.h.  Note that zLib weirdly supports level 0, which just performs
' a bare memory copy with no compression.  We deliberately omit that possibility here.
Private Const ZLIB_MIN_CLEVEL = 1
Private Const ZLIB_MAX_CLEVEL = 9

'This constant was originally declared (or rather, resolved) in deflate.c.
Private Const ZLIB_DEFAULT_CLEVEL = 6

Private Declare Function compress Lib "zlib-ng" Alias "_compress@16" (ByVal ptrToDestBuffer As Long, ByRef dstLen As Long, ByVal ptrToSrcBuffer As Long, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib-ng" Alias "_compress2@20" (ByVal ptrDstBuffer As Long, ByRef dstLen As Long, ByVal ptrSrcBuffer As Any, ByVal srcLen As Long, ByVal cmpLevel As Long) As Long
Private Declare Function uncompress Lib "zlib-ng" Alias "_uncompress@16" (ByVal ptrToDestBuffer As Long, ByRef dstLen As Long, ByVal ptrToSrcBuffer As Long, ByVal srcLen As Long) As Long
Private Declare Function zlibVersion Lib "zlib-ng" Alias "_zlibVersion@0" () As Long

'A single zLib handle is maintained for the life of a class instance; see InitializeZLib and ReleaseZLib, below.
Private m_ZLibHandle As Long

Private Sub Class_Terminate()
    ICompress_ReleaseEngine
End Sub

'Basic init/release functions
Private Function ICompress_InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim zLibPath As String
    zLibPath = pathToDLLFolder & "zlib-ng.dll"
    m_ZLibHandle = LoadLibraryW(StrPtr(zLibPath))
    ICompress_InitializeEngine = (m_ZLibHandle <> 0)
    
    If (Not ICompress_InitializeEngine) Then
        Debug.Print "WARNING!  LoadLibraryW failed to load zLib.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & zLibPath & ")"
    End If
    
End Function

Private Sub ICompress_ReleaseEngine()
    If (m_ZLibHandle <> 0) Then
        FreeLibrary m_ZLibHandle
        m_ZLibHandle = 0
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
    
    'Compress the data using zLib
    ICompress_CompressPtrToDstArray = (compress2(VarPtr(dstArray(0)), dstCompressedSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel) = ZLIB_OK)
    
    'Trim the destination array, as requested
    If trimCompressedArray And ICompress_CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Private Function ICompress_CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    ValidateCompressionLevel compressionLevel
    ICompress_CompressPtrToPtr = (compress2(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel) = ZLIB_OK)
End Function

Private Function ICompress_DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    ICompress_DecompressPtrToDstArray = (uncompress(VarPtr(dstArray(0)), constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes) = ZLIB_OK)
End Function

Private Function ICompress_DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    ICompress_DecompressPtrToPtr = (uncompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes) = ZLIB_OK)
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = srcBufferSizeInBytes + Int(CDbl(srcBufferSizeInBytes) * 0.01) + 12&
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = ZLIB_DEFAULT_CLEVEL
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = ZLIB_MIN_CLEVEL
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = ZLIB_MAX_CLEVEL
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "zLib-ng"
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = (m_ZLibHandle <> 0)
End Function

'***********************************************************************
'Non-ICompress methods follow
Public Function GetCompressorVersion() As String

    If ICompress_IsCompressorReady() Then
        
        'Get a pointer to the version string
        Dim ptrZLibVer As Long
        ptrZLibVer = zlibVersion()
        
        'Convert the char * to a VB string
        If (ptrZLibVer <> 0) Then
            GetCompressorVersion = VBHacks.ConvertCharPointerToVBString(ptrZLibVer, False, 255)
        Else
            GetCompressorVersion = vbNullString
        End If
        
    Else
        GetCompressorVersion = vbNullString
    End If
    
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = ZLIB_DEFAULT_CLEVEL
    ElseIf (inputLevel < ZLIB_MIN_CLEVEL) Then
        inputLevel = ZLIB_MIN_CLEVEL
    ElseIf (inputLevel > ZLIB_MAX_CLEVEL) Then
        inputLevel = ZLIB_MAX_CLEVEL
    End If
End Sub

