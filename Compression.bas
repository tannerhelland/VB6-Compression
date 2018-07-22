Attribute VB_Name = "Compression"
'***************************************************************************
'Unified Compression Interface for PhotoDemon
'Copyright 2016-2017 by Tanner Helland
'Created: 02/December/16
'Last updated: 05/September/17
'Last update: add support for "next-generation zLib" (https://github.com/Dead2/zlib-ng)
'Dependencies: - standalone plugin modules for whatever compression engines you want to use (e.g. the
'              Plugin_ZLib module for zlib compression).  This module simply wraps those dedicated functions,
'              and it performs no library initialization or termination of its own.
'              - VBHacks module (for detecting Windows version and other sundry VB workarounds)
'
'As of v7.0, PhotoDemon performs a *lot* of custom compression work.  There are a lot of different needs in
' image processing - for example, when the user saves a large, multi-layer image, it's okay to take plenty of time
' and squeeze every last bit of compression you can out of the finished file (which is potentially enormous).
' But when saving Undo/Redo data for rapid operations like paint strokes, you want to dump data out to file as
' quickly as humanly possible, with a compression strategy that's as close as possible to HDD performance limits.
'
'For a long time, I used zLib as the program's sole compression interface.  zLib works "well enough" for most
' workloads, with controllable trade-offs between performance and compression, but even at its fastest settings,
' zLib is still one of the slowest available compressors.  (Not surprising, really, given its emphasis on
' portability over all else.)
'
'The past few years have seen a flurry of work on compression algorithms, making it a great time to expand
' PD's compression library coverage.  Zstd came first (http://facebook.github.io/zstd/).  It is basically a
' modernized, superior-in-every-way replacement for zLib.  At its fastest speed setting, it is significantly
' faster than zLib (~4-5x) with only marginally worse compression ratios, while at comparable speed settings,
' it compresses better than zLib across every workload.  Can't argue with that!
'
'Also supported is the lz4 library (http://lz4.github.io/lz4/), developed by the same mad genius as zstd.
' lz4 emphasis real-time compression and decompression speeds, and while its compression ratios are worse
' than both zLib and zstd, it is a full order of magnitude faster.  Its decompression speeds rank among the best
' of any active compression library, making it a useful and unique addition to the corpus.  (It is also the only
' VB-friendly compression library I know of where its performance is good enough to provide concrete benefits
' when reading/writing temp files, because its compression-speed-to-compressed-size ratio is high enough to
' outperform typical disk I/O on a 7200 RPM HDD.)
'
'lz4-hc is also supported.  It is a high-compression variant of lz4, with compression times closer to zLib,
' but the same blazing decompression speeds as stock lz4.  Its support is provided by the stock lz4 library.
'
'Update 05-Sep-17: out of curiosity, I've now added the latest dev build of the "next-generation zLib library"
' (https://github.com/Dead2/zlib-ng/commit/d5a9b75872f15f885cbfa35f08e42faf0cdef5a5).  This project aims to
' modernize and de-cruft the base zLib library, but so far, I'm finding its performance and compression
' results surprisingly meh relative to the base zLib library.  Part of the problem is that many recent zLib
' optimization strategies are built around modern 64-bit processor capabilities, which a 32-bit build
' obviously can't use.  As such, I'm not sure there's any benefit to using current zlib-ng builds over
' stock zlib builds.
'
'I've also included support for the various built-in Windows compression algorithms.  These are only available
' on Win 8 or later, making them poor choices for portability, but if you're only targeting new PCs, they will
' give you compression access without any external dependencies.  (Note that - like most things MS - none of
' the algorithms outperform the 3rd-party solutions, so adjust your expectations accordingly.)
'
'Anyway, the purpose of this module is to simplify code across PD by using standardized compression functions.
' Simply specify the compressor you desire, and this module will silently plug in the right compression or
' decompression code.  (Note that - at present - you *must* request the correct decompressor at decompression
' time, meaning you can't just hand a compressed stream to this module and expect it to magically
' reverse-engineer which decompression engine to use.  That's *your* job.)
'
'All wrapper code in this function is written from scratch by me.  It is not based on any preexisting work.
' This module is, as usual, licensed under the same BSD license governing PD as a whole, so feel free to use
' it in any application, commercial or otherwise.  Bug reports are always welcome.
'
'Licenses for wrapped libraries include:
' zLib: BSD-style license (http://zlib.net/zlib_license.html)
' zLib-ng: BSD-style license (https://github.com/Dead2/zlib-ng/blob/develop/LICENSE.md)
' zstd: BSD 3-clause license (https://github.com/facebook/zstd/blob/dev/LICENSE)
' lz4/lz4-hc: BSD 2-clause license (https://github.com/lz4/lz4/blob/dev/LICENSE)
'
'Copies of these libraries are all custom-built by me as stdcall variants to simplify interop with VB.  Feel free
' to drop-in your own compiled copies, but note that the usual caveats apply if you go with the stock cdecl
' versions - e.g. you will need a safe wrapper around DispCallFunc, such as
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Currently supported compression engines.  Note that you must *always* use the same engine for compression
' and decompression (e.g. there is no way to auto-detect the format of a previously compressed stream).
Public Enum PD_CompressionEngine
    
    'No compression just copies the source bytes to the destination bytes as-is.  It makes for a nice baseline
    ' comparison, especially when testing large data sets.
    PD_CE_NoCompression = 0
    
    'The following compression engines require a 3rd-party DLL
    PD_CE_ZLib = 1
    PD_CE_ZLibNG = 2
    PD_CE_Zstd = 3
    PD_CE_Lz4 = 4
    PD_CE_Lz4HC = 5
    
    'The following compression engines are built-in on Windows 8 or later
    PD_CE_MSZIP = 6
    PD_CE_XPRESS = 7
    PD_CE_XPRESS_HUFF = 8
    PD_CE_LZMS = 9
    
    'The following compression engines are pure VB6 implementation
    PD_CE_ZThunk = 10
    
    [_PD_CE_Last]
End Enum

#If False Then
    Private Const PD_CE_NoCompression = 0, PD_CE_ZLib = 1, PD_CE_ZLibNG = 2, PD_CE_Zstd = 3, PD_CE_Lz4 = 4, PD_CE_Lz4HC = 5, PD_CE_MSZIP = 6, PD_CE_XPRESS = 7, PD_CE_XPRESS_HUFF = 8, PD_CE_LZMS = 9, PD_CE_ZThunk = 10
#End If

'Note that not all compression engines are available on all systems.  Some rely on 3rd-party DLLs; others require Win 8 or later.
Private Const NUM_OF_COMPRESSION_ENGINES = [_PD_CE_Last]

Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)

'When a compression engine is initialized successfully, the matching value in this array will be set to TRUE.
Private m_CompressorAvailable() As Boolean
Private m_Compressor() As ICompress

'Initialize a given compression engine.  The path to the DLL folder *must* include a trailing slash.
'Returns: TRUE if initialization is successful; FALSE otherwise.  FALSE typically means the path to the DLL folder
'         is malformed, or it's correct but the program doesn't have access rights to it.
Public Function InitializeCompressionEngine(ByVal whichEngine As PD_CompressionEngine, ByRef pathToDLLFolder As String) As Boolean
    
    'Keep track of which compression engines have been initialized
    If (Not VBHacks.IsArrayInitialized(m_CompressorAvailable)) Then
        ReDim m_CompressorAvailable(0 To NUM_OF_COMPRESSION_ENGINES - 1) As Boolean
        ReDim m_Compressor(0 To NUM_OF_COMPRESSION_ENGINES - 1) As ICompress
        m_CompressorAvailable(PD_CE_NoCompression) = True
    End If
    
    'Skip initialization if the compressor has already been initialized
    If (Not m_CompressorAvailable(whichEngine)) Then
        
        'Only 3rd-party DLLs need to be initialized.
        If (whichEngine = PD_CE_ZLib) Then
            Set m_Compressor(whichEngine) = New pdCompressZLib
        ElseIf (whichEngine = PD_CE_ZLibNG) Then
            Set m_Compressor(whichEngine) = New pdCompressZLibNG
        ElseIf (whichEngine = PD_CE_Zstd) Then
            Set m_Compressor(whichEngine) = New pdCompressZstd
        ElseIf (whichEngine = PD_CE_Lz4) Then
            Set m_Compressor(whichEngine) = New pdCompressLz4
        ElseIf (whichEngine = PD_CE_Lz4HC) Then
            Set m_Compressor(whichEngine) = New pdCompressLz4HC
        ElseIf (whichEngine = PD_CE_ZThunk) Then
            Set m_Compressor(whichEngine) = New pdCompressZThunk
        ElseIf (whichEngine = PD_CE_MSZIP) Then
            Set m_Compressor(whichEngine) = New pdCompressMSZip
        ElseIf (whichEngine = PD_CE_XPRESS) Then
            Set m_Compressor(whichEngine) = New pdCompressXPress
        ElseIf (whichEngine = PD_CE_XPRESS_HUFF) Then
            Set m_Compressor(whichEngine) = New pdCompressXPressHuff
        ElseIf (whichEngine = PD_CE_LZMS) Then
            Set m_Compressor(whichEngine) = New pdCompressLZMS
        End If
        
        m_CompressorAvailable(whichEngine) = m_Compressor(whichEngine).InitializeEngine(pathToDLLFolder)
        
    End If
    
    InitializeCompressionEngine = m_CompressorAvailable(whichEngine)
    
End Function

'Shut down a compression engine.  You (obviously) cannot use a compression engine once it has been shut down.
' You *must* call this function before your program terminates, and you must call it once for each engine that
' you started this session.
Public Sub ShutDownCompressionEngine(ByVal whichEngine As PD_CompressionEngine)

    'Keep track of which compression engines have been initialized
    If VBHacks.IsArrayInitialized(m_CompressorAvailable) Then
        
        'Skip termination if the compressor has already been shut down
        If m_CompressorAvailable(whichEngine) Then
            m_Compressor(whichEngine).ReleaseEngine
            m_CompressorAvailable(whichEngine) = False
        End If
        
    End If
    
End Sub

'Want to know if a given compression engine is available?  Call this function.  It will (obviously) return FALSE for
' any engines that weren't initialized properly.
Public Function IsCompressionEngineAvailable(ByVal whichEngine As PD_CompressionEngine) As Boolean
    IsCompressionEngineAvailable = m_CompressorAvailable(whichEngine)
End Function

'Compress some arbitrary pointer to a destination array.
'
'Required inputs:
' 1) ByRef Destination array, declared As Byte.  Can be initialized or uninitialized; doesn't matter.
' 2) ByRef final compressed size, as Long.  You generally need to cache this value with your compressed data,
'    so the decompression engine knows how large of a buffer to prepare later on.
' 3) ByVal Pointer to the source data.  This can be any valid pointer, aligned or not.
' 4) ByVal Size of the source data.  This must be byte-accurate, no exceptions.
' 5) ByVal Desired compression engine.  Note that "no compression engine" is a valid option; this module works
'    just fine with uncompressed data, and it will simply perform a fast copy instead (where destination
'    size = source size).
'
'Optional inputs:
' 6) Desired compression level.  This parameter has different meanings for different compression engines.  -1 will use
'    each engine's default setting.  For zLib and zstd, higher values mean *slower but better* compression.  lz4 is the
'    exact opposite; higher values mean *faster but worse* compression.
' 7) If the caller has already prepared the destination array at an appropriate size, pass TRUE for dstArrayIsAlreadySized.
'    This spares us a memory allocation, which can greatly improve performance.  (Note that no verifications are done on the
'    target array, so you *must* have resized the array to a size >= the maximum required size, as calculated by
'    the GetWorstCaseSize() function, ideally.)
' 8) If you want us to trim the destination array to the exact compressed size, pass TRUE for trimCompressedArray.  If you do
'    not specify this, dstArray() will be left at the worst-case size, and it is up to the caller to check the value of
'    dstCompressedSize to see how much size compression actually required.
'
'Returns:
' - TRUE if compression was successful; FALSE otherwise.  Note that a FALSE return will still *always* copy the uncompressed
'   source bytes into the destination array, so you can proceed with processing even if the function fails.
Public Function CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal ptrToSource As Long, ByVal srcBufferSizeInBytes As Long, ByVal compressionEngine As PD_CompressionEngine, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False) As Boolean

    'If the destination array isn't allocated, forcibly initialize it now
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = GetWorstCaseSize(srcBufferSizeInBytes, compressionEngine)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
    'Now that our destination array is guaranteed sized correctly, use naked pointers for compression
    CompressPtrToDstArray = CompressPtrToPtr(VarPtr(dstArray(0)), dstCompressedSizeInBytes, ptrToSource, srcBufferSizeInBytes, compressionEngine, compressionLevel)
    
    'Trim the destination array, as requested
    If trimCompressedArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

'All compression functions ultimately wrap this function.  You can also use it directly, but you *must* size your destination buffer
' correctly to avoid hard crashes.  Also, you *must* pass in the starting destination buffer size as dstSizeInBytes; the compressor
' needs to know this for security reasons.
Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, ByVal compressionEngine As PD_CompressionEngine, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    CompressPtrToPtr = False
    
    If (compressionEngine = PD_CE_NoCompression) Then
        'Do nothing; the catch at the end of the function will handle this case for us
    Else
        CompressPtrToPtr = m_Compressor(compressionEngine).CompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    End If
    
    'If compression failed, perform a direct source-to-dst copy
    If (Not CompressPtrToPtr) Then
        If (compressionEngine <> PD_CE_NoCompression) Then InternalErrorMsg "CompressPtrToPtr failed on compression engine " & compressionEngine
        CopyMemory_Strict constDstPtr, constSrcPtr, constSrcSizeInBytes
        dstSizeInBytes = constSrcSizeInBytes
        CompressPtrToPtr = (compressionEngine = PD_CE_NoCompression)
    End If

End Function

'Decompress some arbitrary pointer (containing compressed data, obviously) to a destination array.
'
'Required inputs:
' 1) ByRef Destination array, declared As Byte.  Can be initialized or uninitialized; doesn't matter.
' 2) ByVal final decompressed size, as Long.  You *must* pass this value to the function, as the decompressed stream
'    may not store this value independently.
' 3) Byval Pointer to the source data.  This can be any valid pointer, aligned or not.
' 4) Byval Size of the source data.  This must be byte-accurate, no exceptions.
' 5) Byval Desired compression engine.  Note that "no compression engine" is a valid option; this module works
'    just fine with uncompressed data, and it will simply perform a fast copy instead (where destination
'    size = source size).
'
'Optional inputs:
' 6) If the caller has already prepared the destination array at an appropriate size, pass TRUE for dstArrayIsAlreadySized.
'    This spares us a memory allocation, which can greatly improve performance.  (Note that no verifications are done on the
'    target array, so you *must* have resized the array to a size >= the original decompressed size.)
'
'Returns:
' - TRUE if decompression was successful; FALSE otherwise.  Note that a FALSE return will still *always* copy the compressed
'   source bytes into the destination array, to mirror the behavior of the matching compression function, above.  (This also
'   allows you to use the compression and decompression functions in "no compression" mode and have them behave as expected.)
'   If FALSE occurs, however, you may need to abandon further processing, as there's currently no way to decompress the
'   bytestream without help from the original decompression library.
Public Function DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal dstDecompressedSizeInBytes As Long, ByVal ptrToSource As Long, ByVal srcBufferSizeInBytes As Long, ByVal compressionEngine As PD_CompressionEngine, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    'If the destination array isn't allocated, forcibly initialize it now
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To dstDecompressedSizeInBytes - 1) As Byte
    
    'Now that our destination array is guaranteed sized correctly, use naked pointers for decompression
    DecompressPtrToDstArray = DecompressPtrToPtr(VarPtr(dstArray(0)), dstDecompressedSizeInBytes, ptrToSource, srcBufferSizeInBytes, compressionEngine)
    
End Function

'All decompression functions ultimately wrap this function.  You can also use it directly, but you *must* size your destination buffer
' correctly to avoid hard crashes.  Also, you *must* pass in the byte-accurate destination buffer size as dstSizeInBytes;
' most decompressors do not store this value independently.
Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, ByVal compressionEngine As PD_CompressionEngine) As Boolean
    
    DecompressPtrToPtr = False
    
    If (compressionEngine = PD_CE_NoCompression) Then
        'Do nothing; the failsafe catch at the end of this function handles this case for us
    Else
        DecompressPtrToPtr = m_Compressor(compressionEngine).DecompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
    End If
    
    'If compression failed, perform a direct source-to-dst copy
    If (Not DecompressPtrToPtr) Then
        If (compressionEngine <> PD_CE_NoCompression) Then InternalErrorMsg "DecompressPtrToPtr failed on compression engine " & compressionEngine
        CopyMemory_Strict constDstPtr, constSrcPtr, constSrcSizeInBytes
        DecompressPtrToPtr = (compressionEngine = PD_CE_NoCompression)
    End If

End Function

'All compression functions require a destination buffer sized to the "worst-case" scenario size, which is the largest size
' the compressed data will consume if it is 100% incompressible.  You can almost always shrink the destination buffer after
' the fact (to the exact compressed size), but you must always start with a buffer at least this large.
'
'Obviously, you must pass the size of your source data, and you must also specify the desired compression engine (as they
' use different rules for formulating a "worst-case" size).
Public Function GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long, ByVal compressionEngine As PD_CompressionEngine) As Long
    
    If (compressionEngine = PD_CE_NoCompression) Then
        GetWorstCaseSize = srcBufferSizeInBytes
    Else
        GetWorstCaseSize = m_Compressor(compressionEngine).GetWorstCaseSize(srcBufferSizeInBytes)
    End If

End Function

'Retrieve default/min/max compression levels from a given library.  Note that these are *not* standardized,
' meaning each library has its own default/min/max levels, and the levels mean different things in different
' libraries.  For example, most libraries use the formula where "larger numbers = slower, better compression",
' but lz4 is the opposite - e.g. "larger lz4 numbers = faster, worse compression".
'
'None of the Windows compression functions support variable compression or acceleration levels, unfortunately.
' (Actually, this isn't *technically* true - the XPRESS and XPRESS_HUFF algorithms support levels of either
'  "0" or "1", but I don't current implement these because the differences are small and these algorithms
'  are terrible compared to lz4 anyway.)
Public Function GetDefaultCompressionLevel(ByVal whichEngine As PD_CompressionEngine) As Long

    If (whichEngine = PD_CE_NoCompression) Then
        GetDefaultCompressionLevel = 0
    Else
        GetDefaultCompressionLevel = m_Compressor(whichEngine).GetDefaultCompressionLevel()
    End If

End Function

Public Function GetMinCompressionLevel(ByVal whichEngine As PD_CompressionEngine) As Long
    
    If (whichEngine = PD_CE_NoCompression) Then
        GetMinCompressionLevel = 0
    Else
        GetMinCompressionLevel = m_Compressor(whichEngine).GetMinCompressionLevel()
    End If
    
End Function

Public Function GetMaxCompressionLevel(ByVal whichEngine As PD_CompressionEngine) As Long

    If (whichEngine = PD_CE_NoCompression) Then
        GetMaxCompressionLevel = 0
    Else
        GetMaxCompressionLevel = m_Compressor(whichEngine).GetMaxCompressionLevel()
    End If

End Function

'This function exists purely for debug purposes.  Feel free to remove it if you find it unnecessary.
Public Function GetCompressorName(ByVal whichEngine As PD_CompressionEngine) As String
    
    If (whichEngine = PD_CE_NoCompression) Then
        GetCompressorName = vbNullString
    Else
        GetCompressorName = m_Compressor(whichEngine).GetCompressorName()
    End If

End Function

Private Sub InternalErrorMsg(ByVal errMsg As String)
    Debug.Print "WARNING! Compression module error: " & errMsg
End Sub
