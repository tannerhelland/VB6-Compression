VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H80000005&
   Caption         =   "Compression demo"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSort 
      Height          =   375
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   10215
   End
   Begin VB.CheckBox chkTestLevels 
      BackColor       =   &H80000005&
      Caption         =   "Test multiple compression levels"
      Height          =   255
      Left            =   7800
      TabIndex        =   5
      Top             =   480
      Width           =   4335
   End
   Begin VB.CheckBox chkTestWindowsLibraries 
      BackColor       =   &H80000005&
      Caption         =   "Test built-in Windows libraries"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   3615
   End
   Begin VB.CheckBox chkTest3rdParty 
      BackColor       =   &H80000005&
      Caption         =   "Test 3rd-party libraries"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.TextBox txtResults 
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   12015
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort test results by: "
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   990
      Width           =   1740
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Want to test your own file?  Drag it onto the text box.)"
      Height          =   255
      Index           =   1
      Left            =   6540
      TabIndex        =   2
      Top             =   120
      Width           =   5565
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Compression test results shown below"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11925
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Modern Compression Demonstration for VB6
'Copyright 2016-2017 by Tanner Helland
'Created: 02/December/16
'Last updated: 05/September/17
'Last update: update zstd and lz4 to latest stable builds; add zlib-ng; make output sortable;
'             other minor bug-fixes and code cleanup
'
'This small project demonstrates a simple interface for three major compression libraries -
' zlib, zstd, and lz4/lz4_hc - as well as the built-in Windows Compression API (if you're on
' Windows 8 or later).  The sample project also includes three small XML files in different
' languages as its test base.  This is not a great way to demonstrate compression capabilities,
' but vbforums.com makes it very difficult to attach larger files.  :/
'
'To test your own files, drag/drop them onto the text box.  Larger files will generally provide
' a more useful comparison.
'
'I know people are curious about "but which compression engine is best?"  The answer is, "it depends
' on your needs."  In general...
'
' - zLib should really only be used in projects that specifically require the DEFLATE algorithm.
'
' - zLib-ng is a "next-generation" version of zLib, developed by an outside team.  While it includes
'   many performance improvements relative to stock zLib, these improvements don't seem to make a
'   huge difference in typical workloads.  YMMV.
'
' - zstd is the best "general purpose" compressor.  It is faster than zLib at both compression and
'   decompression, and it provides better compression at faster speeds.
'
' - lz4 is the fastest compressor/decompressor.  Its compression ratio is not as good as zstd or zLib,
'   but you can use the lz4_hc version of the compressor (provided by the same DLL) to improve
'   compression ratio at some cost to speed.  Decompression speed is the same whether lz4 or lz4_hc is used.
'
' - The Windows Compression API functions are generally worse then zstd and/or lz4, but they require
'   no external dependencies, so as long as you're only targeting Win 8 or later, you get them "for free".
'
'I specifically picked these third-party libraries because they are relatively VB-friendly, and they allow
' for use in any project, commercial or otherwise.  Licenses for the provided libraries include:
'
' zLib: BSD-style license (http://zlib.net/zlib_license.html)
' zLib-ng: BSD-style license (https://github.com/Dead2/zlib-ng/blob/develop/LICENSE.md)
' zstd: BSD 3-clause license (https://github.com/facebook/zstd/blob/dev/LICENSE)
' lz4/lz4-hc: BSD 2-clause license (https://github.com/lz4/lz4/blob/dev/LICENSE)
'
'The copies of these libraries included in this sample project were custom-built by me as stdcall variants
' to simplify interop with VB.  Feel free to drop-in your own compiled copies, but note that the usual
' caveats apply if you go with the stock cdecl versions - e.g. you will need to use a safe wrapper around
' DispCallFunc, such as
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)
'
'Similarly, I've pulled all the Compression wrapper code from my open-source PhotoDemon project
' (www.photodemon.org), which is also BSD-licensed (http://photodemon.org/about/license/).  Just like the
' attached DLLs, you can use this code in any application, commercial or otherwise, as long as you provide
' some small attribution note in your UI or documentation.  (And remember - the main purpose of attribution
' is not vanity.  It's so that people know where to send bug reports if they run into anything unexpected.
' Open-source projects can't be improved without input and testing from others!)
'
'Happy coding,
' Tanner H
' (05 September 2017)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit
'
'***************************************************************************

Option Explicit
Option Base 0

'If the user drags their own file onto the text box, we'll store the filename here.  That way, if they change
' the current compression settings, we can immediately apply the new settings to their sample file.
Private m_LastTestedFile As String

'To make it easier to compare compression engine results, we use a custom struct for storing compression results.
' (This lets us easily sort run data.)
Private Type PD_CompressionResult
    CompressorName As String
    CompressionTime As Double
    DecompressionTime As Double
    CompressionRatio As Double
    CompRatioVsTime As Double
    EngineDetails As String
End Type

'All compression results will be stored in this array; the array can then be sorted before dumping it to the screen.
Private m_CompressionResults() As PD_CompressionResult
Private m_NumOfResults As Long

'The test results can be sorted by several different criteria
Private Enum PD_SortResults
    sr_SortByName = 0
    sr_SortByCompTime = 1
    sr_SortByDecompTime = 2
    sr_SortByCompRatio = 3
    sr_SortByRatioVsTime = 4
End Enum

#If False Then
    Private Const sr_SortByName = 0, sr_SortByCompTime = 1, sr_SortByDecompTime = 2, sr_SortByCompRatio = 3, sr_SortByRatioVsTime = 4
#End If

Private Sub cboSort_Click()
    SortAndDisplayResults
End Sub

Private Sub chkTest3rdParty_Click()
    SetupTest m_LastTestedFile, m_LastTestedFile
End Sub

Private Sub chkTestLevels_Click()
    SetupTest m_LastTestedFile, m_LastTestedFile
End Sub

Private Sub chkTestWindowsLibraries_Click()
    SetupTest m_LastTestedFile, m_LastTestedFile
End Sub

Private Sub Form_Load()
    
    'Add sort options
    cboSort.AddItem "Compressor name", 0
    cboSort.AddItem "Compression time (lower is better)", 1
    cboSort.AddItem "Decompression time (lower is better)", 2
    cboSort.AddItem "Compression ratio (higher is better)", 3
    cboSort.AddItem "Compression Ratio vs Total Time (higher is better)", 4
    cboSort.ListIndex = 3
    
    'Initialize the three compression DLLs.  You can store these DLLs wherever you want, but because these
    ' compression libraries are popular, I *strongly* recommend you deploy them locally, inside your
    ' program folder.
    Dim compressionDLLFolder As String
    compressionDLLFolder = App.Path & "\Plugins\"
    
    'Remember that you only need to initialize the compression engines you'll actually be using.  Most people
    ' won't need to deploy all three of these - I just do it here so you can see how they all work.
    If Compression.InitializeCompressionEngine(PD_CE_ZLib, compressionDLLFolder) Then AddText "zLib initialized successfully!" Else AddText "zLib initialization failed (path = " & compressionDLLFolder & ")"
    If Compression.InitializeCompressionEngine(PD_CE_ZLibNG, compressionDLLFolder) Then AddText "zLib-ng initialized successfully!" Else AddText "zLib-ng initialization failed (path = " & compressionDLLFolder & ")"
    If Compression.InitializeCompressionEngine(PD_CE_Zstd, compressionDLLFolder) Then AddText "zstd initialized successfully!" Else AddText "zstd initialization failed (path = " & compressionDLLFolder & ")"
    If Compression.InitializeCompressionEngine(PD_CE_Lz4, compressionDLLFolder) Then AddText "lz4 initialized successfully!" Else AddText "lz4 initialization failed (path = " & compressionDLLFolder & ")"
    If Compression.InitializeCompressionEngine(PD_CE_Lz4HC, compressionDLLFolder) Then AddText "lz4_hc initialized successfully!" Else AddText "lz4hc initialization failed (path = " & compressionDLLFolder & ")"
    If Compression.InitializeCompressionEngine(PD_CE_ZThunk, compressionDLLFolder) Then AddText "zthunk initialized successfully!" Else AddText "zthunk initialization failed (path = " & compressionDLLFolder & ")"
    
    'The Windows compression engines come as a whole group - as long as you're on Windows 8 or later,
    ' they should always initialize successfully.
    If Compression.InitializeCompressionEngine(PD_CE_MSZIP, vbNullString) Then
        Compression.InitializeCompressionEngine PD_CE_XPRESS, vbNullString
        Compression.InitializeCompressionEngine PD_CE_XPRESS_HUFF, vbNullString
        Compression.InitializeCompressionEngine PD_CE_LZMS, vbNullString
        AddText "built-in Windows engines initialized successfully!"
    Else
        AddText "built-in Windows engines initialization failed - are you on Windows 8 or later?"
    End If
        
    'Please note that you also need to shut down any initialized compression engines when you're finished with them.
    ' This is demonstrated in the Form_Unload event of this sample project.
    
    'You don't need this line of code in your own project, but here, we want to measure how long each
    ' compression test takes.  We do this with help from some Windows APIs that must be initialized in advance.
    VBHacks.EnableHighResolutionTimers
    
    'We're now going to do a quick test of all three compression engines, on three different XML files.
    ' For testing purposes, we're going to measure how well each file is compressed, and how long it takes to
    ' compress/decompress them.
    SetupTest
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'When we're done, we need to shut down all compression engines we started
    Compression.ShutDownCompressionEngine PD_CE_ZLib
    Compression.ShutDownCompressionEngine PD_CE_ZLibNG
    Compression.ShutDownCompressionEngine PD_CE_Zstd
    Compression.ShutDownCompressionEngine PD_CE_Lz4
    
    'Note that you do not need to shut down the built-in Windows compression engines (but if you do, there's no harm)
    
End Sub

'Set up a series of tests.  If no filename is passed, we'll run a series of tests on the small XML files included
' with the project.
Private Sub SetupTest(Optional ByVal srcFilename As String = vbNullString, Optional ByVal testName As String = vbNullString)
    
    'If the caller supplies their own filename, run a test on that file only
    If (Len(srcFilename) <> 0) Then
        StartTestOnFile srcFilename, testName
    
    'Otherwise, run a test on a highly compressible XML file included with this project
    Else
    
        Dim srcTestFile As String
        srcTestFile = App.Path & "\test files\Sample.xml"
        SetupTest srcTestFile, "Sample XML file (UTF-8, multiple languages)"
        
    End If
    
End Sub

Private Sub StartTestOnFile(ByVal srcFilename As String, ByVal testName As String)
    
    AddText vbCrLf & "Results for " & testName & ":"
    
    'Reset our compression results table
    ReDim m_CompressionResults(0 To 15) As PD_CompressionResult
    m_NumOfResults = 0
    
    'Always provide a baseline measurement of "no compression".  This just performs a raw source-to-destination copy,
    ' which tells us how long the "read/write memory" portion of the test takes.  (For very large files, this time
    ' is non-trivial.)
    TestCompressionEngine PD_CE_NoCompression, srcFilename
    
    'If the user wants 3rd-party libraries tested, do them now
    If CBool(chkTest3rdParty.Value) Then
    
        'If the caller wants different compression levels tested, we will run three tests on each library:
        ' the minimum level, (min + max) / 2 level, and maximum level.
        If CBool(chkTestLevels.Value) Then
        
            TestCompressionEngine PD_CE_ZLib, srcFilename, Compression.GetMinCompressionLevel(PD_CE_ZLib)
            TestCompressionEngine PD_CE_ZLib, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_ZLib) + Compression.GetMaxCompressionLevel(PD_CE_ZLib)) \ 2
            TestCompressionEngine PD_CE_ZLib, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_ZLib)
            
            TestCompressionEngine PD_CE_ZLibNG, srcFilename, Compression.GetMinCompressionLevel(PD_CE_ZLibNG)
            TestCompressionEngine PD_CE_ZLibNG, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_ZLibNG) + Compression.GetMaxCompressionLevel(PD_CE_ZLibNG)) \ 2
            TestCompressionEngine PD_CE_ZLibNG, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_ZLibNG)
            
            TestCompressionEngine PD_CE_Zstd, srcFilename, Compression.GetMinCompressionLevel(PD_CE_Zstd)
            TestCompressionEngine PD_CE_Zstd, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_Zstd) + Compression.GetMaxCompressionLevel(PD_CE_Zstd)) \ 2
            TestCompressionEngine PD_CE_Zstd, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_Zstd)
            
            TestCompressionEngine PD_CE_Lz4, srcFilename, Compression.GetMinCompressionLevel(PD_CE_Lz4)
            TestCompressionEngine PD_CE_Lz4, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_Lz4) + Compression.GetMaxCompressionLevel(PD_CE_Lz4)) \ 2
            TestCompressionEngine PD_CE_Lz4, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_Lz4)
            
            TestCompressionEngine PD_CE_Lz4HC, srcFilename, Compression.GetMinCompressionLevel(PD_CE_Lz4HC)
            TestCompressionEngine PD_CE_Lz4HC, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_Lz4HC) + Compression.GetMaxCompressionLevel(PD_CE_Lz4HC)) \ 2
            TestCompressionEngine PD_CE_Lz4HC, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_Lz4HC)
        
            TestCompressionEngine PD_CE_ZThunk, srcFilename, Compression.GetMinCompressionLevel(PD_CE_ZThunk)
            TestCompressionEngine PD_CE_ZThunk, srcFilename, (Compression.GetMinCompressionLevel(PD_CE_ZThunk) + Compression.GetMaxCompressionLevel(PD_CE_ZThunk)) \ 2
            TestCompressionEngine PD_CE_ZThunk, srcFilename, Compression.GetMaxCompressionLevel(PD_CE_ZThunk)
        
        'By default, let's just test the default compression level for each library
        Else
            TestCompressionEngine PD_CE_ZLib, srcFilename
            TestCompressionEngine PD_CE_ZLibNG, srcFilename
            TestCompressionEngine PD_CE_Zstd, srcFilename
            TestCompressionEngine PD_CE_Lz4, srcFilename
            TestCompressionEngine PD_CE_Lz4HC, srcFilename
            TestCompressionEngine PD_CE_ZThunk, srcFilename
        End If
        
    End If
    
    'If the user wants Windows libraries tested, do them last.  Note that these do not support different
    ' compression levels, so each test is only run once.
    If CBool(chkTestWindowsLibraries.Value) Then
        TestCompressionEngine PD_CE_MSZIP, srcFilename
        TestCompressionEngine PD_CE_XPRESS, srcFilename
        TestCompressionEngine PD_CE_XPRESS_HUFF, srcFilename
        TestCompressionEngine PD_CE_LZMS, srcFilename
    End If
    
    'With all tests complete, sort the output and display it to the window
    SortAndDisplayResults
    
End Sub

Private Sub SortAndDisplayResults()
    
    If (m_NumOfResults = 0) Then Exit Sub
    
    'This bunch of gibberish just prints out some nice headers for the testing data
    Dim columnSeparator As String, headerText As String
    columnSeparator = vbTab & "|" & Space$(3)
    headerText = "Engine" & vbTab & columnSeparator & "Comp. time" & columnSeparator & "Decomp. time" & columnSeparator & "Comp. ratio" & columnSeparator & "Ratio / Time" & columnSeparator & "Level tested" & vbCrLf & String$(150, "-")
    AddText vbCrLf & headerText
    
    'Sort the compression test results by the user's current sort choice
    Dim sortType As PD_SortResults
    sortType = cboSort.ListIndex
    
    Dim i As Long, j As Long
    For i = 0 To m_NumOfResults - 1
    For j = 0 To m_NumOfResults - 1
    
        'Sort based on the selected criteria
        If (sortType = sr_SortByName) Then
            If (StrComp(m_CompressionResults(i).CompressorName, m_CompressionResults(j).CompressorName, vbTextCompare) < 0) Then SwapCompResults i, j
        ElseIf (sortType = sr_SortByCompTime) Then
            If (m_CompressionResults(i).CompressionTime < m_CompressionResults(j).CompressionTime) Then SwapCompResults i, j
        ElseIf (sortType = sr_SortByDecompTime) Then
            If (m_CompressionResults(i).DecompressionTime < m_CompressionResults(j).DecompressionTime) Then SwapCompResults i, j
        ElseIf (sortType = sr_SortByCompRatio) Then
            If (m_CompressionResults(i).CompressionRatio > m_CompressionResults(j).CompressionRatio) Then SwapCompResults i, j
        ElseIf (sortType = sr_SortByRatioVsTime) Then
            If (m_CompressionResults(i).CompRatioVsTime > m_CompressionResults(j).CompRatioVsTime) Then SwapCompResults i, j
        End If
        
    Next j
    Next i
    
    'Post the finished list to the text box
    Dim compRatioTimeText As String
    For i = 0 To m_NumOfResults - 1
        
        '(Sorry for the complexity of this block; it's mostly just to format everything nicely before
        ' dumping it out to the text box.)
        With m_CompressionResults(i)
            
            compRatioTimeText = Format$(.CompRatioVsTime, "0.00")
            If (Len(compRatioTimeText) < 7) Then compRatioTimeText = compRatioTimeText & vbTab
            
            AddText .CompressorName & columnSeparator & MakeTimePretty(.CompressionTime) & columnSeparator & MakeTimePretty(.DecompressionTime) & columnSeparator & Format$(.CompressionRatio, "00.00") & "%" & columnSeparator & compRatioTimeText & columnSeparator & .EngineDetails
            
        End With
        
    Next i
    
    'After the test, move the cursor to the end of the results textbox
    txtResults.SelStart = Len(txtResults.Text)
    
End Sub

Private Sub SwapCompResults(ByVal resOneIndex As Long, ByRef resTwoIndex As Long)
    Dim tmpResult As PD_CompressionResult
    tmpResult = m_CompressionResults(resOneIndex)
    m_CompressionResults(resOneIndex) = m_CompressionResults(resTwoIndex)
    m_CompressionResults(resTwoIndex) = tmpResult
End Sub

Private Sub TestCompressionEngine(ByVal whichEngine As PD_CompressionEngine, ByRef srcFile As String, Optional ByVal compressionLevel As Long = -1)

    'Start by loading the source file into a byte array
    Dim fileNum As Integer
    Dim fileBytes() As Byte
    
    fileNum = FreeFile
    Open srcFile For Binary As #fileNum
        ReDim fileBytes(0 To LOF(fileNum) - 1) As Byte
        Get fileNum, , fileBytes
    Close #fileNum
    
    'Note the original size of the file - we'll use this to calculate a "compression ratio", or how well
    ' this compression engine is able to compress the original data.
    Dim originalSize As Long
    originalSize = UBound(fileBytes) + 1
        
    'We also need to declare a second array; this one receives the compressed bytes
    Dim compressedBytes() As Byte
    
    'To demonstrate that compression is perfectly lossless, we're going to decompress the data into a
    ' separate array, then compare the results to the original file.  If every byte matches, we'll know
    ' our functions are working properly!
    Dim testDecompressionBytes() As Byte
    
    'For testing purposes, we want to know how long compression and decompression takes.
    Dim startTime As Currency, totalCompressionTime As Currency, totalDecompressionTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Now, we test compression!  Note that you *must* pass the compression function an extra Long-type variable.
    ' This variable receives the final size of the compressed data.  Compression always starts with a large array
    ' (because we don't know in advance how small our data is going to be compressed), but it will only use some
    ' sub-portion of that array.  After compression finishes, this Long value tells you how many bytes were
    ' actually filled in the final array.
    Dim finalCompressedSize As Long
    
    'For simplicity, the main compression function asks for a source pointer.  I know this is anathema to most
    ' VB developers, but it greatly simplifies the compression interface.  If you have questions about how to
    ' get the pointer of a VB variable, feel free to ask in the original forum thread.
    If Compression.CompressPtrToDstArray(compressedBytes, finalCompressedSize, VarPtr(fileBytes(0)), originalSize, whichEngine, compressionLevel) Then
    
        'Note the time it took to compress the data
        totalCompressionTime = VBHacks.GetTimerDifferenceNow(startTime)
        
        'Start a new time measurement; this one is for decompression time
        VBHacks.GetHighResTime startTime
        
        'Decompress the data
        If Compression.DecompressPtrToDstArray(testDecompressionBytes, originalSize, VarPtr(compressedBytes(0)), finalCompressedSize, whichEngine) Then
        
            'Note how long it took to decompress the data
            totalDecompressionTime = VBHacks.GetTimerDifferenceNow(startTime)
            
            Dim columnSeparator As String
            columnSeparator = vbTab & "|" & vbTab
            
            'As a failsafe, make sure the decompressed data is a byte-for-byte match against the original data
            If VBHacks.MemCmp(VarPtr(fileBytes(0)), VarPtr(testDecompressionBytes(0)), originalSize) Then
                
                'Convert the compression level into something a bit more readable
                Dim cmpLevelText As String
                
                If (whichEngine = PD_CE_NoCompression) Or (whichEngine = PD_CE_MSZIP) Or (whichEngine = PD_CE_XPRESS) Or (whichEngine = PD_CE_XPRESS_HUFF) Or (whichEngine = PD_CE_LZMS) Then
                    cmpLevelText = "n/a"
                Else
                
                    If (compressionLevel = Compression.GetMinCompressionLevel(whichEngine)) Then
                        cmpLevelText = "Minimum (" & CStr(compressionLevel) & ")"
                    ElseIf (compressionLevel = Compression.GetMaxCompressionLevel(whichEngine)) Then
                        cmpLevelText = "Maximum (" & CStr(compressionLevel) & ")"
                    ElseIf (compressionLevel = ((Compression.GetMinCompressionLevel(whichEngine) + Compression.GetMaxCompressionLevel(whichEngine)) \ 2)) Then
                        cmpLevelText = "Middle (" & CStr(compressionLevel) & ")"
                    ElseIf (compressionLevel = -1) Or (compressionLevel = Compression.GetDefaultCompressionLevel(whichEngine)) Then
                        cmpLevelText = "Default (" & CStr(Compression.GetDefaultCompressionLevel(whichEngine)) & ")"
                    Else
                        cmpLevelText = CStr(compressionLevel)
                    End If
                    
                End If
                
                'Add these results to our compression table
                If (m_NumOfResults > UBound(m_CompressionResults)) Then ReDim Preserve m_CompressionResults(0 To m_NumOfResults * 2 - 1) As PD_CompressionResult
                
                With m_CompressionResults(m_NumOfResults)
                    .CompressorName = GetCompressorName(whichEngine)
                    .CompressionTime = totalCompressionTime
                    .DecompressionTime = totalDecompressionTime
                    .CompressionRatio = 100# * (1# - (finalCompressedSize / originalSize))
                    If (.CompressionTime <> 0#) Then
                        .CompRatioVsTime = .CompressionRatio / (.DecompressionTime + .CompressionTime)
                    Else
                        .CompRatioVsTime = 0#
                    End If
                    .EngineDetails = cmpLevelText
                End With
                
                m_NumOfResults = m_NumOfResults + 1
                
            Else
                AddText "WARNING! " & GetCompressorName(whichEngine) & " compression/decompression cycle was not lossless."
            End If
        
        Else
            AddText "WARNING! " & GetCompressorName(whichEngine) & " decompression test failed for unknown reasons."
        End If
    
    Else
        AddText "WARNING! " & GetCompressorName(whichEngine) & " compression test failed for unknown reasons."
    End If
    
    'After the test, move the cursor to the end of the results textbox
    txtResults.SelStart = Len(txtResults.Text)
    
End Sub

'This helper function just reports the name of a given compression engine.
Private Function GetCompressorName(ByVal whichEngine As PD_CompressionEngine) As String
    If (whichEngine = PD_CE_NoCompression) Then
        GetCompressorName = "None" & vbTab
    ElseIf (whichEngine = PD_CE_ZLib) Then
        GetCompressorName = "Zlib" & vbTab
    ElseIf (whichEngine = PD_CE_ZLibNG) Then
        GetCompressorName = "Zlib-ng" & vbTab
    ElseIf (whichEngine = PD_CE_Zstd) Then
        GetCompressorName = "Zstd" & vbTab
    ElseIf (whichEngine = PD_CE_Lz4) Then
        GetCompressorName = "Lz4" & vbTab
    ElseIf (whichEngine = PD_CE_Lz4HC) Then
        GetCompressorName = "Lz4_HC" & vbTab
    ElseIf (whichEngine = PD_CE_ZThunk) Then
        GetCompressorName = "ZipThunk" & vbTab
    ElseIf (whichEngine = PD_CE_MSZIP) Then
        GetCompressorName = "MSZIP" & vbTab
    ElseIf (whichEngine = PD_CE_XPRESS) Then
        GetCompressorName = "XPRESS" & vbTab
    ElseIf (whichEngine = PD_CE_XPRESS_HUFF) Then
        GetCompressorName = "XPRESS_HUFF"
    ElseIf (whichEngine = PD_CE_LZMS) Then
        GetCompressorName = "LZMS" & vbTab
    End If
End Function

'This helper function just formats a timing result in a human-readable way
Private Function MakeTimePretty(ByRef srcTime As Double) As String
    MakeTimePretty = Format$(srcTime * 1000, "#0000.00") & " ms"
End Function

Private Sub AddText(ByRef srcString As String)
    txtResults.Text = txtResults.Text & srcString & vbCrLf
End Sub

Private Sub txtResults_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        m_LastTestedFile = Data.Files(1)
        SetupTest m_LastTestedFile, m_LastTestedFile
    End If
End Sub
