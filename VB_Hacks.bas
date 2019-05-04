Attribute VB_Name = "VBHacks"
'***************************************************************************
'Miscellaneous VB6 Hacks
'Copyright 2016 by Tanner Helland
'Created: 06/January/16
'Last updated: 10/December/16
'Last update: throw together some hacks to make the standalone "Compression" module easier to use.
'
'PhotoDemon relies on a lot of "not officially sanctioned" VB6 behavior to enable various optimizations and C-style
' code techniques. If a function's primary purpose is a VB6-specific workaround, I prefer to move it here, so I
' don't clutter up purposeful modules with obscure, VB-specific hackery.
'
'Note that some code here may seem redundant (e.g. identical functions suffixed by data type, instead of declared
' "As Any") but that's by design, to improve safety since these techniques are crash-prone if used incorrectly or
' imprecisely.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long

Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDestPtr As Long, ByVal lpvSourcePtr As Long, ByVal cbCopy As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByVal lpVersionInformation As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal Addr As Long, ByRef dstValue As Long) As Long

Private Declare Function RtlCompareMemory Lib "ntdll" (ByVal ptrSource1 As Long, ByVal ptrSource2 As Long, ByVal Length As Long) As Long

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As String

'Higher-performance timing functions are also handled by this class.  Note that you *must* initialize the timer engine
' before requesting any time values, or crashes will occurs because the frequency timer is 0.
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private m_TimerFrequency As Currency

'OS-level compression APIs are only available on Win 8 or later; we now check for this automatically
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

'Check array initialization.  All array types supported.  Thank you to http://stackoverflow.com/questions/183353/how-do-i-determine-if-an-array-is-initialized-in-vb6
Public Function IsArrayInitialized(arr) As Boolean
    Dim saAddress As Long
    GetMem4 VarPtr(arr) + 8, saAddress
    GetMem4 saAddress, saAddress
    IsArrayInitialized = (saAddress <> 0)
    If IsArrayInitialized Then IsArrayInitialized = UBound(arr) >= LBound(arr)
End Function

Public Sub EnableHighResolutionTimers()
    QueryPerformanceFrequency m_TimerFrequency
    If (m_TimerFrequency = 0) Then m_TimerFrequency = 1
End Sub

Public Function GetTimerDifference(ByRef startTime As Currency, ByRef stopTime As Currency) As Double
    GetTimerDifference = (stopTime - startTime) / m_TimerFrequency
End Function

Public Function GetTimerDifferenceNow(ByRef startTime As Currency) As Double
    Dim tmpTime As Currency
    QueryPerformanceCounter tmpTime
    GetTimerDifferenceNow = (tmpTime - startTime) / m_TimerFrequency
End Function

Public Sub GetHighResTime(ByRef dstTime As Currency)
    QueryPerformanceCounter dstTime
End Sub

Public Function MemCmp(ByVal ptr1 As Long, ByVal ptr2 As Long, ByVal bytesToCompare As Long, Optional ByRef dstPosMismatch As Long) As Boolean
    Dim bytesEqual As Long
    bytesEqual = RtlCompareMemory(ptr1, ptr2, bytesToCompare)
    MemCmp = CBool(bytesEqual = bytesToCompare)
    If (Not MemCmp) Then dstPosMismatch = bytesEqual + 1 Else dstPosMismatch = -1
End Function

'Given an arbitrary pointer to a null-terminated CHAR or WCHAR run, measure the resulting string and copy the results
' into a VB string.
'
'For security reasons, if an upper limit of the string's length is known in advance (e.g. MAX_PATH), pass that limit
' via the optional maxLength parameter to avoid a buffer overrun.  This function has a hard-coded limit of 65k chars,
' a limit you can easily lift but which makes sense for PD.  If a string exceeds the limit (whether passed or
' hard-coded), *a string will still be created and returned*, but it will be clamped to the maximum length.
Public Function ConvertCharPointerToVBString(ByVal srcPointer As Long, Optional ByVal stringIsUnicode As Boolean = True, Optional ByVal maxLength As Long = -1) As String
    
    'Check string length
    Dim strLength As Long
    If stringIsUnicode Then strLength = lstrlenW(srcPointer) Else strLength = lstrlenA(srcPointer)
    
    'Make sure the length/pointer isn't null
    If (strLength <= 0) Then
        ConvertCharPointerToVBString = ""
        Exit Function
    End If
    
    'Make sure the string's length is valid.  A magic number of 65k is used for the purposes of this demo.
    Dim maxAllowedLength As Long
    If (maxLength = -1) Then maxAllowedLength = 65535 Else maxAllowedLength = maxLength
    If (strLength > maxAllowedLength) Then strLength = maxAllowedLength
    
    'Create the target string and copy the bytes over
    If stringIsUnicode Then
        ConvertCharPointerToVBString = String$(strLength, 0)
        CopyMemoryStrict StrPtr(ConvertCharPointerToVBString), srcPointer, strLength * 2
    Else
        ConvertCharPointerToVBString = SysAllocStringByteLen(srcPointer, strLength)
    End If
    
End Function

'Many places in PD need to know the current Windows version, so they can enable/disable features accordingly.  To avoid
' constantly retrieving that info via APIs, we retrieve it once - at first request - then cache it locally.
Private Sub CacheOSVersion()
    If (Not m_VersionInfoCached) Then
        m_OSVI.dwOSVersionInfoSize = Len(m_OSVI)
        GetVersionEx VarPtr(m_OSVI)
        m_VersionInfoCached = True
    End If
End Sub

Public Function IsWin8OrLater() As Boolean
    CacheOSVersion
    IsWin8OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 2))
End Function
