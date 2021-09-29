VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdCompressZThunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'ZipArchive Project Thunks Interface
'Copyright 2018 by wqweto@gmail.com
'Created: 22/July/18
'***************************************************************************
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "pdCompressZThunk"
Implements ICompress

'=========================================================================
' API
'=========================================================================

'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for VirtualAlloc
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const MEM_COMMIT                    As Long = &H1000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, Optional ByVal Msg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

'=========================================================================
' Thunk data
'=========================================================================

' Auto-generated on 5.4.2018 16:38:12, CodeSize=6736, DataSize=984, ALIGN_SIZE=16
Private Const STR_THUNK1 As String = _
    "UYtEJAhTi1wkEFWLbCQYVleLeEQD64tEJCSJbCQQiwD2wwN0HDvddBQPthNDD7bIM9HB6AgzBJf2wwN16IlcJByL1cdEJBgAAAAAK9OD4vyNDBqDwgPB6gI7yxvJ99EjyolMJCB0eYvpjaQkAAAAAIsbi8vB6QgPtvGLyMHpCA+2yTPxi8vB6RAPttGLyMHpEIu0twAIAAAPtskz0YvIwekYM7SXAAQAAIvTweoYM9EPtsgPtsOLXCQcM8iDwwQzNJeJXCQcM7SPAAwAAItMJBiLxkGJTCQYO811lItsJBCLzTP2K8s76xvt99Uj6XQWD7YTjVsBD7bIRjPRwegIMwSXO/V16otMJCRfXl2JAVtZwhAAzMzMzMzMzMzMzMzMVot0JBCF9n41i1QkDE6LRCQIwe4ERoMCAYsKdQP/QgSJCItKBIlIBMdACAAAAADHQAwAAAAAg8AQg+4BddhewhAAzMzMzMzMzMzMzMzMzMyLVCQMhdJ0GotEJAhWi3Qk" & _
    "CCvwigwGjUABMEj/g+oBdfJewhAAzMzMzMzMzMzMzMyDfCQMAItEJASLSAyJTCQED46OAAAAU1VWi3QkFFeLfCQgi++D5QGD5wKNmwAAAACKHoXtdBWLSAiDyQKL0YPyAQ+v0cHqCDLTiBaF/3QCih4PthBGD7bLi1wkFDPRiwjB6QiLFJOLWAgz0Q+2ygNIBGnJBYQICIkQD7bTwesIQYlIBA+2SAcz0YtMJBQzHJGLTCQcSYlYCIlMJByFyX+QX15dW8IQAMzMzMzMzMzMzFaLdCQIV2oMi0Yo/9CL+FeJN+hpCQAAi04oahT/0cdAEAAAAADHQAwAAAAAiUcIi8dfXsIQAMzMzMzMzMzMzMxWi3QkCFeLPv92CItHMP/Q/3YEi0cw/9CLRzBW/9BfXsIQAMzMzMzMzMzMzMzMzMyLRCQMVot0JAxXi3wkDIXAdA1Q/3YE/zb/N+gh/f///3YcjUYM/3YY/3YU/3YQUI1GCFD/dgT/NlfoggkAAF9e" & _
    "whAAzMzMzMzMzMzMzMzMzP90JAToRxAAAMIQAMzMzMxWi3QkCItODI1GDFeLPoXJdAw7TgR0B1BW6OIRAACLThCNRhCFyXQMO04IdAdQVujMEQAAg34UAI1GFHQHUFbovBEAAI1GBFBW6LIRAACNRghQVuioEQAAi0cwVv/QX17CEADMzMzMzMzMzMzMzMzMi0QkCFNVi2wkDFZXjVgMU414CFf/cAT/MFXogQkAAItMJByL8IXJdBiF9nQUiweFwHQMUf8zUP91AOgx/P//i8ZfXl1bwhAAzMzMzMzMzMxWi3QkCIPK/4uGAIAEAA+3TMYCjQTGZjvKdAkPv8FmiRTG6xCLQASD+P90CGaJlEYEgAQAi4YAgAQAi0wkEIlMxgSLhgCABABmiVTGAo0UTouOAIAEAA+3ggSABABmiQTOi4YAgAQAD78Mxg+3wGaJggSABACD+f90DGaLhgCABABmiUTOAouOAIAEAIpEJAyIhA4AAAQAi4YAgAQAQCX/" & _
    "fwAAiYYAgAQAXsIMAMzMzIPsDFOLXCQYVVZXi3wkIDP2DxgLi38EiXwkGDm3CIAGAA+O3wAAAItsJCjrA41JAIuPCIAGAIvBK8YDxYP4BA+MmAAAADvxfQoPtoQ+BIAGAOsIi8YrwQ+2BBiNVgKIRCQUjUL/O8F9Cg+2hD4FgAYA6wmLxivBD7ZEGAGIRCQVO9F9Cg+2hD4GgAYA6wmLxivBD7ZEGAKIRCQWjUIBO8F9Cg+2hD4HgAYA6wmLxivBD7ZEGAOIRCQXi0QkFGnIsXk3nsHpEFFQV+iC/v//Rju3CIAGAA+MVf///+spi847twiABgB9H42XBIAGAI1JAIqEOQSABgCNUgGIQv9BO48IgAYAfOoptwiABgCLRCQohcAPjiUCAADrBo2bAAAAADPtg/gED4yLAQAAaQOxeTeewegQD7+URwSABACD+v8PhHEBAACLjwCABACLwivBvgCAAAAl/38AACvwK86B4f9/AACLhDkAAAQAOwN1Dom0" & _
    "rwyABgBFO2wkLH0JD78U14P6/3XAhe0PjikBAACLRCQoOUQkMH4GiUQkMOsEi0QkML4EAAAAg/0Bfm47xn5qkIoEHjPJM9uIRCQThe1+VpCLlI8MgAYAi8YrwnkYi4cAgAQAK8IDxiX/fwAAioQ4AAAEAOsLi3wkJIoEOIt8JBg4RCQTdQiJlJ8MgAYAQ0E7zXy9g/sBfg1Gi+uLXCQkO3QkMHybi1wkJItsJCiNTgOLlwyABgA7zX0wjZsAAAAAO8p9JouHAIAEACvCA8Yl/38AAIuEOAAABAAzRAv9dTGDwQSDxgQ7zXzWO/V9Q4vOjSwaK8qFyXkki4cAgAQAK8IDxiX/fwAAioQ4AAAEAOsPD7zAdBvB+AMD8OsUigQZOAQpdQhGQTt0JCh8yItsJChWUv90JCjo1A4AAIX2D459AAAA6ziLTCQgigM8j4sRi3EID7bIi0I8dwkPtkQIMGoI6w0PtgQIagmNBEUBAAAAUFZS6HcSAAC+AQAAAIts" & _
    "JCiD/QR8FmkDsXk3nsHoEFAPtgNQV+hF/P//6xWLjwiABgCKA4iEDwSABgD/hwiABgBOQ02F9n/JiWwkKIlcJCSLRCQohe0Pj+P9//9fXl1bg8QMwhQAzMzMzMzMzMzMg+wUU4tcJCBVVleLfCQoM/YPGAuLfwSJfCQgObcIgAYAD47fAAAAi2wkMOsDjUkAi48IgAYAi8ErxgPFg/gED4yYAAAAO/F9Cg+2hD4EgAYA6wiLxivBD7YEGI1WAohEJBiNQv87wX0KD7aEPgWABgDrCYvGK8EPtkQYAYhEJBk70X0KD7aEPgaABgDrCYvGK8EPtkQYAohEJBqNQgE7wX0KD7aEPgeABgDrCYvGK8EPtkQYA4hEJBuLRCQYacixeTeewekQUVBX6DL7//9GO7cIgAYAD4xV////6ymLzju3CIAGAH0fjZcEgAYAjUkAioQ5BIAGAI1SAYhC/0E7jwiABgB86im3CIAGADPAM8mJRCQUiUQkHItEJDCJTCQY" & _
    "hcAPjjQCAADrBo2bAAAAADPtg/gED4ztAQAAaQOxeTeewegQD7+URwSABACD+v8PhNMBAACLjwCABACLwivBvgCAAAAl/38AACvwK86B4f9/AACLhDkAAAQAOwN1Dom0rwyABgBFO2wkNH0JD78U14P6/3XAhe0PjocBAACLRCQwOUQkOH4GiUQkOOsEi0QkOL4EAAAAg/0Bfm47xn5qkIoEHjPJM9uIRCQThe1+VpCLlI8MgAYAi8YrwnkYi4cAgAQAK8IDxiX/fwAAioQ4AAAEAOsLi3wkLIoEOIt8JCA4RCQTdQiJlJ8MgAYAQ0E7zXy9g/sBfg1Gi+uLXCQsO3QkOHybi1wkLItsJDCNTgOLlwyABgA7zX0wjZsAAAAAO8p9JouHAIAEACvCA8Yl/38AAIuEOAAABAAzRAv9dTGDwQSDxgQ7zXzWO/V9P4vOjSwaK8qFyXkki4cAgAQAK8IDxiX/fwAAioQ4AAAEAOsPD7zAdBfB+AMD8OsQigQZ" & _
    "OAQpdQhGQTt0JDB8yItsJBSF7X5ejUUBO/B+T4tEJCiLEItoCItEJBw8jw+2yItCPHcJD7ZECDBqCOsND7YECGoJjQRFAQAAAFBVUugfDwAAD7YDi48MgAYAiUwkGIl0JBSJRCQc6Y4AAACLVCQYVVLrHw+2A4lUJBiJdCQUiUQkHOt1i0wkGItsJBSF7X42VVH/dCQw6PUKAAAzwI11/4lEJBSF9n9Wi2wkMItEJDCLTCQYhe0Pj9T9//9fXl1bg8QUwhQAi0wkKIoDPI+LEYtxCA+2yItCPHcJD7ZECDBqCOsND7YECGoJjQRFAQAAAFBWUuh3DgAAvgEAAACLbCQwg/0EfBZpA7F5N57B6BBQD7YDUFfoRfj//+sVi48IgAYAigOIhA8EgAYA/4cIgAYATkNNhfZ/yYlsJDCJXCQs6Wb////MzMzMzMzMzMzMVleLfCQMaKyPBgCLB4tAKP/Qi/CF9nUFX17CBACJdwSNTgK6AIAAAIPI/+sDjUkA" & _
    "x0EC/////41JCGaJQfhmiUH2g+oBdemNvgSABAC5AIAAAPOrX4mWAIAEAI1CAYmWCIAGAF7CBADMzMzMzMzMzFZXi3wkDLkSAAAAvgBQeG248Cl4bfOli0wkEIPAYF9exwEAEHhtiUEEx0EIoEB4bcdBDHhEeG3CCADMzMzMzMxTi1wkCFZXaACAAACLO4tzCItHKP/QiQYzwDlEJCRqAw+VwMdGCACAAACDwALHRgQAAAAAUFZX6DQNAACDfCQoAP90JDD/dCQw/3QkIP90JCBTdAfot/f//+sF6AD7//9qB2oAVlfoBQ0AAIN8JCQAdBiLThCFyXQRuAgAAAArwVBqAFZX6OYMAACLRCQciw5fiQiLRCQci04EXluJCLgBAAAAwiQAzMzMzMzMzMzMzFFTVVaLdCQUV2gAgAAAix6JXCQci0Mo/9CLbCQgiYakgQAAx4asgQAAAIAAAMeGqIEAAAAAAACF7X8Ng76cAQAAAA+OtgUAAIO+nAEAABiN" & _
    "vpwBAAB9KotUJBzrA41JAIXtfhkPtgJNiw9C0+AJhpgBAACNQQiJB4P4GHzjiVQkHItGGIXAdQnHRhgBAAAA66aD+AF1c4sXg/oDD4xdBQAAi4aYAQAAg8L90eiLyIkXwegCiYaYAQAAg+EDdR2LysdGGAoAAACD4Qcr0dPoiReJhpgBAADpXP///4P5AXUYi0YEiUYMi0YIiUYQx0YYBgAAAOk/////g/kCD4U2////iU4Y6S7///+D+AJ1bosXg/oOD4zlBAAAi46YAQAAi8GD4B/B6QUFAQEAAMdGLAAAAACJRiCLwYPgH8HpBUDHRhgDAAAAiUYki8GD4A/B6QSDwASJjpgBAACJRiiNQvKJBzPAiUZEiUZIiUZMiUZQZolGVIhGVum7/v//g/gDdW85Bw+MdQQAAItGLDtGKH0xgz8DfCyLS0CLRiyKlpgBAACA4gcPtgQBiFQwRP9GLItGLIMH/cGumAEAAAM7Rih8z4tGLDtGKA+Faf7//2oT"
Private Const STR_THUNK2 As String = _
    "jUZEUFboxgkAAIlGFMdGGAQAAADHRiwAAAAA6Uf+//+D+AQPheIAAACLRiSLTiADwTlGLHw/UY1GV1BW6I8JAAD/diSJRgyLRiCDwFcDxlBW6HoJAACNfhSJRhBXVui9BQAAxwcAAAAAx0YYBgAAAOny/f///3YUjYaYAQAAV1DoGwYAAIvIg/n/D4SeAwAAg/n+D4RiAwAAg/kQfQ+LRiyITDBX/0Ys6br9//91B7gCAAAA6w8zwIP5EQ+VwI0EhQMAAACJRjAzwIP5Eg+UwI0ExQMAAACJRjSD+RB1G4tGLIXAfhQPtkQwVolGPMdGGAUAAADpbf3//zPAx0YYBQAAAIlGPOlc/f//g/gFdVyLH4tOMDvZD4wRAwAAi4aYAQAAugEAAADT4ivZSokfI9DT6ANWNImGmAEAAIXSfhyLRiQDRiCLTiw7yH0PikY8SohEMVf/RiyF0n/ki1wkGMdGGAQAAADp+/z//4P4Bg+FnAAAAP92DI1eDFeNhpgB" & _
    "AABQ6BgFAACD+P8PhJ0CAACD+P4PhF0CAAA9AAEAAH0QUFboCAQAAItcJBjptvz//3VAiwPHRhgBAAAAO0YEdA1TVuhYBAAAxwMAAAAAi0YQjX4Qi1wkGDtGCA+Ehvz//1dW6DgEAADHBwAAAADpdPz//4tcJBg9HgEAAA+NZfz//8dGGAcAAACJRhzpVvz//4P4B3VUi0YcLQEBAACNDECLQzSNHIgPv0sCOQ8PjPsBAACLlpgBAAC4AQAAANPgSCPCA0MEiUY4D79DAikHiksCi1wkGNPqiZaYAQAAx0YYCAAAAOn9+///g/gIdTr/dhCNhpgBAABXUOghBAAAg/j/D4SmAQAAg/j+D4RqAQAAg/geD41hAQAAx0YYCQAAAIlGHOm++///g/gJD4WPAAAAi0YcixeNDECLQziNBIgPv0gCiUQkIIlMJBA70Q+MWgEAAIuGmAEAALsBAAAAK1QkENPji0wkIEsj2ANZBIkXikkC0+iDfjgAiYaYAQAA" & _
    "x0YYBgAAAHQsjaQkAAAAAIuGoIEAAP9OOCvDJf9/AAAPtoQwoAEAAFBW6IECAACDfjgAddv/TjiLXCQY6Sb7//+D+Ap1MYsXg/oQD4zdAAAAi46YAQAAD7fBiUZAjULwwekQiQeJjpgBAADHRhgLAAAA6fD6//+D+At1P4sHg/gQD4ynAAAAi46YAQAAg8DwD7fRiQeB8v//AACLRkDB6RCJjpgBAAA7wnVQ99gbwIPgC0CJRhjprPr//4P4DA+Fo/r//4M/CHxlD7aGmAEAAFBW6NkBAACDB/jBrpgBAAAIg0ZA/w+FfPr//8dGGAEAAADpcPr//4tcJBj/tqSBAACLQzD/0ItEJCTHhqSBAAAAAAAAX17HAAAAAACLRCQgXVvHAAAAAAAzwFnCFACLTCQkuAEAAACLlqSBAABfiRGLlqiBAACLTCQkXl1biRFZwhQAzMzMzMzMzMzMzMzMzIHsIAEAAFNWi7QkLAEAAFdosIEAAItGKP/Qi9iNfCQM" & _
    "uAgICAi5JAAAAGggAQAAiTPzq7gJCQkJjbwkoAAAALkcAAAA86uNRCQQx4QkEAEAAAcHBwdQU8eEJBwBAAAHBwcHx4QkIAEAAAcHBwfHhCQkAQAABwcHB8eEJCgBAAAHBwcHx4QkLAEAAAcHBwfHhCQwAQAACAgICMeEJDQBAAAICAgI6McEAACJQwSNRCQMaiBQU8dEJBgFBQUFx0QkHAUFBQXHRCQgBQUFBcdEJCQFBQUFx0QkKAUFBQXHRCQsBQUFBcdEJDAFBQUFx0QkNAUFBQXodwQAAF+JQwiLw17HQxgAAAAAx0MUAAAAAMdDEAAAAADHQwwAAAAAx4OYAQAAAAAAAMeDnAEAAAAAAADHg6CBAAAAAAAAW4HEIAEAAMIEAMzMzMzMzMzMzMzMzFOLXCQMVot0JAyLhqCBAACLDoicMKABAACLhqCBAABAJf9/AACJhqCBAACLhqiBAAA7hqyBAAB8GgPAUP+2pIEAAImGrIEAAItBLP/QiYak" & _
    "gQAAi4aogQAAi46kgQAAiBwB/4aogQAAXlvCCADMzMxTi1wkCFWLbCQQiwOJRCQMhe11CF2DyP9bwggAV4t9AIX/dEtWM/Y5N3wj6wONSQCLRwSNBPCDwASDOAB0B1BT6Lv///9GOzd+5otEJBT/dwSLQDD/0ItEJBTHRwQAAAAAV4tAMP/Qx0UAAAAAAF5fXTPAW8IIAMzMzMzMzMzMzItEJAxTi1wkDFaLCItABIszV4t8JBCLFyPKjQTID7YIO85/Jg+2CNPqK/Fmg3gC/3Usi0AEhcB0GosIi0AEI8qNBMgPtgg7zn7aX16DyP9bwgwAX164/v///1vCDACJF4kzD79AAl9eW8IMAMzMzMyD7AiLRCQMVYsoi0AIiUQkEItEJBiJbCQEhcAPji0BAABTVlc9BAEAAH4HvgIBAADrDIvwPQIBAAB+A41w/SvGux0AAACJRCQkg8//i0U0i+iJRCQUjQQ7mSvC0fiNDEA7dI0EfQSL2OvrO3SNCH4E" & _
    "i/jr4YtcJByNBECNPIUAAAAAuRcBAAAPtwQvA/2LbCQQZjvBD7/Ii0U8fwwPtoRIAP7//2oH6wcPtkQBqGoIUFNV6CkDAAAPt0cCZoXAdA0rdwSYUFZTVegTAwAAi0U4g87/i1wkIL8eAAAAi+iL/40EN5krwtH4jQxAO1yNBH0Ei/jr6ztcjQh+BIvw6+GLfCQcjQRAjTSFAAAAAAP1i2wkEGoFD78Oi0U8D7YEyFBXVei5AgAAD7dGAmaFwHQPmFCLwytGBFBXVeihAgAAi0QkJIXAD4/Z/v//X15bXYPECMIMAMzMzMzMzMzMzMzMU1WLbCQMVldqCItFKP/Qi1wkKIv4i8u4AQAAANPgvgEAAACLTCQsSNPmiUQkKI0E9QAAAABQi0Uo/9CJRwQzyY1G/4PK/4kHhcB4I41kJACLRwRmiVTIAotHBMYEyACLRwTHRMgEAAAAAEE7D37hM+05bCQgfmmLdCQc6wONSQAPtgQuO8N+UYtEJBiLBKiL" & _
    "yCNMJCg7TCQkdT6Ly9P4Iwc7B380i08EZolswQKLTwQPthQuK9ONNMEPtg47yn0CiBaLdCQcugEAAAAPtgwuK8vT4gPCOwd+zEU7bCQgfKAz9jk3fGeLVCQsi2wkHIv/i0cEjQzwD7YBO8J+S4PI/2aJQQKLRwSNDPAPtgErwoP4B34FuAcAAABQjQQTiBFQi8uLxtPgC0QkLFD/dCQsVf90JCz/dCQs6MP+//+LTwSLVCQsiUTxBEY7N36ji8dfXl1bwhwAzMzMzMzMzMzMzIHsBAUAADPAM9JVi6wkFAUAAIlUJASJRCQMiUQkEIlEJBSJRCQYiUQkHIlEJCCJRCQkiUQkKIlEJCyJRCQwiUQkNIlEJDiJRCQ8iUQkQIlEJERWi7QkFAUAAIXtfhmL/w+2DDD/RIwMO9F9AovRQDvFfO2JVCQIM9IzyesDjUkAiVQMUANUDBCDwQQD0oP5PHzuVzP/he1+SVONnCSUAAAAD7Y0N8cDAAAAAItUtFSN" & _
    "QgGJRLRUhfZ0GDPAjWQkAIvKA8CD4QHR+gvBg+4BdfCJA4u0JBwFAABHg8MEO/18wFuLRCQMX4P4CXwFuAkAAABQagBqAFX/tCQkBQAAjYQkoAAAAFCLhCQoBQAA/zDojP3//15dgcQEBQAAwgwAzItEJAxWi3QkDItOENPgCUYMi0QkFAPBiUYQg/gIfDlXi3wkDItGCDlGBHwPA8BQ/zaJRgiLRyz/0IkGi1YEiw6KRgyIBAr/RgTBbgwIg0YQ+IN+EAh9zV9ewhAAVYvsiwCAQMAgoGDgEJBQ0DCwcPAIiEjIKKho6BiYWNg4uHj4BIRExCSkZOQUlFTUNLR09AyMTMwsrGzsHJxc3Dy8fPwCgkLCIqJi4hKSUtIysnLyCopKyiqqauoamlraOrp6+gaGRsYmpmbmFpZW1ja2dvYOjk7OLq5u7h6eXt4+vn7+AYFBwSGhYeERkVHRMbFx8QmJSckpqWnpGZlZ2Tm5efkFhUXFJaVl5RWVVdU1tXX1" & _
    "DY1NzS2tbe0dnV3dPb19/QODQ8Mjo2PjE5NT0zOzc/MLi0vLK6tr6xubW9s7u3v7B4dHxyenZ+cXl1fXN7d39w+PT88vr2/vH59f3z+/f/8BAQAAAwAAAAMAAAACAQAABAAAAAQAAAADAQAABQAAAAUAAAAEAQAABgAAAAYAAAAFAQAABwAAAAcAAAAGAQAACAAAAAgAAAAHAQAACQAAAAkAAAAIAQAACgAAAAoAAAAJAQEACwAAAAwAAAAKAQEADQAAAA4AAAALAQEADwAAABAAAAAMAQEAEQAAABIAAAANAQIAEwAAABYAAAAOAQIAFwAAABoAAAAPAQIAGwAAAB4AAAAQAQIAHwAAACIAAAARAQMAIwAAACoAAAASAQMAKwAAADIAAAATAQMAMwAAADoAAAAUAQMAOwAAAEIAAAAVAQQAQwAAAFIAAAAWAQQAUwAAAGIAAAAXAQQAYwAAAHIAAAAYAQQAcwAAAIIAAAAZAQUAgwAAAKIAAAAaAQUA" & _
    "owAAAMIAAAAbAQUAwwAAAOIAAAAcAQUA4wAAAAEBAAAdAQAAAgEAAAIBAAAQERIACAcJBgoFCwQMAw0CDgEPAAAAAAABAAAAAQAAAAEAAAACAAAAAgAAAAIAAAADAAAAAwAAAAMAAAAEAAAABAAAAAQAAQAFAAAABgAAAAUAAQAHAAAACAAAAAYAAgAJAAAADAAAAAcAAgANAAAAEAAAAAgAAwARAAAAGAAAAAkAAwAZAAAAIAAAAAoABAAhAAAAMAAAAAsABAAxAAAAQAAAAAwABQBBAAAAYAAAAA0ABQBhAAAAgAAAAA4ABgCBAAAAwAAAAA8ABgDBAAAAAAEAABAABwABAQAAgAEAABEABwCBAQAAAAIAABIACAABAgAAAAMAABMACAABAwAAAAQAABQACQABBAAAAAYAABUACQABBgAAAAgAABYACgABCAAAAAwAABcACgABDAAAABAAABgACwABEAAAABgAABkACwABGAAAACAAABoADAABIAAA"
Private Const STR_THUNK3 As String = _
    "ADAAABsADAABMAAAAEAAABwADQABQAAAAGAAAB0ADQABYAAAAIAAAA=="
Private Const STR_THUNK_OFFSETS As String = "592|656|704|784|800|912|0|288|368|416|0|0|0|6992|7360|6736|7340"
Private Const STR_THUNK_BUILDDATE As String = "5.4.2018 16:38:12"
' end of generated code

'Private Const STR_THUNK3 As String = ""

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_uRtbl                 As UcsZlibRelocTableType

Private Type UcsZlibRelocTableType
    CompressInit        As Long
    CompressCleanup     As Long
    CompressBlock       As Long
    DecompressInit      As Long
    DecompressCleanup   As Long
    DecompressBlock     As Long
    CalcCrc32           As Long
    MemNonce            As Long
    MemXor              As Long
    ZipCrypt            As Long
    MallocImpl          As Long
    ReallocImpl         As Long
    FreeImpl            As Long
    LenCodes            As Long
    DistCodes           As Long
    MirrorBytes         As Long
    LenLenMap           As Long
    Crc32Table          As Long
End Type

Private Type UcsZlibBuffersType
    InBlock             As Long
    InLen               As Long
    OutBlock            As Long
    OutLen              As Long
    Final               As Long
    Greedy              As Long
    MaxMatch            As Long
    NiceLen             As Long
End Type

Private Enum UcsRelocIndexesEnum
    ucsIdx_CompressInit = 0
    ucsIdx_CompressCleanup
    ucsIdx_CompressBlock
    ucsIdx_DecompressInit
    ucsIdx_DecompressCleanup
    ucsIdx_DecompressBlock
    ucsIdx_CalcCrc32
    ucsIdx_MemNonce
    ucsIdx_MemXor
    ucsIdx_ZipCrypt
    ucsIdx_MallocImpl
    ucsIdx_ReallocImpl
    ucsIdx_FreeImpl
    ucsIdx_LenCodes
    ucsIdx_DistCodes
    ucsIdx_MirrorBytes
    ucsIdx_LenLenMap
End Enum

'=========================================================================
' Methods
'=========================================================================

'= zlib thunks ===========================================================

Private Function pvInitRelocTable(uRtbl As UcsZlibRelocTableType) As Long
    Dim lpThunk         As Long
    Dim vSplit          As Variant
    
    lpThunk = pvGetThunkAddress()
    vSplit = Split(STR_THUNK_OFFSETS, "|")
    With uRtbl
        .CompressInit = lpThunk + vSplit(ucsIdx_CompressInit)
        .CompressCleanup = lpThunk + vSplit(ucsIdx_CompressCleanup)
        .CompressBlock = lpThunk + vSplit(ucsIdx_CompressBlock)
        .DecompressInit = lpThunk + vSplit(ucsIdx_DecompressInit)
        .DecompressCleanup = lpThunk + vSplit(ucsIdx_DecompressCleanup)
        .DecompressBlock = lpThunk + vSplit(ucsIdx_DecompressBlock)
        .CalcCrc32 = lpThunk + vSplit(ucsIdx_CalcCrc32)
        .MemNonce = lpThunk + vSplit(ucsIdx_MemNonce)
        .MemXor = lpThunk + vSplit(ucsIdx_MemXor)
        .ZipCrypt = lpThunk + vSplit(ucsIdx_ZipCrypt)
        .MallocImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemAlloc")
        .ReallocImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemRealloc")
        .FreeImpl = GetProcAddress(GetModuleHandle("ole32.dll"), "CoTaskMemFree")
        .LenCodes = lpThunk + vSplit(ucsIdx_LenCodes)
        .DistCodes = lpThunk + vSplit(ucsIdx_DistCodes)
        .MirrorBytes = lpThunk + vSplit(ucsIdx_MirrorBytes)
        .LenLenMap = lpThunk + vSplit(ucsIdx_LenLenMap)
'        .Crc32Table = pvGetCrc32Table()
    End With
End Function

Private Function pvGetThunkAddress() As Long
    Static lpThunk      As Long
    Dim baThunk()       As Byte
    
    If lpThunk = 0 Then
        baThunk = FromBase64Array(STR_THUNK1 & STR_THUNK2 & STR_THUNK3)
        lpThunk = VirtualAlloc(0, UBound(baThunk) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CopyMemory(ByVal lpThunk, baThunk(0), UBound(baThunk) + 1)
    End If
    pvGetThunkAddress = lpThunk
End Function

Private Function FromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim dwDummy         As Long
    Dim baOutput()      As Byte
    
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, 0, lSize, 0, dwDummy)
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, dwDummy)
    FromBase64Array = baOutput
End Function

Private Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    At = vArray(lIdx)
QH:
End Function

'=========================================================================
' ICompress interface
'=========================================================================

Private Function ICompress_InitializeEngine(pathToDLLFolder As String) As Boolean
    pvInitRelocTable m_uRtbl
    '--- success
    ICompress_InitializeEngine = True
End Function

Private Sub ICompress_ReleaseEngine()

End Sub

Private Function ICompress_CompressPtrToDstArray( _
            dstArray() As Byte, dstCompressedSizeInBytes As Long, _
            ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, _
            Optional ByVal compressionLevel As Long = -1&, _
            Optional ByVal dstArrayIsAlreadySized As Boolean = False, _
            Optional ByVal trimCompressedArray As Boolean = False) As Boolean
    'Prep the destination array, as necessary
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = ICompress_GetWorstCaseSize(constSrcSizeInBytes)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    '--- do compress
    If Not ICompress_CompressPtrToPtr(VarPtr(dstArray(0)), dstCompressedSizeInBytes, _
            constSrcPtr, constSrcSizeInBytes, compressionLevel) Then
        GoTo QH
    End If
    'Trim the destination array, as requested
    If trimCompressedArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then
            ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
        End If
    End If
    '--- success
    ICompress_CompressPtrToDstArray = True
QH:
End Function

Private Function ICompress_CompressPtrToPtr( _
            ByVal constDstPtr As Long, dstSizeInBytes As Long, _
            ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, _
            Optional ByVal compressionLevel As Long = -1&) As Boolean
    Const FUNC_NAME     As String = "ICompress_CompressPtrToPtr"
    Dim hCompressCtx    As Long
    Dim uBuf            As UcsZlibBuffersType
    
    hCompressCtx = CallWindowProc(m_uRtbl.CompressInit, VarPtr(m_uRtbl))
    If hCompressCtx = 0 Then
        GoTo QH
    End If
    '-- setup compress params
    If compressionLevel < 0 Then
        compressionLevel = ICompress_GetDefaultCompressionLevel()
    End If
    uBuf.Greedy = (compressionLevel <= 4)
    uBuf.MaxMatch = At(Array(0, 2, 6, 12, 24, 8, 16, 32, 64, 1000), compressionLevel)
    uBuf.NiceLen = At(Array(0, 8, 10, 14, 24, 30, 65, 130, 200, 32768), compressionLevel)
    '--- setup input buffer & size
    uBuf.InBlock = constSrcPtr
    uBuf.InLen = constSrcSizeInBytes
    uBuf.Final = 1
    Call CallWindowProc(m_uRtbl.CompressBlock, hCompressCtx, VarPtr(uBuf))
    If uBuf.OutBlock = 0 Then
        Debug.Print "CompressBlock failed [" & STR_MODULE_NAME & "." & FUNC_NAME & "]", Timer
        GoTo QH
    End If
    If uBuf.OutLen > dstSizeInBytes Then
        Debug.Print "Will trim compression output, dstSizeInBytes=" & dstSizeInBytes & _
            ", uBuf.OutLen=" & uBuf.OutLen & " [" & STR_MODULE_NAME & "." & FUNC_NAME & "]", Timer
        uBuf.OutLen = dstSizeInBytes
    End If
    '--- copy & release output buffer
    dstSizeInBytes = uBuf.OutLen
    Call CopyMemory(ByVal constDstPtr, ByVal uBuf.OutBlock, dstSizeInBytes)
    Call CoTaskMemFree(uBuf.OutBlock)
    '--- success
    ICompress_CompressPtrToPtr = True
QH:
    If hCompressCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.CompressCleanup, hCompressCtx)
    End If
End Function

Private Function ICompress_DecompressPtrToDstArray( _
            dstArray() As Byte, ByVal constDstSizeInBytes As Long, _
            ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, _
            Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    If (Not dstArrayIsAlreadySized) Then
        ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    End If
    If Not ICompress_DecompressPtrToPtr(VarPtr(dstArray(0)), constDstSizeInBytes, _
            constSrcPtr, constSrcSizeInBytes) Then
        GoTo QH
    End If
    '--- success
    ICompress_DecompressPtrToDstArray = True
QH:
End Function

Private Function ICompress_DecompressPtrToPtr( _
            ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, _
            ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Boolean
    Const FUNC_NAME     As String = "ICompress_DecompressPtrToPtr"
    Dim hDecompressCtx  As Long
    Dim uBuf            As UcsZlibBuffersType
    
    hDecompressCtx = CallWindowProc(m_uRtbl.DecompressInit, VarPtr(m_uRtbl))
    If hDecompressCtx = 0 Then
        Debug.Print "DecompressInit failed [" & STR_MODULE_NAME & "." & FUNC_NAME & "]", Timer
        GoTo QH
    End If
    uBuf.InBlock = constSrcPtr
    uBuf.InLen = constSrcSizeInBytes
    Call CallWindowProc(m_uRtbl.DecompressBlock, hDecompressCtx, VarPtr(uBuf), 0)
    If uBuf.OutBlock = 0 Then
        Debug.Print "DecompressBlock failed [" & STR_MODULE_NAME & "." & FUNC_NAME & "]", Timer
        GoTo QH
    End If
    If uBuf.OutLen > constDstSizeInBytes Then
        Debug.Print "Will trim decompress output, constDstSizeInBytes=" & constDstSizeInBytes & _
            ", uBuf.OutLen=" & uBuf.OutLen & " [" & STR_MODULE_NAME & "." & FUNC_NAME & "]", Timer
        uBuf.OutLen = constDstSizeInBytes
    End If
    Call CopyMemory(ByVal constDstPtr, ByVal uBuf.OutBlock, uBuf.OutLen)
    Call CoTaskMemFree(uBuf.OutBlock)
    '--- success
    ICompress_DecompressPtrToPtr = True
QH:
    If hDecompressCtx <> 0 Then
        Call CallWindowProc(m_uRtbl.DecompressCleanup, hDecompressCtx)
    End If
End Function

Private Function ICompress_GetCompressorName() As String
    ICompress_GetCompressorName = "ZipArchive Thunks " & STR_THUNK_BUILDDATE
End Function

Private Function ICompress_GetDefaultCompressionLevel() As Long
    ICompress_GetDefaultCompressionLevel = 6
End Function

Private Function ICompress_GetMaxCompressionLevel() As Long
    ICompress_GetMaxCompressionLevel = 9
End Function

Private Function ICompress_GetMinCompressionLevel() As Long
    ICompress_GetMinCompressionLevel = 1
End Function

Private Function ICompress_GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long) As Long
    ICompress_GetWorstCaseSize = srcBufferSizeInBytes * 2
End Function

Private Function ICompress_IsCompressorReady() As Boolean
    ICompress_IsCompressorReady = True
End Function

