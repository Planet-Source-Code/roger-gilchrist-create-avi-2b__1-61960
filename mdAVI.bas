Attribute VB_Name = "mdAVI"
Option Explicit
'this is original to 'Create AVI' by Ron Hoebe
'all I have done is a bit of Code Fixer tidying up and removed unused bits
'Ron's comments
'note! - the ppfile argument is ByRef because it is a pointer to a pointer :-)
'Careful! this function is awkward to use in VB
' the only way to make it work is to pass a pointer to a AVI_COMPRESS_OPTIONS UDT (last parameter) ByRef
' This means in your code you should Dim a long variable and get a pointer to your AVI_COMPRESS_OPTIONS UDT
' using VarPtr() - e.g. mylong = VarPtr(myUDT) - and then pass the mylong BYREF
' this will give you a pointer to a pointer to an (array of) UDT (yech!) -Ray
'This is actually the AVISaveV function aliased to be called as AVISave from VB because
'AVISave seems to be compiled using CDECL calling convention ;-(
'ALSO - see note above AVISaveOptions declare - this function also requires a pointer to a pointer to (an array of) UDT
'See note above AVISaveOptions declare - this function also requires a pointer to a pointer to (an array of) UDT
'/****************************************************************************
' *
' *  Clipboard routines
' *
' ***************************************************************************/
Public Const BMP_MAGIC_COOKIE           As Integer = 19778
'this is equivalent to ascii string "BM"
Public Type BITMAPFILEHEADER '14 bytes
  bfType                                As Integer
'"magic cookie" - must be "BM" (19778)
  bfSize                                As Long
  bfReserved1                           As Integer
  bfReserved2                           As Integer
  bfOffBits                             As Long
End Type
Public Type BITMAPINFOHEADER '40 bytes
  biSize                                As Long
  biWidth                               As Long
  biHeight                              As Long
  biPlanes                              As Integer
  biBitCount                            As Integer
  biCompression                         As Long
  biSizeImage                           As Long
  biXPelsPerMeter                       As Long
  biYPelsPerMeter                       As Long
  biClrUsed                             As Long
  biClrImportant                        As Long
End Type
Public Type BITMAPINFOHEADER_MJPEG '68 bytes
  biSize                                As Long
  biWidth                               As Long
  biHeight                              As Long
  biPlanes                              As Integer
  biBitCount                            As Integer
  biCompression                         As Long
  biSizeImage                           As Long
  biXPelsPerMeter                       As Long
  biYPelsPerMeter                       As Long
  biClrUsed                             As Long
  biClrImportant                        As Long
'/* extended BITMAPINFOHEADER fields */
  biExtDataOffset                       As Long
'/* compression-specific fields */
'/* these fields are defined for 'JPEG' and 'MJPG' */
  JPEGSize                              As Long
  JPEGProcess                           As Long
'/* Process specific fields */
  JPEGColorSpaceID                      As Long
  JPEGBitsPerSample                     As Long
  JPEGHSubSampling                      As Long
  JPEGVSubSampling                      As Long
End Type
Public Type AVI_RECT
  left                                  As Long
  top                                   As Long
  right                                 As Long
  bottom                                As Long
End Type
Public Type AVI_STREAM_INFO
  fccType                               As Long
  fccHandler                            As Long
  dwFlags                               As Long
  dwCaps                                As Long
  wPriority                             As Integer
  wLanguage                             As Integer
  dwScale                               As Long
  dwRate                                As Long
  dwStart                               As Long
  dwLength                              As Long
  dwInitialFrames                       As Long
  dwSuggestedBufferSize                 As Long
  dwQuality                             As Long
  dwSampleSize                          As Long
  rcFrame                               As AVI_RECT
  dwEditCount                           As Long
  dwFormatChangeCount                   As Long
  szName                                As String * 64
End Type
'for use with AVIFIleInfo
Public Type AVI_FILE_INFO  '108 bytes?
  dwMaxBytesPerSecond                   As Long
  dwFlags                               As Long
  dwCaps                                As Long
  dwStreams                             As Long
  dwSuggestedBufferSize                 As Long
  dwWidth                               As Long
  dwHeight                              As Long
  dwScale                               As Long
  dwRate                                As Long
  dwLength                              As Long
  dwEditCount                           As Long
  szFileType                            As String * 64
End Type
Public Type AVI_COMPRESS_OPTIONS
  fccType                               As Long
'/* stream type, for consistency */
  fccHandler                            As Long
'/* compressor */
  dwKeyFrameEvery                       As Long
'/* keyframe rate */
  dwQuality                             As Long
'/* compress quality 0-10,000 */
  dwBytesPerSecond                      As Long
'/* bytes per second */
  dwFlags                               As Long
'/* flags... see below */
  lpFormat                              As Long
'/* save format */
  cbFormat                              As Long
  lpParms                               As Long
'/* compressor options */
  cbParms                               As Long
  dwInterleaveEvery                     As Long
'/* for non-video streams only */
End Type
' /**************************************************************************
' *
' *  AVIFile* Constants (converted from C defines)
' *
' ***************************************************************************/
Public Const AVIERR_OK                  As Long = 0&
Private Const SEVERITY_ERROR            As Long = &H80000000
Private Const FACILITY_ITF              As Long = &H40000
Private Const AVIERR_BASE               As Long = &H4000
Private Const AVIERR_UNSUPPORTED        As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 101)                          '-2147205019
Public Const AVIERR_BADFORMAT           As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 102)                          '-2147205018
Private Const AVIERR_MEMORY             As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 103)                          '-2147205017
Public Const AVIERR_INTERNAL            As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 104)                          '-2147205016
Private Const AVIERR_BADFLAGS           As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 105)                          '-2147205015
Private Const AVIERR_BADPARAM           As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 106)                          '-2147205014
Private Const AVIERR_BADSIZE            As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 107)                          '-2147205013
'#define AVIERR_BADHANDLE        MAKE_AVIERR(108)
'#define AVIERR_FILEREAD         MAKE_AVIERR(109)
'#define AVIERR_FILEWRITE        MAKE_AVIERR(110)
'#define AVIERR_FILEOPEN         MAKE_AVIERR(111)
'#define AVIERR_COMPRESSOR       MAKE_AVIERR(112)
'#define AVIERR_NOCOMPRESSOR     MAKE_AVIERR(113)
'#define AVIERR_READONLY     MAKE_AVIERR(114)
'#define AVIERR_NODATA       MAKE_AVIERR(115)
'#define AVIERR_BUFFERTOOSMALL   MAKE_AVIERR(116)
'#define AVIERR_CANTCOMPRESS MAKE_AVIERR(117)
Private Const AVIERR_USERABORT          As Long = SEVERITY_ERROR Or FACILITY_ITF Or (AVIERR_BASE + 198)                          '-2147204922
'#define AVIERR_ERROR            MAKE_AVIERR(199)
'// Flags for dwFlags
''Private Const AVIFILEINFO_HASINDEX              As Long = &H10
''Private Const AVIFILEINFO_MUSTUSEINDEX          As Long = &H20
''Private Const AVIFILEINFO_ISINTERLEAVED         As Long = &H100
''Private Const AVIFILEINFO_WASCAPTUREFILE        As Long = &H10000
''Private Const AVIFILEINFO_COPYRIGHTED           As Long = &H20000
'// Flags for dwCaps
''Private Const AVIFILECAPS_CANREAD               As Long = &H1
''Private Const AVIFILECAPS_CANWRITE              As Long = &H2
''Private Const AVIFILECAPS_ALLKEYFRAMES          As Long = &H10
''Private Const AVIFILECAPS_NOCOMPRESSION         As Long = &H20
''Private Const AVICOMPRESSF_INTERLEAVE          As Long = &H1
'// interleave
''Private Const AVICOMPRESSF_DATARATE            As Long = &H2
'// use a data rate
''Private Const AVICOMPRESSF_KEYFRAMES           As Long = &H4
'// use keyframes
''Private Const AVICOMPRESSF_VALID               As Long = &H8
'// has valid data?
''Private Const OF_READ                           As Long = &H0
Public Const OF_WRITE                   As Long = &H1
''Public Const OF_SHARE_DENY_WRITE As Long = &H20
Public Const OF_CREATE                  As Long = &H1000
Public Const AVIIF_KEYFRAME             As Long = &H10
'/* DIB color table identifiers */
''Private Const DIB_RGB_COLORS                   As Long = 0
'/* color table in RGBs */
''Private Const DIB_PAL_COLORS                   As Long = 1
'/* color table in palette indices */
'/* constants for the biCompression field */
Public Const BI_RGB                     As Long = 0
''Private Const BI_RLE8                           As Long = 1
''Private Const BI_RLE4                           As Long = 2
''Private Const BI_BITFIELDS                      As Long = 3
'Stream types for use in VB (translated from C macros)
''Private Const streamtypeVIDEO                  As Long = 1935960438
'equivalent to: mmioStringToFOURCC("vids", 0&)
''Private Const streamtypeAUDIO                  As Long = 1935963489
'equivalent to: mmioStringToFOURCC("auds", 0&)
''Private Const streamtypeMIDI                   As Long = 1935960429
'equivalent to: mmioStringToFOURCC("mids", 0&)
''Private Const streamtypeTEXT                   As Long = 1937012852
'equivalent to: mmioStringToFOURCC("txts", 0&)
'// For GetFrame::SetFormat - use the best format for the display
''Private Const AVIGETFRAMEF_BESTDISPLAYFMT       As Long = 1
'// defines for uiFlags (AVISaveOptions)
Public Const ICMF_CHOOSE_KEYFRAME       As Long = &H1
'// show KeyFrame Every box
Public Const ICMF_CHOOSE_DATARATE       As Long = &H2
'// show DataRate box
''Private Const ICMF_CHOOSE_PREVIEW              As Long = &H4
'// allow expanded preview dialog
''Private Const ICMF_CHOOSE_ALLCOMPRESSORS       As Long = &H8
'// don't only show those that
'// can handle the input format
'// or input data
' /**************************************************************************
' *
' *  'STANDARD WIN32 API DECLARES, UDTs and Constants
' *
' ***************************************************************************/
Public Const HEAP_ZERO_MEMORY           As Long = &H8
'THESE ARE OTHER AVI API you might like to include but are not used in this program
''Public Declare Function VideoForWindowsVersion Lib "msvfw32.dll" () As Long
''Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
''Public Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pfile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
''Public Declare Function AVISave Lib "avifil32.dll" Alias "AVISaveVA" (ByVal szFile As String, ByVal pclsidHandler As Long, ByVal lpfnCallback As Long, ByVal nStreams As Long, ByRef ppaviStream As Long, ByRef ppCompOptions As Long) As Long
''Public Declare Function AVIStreamReadFormat Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lPos As Long, ByVal lpFormatBuf As Long, ByRef sizeBuf As Long) As Long
''Public Declare Function AVIStreamRead Lib "avifil32.dll" (ByVal pAVIStream As Long, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long, ByRef pBytesWritten As Long, ByRef pSamplesWritten As Long) As Long
''Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef bih As Any) As Long 'returns pointer to GETFRAME object on success (or NULL on error)
''Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, ByVal lPos As Long) As Long 'returns pointer to packed DIB on success (or NULL on error)
''Public Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long ' returns zero on success (error number) after calling this function the GETFRAME object pointer is invalid
''Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
''Public Declare Function AVIMakeFileFromStreams Lib "avifil32.dll" (ByRef ppfile As Long, ByVal nStreams As Long, ByVal pAVIStreamArray As Long) As Long
''Public Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
''Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
''Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
''Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
''Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long
''Public Declare Function AVIMakeStreamFromClipboard Lib "avifil32.dll" (ByVal cfFormat As Long, ByVal hpublic As Long, ByRef ppstream As Long) As Long
''Public Declare Function AVIPutFileOnClipboard Lib "avifil32.dll" (ByVal pAVIFile As Long) As Long
''Public Declare Function AVIGetFromClipboard Lib "avifil32.dll" (ByRef ppAVIFile As Long) As Long
''Public Declare Function AVIClearClipboard Lib "avifil32.dll" () As Long
'handle
Public Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, _
                                                                                        ByVal uFlags As Long) As Long 'returns fourcc
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, _
                                                        ByVal szFile As String, _
                                                        ByVal uMode As Long, _
                                                        ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVIFileCreateStream Lib "avifil32.dll" Alias "AVIFileCreateStreamA" (ByVal pfile As Long, _
                                                                                             ByRef ppavi As Long, _
                                                                                             ByRef psi As AVI_STREAM_INFO) As Long
Public Declare Function AVISaveOptions Lib "avifil32.dll" (ByVal hWnd As Long, _
                                                           ByVal uiFlags As Long, _
                                                           ByVal nStreams As Long, _
                                                           ByRef ppavi As Long, _
                                                           ByRef ppOptions As Long) As Long 'TRUE if user pressed OK, False if cancel, or error if error
Public Declare Function AVISaveOptionsFree Lib "avifil32.dll" (ByVal nStreams As Long, _
                                                               ByRef ppOptions As Long) As Long
Public Declare Function AVIMakeCompressedStream Lib "avifil32.dll" (ByRef ppsCompressed As Long, _
                                                                    ByVal psSource As Long, _
                                                                    ByRef lpOptions As AVI_COMPRESS_OPTIONS, _
                                                                    ByVal pclsidHandler As Long) As Long '
Public Declare Function AVIStreamSetFormat Lib "avifil32.dll" (ByVal pavi As Long, _
                                                               ByVal lPos As Long, _
                                                               ByRef lpFormat As Any, _
                                                               ByVal cbFormat As Long) As Long
Public Declare Function AVIStreamWrite Lib "avifil32.dll" (ByVal pavi As Long, _
                                                           ByVal lStart As Long, _
                                                           ByVal lSamples As Long, _
                                                           ByVal lpBuffer As Long, _
                                                           ByVal cbBuffer As Long, _
                                                           ByVal dwFlags As Long, _
                                                           ByRef plSampWritten As Long, _
                                                           ByRef plBytesWritten As Long) As Long
Public Declare Function AVIStreamClose Lib "avifil32.dll" Alias "AVIStreamRelease" (ByVal pavi As Long) As Long                                                                                                                                                                                                                                                                                                                                                                                       'ULONG
Public Declare Function AVIFileClose Lib "avifil32.dll" Alias "AVIFileRelease" (ByVal pfile As Long) As Long
Public Declare Sub AVIFileExit Lib "avifil32.dll" ()
Public Declare Function SetRect Lib "user32.dll" (ByRef lprc As AVI_RECT, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal xRight As Long, _
                                                  ByVal yBottom As Long) As Long 'BOOL
Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Public Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, _
                                                      ByVal dwFlags As Long, _
                                                      ByVal dwBytes As Long) As Long 'Pointer to mem
Public Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, _
                                                     ByVal dwFlags As Long, _
                                                     ByVal lpMem As Long) As Long                                                                   'BOOL
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, _
                                                                        ByRef src As Any, _
                                                                        ByVal dwLen As Long)


':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:24) 261 + 0 = 261 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

