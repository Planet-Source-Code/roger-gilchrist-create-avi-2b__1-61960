VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this is original to 'Create AVI' by Ron Hoebe
'all I have done is a bit of Code Fixer tidying up
Private Const BMP_MAGIC_COOKIE      As Integer = 19778
'this is equivalent to ascii string "BM"
Private Type BITMAPFILEHEADER '14 bytes
  bfType                            As Integer
'"magic cookie" - must be "BM"
  bfSize                            As Long
  bfReserved1                       As Integer
  bfReserved2                       As Integer
  bfOffBits                         As Long
End Type
Private Type BITMAPINFOHEADER '40 bytes
  biSize                            As Long
  biWidth                           As Long
  biHeight                          As Long
  biPlanes                          As Integer
  biBitCount                        As Integer
  biCompression                     As Long
  biSizeImage                       As Long
  biXPelsPerMeter                   As Long
  biYPelsPerMeter                   As Long
  biClrUsed                         As Long
  biClrImportant                    As Long
End Type
Private Type RGBQUAD
  Red                               As Byte
  Green                             As Byte
  Blue                              As Byte
  Reserved                          As Byte
End Type
Private Type BITMAP
  bmType                            As Long
  bmWidth                           As Long
  bmHeight                          As Long
  bmWidthBytes                      As Long
  bmPlanes                          As Integer
  bmBitsPixel                       As Integer
  bmBits                            As Long
End Type
'/* constants for the biCompression field */
Private Const BI_RGB                As Long = 0&
Private Const HEAP_ZERO_MEMORY      As Long = &H8
Private m_memBits()                 As Byte
Private m_memBitmapInfo()           As Byte
Private m_bih                       As BITMAPINFOHEADER
Private m_bfh                       As BITMAPFILEHEADER
'handle
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, _
                                                       ByVal dwFlags As Long, _
                                                       ByVal dwBytes As Long) As Long 'Pointer to mem
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, _
                                                      ByVal dwFlags As Long, _
                                                      ByVal lpMem As Long) As Long                                                                   'BOOL
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, _
                                                                         ByRef src As Any, _
                                                                         ByVal dwLen As Long)
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long


Public Property Get BitCount() As Long

  BitCount = m_bih.biBitCount

End Property

Public Property Get BitOffset() As Long

  BitOffset = m_bfh.bfOffBits

End Property

Public Property Get ClrImportant() As Long

  ClrImportant = m_bih.biClrImportant

End Property

Public Property Get ClrUsed() As Long

  ClrUsed = m_bih.biClrUsed

End Property

Public Property Get Compression() As Long

  Compression = m_bih.biCompression

End Property

Public Function CreateFromFile(ByVal strfilename As String) As Boolean

  Dim hFile As Long

  If Not ExistFile(strfilename) Then
    MsgBox "File does not exist:" & vbNewLine & strfilename, vbCritical, App.Title
   Else
    hFile = FreeFile()
'<====ERROR TRAP ON
    On Error Resume Next
    Open strfilename For Binary Access Read As #hFile
    If Err Then
      If Err.Number = 70 Then
        MsgBox "File is locked - cannot access:" & vbNewLine & strfilename, vbCritical, App.Title
       Else
        MsgBox Err.Description, vbInformation, App.Title
      End If
      Exit Function 'assume file was not opened
    End If
    On Error GoTo 0
'====>ERROR TRAP OFF
'OK, file is opened - now for the real algorithm...
    Get #hFile, , m_bfh 'get the BITMAPFILEHEADER this identifies the bitmap
    If m_bfh.bfType <> BMP_MAGIC_COOKIE Then 'this is not a BMP file
      MsgBox "File is not a supported bitmap format:" & vbNewLine & strfilename, vbInformation, App.Title
      Close #hFile
     Else
'now get the info header
      Get #hFile, Len(m_bfh) + 1, m_bih 'start at the 15th byte
      If m_bih.biSizeImage = 0 Then
        m_bih.biSizeImage = m_bih.biHeight * m_bih.biWidth
      End If
'now get the bitmap bits
      ReDim m_memBits(0 To m_bih.biSizeImage - 1)
      Get #hFile, m_bfh.bfOffBits + 1, m_memBits
'and BitmapInfo variable-length UDT
      ReDim m_memBitmapInfo(0 To m_bfh.bfOffBits - 14) 'don't need first 14 bytes (fileinfo)
      Get #hFile, Len(m_bfh) + 1, m_memBitmapInfo
      Close #hFile   'Close file
      CreateFromFile = True 'indicate success
    End If
  End If

End Function

Public Function CreateFromPackedDIBPointer(ByRef pDIB As Long) As Boolean

'Debug.Assert pDIB <> 0
'Creates a full-color (no palette) DIB from a pointer to a full-color memory DIB
'get the BitmapInfoHeader

  CopyMemory ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih)
  If m_bih.biBitCount >= 16 Then
'now get the bitmap bits
'return False
    If m_bih.biSizeImage >= 1 Then
      ReDim m_memBits(0 To m_bih.biSizeImage - 1)
      CopyMemory m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage
'and BitmapInfo variable-length UDT
      ReDim m_memBitmapInfo(0 To 39) 'don't need first 14 bytes (fileinfo)
      CopyMemory m_memBitmapInfo(0), m_bih, Len(m_bih)
'create a file header
      With m_bfh
        .bfType = BMP_MAGIC_COOKIE
        .bfSize = 55 + m_bih.biSizeImage 'size of file as written to disk
        .bfReserved1 = 0&
        .bfReserved2 = 0&
        .bfOffBits = 54 'BitmapInfoHeader + BitmapFileHeader
      End With
'and return True
      CreateFromPackedDIBPointer = True
    End If
  End If

End Function

Private Function ExistFile(ByVal sSpec As String) As Boolean

  On Error Resume Next
  FileLen sSpec
  ExistFile = (Err.Number = 0)
  On Error GoTo 0

End Function

Public Property Get FileType() As Long

  FileType = m_bfh.bfType

End Property

Public Property Get Height() As Long

  Height = m_bih.biHeight

End Property

Public Property Get Planes() As Long

  Planes = m_bih.biPlanes

End Property

Public Property Get PointerToBitmapInfo() As Long

  PointerToBitmapInfo = VarPtr(m_memBitmapInfo(0))

End Property

Public Property Get PointerToBits() As Long

  PointerToBits = VarPtr(m_memBits(0))

End Property

Public Property Get SizeBitmapInfo() As Long

  SizeBitmapInfo = UBound(m_memBitmapInfo()) + 1

End Property

Public Property Get SizeFileHeader() As Long

  SizeFileHeader = m_bfh.bfSize

End Property

Public Property Get SizeImage() As Long

  SizeImage = m_bih.biSizeImage

End Property

Public Property Get SizeInfoHeader() As Long

  SizeInfoHeader = m_bih.biSize

End Property

Public Property Get Width() As Long

  Width = m_bih.biWidth

End Property

Public Function WriteToFile(ByVal strfilename As String) As Boolean

  Dim hFile As Integer

  On Error Resume Next
  hFile = FreeFile()
  Open strfilename For Binary As hFile
  Put hFile, 1, m_bfh
  Put hFile, Len(m_bfh) + 1, m_memBitmapInfo
  Put hFile, , m_memBits
  Close hFile
  WriteToFile = True
  On Error GoTo 0

End Function

Public Property Get XPPM() As Long

  XPPM = m_bih.biXPelsPerMeter

End Property

Public Property Get YPPM() As Long

  YPPM = m_bih.biYPelsPerMeter

End Property

':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:27) 53 + 209 = 262 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

