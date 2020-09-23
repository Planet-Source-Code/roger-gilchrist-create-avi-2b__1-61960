Attribute VB_Name = "ModFolderBrowser"
'VERSION 2
'FolderBrowser code added
Option Explicit
Private Const BFFM_INITIALIZED            As Long = 1
Private Const WM_USER                     As Long = &H400
Private Const BFFM_SETSELECTIONA          As Long = (WM_USER + 102)
Private Const BIF_RETURNONLYFSDIRS        As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN       As Long = 2
Private Const MAX_PATH                    As Long = 260
Private Const LMEM_FIXED                  As Long = &H0
'added
Private Const LMEM_ZEROINIT               As Long = &H40
'added
Private Const LPTR                        As Double = (LMEM_FIXED Or LMEM_ZEROINIT)
Public Type BROWSEINFO
  hOwner                                  As Long
  pidlRoot                                As Long
  pszDisplayName                          As String
  lpszTitle                               As String
  ulFlags                                 As Long
  lpfn                                    As Long
  lParam                                  As Long
  iImage                                  As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal dwLength As Long)                                                                                                                                                           'added
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                                                                             ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, _
                                                    ByVal uBytes As Long) As Long                                                                                                                                                                                                                                                                            'added
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, _
                                                  lpString2 As Any) As Long                                                                                                                                                                                                                                                                                                                                                 'added
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long                         'added


Public Function Browse_Folder(sPath As String, _
                              StrTitle As String) As String   'added

  Dim lpIDList    As Long 'Declare Varibles
  Dim sBuffer     As String
  Dim lpPath      As Long 'added
  Dim tBrowseInfo As BROWSEINFO

'Text to appear in the the gray area under the title bar
'telling you what to do
  With tBrowseInfo
    .hOwner = 0 'UserControl.hWnd '.hWnd 'Owner Form
    .pidlRoot = 0
    .lpszTitle = lstrcat(StrTitle, vbNullString)
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    .lpfn = FARPROC(AddressOf BrowseCallbackProcStr) 'added
    If Len(sPath) Then
      lpPath = LocalAlloc(LPTR, Len(sPath) + 1)   'added
      CopyMemory ByVal lpPath, ByVal sPath, Len(sPath) + 1    'added
      .lParam = lpPath    'added
     Else
      .lParam = 0 '
    End If
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If lpIDList Then
    sBuffer = Space$(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  End If
  Browse_Folder = sBuffer

End Function

Public Function BrowseCallbackProcStr(ByVal lngHWnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long

'Callback for the Browse STRING method.
'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.

  Select Case uMsg
   Case BFFM_INITIALIZED
    SendMessage lngHWnd, BFFM_SETSELECTIONA, True, ByVal lpData
  End Select

End Function

Public Function FARPROC(pfn As Long) As Long

'A dummy procedure that receives and returns
'the value of the AddressOf operator.
'This workaround is needed as you can't assign
'AddressOf directly to a member of a user-
'defined type, but you can assign it to another
'long and use that (as returned here)

  FARPROC = pfn

End Function

':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:38) 34 + 65 = 99 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

