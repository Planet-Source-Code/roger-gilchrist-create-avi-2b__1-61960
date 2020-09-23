VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAVI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create AVI from Picture Files"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmAvi.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleMode       =   0  'User
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CheckBox chkShowImage 
      Caption         =   "Show Images while creating"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      ToolTipText     =   "Slightly slower"
      Top             =   7560
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.ListBox lstAVIPlayers 
      Height          =   450
      Left            =   3840
      TabIndex        =   26
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Frame fraCreateAVI 
      Caption         =   "AVI Style"
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   1215
      Begin VB.PictureBox picCFXPBugFixfrmAVI 
         BorderStyle     =   0  'None
         Height          =   758
         Index           =   2
         Left            =   100
         ScaleHeight     =   765
         ScaleWidth      =   1020
         TabIndex        =   22
         Top             =   276
         Width           =   1015
         Begin VB.OptionButton optAVIStyle 
            Caption         =   "Bounce"
            Height          =   255
            Index           =   2
            Left            =   20
            TabIndex        =   25
            Top             =   492
            Width           =   975
         End
         Begin VB.OptionButton optAVIStyle 
            Caption         =   "Reverse"
            Height          =   255
            Index           =   1
            Left            =   20
            TabIndex        =   24
            Top             =   237
            Width           =   975
         End
         Begin VB.OptionButton optAVIStyle 
            Caption         =   "Standard"
            Height          =   255
            Index           =   0
            Left            =   20
            TabIndex        =   23
            Top             =   -18
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdPlayAVI 
      Caption         =   "Play Using"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   6960
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Frame fraCreateAVI 
      Caption         =   "Image"
      Height          =   5145
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1740
      Width           =   6765
      Begin VB.PictureBox picCFXPBugFixfrmAVI 
         BorderStyle     =   0  'None
         Height          =   4815
         Index           =   1
         Left            =   100
         ScaleHeight     =   4815
         ScaleWidth      =   6570
         TabIndex        =   17
         Top             =   276
         Width           =   6570
         Begin VB.PictureBox imShow 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   5000
            Left            =   -45
            ScaleHeight     =   4935
            ScaleWidth      =   6615
            TabIndex        =   19
            Top             =   -315
            Width           =   6675
         End
         Begin VB.PictureBox imL 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   315
            Left            =   4320
            ScaleHeight     =   255
            ScaleWidth      =   585
            TabIndex        =   18
            Top             =   4440
            Visible         =   0   'False
            Width           =   645
         End
      End
   End
   Begin VB.Frame fraCreateAVI 
      Caption         =   "Options"
      Height          =   1095
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   6765
      Begin VB.PictureBox picCFXPBugFixfrmAVI 
         BorderStyle     =   0  'None
         Height          =   758
         Index           =   0
         Left            =   100
         ScaleHeight     =   765
         ScaleWidth      =   6570
         TabIndex        =   8
         Top             =   276
         Width           =   6570
         Begin VB.CommandButton cmdFastAutoName 
            Caption         =   "Fast Auto Name"
            Height          =   735
            Left            =   1680
            TabIndex        =   28
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdFiles 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   5985
            TabIndex        =   11
            Top             =   -18
            Width           =   525
         End
         Begin VB.CommandButton cmdFiles 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   5985
            TabIndex        =   10
            Top             =   402
            Width           =   525
         End
         Begin VB.TextBox txtFPS 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "10"
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lblCreateAVI 
            Alignment       =   1  'Right Justify
            Caption         =   "Destination:"
            Height          =   255
            Index           =   1
            Left            =   2775
            TabIndex        =   16
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lblCreateAVI 
            Alignment       =   1  'Right Justify
            Caption         =   "Source:"
            Height          =   255
            Index           =   2
            Left            =   2865
            TabIndex        =   15
            Top             =   45
            Width           =   765
         End
         Begin VB.Label lblCreateAVI 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "First Picture File -> Browse ->"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   3735
            TabIndex        =   14
            Top             =   15
            Width           =   2175
         End
         Begin VB.Label lblCreateAVI 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AVI File -> Browse ->"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   3735
            TabIndex        =   13
            Top             =   405
            Width           =   2175
         End
         Begin VB.Label lblCreateAVI 
            Caption         =   "FPS: ( 0.1 to 50) Small = slow"
            Height          =   375
            Index           =   0
            Left            =   15
            TabIndex        =   12
            Top             =   45
            Width           =   1155
         End
      End
   End
   Begin MSComDlg.CommonDialog cdlCreateAVI 
      Left            =   3000
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create AVI"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   7080
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   8085
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCreateAVI 
      Caption         =   "Progress"
      Height          =   555
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1170
      Width           =   6765
      Begin VB.PictureBox picCreateAVI 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   6465
         TabIndex        =   1
         Top             =   240
         Width           =   6525
      End
   End
End
Attribute VB_Name = "frmAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'PURPOSE
' Convert ordered collections of bmp, jpg or gif files into AVI files
'
'Based on
'Ron Hoebe's 'Create AVI'
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=29517&lngWId=1
'AND
'NeO78's  'PDF Printer Class'
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61936&lngWId=1
'which showed mee how to search disk for specific files
'
'ProgressBar & Estimated time
'__merlin__'s 'ProgressBar2Class (8 DrawDirections, XOR Caption,Time2End Display)'
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=23431&lngWId=1
'
'MODIFICATIONS by roger gilchrist
'(apart from a little tidying up all the modifications are in the code on this form and ModFindFile)
'1 Ignores non-graphic files in the folder (original tried to load all files in folder and crashed if it found non-graphics)
'2 Side effect of 1 is to allow writing the AVI to the same folder as the source files.
'3 Added Reverse Style(Images inserted in reverse order of alphanumeric sort)
'4 Added Bounce  Style(Standard + Reverse but turnaround image not doubled and last frame = 2nd image)
'  This means that you can run the AVI as a continuous loop without an apparent pause as it wraps around
'5 Added Play AVI from program using AVI Players on the system.
'  Program is coded to search for Windows Media Player and WinAmp
'  You can add others by modifying the list in GetAVIPlayers (details in the procedure)
'
'VERSION 2
'6 Added 'Fast Auto Name' Uses FolderBrowser for simpler folder selection and automatic AVI file name
'   Use if you are sure of where the files are
'7 Modified the FPS to allow fractions. At 0.1 each picture will last about 3 seconds
'8 improved button activation and deactivation
'
'USAGE
'1 Place all files you wish to make into an AVI in single folder.
'2 All files need to be same size as each other (typical of captures from film or screen captures)
'3 Sort the folder to create the base order (normally sorted by name if using captures with auto numbering schemes)
'4 Try out the FPS, different codec and compression schemes offered(See WARNING 2 below) on small sets of images first until you find what suits you best.
'5 Select the first file in the folder (this will be used to set the basic size of the AVI frames)
'6 Create a name for the target AVI file.
'7 Select a Style (default Standard)
'8 Steps 5 & 6 activate the 'Create AVI' button
'9 Click the button and go have a coffee.
'10 You can launch the AVI in AVI player(s) from the program (Modification 5)
'
'<<<<<<<<<<<<<<<<<<<<<<<<<<  WARNINGS   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'1 If any file is not the same size as others it will distort badly (and may corrupt the AVI)
'2 The codecs listed are those detected on your machine however you may not be able to use them all
'  due to licencing restrictions, incomplete installation or uninstalling errors.
'3 Building AVI's can be slow (the more compression you use the worse it will be)
'4 Interrupting constuction may create an unusable AVI
'5 Corrupt AVI's may cause the Folder Explorer to crash.
'   (Especially if the folder is set to display thumbnails)
'6 If 5 occurs restart Windows, open the folder containing the AVI
'  (You may have to force it out of thumbnail view first
'     a. select another folder
'     b. set to details view
'     c. Use the Tools|Folder Options menu View Tab 'Apply to All Folders' button
'   Delete the AVI
'7 If you build your image set by frame capture from a DVD you will almost certainly be violating copyright laws.
'<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'<<<<<< DISCLAIMER >>>>>>>>>>>>>>>>
'All care no responsibility
'Code is supplied 'as is'
'Code has been tested and is capable of produicng AVI files
'BUT is also capable of disrupting (temporarily) your system (See WARNINGS 4,5,6)
'<<<<<< DISCLAIMER >>>>>>>>>>>>>>>>
'
Private arrFile              As Variant
Private strLastFolder        As String
Private cBar                 As ClsProgressBar2
Private bmp                  As cDIB
Private fpath                As String
Private szOutputAVIFile      As String
Private szInputFile          As String
Private docancel             As Boolean
Private fDest                As Boolean
Private fSrc                 As Boolean
Private StrWMP               As String
Private Players()            As String
Private PlayerCount          As Long


Private Function AddAFrame(psCompressed As Long, _
                           VarFileName As Variant, _
                           Counter As Long, _
                           res As Long) As Long

  With cBar 'update the progbar
    .Value = Counter
    fraCreateAVI(0).Caption = "Estimated Time Remaining " & .Time2End
    .ShowBar
  End With 'cBar
  DoEvents ' this allows you to interupt the process
  sBar.SimpleText = "Adding " & FlleNameOnly(CStr(VarFileName))
  bmp.CreateFromFile (loadPic(strLastFolder & "\" & VarFileName, chkShowImage = vbChecked, False, True))
  res = AVIStreamWrite(psCompressed, Counter, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
  If res = AVIERR_OK Then
    AddAFrame = True
  End If
  If docancel Then
    AddAFrame = False
  End If

End Function

Private Sub ButtonSet(ByVal bEnabled As Boolean)

  Dim I As Long

  docancel = bEnabled
  cmdCancel.Visible = Not bEnabled
  cmdCancel.Enabled = Not bEnabled
  cmdClose.Enabled = bEnabled
  For I = 0 To 1
    cmdFiles(I).Enabled = bEnabled
  Next I
  cmdCreate.Visible = bEnabled
  fraCreateAVI(0).Caption = vbNullString

End Sub

Private Sub CloseAviWriting(psCompressed As Long, _
                            pfile As Long, _
                            ps As Long, _
                            ByVal res As Long)

  If ps <> 0 Then
    AVIStreamClose ps
  End If
  If psCompressed <> 0 Then
    AVIStreamClose psCompressed
  End If
  If pfile <> 0 Then
    AVIFileClose pfile
  End If
  AVIFileExit
  If res <> AVIERR_OK Then
    If res = AVIERR_BADFORMAT Or res = AVIERR_INTERNAL Then
      MsgBox "There was an error creating the AVI File." & vbNewLine & _
       "Probably the choosen Video Compression does not support the File format or the input File is corrupt", vbInformation, App.Title
      sBar.SimpleText = "There was an error writing the file."
     Else
      MsgBox "There was an error creating the AVI File.", vbInformation, App.Title
      sBar.SimpleText = "There was an error writing the file."
    End If
   Else
    sBar.SimpleText = "Created " & FlleNameOnly(CStr(szOutputAVIFile))
    If cmdPlayAVI.Visible Then
      cmdPlayAVI.Enabled = True
    End If
  End If

End Sub

Private Sub cmdCancel_Click()

  docancel = True

End Sub

Private Sub cmdClose_Click()

  Unload Me

End Sub

Private Sub cmdCreate_Click()

  Dim res          As Long
  Dim pfile        As Long   'ptr PAVIFILE
  Dim ps           As Long   'ptr PAVISTREAM
  Dim psCompressed As Long   'ptr PAVISTREAM
  Dim FrameTime    As Long   'Frame counter and time remaining counter
  Dim J            As Long   'loop through file array

  ButtonSet False
  If InitAviWriting(psCompressed, pfile, ps) Then
'  arrFile = GetFileArray(fpath)
'initialize the progbar
    With cBar
      .SetParamFast 1, UBound(arrFile) * IIf(optAVIStyle(2).Value, 2, 1), Left2Right, True, ShowCaptionL
      .Value = 1
      .StartTimer
    End With 'cBar
    Select Case True
     Case optAVIStyle(0).Value  'Standard
      cBar.Caption = "Standard"
      For J = 0 To UBound(arrFile)
        FrameTime = FrameTime + 1
        If Not AddAFrame(psCompressed, arrFile(J), FrameTime, res) Then
          Exit For
        End If
      Next J
     Case optAVIStyle(1).Value  'Reverse
      cBar.Caption = "Reverse"
      For J = UBound(arrFile) To 0 Step -1
        FrameTime = FrameTime + 1
        If Not AddAFrame(psCompressed, arrFile(J), FrameTime, res) Then
          Exit For
        End If
      Next J
     Case optAVIStyle(2).Value  'Bounce
      cBar.Caption = "Bounce"
      For J = 0 To UBound(arrFile)
        FrameTime = FrameTime + 1
        If Not AddAFrame(psCompressed, arrFile(J), FrameTime, res) Then
          GoTo error 'Exit For
        End If
      Next J
      For J = UBound(arrFile) - 1 To 1 Step -1
        FrameTime = FrameTime + 1
        If Not AddAFrame(psCompressed, arrFile(J), FrameTime, res) Then
          Exit For
        End If
      Next J
    End Select
  End If
error:
  CloseAviWriting psCompressed, pfile, ps, res
  cBar.Value = 1
  cBar.ShowBar
  ButtonSet True
  cmdCreate.Enabled = False

End Sub

Private Sub cmdFastAutoName_Click()

  Dim strfolder As String
  Dim Counter   As Long

  strfolder = Browse_Folder(strLastFolder, "Select a folder") ', 0)
  If LenB(strfolder) Then
    strLastFolder = strfolder
    arrFile = GetFileArray(strLastFolder)
    If LenB(arrFile(0)) Then
      szInputFile = strLastFolder & "\" & arrFile(0)
      loadPic szInputFile, chkShowImage = vbChecked, True, False
      lblCreateAVI(4).Caption = FlleNameOnly(szInputFile)
'avi name based on foldername
      szOutputAVIFile = strfolder & "\" & FlleNameOnly(strfolder) & ".avi"
      If FileExists(szOutputAVIFile) Then
'if already is an AVI of that name then create a numbered name
        Do While FileExists(strfolder & "\" & FlleNameOnly(strfolder) & Format$(Counter, "000") & ".avi")
          Counter = Counter + 1
        Loop
        szOutputAVIFile = strfolder & "\" & FlleNameOnly(strfolder) & Format$(Counter, "000") & ".avi"
      End If
      lblCreateAVI(3).Caption = FlleNameOnly(szOutputAVIFile)
      cmdCreate.Enabled = True
     Else
      MsgBox "No suitable source files found in folder '" & FlleNameOnly(strfolder) & "'.", vbInformation
    End If
    
  End If
  cmdPlayAVI.Enabled = False
End Sub

Private Sub cmdFiles_Click(Index As Integer)

  Select Case Index
   Case 0 '...
    fSrc = SrcFile
    If fSrc Then
      loadPic szInputFile, chkShowImage = vbChecked, True, False
      strLastFolder = GetFilePath(szInputFile)
    End If
   Case 1 '...
    fDest = DestFile
  End Select
  cmdCreate.Enabled = fSrc And fDest

End Sub

Private Sub cmdPlayAVI_Click()

  Shell Players(lstAVIPlayers.ListIndex) & " " & Chr$(34) & szOutputAVIFile & Chr$(34), vbNormalFocus

End Sub

Private Function DestFile() As Boolean

  With cdlCreateAVI
    .DialogTitle = "Save AVI File"
    .CancelError = False
    .Filter = "AVI Files (*.avi)|*.avi"
    .DefaultExt = "avi"
    .FileName = vbNullString
    .ShowSave
  End With
  If Len(cdlCreateAVI.FileName) = 0 Then
    DestFile = False
   Else
    szOutputAVIFile = cdlCreateAVI.FileName
    lblCreateAVI(3).Caption = FlleNameOnly(szOutputAVIFile)
    DestFile = True
  End If

End Function

Private Sub Form_Load()

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set cBar = New ClsProgressBar2
  cBar.SetPictureBox = picCreateAVI
  fSrc = False
  fDest = False
  GetAVIPlayers

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set bmp = Nothing
  Set cBar = Nothing
  Set fso = Nothing

End Sub

Private Sub GetAVIPlayers()

  Dim strTmp      As String
  Dim PlayerCount As Long
  Dim I           As Long
  Dim ArrPlayers  As Variant
  Dim aPlayer     As Variant

'NOTE if any players are found in standard locations ("C:\Program Files....) then the all disks search will not be preformed
'You can either recode this or open the Registry and edit the values by hand to include the non-standard paths
'
'MODIFICATION POINT 1
'to add/delete players modify the next line to include/exclude players
  ArrPlayers = Array("wmplayer.exe", "winamp.exe")
'MODIFICATION POINT 2
'if you have already run the program and wish to change the player list uncomment the next line, run in VB then comment it
' SaveSetting App.EXEName, "OPTIONS", "PlayerCount", 0
'line above this deletes the PlayerCount Value of the key and forces program to search for players
'
  PlayerCount = GetSetting(App.EXEName, "OPTIONS", "PlayerCount", 0)
  If PlayerCount = 0 Then
    MsgBox "This is the first run of this program" & vbNewLine & _
       "There will be a short delay while the program searches for AVI Players", vbInformation
    For Each aPlayer In ArrPlayers
'yeah I know after all those comments about hard-coded paths but there you go...
'should be OK for about 99.9% of machines  RG
      strTmp = FileFind("C:\Program Files", CStr(aPlayer), 0)
      If LenB(strTmp) Then
        PlayerCount = PlayerCount + 1
        SaveSetting App.EXEName, "OPTIONS", "Player" & PlayerCount, strTmp
        SaveSetting App.EXEName, "OPTIONS", "PlayerCount", PlayerCount
      End If
    Next aPlayer
  End If
  If PlayerCount = 0 Then
    MsgBox "Program couldn't find AVI Players in the standard locations." & vbNewLine & _
       "There will be a longer delay while all your disks are searched for AVI Players", vbInformation
    For Each aPlayer In ArrPlayers
      strTmp = SystemFilePathScan(CStr(aPlayer))
      If LenB(strTmp) Then
        PlayerCount = PlayerCount + 1
        SaveSetting App.EXEName, "OPTIONS", "Player" & PlayerCount, strTmp
        SaveSetting App.EXEName, "OPTIONS", "PlayerCount", PlayerCount
      End If
    Next aPlayer
  End If
  PlayerCount = GetSetting(App.EXEName, "OPTIONS", "PlayerCount", 0)
  If PlayerCount = 0 Then
    MsgBox "Unable to find any AVI players." & vbNewLine & _
       "Please change the constants 'AVIPlayer#' to the name of your AVI player and restart the program." & vbNewLine & _
       " You can continue to run and create an AVI but the Play button will not appear."
   Else
    ReDim Players(PlayerCount) As String
    For I = 0 To PlayerCount - 1
      Players(I) = GetSetting(App.EXEName, "OPTIONS", "Player" & I + 1, vbNullString)
      lstAVIPlayers.AddItem FlleNameOnly(CStr(Players(I)))
    Next I
  End If
  cmdPlayAVI.Visible = lstAVIPlayers.ListCount > 0
  lstAVIPlayers.Visible = lstAVIPlayers.ListCount > 0
  If lstAVIPlayers.ListCount > 0 Then
    lstAVIPlayers.ListIndex = GetSetting(App.EXEName, "OPTIONS", "PlayerPref", 0)
  End If

End Sub

Private Function GetFileArray(ByVal fpath As String) As Variant

  Dim arrTmp() As Variant
  Dim sFile    As String

'Get 1st file in folder
  sFile = Dir(fpath & "\*.*")
  Do Until InStr("|bmp|jpg|gif|", GetFileExtension(sFile, True))
    sFile = Dir()
    If LenB(sFile) = 0 Then
      GoTo NoFiles 'no files found
    End If
  Loop
'set first member of array
  ReDim arrTmp(0) As Variant
  arrTmp(UBound(arrTmp)) = sFile
'get rest of acceptable files
  Do
    Do
      sFile = Dir()
      If LenB(sFile) = 0 Then
        GoTo Finished:
      End If
    Loop Until InStr("|bmp|jpg|gif|", GetFileExtension(sFile, True))
    ReDim Preserve arrTmp(UBound(arrTmp) + 1) As Variant
    arrTmp(UBound(arrTmp)) = sFile
  Loop
Finished:
  GetFileArray = arrTmp
NoFiles:

End Function

Private Function InitAviWriting(psCompressed As Long, _
                                pfile As Long, _
                                ps As Long) As Boolean

  Dim res    As Long
  Dim strhdr As AVI_STREAM_INFO
  Dim BI     As BITMAPINFOHEADER
  Dim opts   As AVI_COMPRESS_OPTIONS
  Dim pOpts  As Long

'    Open the file for writing
  res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
  If res <> AVIERR_OK Then
    InitAviWriting = False
    GoTo error
  End If
'Get the first bmp in the list for setting format
  Set bmp = New cDIB
''bmpFile = loadPic(szInputFile, False, False, True)
  If bmp.CreateFromFile(loadPic(szInputFile, False, False, True)) <> True Then
    MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
    GoTo error
  End If
'   Fill in the header for the video stream
  If CSng(txtFPS.Text) < 0.1 Or CLng(txtFPS) > 50 Then
    txtFPS.Text = "10"
  End If
  With strhdr
    .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
    .fccHandler = 0&                             '// default AVI handler
    .dwScale = 1
    .dwRate = CSng(txtFPS.Text)                          '// fps
    .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
    SetRect .rcFrame, 0, 0, bmp.Width, bmp.Height       '// rectangle for stream
  End With
'validate user input
  If strhdr.dwRate < 1 Then
    strhdr.dwRate = 1
  End If
  If strhdr.dwRate > 30 Then
    strhdr.dwRate = 30
  End If
'   And create the stream
  res = AVIFileCreateStream(pfile, ps, strhdr)
  If res <> AVIERR_OK Then
    GoTo error
  End If
'get the compression options from the user
'Careful! this API requires a pointer to a pointer to a UDT
  pOpts = VarPtr(opts)
  res = AVISaveOptions(frmAVI.hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
'returns TRUE if User presses OK, FALSE if Cancel, or error code
  If res <> 1 Then 'In C TRUE = 1
    AVISaveOptionsFree 1, pOpts
    GoTo error
  End If
'make compressed stream
  res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
  If res <> AVIERR_OK Then
    GoTo error
  End If
'set format of stream according to the bitmap
  With BI
    .biBitCount = bmp.BitCount
    .biClrImportant = bmp.ClrImportant
    .biClrUsed = bmp.ClrUsed
    .biCompression = bmp.Compression
    .biHeight = bmp.Height
    .biWidth = bmp.Width
    .biPlanes = bmp.Planes
    .biSize = bmp.SizeInfoHeader
    .biSizeImage = bmp.SizeImage
    .biXPelsPerMeter = bmp.XPPM
    .biYPelsPerMeter = bmp.YPPM
  End With
'set the format of the compressed stream
  res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
  If res <> AVIERR_OK Then
    GoTo error
  End If
  InitAviWriting = True
error:

End Function

Private Function loadPic(sFile As String, _
                         ByVal showIm As Boolean, _
                         ByVal clearImFirst As Boolean, _
                         ByVal saveBMP As Boolean) As String

  Dim xw      As Single
  Dim yh      As Single
  Dim xyi     As Single
  Dim bmpFile As String

  imL.Picture = LoadPicture(sFile)
  If showIm Then
    xyi = imL.ScaleWidth / imL.ScaleHeight
    If clearImFirst Then
      imShow.Picture = LoadPicture("")
      imShow.Refresh
    End If
    With imShow
      If xyi < .ScaleWidth / .ScaleHeight Then
        xw = xyi * .Width
        .PaintPicture imL.Picture, (.Width - xw) / 2, 0, xw, .Height
       Else
        yh = .Height / xyi
        .PaintPicture imL.Picture, 0, (.Height - yh) / 2, .Width, yh
      End If
      .Refresh
    End With 'imShow
  End If
  If saveBMP Then
    If GetFileExtension(sFile, True) <> "bmp" Then
      bmpFile = App.Path & "\temp.bmp"
      SavePicture imL.Picture, bmpFile
     Else
      bmpFile = sFile
    End If
    loadPic = bmpFile
   Else
    loadPic = vbNullString
  End If

End Function

Private Sub lstAVIPlayers_Click()

  StrWMP = Players(lstAVIPlayers.ListIndex)
'this reset the default player to use
  SaveSetting App.EXEName, "OPTIONS", "PlayerPref", lstAVIPlayers.ListIndex

End Sub

Private Function SrcFile() As Boolean

  With cdlCreateAVI
    .DialogTitle = "Choose First Image File in directory"
    .CancelError = False
    .Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
    .DefaultExt = vbNullString
    .FileName = vbNullString
    .ShowOpen
    If Len(.FileName) = 0 Then
      SrcFile = False
     Else
      szInputFile = .FileName
      lblCreateAVI(4).Caption = FlleNameOnly(szInputFile)
      SrcFile = True
    End If
  End With

End Function

':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:32) 78 + 479 = 557 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

