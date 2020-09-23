Attribute VB_Name = "ModFindFile"
Option Explicit
'replaced Ron Hoebe's mdFile.bas with this
'it used the FileSystemObject to manipulate file paths
'AND
' modifed from NeO78's  'PDF Printer Class'
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61936&lngWId=1
'code to search for specific files on disks
Private bFound     As Boolean
Public fso         As FileSystemObject


Public Function FileExists(strFullPath As String) As Boolean

  FileExists = fso.FileExists(strFullPath)

End Function

Public Function FileFind(ByVal sPathBegin As String, _
                         ByVal strfilename As String, _
                         iIndexFolder As Long) As String

'This is relavively fast because you can restrict the search range using the sPathBegin parameter

  Dim oRep    As Folder
  Dim oSubRep As Folder
  Dim oFolder As Folder
  Dim oFiles  As Object

  If iIndexFolder = 0 Then
    bFound = False
  End If
  Set oRep = fso.GetFolder(sPathBegin)
  For Each oFolder In oRep.SubFolders
    iIndexFolder = iIndexFolder + 1
    If oFolder.Attributes <> 22 Then
      For Each oFiles In oFolder.Files
        If InStr(1, oFiles.Path, strfilename) <> 0 Then
          FileFind = oFiles.Path
          bFound = True
          Exit For
        End If
      Next oFiles
    End If
    If bFound Then
      Exit For
    End If
  Next oFolder
  For Each oSubRep In oRep.SubFolders
    If bFound Then
      Exit For
    End If
    FileFind = FileFind(oSubRep.Path, strfilename, iIndexFolder)
  Next oSubRep

End Function

Public Function FlleNameOnly(strFullPath As String) As String

  FlleNameOnly = fso.GetFileName(strFullPath)

End Function

Public Function GetFileExtension(ByVal strfilename As String, _
                                 Optional ByVal bLcase As Boolean = False) As String

  GetFileExtension = fso.GetExtensionName(strfilename)
  If bLcase Then
    GetFileExtension = LCase$(GetFileExtension)
  End If

End Function

Public Function GetFilePath(strfilename As String) As String

  GetFilePath = fso.GetParentFolderName(strfilename)

End Function

Public Function SystemFilePathScan(strfilename As String) As String

'this is slower as it has to scan all disks

  Dim iFolder As Long
  Dim D       As Drive

  For Each D In fso.Drives
    If D.IsReady Then
      SystemFilePathScan = FileFind(D.DriveLetter & ":", strfilename, iFolder)
      If Len(SystemFilePathScan) Then
        Exit For
      End If
    End If
  Next D

End Function

':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:37) 9 + 85 = 94 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

