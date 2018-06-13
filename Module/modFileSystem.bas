Attribute VB_Name = "modFileSystem"
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  File System Operations Module                                               '
'  [modFileSystem]                                                             '
'                                                                              '
'------------------------------------------------------------------------------'
'                                                                              '
'  Copyright (C) 2003 Mark David Gan                                           '
'                                                                              '
'  This file is part of Chameleon.                                             '
'                                                                              '
'  Chameleon is free software; you can redistribute it and/or modify           '
'  it under the terms of the GNU General Public License as published by        '
'  the Free Software Foundation; either version 2 of the License, or           '
'  (at your option) any later version.                                         '
'                                                                              '
'  Chameleon is distributed in the hope that it will be useful,                '
'  but WITHOUT ANY WARRANTY; without even the implied warranty of              '
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               '
'  GNU General Public License for more details.                                '
'                                                                              '
'  You should have received a copy of the GNU General Public License           '
'  along with Chameleon; if not, write to the Free Software                    '
'  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA   '
'                                                                              '
'------------------------------------------------------------------------------'


Option Explicit


'------------------------------------------------------------------------------'
'  Windows API Structure Data Types                                            '
'------------------------------------------------------------------------------'


'[  64-bit time structure  ]'
Private Type FILETIME
  dwLowDateTime  As Long
  dwHighDateTime As Long
End Type

'[  multiple-field time structure  ]'
Private Type SYSTEMTIME
  wYear         As Integer
  wMonth        As Integer
  wDayOfWeek    As Integer
  wDay          As Integer
  wHour         As Integer
  wMinute       As Integer
  wSecond       As Integer
  wMilliseconds As Integer
End Type


'------------------------------------------------------------------------------'
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constants for "CreateFile" function "dwDesiredAccess" parameter  ]'
Private Const GENERIC_READ  As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000

'[  constants for "CreateFile" function "dwShareMode" parameter  ]'
Private Const FILE_SHARE_READ  As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

'[  constants for "CreateFile" function "dwCreateDisposition" parameter  ]'
Private Const CREATE_ALWAYS     As Long = 2
Private Const CREATE_NEW        As Long = 1
Private Const OPEN_ALWAYS       As Long = 4
Private Const OPEN_EXISTING     As Long = 3
Private Const TRUNCATE_EXISTING As Long = 5

'[  constants for "CreateFile" function "dwFlagsAndAttributes" parameter  ]'
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const FILE_FLAG_NO_BUFFERING    As Long = &H20000000
Private Const FILE_FLAG_OVERLAPPED      As Long = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS As Long = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS   As Long = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH   As Long = &H80000000

'[  constants for "SetFilePointer" function "dwMoveMethod" parameter  ]'
Private Const FILE_BEGIN   As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END     As Long = 2


'------------------------------------------------------------------------------'
'  FileSystemObject Constants                                                  '
'------------------------------------------------------------------------------'


'[  constants for "GetSpecialFolder" function "folderspec" parameter  ]'
Private Const GSF_WINDOWSFOLDER As Long = 0
Private Const GSF_SYSTEMFOLDER As Long = 1
Private Const GSF_TEMPORARYFOLDER As Long = 2


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function CloseHandle _
        Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function CreateFile _
        Lib "kernel32.dll" Alias "CreateFileA" ( _
          ByVal lpFileName As String, _
          ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, _
          ByVal lpSecurityAttributes As Long, _
          ByVal dwCreationDisposition As Long, _
          ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long _
        ) As Long

Private Declare Function SetFilePointer _
        Lib "kernel32.dll" ( _
          ByVal iFileHandler As Long, _
          ByVal lDistanceToMove As Long, _
          ByRef lpDistanceToMoveHigh As Long, _
          ByVal dwMoveMethod As Long _
        ) As Long

Private Declare Function SetFileTime _
        Lib "kernel32" ( _
          ByVal hFile As Long, _
          ByVal lpCreationTime As Long, _
          ByRef lpLastAccessTime As FILETIME, _
          ByRef lpLastWriteTime As FILETIME _
        ) As Long

Private Declare Function SystemTimeToFileTime _
        Lib "kernel32" ( _
          lpSystemTime As SYSTEMTIME, _
          lpFileTime As FILETIME _
        ) As Long

Private Declare Function WriteFile _
        Lib "kernel32.dll" ( _
          ByVal iFileHandler As Long, _
          ByRef lpBuffer As Any, _
          ByVal nNumberOfBytesToWrite As Long, _
          ByRef lpNumberOfBytesWritten As Long, _
          ByVal lpOverlapped As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub CreateFolder(ByVal Path As String)

  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  If Len(FSO.GetParentFolderName(Path)) > 0 Then
    CreateFolder FSO.GetParentFolderName(Path)
  End If

  On Error Resume Next
  MkDir Path

End Sub


Public Function FileExists(ByVal Path As String) As Boolean
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  FileExists = FSO.FileExists(Path)
End Function


Public Function FolderExists(ByVal Path As String) As Boolean
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  FolderExists = FSO.FolderExists(Path)
End Function


Public Function GenerateTempFile() As String

  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  '[  set current drive to that of the windows folder  ]'
  ChDir Left$(FSO.GetSpecialFolder(GSF_WINDOWSFOLDER), 2) & "\"

  '[  get path of temporary folder  ]'
  GenerateTempFile = FSO.GetSpecialFolder(GSF_TEMPORARYFOLDER) & "\" & _
                     UCase$(FSO.GetTempName)

End Function


Public Function GetAbsolutePath(ByVal Path As String) As String

  If Len(Path) = 0 Then
    GetAbsolutePath = ""
    Exit Function
  End If

  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")

  '[  replace illegal characters  ]'
  Path = Replace$(Path, ">", "")
  Path = Replace$(Path, "<", "")
  Path = Replace$(Path, "*", "")
  Path = Replace$(Path, "?", "")
  Path = Replace$(Path, "|", "")
  Path = Replace$(Path, Chr(34), "")

  '[  ensure that the colon character is in place  ]'
  If InStr(1, Path, ":") = 2 Then
    Path = Left$(Path, 2) & Replace$(Path, ":", "", 3)
  Else
    Path = Replace$(Path, ":", "")
  End If

  '[  set current drive to that of the windows folder  ]'
  ChDir Left$(FSO.GetSpecialFolder(GSF_WINDOWSFOLDER), 2) & "\"

  GetAbsolutePath = FSO.GetAbsolutePathName(Path)

End Function


Public Function GetFileDateTime(ByVal FileName As String) As String

  If FileExists(FileName) Then
    GetFileDateTime = FormatDate(FileDateTime(FileName))
  Else
    GetFileDateTime = vbNullString
  End If

End Function


Public Function GetFileSize(ByVal FileName As String) As String

  If FileExists(FileName) Then
    GetFileSize = FormatSize(FileLen(FileName))
  Else
    GetFileSize = vbNullString
  End If

End Function


Public Function GetPathFileName(ByVal Path As String) As String
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  GetPathFileName = FSO.GetFileName(Path)
End Function


Public Function GetPathFolderName(ByVal Path As String) As String
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  GetPathFolderName = FSO.GetParentFolderName(Path)
End Function


Public Function GetPrimaryDrive() As String
  Dim FSO As Object
  Set FSO = CreateObject("Scripting.FileSystemObject")
  GetPrimaryDrive = Left$(FSO.GetSpecialFolder(GSF_WINDOWSFOLDER), 2) & "\"
End Function


Public Function SetFileDate(ByVal FileName As String, _
                            ByVal FileDate As Date) As Boolean

  Dim STime As SYSTEMTIME
  Dim FTime As FILETIME
  Dim hFile As Long

  '[  get data file date and time  ]'
  STime.wYear = Year(FileDate)
  STime.wMonth = Month(FileDate)
  STime.wDay = Day(FileDate)
  STime.wDayOfWeek = Weekday(FileDate) - 1
  STime.wHour = Hour(FileDate)
  STime.wMinute = Minute(FileDate)
  STime.wSecond = Second(FileDate)
  STime.wMilliseconds = 0
  SystemTimeToFileTime STime, FTime

  '[  set file attribute to normal  ]'
  On Error Resume Next
  SetAttr FileName, vbNormal

  '[  open source file  ]'
  hFile = CreateFile(FileName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)

  '[  get file date and time  ]'
  If hFile <> -1 Then
    SetFileTime hFile, 0, FTime, FTime
    CloseHandle hFile
    SetFileDate = True
  Else
    SetFileDate = False
  End If

End Function


Public Function WipeFile(ByVal FileName As String, ByVal Rename As Boolean) _
                         As Boolean

  Dim FileName2    As String
  Dim FileSize     As Long
  Dim hFile        As Long
  Dim Ctr          As Long
  Dim Pattern()    As Byte
  Dim BytesWritten As Long

  On Error Resume Next

  '[  set file attribute to normal and retrieve file size  ]'
  SetAttr FileName, vbNormal
  FileSize = FileLen(FileName)

  '[  rename file and move to windows temporary folder  ]'
  If Rename Then
    FileName2 = GenerateTempFile
    Kill FileName2
    Name FileName As FileName2
  Else
    FileName2 = FileName
  End If

  On Error GoTo FileError

  '[  open file with disk caching disabled  ]'
  hFile = CreateFile(FileName2, GENERIC_WRITE, 0, 0, OPEN_EXISTING, _
                     FILE_FLAG_WRITE_THROUGH + FILE_FLAG_DELETE_ON_CLOSE + _
                     FILE_FLAG_SEQUENTIAL_SCAN, 0)

  '[  if file opened successfully, then wipe file  ]'
  If hFile <> -1 Then

    ReDim Pattern(1 To FileSize, 1 To 3)

    '[  assign bit patterns  ]'
    For Ctr = 1 To FileSize
      Pattern(Ctr, 1) = &H55  '[  bit pattern 01010101  ]'
      Pattern(Ctr, 2) = &HAA  '[  bit pattern 10101010  ]'
      Pattern(Ctr, 3) = &H0   '[  bit pattern 00000000  ]'
    Next Ctr

    '[  write bit patterns to file  ]'
    For Ctr = 1 To 3
      SetFilePointer hFile, 0, 0, FILE_BEGIN
      WriteFile hFile, Pattern(1, Ctr), FileSize, BytesWritten, 0
    Next Ctr

    '[  close and delete file  ]'
    CloseHandle hFile
    WipeFile = True

    Exit Function

  End If

FileError:
  WipeFile = False

End Function
