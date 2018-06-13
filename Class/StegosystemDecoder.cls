VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StegosystemDecoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Stegosystem Decoder Class                                                   '
'  [StegosystemDecoder]                                                        '
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


'------------------------------------------------------------------------------'
'                                                                              '
'  Requires:  FreeImageWrapper.cls                                             '
'             FreeImage.dll                                                    '
'                                                                              '
'------------------------------------------------------------------------------'


'------------------------------------------------------------------------------'
'                                                                              '
'  Metadata Format:                                                            '
'    Password SHA Hash Value - String   - 20 Bytes                             '
'    Data File Name          - String   - 255 Bytes                            '
'    Data File Date & Time   - FILETIME - 8 Bytes                              '
'    Data File Size          - Long     - 4 Bytes                              '
'    Data File MD5 Checksum  - String   - 16 Bytes                             '
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
'  Private Constants                                                           '
'------------------------------------------------------------------------------'


'[  metadata bit length constants  ]'
Private Const STOREDPASSWORD_BITS As Long = 160
Private Const FILENAME_BITS       As Long = 2040
Private Const FILEDATEHIGH_BITS   As Long = 32
Private Const FILEDATELOW_BITS    As Long = 32
Private Const FILESIZE_BITS       As Long = 32
Private Const FILECHECKSUM_BITS   As Long = 128
Private Const METADATA_BITS       As Long = STOREDPASSWORD_BITS + _
                                            FILENAME_BITS + _
                                            FILEDATEHIGH_BITS + _
                                            FILEDATELOW_BITS + _
                                            FILESIZE_BITS + _
                                            FILECHECKSUM_BITS

'[  metadata byte length constants  ]'
Private Const STOREDPASSWORD_BYTES As Long = 20
Private Const FILENAME_BYTES       As Long = 255
Private Const FILEDATEHIGH_BYTES   As Long = 4
Private Const FILEDATELOW_BYTES    As Long = 4
Private Const FILESIZE_BYTES       As Long = 4
Private Const FILECHECKSUM_BYTES   As Long = 16
Private Const METADATA_BYTES       As Long = STOREDPASSWORD_BYTES + _
                                             FILENAME_BYTES + _
                                             FILEDATEHIGH_BYTES + _
                                             FILEDATELOW_BYTES + _
                                             FILESIZE_BYTES + _
                                             FILECHECKSUM_BYTES

'[  metadata progress report interval  ]'
Private Const METADATA_INTERVAL As Long = METADATA_BITS \ 100


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function FileTimeToSystemTime _
        Lib "kernel32" ( _
          ByRef lpFileTime As FILETIME, _
          ByRef lpSystemTime As SYSTEMTIME _
        ) As Long

Private Declare Function GetInputState Lib "user32" () As Long


'------------------------------------------------------------------------------'
'  Public Events                                                               '
'------------------------------------------------------------------------------'


Public Event DataFileProgress(ByVal Processed As Long, ByVal Total As Long)
Public Event MetadataProgress(ByVal Processed As Long, ByVal Total As Long)


'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


'[  freeimage wrapper object  ]'
Private Imager As FreeImageWrapper

'[  dib handles  ]'
Private m_StegoImageDIB As Long

'[  image properties  ]'
Private m_ImageBits() As Byte
Private m_ImageWidth  As Long
Private m_ImageHeight As Long
Private m_ImagePixels As Long

'[  data file properties  ]'
Private m_DataFileName     As String
Private m_DataFileDate     As Date
Private m_DataFileTime     As FILETIME
Private m_DataFileSize     As Long
Private m_DataFileChecksum As String
Private m_DataFileBits()   As Byte

'[  password hash value stored with metadata  ]'
Private m_StoredPassword As String

'[  highest suitable address for metadata  ]'
Private m_MetadataMaxPosX As Long
Private m_MetadataMaxPosY As Long

'[  starting address for metadata (color channel, horizontal, vertical)  ]'
Private m_MetadataPosC As Long
Private m_MetadataPosX As Long
Private m_MetadataPosY As Long

'[  current pixel position pointers (color channel, horizontal, vertical)  ]'
Private m_PosC         As Long
Private m_PosX(0 To 2) As Long
Private m_PosY(0 To 2) As Long

'[  abort process flag  ]'
Private m_Abort As Boolean


'------------------------------------------------------------------------------'
'  Public Properties                                                           '
'------------------------------------------------------------------------------'


Public Property Get DataFileChecksum() As String
  DataFileChecksum = m_DataFileChecksum
End Property


Public Property Get DataFileDate() As Date
  DataFileDate = m_DataFileDate
End Property


Public Property Get DataFileName() As String
  DataFileName = m_DataFileName
End Property


Public Property Get DataFileSize() As Long
  DataFileSize = m_DataFileSize
End Property


Public Property Get StegoImageDIB() As Long
  StegoImageDIB = m_StegoImageDIB
End Property


Public Property Get StoredPassword() As String
  StoredPassword = m_StoredPassword
End Property


'------------------------------------------------------------------------------'
'  Private Properties                                                          '
'------------------------------------------------------------------------------'


Private Property Get Pixel0() As Long
  '[  color value of current pixel  ]'
  Pixel0 = m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC))
End Property


Private Property Get Pixel1() As Long
  '[  color value of middle-left pixel  ]'
  Pixel1 = m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC))
End Property


Private Property Get Pixel2() As Long
  '[  color value of top-left pixel  ]'
  Pixel2 = m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC) - 1)
End Property


Private Property Get Pixel3() As Long
  '[  color value of top-middle pixel  ]'
  Pixel3 = m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC) - 1)
End Property


Private Property Get Pixel4() As Long
  '[  color value of top-right pixel  ]'
  Pixel4 = m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC) - 1)
End Property


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Function Abort()
  m_Abort = True
End Function


Public Function Decode(ByVal PasswordHashMD5 As String) As Boolean

  Dim hBmp As Long  '[  temporary bitmap handle  ]'
  Dim Bits As Long  '[  numeric value of bits extracted from pixel  ]'
  Dim C    As Long  '[  pixel capacity  ]'

  Dim MetadataCtr As Long  '[  number of extracted metadata bits  ]'
  Dim WriteCtr    As Long  '[  number of bits written on the data array  ]'
  Dim Interval    As Long  '[  progress event interval  ]'

  Dim Buffer As BitStream  '[  buffer for extracted bits  ]'

  '[  compute highest address in image where metadata can fit  ]'
  m_MetadataMaxPosX = m_ImageWidth - METADATA_BITS
  m_MetadataMaxPosY = m_ImageHeight - 1

  Do While (m_MetadataMaxPosX < 0)
    m_MetadataMaxPosX = m_ImageWidth + m_MetadataMaxPosX
    m_MetadataMaxPosY = m_MetadataMaxPosY - 1
  Loop

  '[  set progress event interval  ]'
  If m_DataFileSize < 200 Then
    Interval = 1
  Else
    Interval = m_DataFileSize \ 100
  End If

  '[  create temporary bitmap from the stego image dib  ]'
  hBmp = Imager.CreateBitmapFromDIB(m_StegoImageDIB)

  '[  copy bitmap bits to pixel array  ]'
  Imager.GetBitmapPixels hBmp, m_ImageBits

  '[  initialize prng with MD5 hash of password  ]'
  InitializePRNG PasswordHashMD5

  '[  select starting address for the metadata bits  ]'
  m_MetadataPosC = Int(Rnd * 3)
  m_MetadataPosX = Int(Rnd * (m_MetadataMaxPosX + 1))
  m_MetadataPosY = Int(Rnd * (m_MetadataMaxPosY + 1))

  '[  initialize bit buffer  ]'
  Set Buffer = New BitStream

  '[  clear and resize data array  ]'
  ReDim m_DataFileBits(1 To m_DataFileSize)

  '[  reset counters  ]'
  MetadataCtr = 0
  WriteCtr = 0

  ResetPixelPosition

  '[  start extraction  ]'
  Do While WriteCtr < m_DataFileSize

    If m_Abort Then
      m_Abort = False
      GoTo Cleanup
    End If

    '[  if no metadata is in the buffer  ]'
    If MetadataCtr < 1 Then

      '[  select a random color channel  ]'
      SelectPixelChannel

      '[  initialize countdown variable for metadata  ]'
      If (m_PosC = m_MetadataPosC) And _
         (m_PosY(m_PosC) = m_MetadataPosY) And _
         (m_PosX(m_PosC) = m_MetadataPosX) Then
        MetadataCtr = METADATA_BITS
      End If

    End If

    '[  get capacity  ]'
    PerformCapacityEvaluation C

    If C > 0 Then

      '[  store extracted bits in the bit buffer  ]'
      Bits = Pixel0 And ((2 ^ C) - 1)

      If MetadataCtr > 0 Then
        MetadataCtr = MetadataCtr - C
        If MetadataCtr < 0 Then
          C = Abs(MetadataCtr)
          Buffer.InsertBitsAtEnd (Bits And ((2 ^ C) - 1)), C
          MetadataCtr = 0
        End If
      Else
        Buffer.InsertBitsAtEnd Bits, C
      End If

      Do While Buffer.Length > 7
        WriteCtr = WriteCtr + 1
        m_DataFileBits(WriteCtr) = CByte(Buffer.ExtractByte)
      Loop

    End If

    SelectPixelCoordinate

    If (WriteCtr Mod Interval) = 0 Then
      '[  signal progress  ]'
      RaiseEvent DataFileProgress(WriteCtr, m_DataFileSize)
      DoEvents
    ElseIf GetInputState() Then
      '[  allow other processes to execute  ]'
      DoEvents
    End If

  Loop

  '[  signal completion  ]'
  If (WriteCtr Mod Interval) > 0 Then
    RaiseEvent DataFileProgress(WriteCtr, m_DataFileSize)
  End If

  '[  return whether all bits have been extracted  ]'
  Decode = (WriteCtr >= m_DataFileSize)

Cleanup:

  '[  cleanup temporary bitmap  ]'
  Imager.DeleteBitmap hBmp
  Erase m_ImageBits

End Function


Public Sub DecodeMetadata(ByVal PasswordHashMD5 As String)

  Dim hBmp    As Long        '[  temporary bitmap handle  ]'
  Dim Bits    As Long        '[  numeric value of bits extracted from pixel  ]'
  Dim C       As Long        '[  pixel capacity  ]'
  Dim Buffer  As BitStream   '[  buffer for extracted bits  ]'
  Dim SysTime As SYSTEMTIME  '[  temporary variable for file date and time  ]'

  Dim MetadataCtr As Long  '[  number of extracted metadata bits  ]'

  '[  compute highest address in image where metadata can fit  ]'
  m_MetadataMaxPosX = m_ImageWidth - METADATA_BITS
  m_MetadataMaxPosY = m_ImageHeight - 1

  Do While (m_MetadataMaxPosX < 0)
    m_MetadataMaxPosX = m_ImageWidth + m_MetadataMaxPosX
    m_MetadataMaxPosY = m_MetadataMaxPosY - 1
  Loop

  '[  create temporary bitmap from the stego image dib  ]'
  hBmp = Imager.CreateBitmapFromDIB(m_StegoImageDIB)

  '[  copy bitmap bits to pixel array  ]'
  Imager.GetBitmapPixels hBmp, m_ImageBits

  '[  initialize prng with MD5 hash of password  ]'
  InitializePRNG PasswordHashMD5

  '[  select starting address for the metadata bits  ]'
  m_MetadataPosC = Int(Rnd * 3)
  m_MetadataPosX = Int(Rnd * (m_MetadataMaxPosX + 1))
  m_MetadataPosY = Int(Rnd * (m_MetadataMaxPosY + 1))

  '[  initialize bit buffer  ]'
  Set Buffer = New BitStream

  '[  reset counters  ]'
  MetadataCtr = 0

  '[  set pixel position to starting address of metadata bits  ]'
  m_PosC = m_MetadataPosC
  m_PosX(m_PosC) = m_MetadataPosX
  m_PosY(m_PosC) = m_MetadataPosY

  '[  extract metadata bits From the stego image  ]'
  Do While MetadataCtr < METADATA_BITS

    If m_Abort Then
      m_Abort = False
      GoTo Cleanup
    End If

    '[  get capacity  ]'
    PerformCapacityEvaluation C

    If C > 0 Then

      '[  store extracted metadata bits in the bit buffer  ]'
      Bits = Pixel0 And ((2 ^ C) - 1)
      Buffer.InsertBitsAtEnd Bits, C

      MetadataCtr = MetadataCtr + C
      If MetadataCtr > METADATA_BITS Then MetadataCtr = METADATA_BITS

      If (MetadataCtr Mod METADATA_INTERVAL) = 0 Then
        '[  report progress  ]'
        RaiseEvent MetadataProgress(MetadataCtr, METADATA_BITS)
        DoEvents
      ElseIf GetInputState() Then
        '[  allow other processes to execute  ]'
        DoEvents
      End If

    End If

    SelectPixelCoordinate

  Loop

  '[  report completion  ]'
  If (MetadataCtr Mod METADATA_INTERVAL) > 0 Then
    RaiseEvent MetadataProgress(MetadataCtr, METADATA_BITS)
  End If

  On Error Resume Next

  '[  get information from metadata  ]'
  m_StoredPassword = Buffer.ExtractString(STOREDPASSWORD_BYTES)
  m_DataFileName = Trim$(Buffer.ExtractString(FILENAME_BYTES))
  m_DataFileTime.dwHighDateTime = Buffer.ExtractBits(FILEDATEHIGH_BITS)
  m_DataFileTime.dwLowDateTime = Buffer.ExtractBits(FILEDATELOW_BITS)
  FileTimeToSystemTime m_DataFileTime, SysTime
  m_DataFileDate = DateSerial(SysTime.wYear, SysTime.wMonth, SysTime.wDay) + _
                   TimeSerial(SysTime.wHour, SysTime.wMinute, SysTime.wSecond)
  m_DataFileSize = Buffer.ExtractBits(FILESIZE_BITS)
  m_DataFileChecksum = Buffer.ExtractString(FILECHECKSUM_BYTES)

Cleanup:

  '[  cleanup temporary bitmap  ]'
  Imager.DeleteBitmap hBmp
  Erase m_ImageBits

End Sub


Public Function LoadStegoImage(ByVal FileName As String) As Boolean

  '[  create new dib for stego image  ]'
  m_StegoImageDIB = Imager.LoadDIB(FileName)

  If m_StegoImageDIB = 0 Then
    LoadStegoImage = False
  Else
    '[  get stego image dimensions  ]'
    m_ImageWidth = Imager.GetWidth(m_StegoImageDIB)
    m_ImageHeight = Imager.GetHeight(m_StegoImageDIB)
    m_ImagePixels = m_ImageWidth * m_ImageHeight * 3
    LoadStegoImage = (m_ImagePixels > METADATA_BITS)
  End If

End Function


Public Function SaveDataFile(ByVal FileName As String) As Boolean

  Dim FileIdx As Long
  On Error GoTo FileError

  '[  open data file  ]'
  FileIdx = FreeFile
  Open FileName For Binary Access Write As FileIdx

  '[  copy data file bits from data array  ]'
  Put FileIdx, , m_DataFileBits

  '[  close data file  ]'
  Close FileIdx

  SaveDataFile = True
  Exit Function

FileError:
  SaveDataFile = False

End Function


'------------------------------------------------------------------------------'
'  Private Procedures                                                          '
'------------------------------------------------------------------------------'


Private Sub InitializePRNG(ByVal Seed As String)

  Dim Ctr As Long
  Dim Tmp As Single

  Rnd -1
  Randomize AscB(Left$(Seed, 1))

  For Ctr = 2 To Len(Seed)
    Tmp = Rnd
    Rnd -1
    Randomize Int(AscB(Mid$(Seed, Ctr, 1)) * Tmp)
  Next Ctr

End Sub


Private Sub PerformCapacityEvaluation(ByRef Capacity As Long)

  '[  if current pixel is not at topmost row, then compute capacity  ]'
  If (m_PosY(m_PosC) > 0) Then

    Dim Dif1 As Long
    Dim Dif2 As Long
    Dim Dif3 As Long
    Dim Dif4 As Long
    Dim Avg  As Long

    '[  get difference of adjacent pairs formed by pixel1, pixel2, & pixel3  ]'
    If m_PosX(m_PosC) > 0 Then
      Dif1 = Abs(Pixel3 - Pixel1)
      Dif2 = Abs(Pixel1 - Pixel2)
      Dif3 = Abs(Pixel2 - Pixel3)
    Else
      Dif1 = 0
      Dif2 = 0
      Dif3 = 0
    End If

    '[  get difference of pixel3 and pixel4  ]'
    If m_PosX(m_PosC) < (m_ImageWidth - 1) Then
      Dif4 = Abs(Pixel3 - Pixel4)
    Else
      Dif4 = 0
    End If

    '[  compute average difference in color intensity  ]'
    Avg = Round((Dif1 + Dif2 + Dif3 + Dif4) / 4)

    '[  if c would not be equal to 0, compute hiding capacity of pixel  ]'
    If Avg > 1 Then

      Dim C As Long
      Dim U As Long

      '[  compute capacity (base-2 logarithm of average difference)  ]'
      C = Int(Log(Avg) / Log(2))

      If C > 4 Then

        '[  compute upper boundary  ]'
        If Pixel0 > 191 Then
          U = 5
        Else
          U = 4
        End If

        '[  limit capacity by the upper boundary  ]'
        If C < U Then
          Capacity = C
        Else
          Capacity = U
        End If

      Else

        Capacity = C

      End If

    '[  if pixel in color channel of reserved area for metadata  ]'
    ElseIf m_PosC = m_MetadataPosC Then

      If m_PosY(m_PosC) > m_MetadataMaxPosY Then
        Capacity = 1
      ElseIf (m_PosY(m_PosC) = m_MetadataMaxPosY) And _
             (m_PosX(m_PosC) >= m_MetadataMaxPosX) Then
        Capacity = 1
      Else
        Capacity = 0
      End If

    Else

      Capacity = 0

    End If

  '[  if current pixel is at topmost row  ]'
  Else

    '[  if pixel in color channel of reserved area for metadata  ]'
    If m_PosC = m_MetadataPosC Then

      If m_PosY(m_PosC) > m_MetadataMaxPosY Then
        Capacity = 1
      ElseIf (m_PosY(m_PosC) = m_MetadataMaxPosY) And _
             (m_PosX(m_PosC) >= m_MetadataMaxPosX) Then
        Capacity = 1
      Else
        Capacity = 0
      End If

    Else

      Capacity = 0

    End If

  End If

End Sub


Private Sub ResetPixelPosition()
  m_PosC = 0
  Erase m_PosX
  Erase m_PosY
End Sub


Private Sub SelectPixelChannel()

  '[  select random color channel  ]'
  m_PosC = Int(Rnd * 3)

  '[  if selected color channel is full, then select next  ]'
  If m_PosY(m_PosC) >= m_ImageHeight Then
    m_PosC = (m_PosC + 1) Mod 3
    If m_PosY(m_PosC) >= m_ImageHeight Then
      m_PosC = (m_PosC + 1) Mod 3
    End If
  End If

End Sub


Private Sub SelectPixelCoordinate()

  '[  increment horizontal position  ]'
  m_PosX(m_PosC) = m_PosX(m_PosC) + 1

  '[  increment vertical position  ]'
  If m_PosX(m_PosC) >= m_ImageWidth Then
    m_PosX(m_PosC) = 0
    m_PosY(m_PosC) = m_PosY(m_PosC) + 1
  End If

End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub Class_Initialize()
  Set Imager = New FreeImageWrapper
  m_StegoImageDIB = 0
  m_Abort = False
End Sub


Private Sub Class_Terminate()
  On Error Resume Next
  Imager.UnloadDIB m_StegoImageDIB
  Set Imager = Nothing
End Sub
