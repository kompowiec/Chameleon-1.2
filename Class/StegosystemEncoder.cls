VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StegosystemEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Stegosystem Encoder Class                                                   '
'  [StegosystemEncoder]                                                        '
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


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function GetInputState Lib "user32" () As Long

Private Declare Function SystemTimeToFileTime _
        Lib "kernel32" ( _
          lpSystemTime As SYSTEMTIME, _
          lpFileTime As FILETIME _
        ) As Long


'------------------------------------------------------------------------------'
'  Public Events                                                               '
'------------------------------------------------------------------------------'


Public Event Progress(ByVal Processed As Long, ByVal Total As Long)


'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


'[  freeimage wrapper object  ]'
Private Imager As FreeImageWrapper

'[  dib handles  ]'
Private m_CoverImageDIB As Long
Private m_StegoImageDIB As Long

'[  image properties  ]'
Private m_ImageBits()   As Byte
Private m_ImageWidth    As Long
Private m_ImageHeight   As Long
Private m_ImagePixels   As Long
Private m_ImageCapacity As Long

'[  data file properties  ]'
Private m_DataFileName     As String
Private m_DataFileDate     As Date
Private m_DataFileTime     As FILETIME
Private m_DataFileSize     As Long
Private m_DataFileChecksum As String
Private m_DataFileBits()   As Byte

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


Public Property Get CoverImageDIB() As Long
  CoverImageDIB = m_CoverImageDIB
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


Public Property Get ImageCapacity() As Long
  ImageCapacity = m_ImageCapacity
End Property


Public Property Get StegoImageDIB() As Long
  StegoImageDIB = m_StegoImageDIB
End Property


'------------------------------------------------------------------------------'
'  Private Properties                                                          '
'------------------------------------------------------------------------------'


Private Property Get Pixel0() As Long
  '[  color value of current pixel  ]'
  Pixel0 = m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC))
End Property


Private Property Let Pixel0(New_Pixel0 As Long)
  '[  color value of current pixel  ]'
  m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC)) = New_Pixel0
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


Private Property Get Pixel5() As Long
  '[  color value of middle-right pixel  ]'
  Pixel5 = m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC))
End Property


Private Property Let Pixel5(New_Pixel5 As Long)
  '[  color value of middle-right pixel  ]'
  If New_Pixel5 > 255 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC)) = 255
  ElseIf New_Pixel5 < 0 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC)) = 0
  Else
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC)) = New_Pixel5
  End If
End Property


Private Property Get Pixel6() As Long
  '[  color value of current bottom-right pixel  ]'
  Pixel6 = m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC) + 1)
End Property


Private Property Let Pixel6(New_Pixel6 As Long)
  '[  color value of current bottom-right pixel  ]'
  If New_Pixel6 > 255 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC) + 1) = 255
  ElseIf New_Pixel6 < 0 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC) + 1) = 0
  Else
    m_ImageBits(m_PosC, m_PosX(m_PosC) + 1, m_PosY(m_PosC) + 1) = New_Pixel6
  End If
End Property


Private Property Get Pixel7() As Long
  '[  color value of bottom-middle pixel  ]'
  Pixel7 = m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC) + 1)
End Property


Private Property Let Pixel7(New_Pixel7 As Long)
  '[  color value of bottom-middle pixel  ]'
  If New_Pixel7 > 255 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC) + 1) = 255
  ElseIf New_Pixel7 < 0 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC) + 1) = 0
  Else
    m_ImageBits(m_PosC, m_PosX(m_PosC), m_PosY(m_PosC) + 1) = New_Pixel7
  End If
End Property


Private Property Get Pixel8() As Long
  '[  color value of bottom-left pixel  ]'
  Pixel8 = m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC) + 1)
End Property


Private Property Let Pixel8(New_Pixel8 As Long)
  '[  color value of bottom-left pixel  ]'
  If New_Pixel8 > 255 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC) + 1) = 255
  ElseIf New_Pixel8 < 0 Then
    m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC) + 1) = 0
  Else
    m_ImageBits(m_PosC, m_PosX(m_PosC) - 1, m_PosY(m_PosC) + 1) = New_Pixel8
  End If
End Property


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Function Abort()
  m_Abort = True
End Function


Public Function Encode(ByVal PasswordHashMD5 As String, _
                       ByVal PasswordHashSHA As String) As Boolean

  Dim hBmp As Long  '[  temporary bitmap handle  ]'
  Dim Bits As Long  '[  numeric value of bits to be embedded  ]'
  Dim C    As Long  '[  pixel capacity  ]'
  Dim E    As Long  '[  embedding error  ]'

  Dim MetadataCtr As Long  '[  number of metadata bits in the buffer  ]'
  Dim ReadCtr     As Long  '[  number of bytes read from the data file  ]'
  Dim EmbedCtr    As Long  '[  number of bits embedded in the image  ]'
  Dim PixelCtr    As Long  '[  number of pixels processed  ]'
  Dim Interval    As Long  '[  progress event interval  ]'

  Dim Buffer As BitStream  '[  buffer for bits to be embedded  ]'

  '[  compute highest address in image where metadata can fit  ]'
  m_MetadataMaxPosX = m_ImageWidth - METADATA_BITS
  m_MetadataMaxPosY = m_ImageHeight - 1

  Do While (m_MetadataMaxPosX < 0)
    m_MetadataMaxPosX = m_ImageWidth + m_MetadataMaxPosX
    m_MetadataMaxPosY = m_MetadataMaxPosY - 1
  Loop

  '[  set progress event interval  ]'
  If m_ImagePixels < 200 Then
    Interval = 1
  Else
    Interval = m_ImagePixels \ 100
  End If

  '[  create temporary bitmap from the cover image dib  ]'
  hBmp = Imager.CreateBitmapFromDIB(m_CoverImageDIB)

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
  ReadCtr = 0
  EmbedCtr = 0
  PixelCtr = 0

  ResetPixelPosition

  '[  start embedding  ]'
  For PixelCtr = 1 To m_ImagePixels

    If m_Abort Then
      m_Abort = False
      GoTo Cleanup
    End If

    '[  if no metadata is in the buffer  ]'
    If MetadataCtr < 1 Then

      '[  select a random color channel  ]'
      SelectPixelChannel

      '[  insert metadata bits at the start of the bit buffer  ]'
      If (m_PosC = m_MetadataPosC) And _
         (m_PosY(m_PosC) = m_MetadataPosY) And _
         (m_PosX(m_PosC) = m_MetadataPosX) Then
        MetadataCtr = METADATA_BITS
        Buffer.InsertStringAtTop m_DataFileChecksum
        Buffer.InsertBitsAtTop m_DataFileSize, FILESIZE_BITS
        Buffer.InsertBitsAtTop m_DataFileTime.dwLowDateTime, FILEDATELOW_BITS
        Buffer.InsertBitsAtTop m_DataFileTime.dwHighDateTime, FILEDATEHIGH_BITS
        Buffer.InsertStringAtTop m_DataFileName
        Buffer.InsertStringAtTop PasswordHashSHA
      End If

    End If

    '[  get capacity  ]'
    PerformCapacityEvaluation C

    If C > 0 Then

      If Buffer.Length < C Then

        If ReadCtr < m_DataFileSize Then
          '[  insert data bits at the end of the bit buffer  ]'
          ReadCtr = ReadCtr + 1
          Buffer.InsertByteAtEnd m_DataFileBits(ReadCtr)
        Else
          '[  insert random bits at the end of the bit buffer  ]'
          Buffer.InsertByteAtEnd Int(Rnd * 256)
        End If

      End If

      '[  select bits to embed  ] '
      Bits = Buffer.ExtractBits(C)

      '[  embed with minimum error  ] '
      E = PerformMinimumErrorReplacement(Bits, C)

      '[  diffuse embedding error  ] '
      PerformErrorDiffusion E

      '[  increment embedded bits counter  ] '
      EmbedCtr = EmbedCtr + C

      '[  decrement metadata counter  ] '
      If MetadataCtr > 0 Then MetadataCtr = MetadataCtr - C

    End If

    SelectPixelCoordinate

    If (PixelCtr Mod Interval) = 0 Then
      '[  report progress  ]'
      RaiseEvent Progress(PixelCtr, m_ImagePixels)
      DoEvents
    ElseIf GetInputState() Then
      '[  allow other processes to execute  ]'
      DoEvents
    End If

  Next PixelCtr

  '[  report completion  ]'
  If (PixelCtr Mod Interval) > 0 Then
    RaiseEvent Progress(PixelCtr, m_ImagePixels)
  End If

  '[  update temporary bitmap with contents of pixel array  ]'
  Imager.SetBitmapPixels hBmp, m_ImageBits

  '[  create stego dib from temporary bitmap  ]'
  m_StegoImageDIB = Imager.CreateDIBFromBitmap(hBmp)

  '[  compute net image capacity in bytes  ]'
  m_ImageCapacity = (EmbedCtr - METADATA_BITS) \ 8

  '[  return whether all data file bits have been embedded  ]'
  Encode = (m_ImageCapacity >= m_DataFileSize)

Cleanup:

  '[  unload temporary bitmap  ]'
  Imager.DeleteBitmap hBmp
  Erase m_ImageBits

End Function


Public Function LoadCoverImage(ByVal FileName As String) As Boolean

  '[  create new dib for cover image  ]'
  m_CoverImageDIB = Imager.LoadDIB(FileName)

  If m_CoverImageDIB = 0 Then
    LoadCoverImage = False
  Else
    '[  get cover image dimensions  ]'
    m_ImageWidth = Imager.GetWidth(m_CoverImageDIB)
    m_ImageHeight = Imager.GetHeight(m_CoverImageDIB)
    m_ImagePixels = m_ImageWidth * m_ImageHeight * 3
    LoadCoverImage = (m_ImagePixels > METADATA_BITS)
  End If

  m_ImageCapacity = 0

End Function


Public Function LoadDataFile(ByVal FileName As String, _
                             ByVal StoredFileName As String, _
                             ByVal StoredFileDate As Date, _
                             ByVal Checksum As String)

  Dim FileIdx As Long
  Dim SysTime As SYSTEMTIME

  '[  get data file stored file name  ]'
  m_DataFileName = Format$(StoredFileName, "!" & String$(255, "@"))

  '[  get data file date and time  ]'
  m_DataFileDate = StoredFileDate
  SysTime.wYear = Year(m_DataFileDate)
  SysTime.wMonth = Month(m_DataFileDate)
  SysTime.wDay = Day(m_DataFileDate)
  SysTime.wDayOfWeek = Weekday(m_DataFileDate) - 1
  SysTime.wHour = Hour(m_DataFileDate)
  SysTime.wMinute = Minute(m_DataFileDate)
  SysTime.wSecond = Second(m_DataFileDate)
  SysTime.wMilliseconds = 0
  SystemTimeToFileTime SysTime, m_DataFileTime

  '[  get data file checksum  ]'
  m_DataFileChecksum = Checksum

  '[  open data file  ]'
  FileIdx = FreeFile
  On Error GoTo FileError
  Open FileName For Binary Access Read As FileIdx

  '[  get data file size  ]'
  m_DataFileSize = LOF(FileIdx)

  '[  copy data file bits to data array  ]'
  ReDim m_DataFileBits(1 To m_DataFileSize)
  Get FileIdx, , m_DataFileBits

  '[  close data file  ]'
  Close FileIdx

  LoadDataFile = True
  Exit Function

FileError:
  LoadDataFile = False

End Function


Public Function SaveStegoImage(ByVal FileName As String, _
                               ByVal Format As FREE_IMAGE_FORMAT) As Boolean
  SaveStegoImage = Imager.SaveDIB(m_StegoImageDIB, FileName, Format)
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


Private Sub PerformErrorDiffusion(ByVal EmbeddingError As Long)

  If EmbeddingError > 3 Then

    Dim E As Long

    '[  get fraction of embedding error  ]'
    E = EmbeddingError \ 4

    '[  distribute embedding error among neighboring pixels  ]'
    If m_PosY(m_PosC) < (m_ImageHeight - 1) Then

      If m_PosX(m_PosC) < (m_ImageWidth - 1) Then
        Pixel5 = Pixel5 - E
        Pixel6 = Pixel6 - E
      End If

      Pixel7 = Pixel7 - E

      If m_PosX(m_PosC) > 0 Then
        Pixel8 = Pixel8 - E
      End If

    ElseIf m_PosX(m_PosC) < (m_ImageWidth - 1) Then

      Pixel5 = Pixel5 - E

    End If

  End If

End Sub


Private Function PerformMinimumErrorReplacement(ByVal Bits As Long, _
                                                ByVal Capacity As Long)

  Dim Mask   As Long
  Dim Value0 As Long
  Dim Value1 As Long
  Dim Value2 As Long
  Dim Error1 As Long
  Dim Error2 As Long

  Value0 = Pixel0

  '[  compute new color value with rightmost Unreplaced bit set to 0  ]'
  Mask = Not ((2 ^ (Capacity + 1)) - 1)
  Value1 = (Value0 And Mask) Or Bits
  Error1 = Value1 - Value0

  '[  compute new color value with rightmost Unreplaced bit set to 1  ]'
  Value2 = Value1 Or (2 ^ Capacity)
  Error2 = Value2 - Value0

  '[  select color value with less embedding Error  ]'
  If Abs(Error1) < Abs(Error2) Then
    Pixel0 = Value1
    PerformMinimumErrorReplacement = Error1
  Else
    Pixel0 = Value2
    PerformMinimumErrorReplacement = Error2
  End If

End Function


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

  '[  if x exceeds width, then increment vertical position and reset x  ]'
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
  m_CoverImageDIB = 0
  m_StegoImageDIB = 0
  m_ImageCapacity = 0
  m_DataFileSize = 0
  m_Abort = False
End Sub


Private Sub Class_Terminate()
  On Error Resume Next
  Imager.UnloadDIB m_CoverImageDIB
  Imager.UnloadDIB m_StegoImageDIB
  Set Imager = Nothing
End Sub
