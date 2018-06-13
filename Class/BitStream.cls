VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Bit Stream Class                                                            '
'  [BitStream]                                                                 '
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
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Sub CopyMemory _
        Lib "kernel32" Alias "RtlMoveMemory" ( _
          pDst As Any, _
          pSrc As Any, _
          ByVal ByteLen As Long)


'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


Private m_Stream() As Byte  '[  bit stream array  ]'
Private m_Length   As Long  '[  bit stream length  ]'


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub Clear()
  Erase m_Stream
  m_Length = 0
End Sub


Public Function ExtractBits(ByVal BitCount As Long) As Long

  If BitCount > 32 Then BitCount = 32

  If BitCount <= m_Length Then

    '[  decrement length  ]'
    m_Length = m_Length - BitCount

    ExtractBits = 0

    Dim Ctr As Long
    For Ctr = 1 To BitCount

      If Ctr < 32 Then
        '[  add bit value  ]'
        If m_Stream(m_Length + Ctr) > 0 Then
          ExtractBits = ExtractBits + Int(2 ^ (Ctr - 1))
        End If
      Else
        '[  if negative, compute two's complement  ]'
        If m_Stream(m_Length + Ctr) > 0 Then
          ExtractBits = (Not ExtractBits) + 1
        End If
      End If

    Next Ctr

    '[  resize stream  ]'
    If m_Length > 0 Then
      ReDim Preserve m_Stream(1 To m_Length)
    Else
      Erase m_Stream
    End If

  End If

End Function


Public Function ExtractByte() As Byte

  If m_Length > 7 Then

    '[  decrement length  ]'
    m_Length = m_Length - 8

    '[  add bit values  ]'
    ExtractByte = 0
    If m_Stream(m_Length + 1) > 0 Then ExtractByte = ExtractByte + 1
    If m_Stream(m_Length + 2) > 0 Then ExtractByte = ExtractByte + 2
    If m_Stream(m_Length + 3) > 0 Then ExtractByte = ExtractByte + 4
    If m_Stream(m_Length + 4) > 0 Then ExtractByte = ExtractByte + 8
    If m_Stream(m_Length + 5) > 0 Then ExtractByte = ExtractByte + 16
    If m_Stream(m_Length + 6) > 0 Then ExtractByte = ExtractByte + 32
    If m_Stream(m_Length + 7) > 0 Then ExtractByte = ExtractByte + 64
    If m_Stream(m_Length + 8) > 0 Then ExtractByte = ExtractByte + 128

    '[  resize stream  ]'
    If m_Length > 0 Then
      ReDim Preserve m_Stream(1 To m_Length)
    Else
      Erase m_Stream
    End If

  End If

End Function


Public Function ExtractString(ByVal CharCount As Long) As String
  Dim Ctr As Long
  ExtractString = vbNullString
  For Ctr = 1 To CharCount
    ExtractString = ExtractString & Chr(ExtractByte)
  Next Ctr
End Function


Public Sub InsertBitsAtEnd(ByVal Bits As Long, ByVal BitCount As Long)

  If BitCount > 32 Then BitCount = 32

  '[  resize stream  ]'
  ReDim Preserve m_Stream(1 To (m_Length + BitCount))

  '[  shift bits towards the top  ]'
  If m_Length > 0 Then
    CopyMemory ByVal VarPtr(m_Stream(BitCount + 1)), _
               ByVal VarPtr(m_Stream(1)), _
               m_Length
  End If

  '[  insert bits in the stream  ]'
  Dim Ctr As Long
  For Ctr = 1 To BitCount
    If Ctr < 32 Then
      If (Bits And Int(2 ^ (Ctr - 1))) > 0 Then
        m_Stream(Ctr) = 1
      Else
        m_Stream(Ctr) = 0
      End If
    Else
      If Bits < 0 Then
        m_Stream(Ctr) = 1
      Else
        m_Stream(Ctr) = 0
      End If
    End If
  Next Ctr

  '[  increment length  ]'
  m_Length = m_Length + BitCount

End Sub


Public Sub InsertBitsAtTop(ByVal Bits As Long, ByVal BitCount As Long)

  If BitCount > 32 Then BitCount = 32

  '[  resize stream  ]'
  ReDim Preserve m_Stream(1 To (m_Length + BitCount))

  '[  insert bits in the stream  ]'
  Dim Ctr As Long
  For Ctr = 1 To BitCount
    If Ctr < 32 Then
      If (Bits And Int(2 ^ (Ctr - 1))) > 0 Then
        m_Stream(m_Length + Ctr) = 1
      Else
        m_Stream(m_Length + Ctr) = 0
      End If
    Else
      If Bits < 0 Then
        m_Stream(m_Length + Ctr) = 1
      Else
        m_Stream(m_Length + Ctr) = 0
      End If
    End If
  Next Ctr

  '[  increment length  ]'
  m_Length = m_Length + BitCount

End Sub


Public Sub InsertByteAtEnd(ByVal Bits As Byte)

  '[  resize stream  ]'
  ReDim Preserve m_Stream(1 To (m_Length + 8))

  '[  shift bits towards the top  ]'
  If m_Length > 0 Then
    CopyMemory ByVal VarPtr(m_Stream(9)), ByVal VarPtr(m_Stream(1)), m_Length
  End If

  '[  insert bits in the stream  ]'
  m_Stream(8) = (Bits And &H80) \ &H80
  m_Stream(7) = (Bits And &H40) \ &H40
  m_Stream(6) = (Bits And &H20) \ &H20
  m_Stream(5) = (Bits And &H10) \ &H10
  m_Stream(4) = (Bits And &H8) \ &H8
  m_Stream(3) = (Bits And &H4) \ &H4
  m_Stream(2) = (Bits And &H2) \ &H2
  m_Stream(1) = Bits And &H1

  '[  increment length  ]'
  m_Length = m_Length + 8

End Sub


Public Sub InsertByteAtTop(ByVal Bits As Byte)

  '[  resize stream  ]'
  ReDim Preserve m_Stream(1 To (m_Length + 8))

  '[  insert bits in the stream  ]'
  m_Stream(m_Length + 8) = (Bits And &H80) \ &H80
  m_Stream(m_Length + 7) = (Bits And &H40) \ &H40
  m_Stream(m_Length + 6) = (Bits And &H20) \ &H20
  m_Stream(m_Length + 5) = (Bits And &H10) \ &H10
  m_Stream(m_Length + 4) = (Bits And &H8) \ &H8
  m_Stream(m_Length + 3) = (Bits And &H4) \ &H4
  m_Stream(m_Length + 2) = (Bits And &H2) \ &H2
  m_Stream(m_Length + 1) = Bits And &H1

  '[  increment length  ]'
  m_Length = m_Length + 8

End Sub


Public Sub InsertStringAtEnd(ByVal Bits As String)
  Dim Ctr As Long
  For Ctr = 1 To Len(Bits)
    InsertByteAtEnd Asc(Mid$(Bits, Ctr, 1))
  Next Ctr
End Sub


Public Sub InsertStringAtTop(ByVal Bits As String)
  Dim Ctr As Long
  For Ctr = Len(Bits) To 1 Step -1
    InsertByteAtTop Asc(Mid$(Bits, Ctr, 1))
  Next Ctr
End Sub


'------------------------------------------------------------------------------'
'  Public Properties                                                           '
'------------------------------------------------------------------------------'


Public Property Get Length() As Long
  Length = m_Length
End Property


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub Class_Initialize()
  Clear
End Sub


Private Sub Class_Terminate()
  Clear
End Sub
