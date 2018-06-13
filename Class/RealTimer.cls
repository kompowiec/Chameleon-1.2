VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  High Frequency Timer Class                                                  '
'  [RealTimer]                                                                 '
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
'  Windows Timer Function Declarations                                         '
'------------------------------------------------------------------------------'


Private Declare Function QueryPerformanceCounter _
        Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency _
        Lib "kernel32" (lpFrequency As Currency) As Long


'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


Private m_Elapsed   As Currency
Private m_Frequency As Currency
Private m_Start     As Currency
Private m_Stop      As Currency
Private m_Supported As Boolean


'------------------------------------------------------------------------------'
'  Public Properties                                                           '
'------------------------------------------------------------------------------'


Public Property Get CounterSupported() As Boolean
  CounterSupported = m_Supported
End Property


Public Property Get Frequency() As Currency
  Frequency = m_Frequency
End Property


Public Property Get ElapsedTime() As Currency
  ElapsedTime = m_Elapsed
End Property


Public Property Get ElapsedTimeInHours() As String
  ElapsedTimeInHours = Format$((m_Elapsed \ 360000) Mod 100, "#00:") & _
                       Format$((m_Elapsed \ 60000) Mod 60, "00:") & _
                       Format$((m_Elapsed \ 1000) Mod 60, "00.") & _
                       Format$(m_Elapsed Mod 1000, "000")
End Property


Public Property Get ElapsedTimeInMinutes() As String
  ElapsedTimeInMinutes = Format$((m_Elapsed \ 60000) Mod 100, "#00:") & _
                         Format$((m_Elapsed \ 1000) Mod 60, "00.") & _
                         Format$(m_Elapsed Mod 1000, "000")
End Property


Public Property Get ElapsedTimeInSeconds() As String
  ElapsedTimeInSeconds = Format$((m_Elapsed \ 1000) Mod 100, "#00.") & _
                         Format$(m_Elapsed Mod 1000, "000")
End Property


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub Reset()
  QueryPerformanceCounter m_Start
End Sub


Public Sub Mark()
  QueryPerformanceCounter m_Stop
  m_Elapsed = ((m_Stop - m_Start) / m_Frequency) * 1000
End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub Class_Initialize()
  m_Supported = QueryPerformanceFrequency(m_Frequency)
  QueryPerformanceCounter m_Start
End Sub
