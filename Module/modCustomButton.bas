Attribute VB_Name = "modCustomButton"
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Custom Button Subclassing Module                                            '
'  [modCustomButton]                                                           '
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
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constants for "SetWindowLong" function "nIndex" parameter  ]'
Private Const GWL_WNDPROC As Long = (-4)

'[  constants for "TRACTMOUSEEVENTTYPE" structure "dwFlags" member  ]'
Private Const TME_HOVER  As Long = &H1
Private Const TME_LEAVE  As Long = &H2
Private Const TME_QUERY  As Long = &H40000000
Private Const TME_CANCEL As Long = &H80000000

'[  window messages  ]'
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN   As Long = &H201
Private Const WM_LBUTTONUP     As Long = &H202
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_MOUSEHOVER    As Long = &H2A1
Private Const WM_MOUSELEAVE    As Long = &H2A3
Private Const WM_MOUSEMOVE     As Long = &H200
Private Const WM_MOUSEWHEEL    As Long = &H20A


'------------------------------------------------------------------------------'
'  Public Structure Data Types                                                 '
'------------------------------------------------------------------------------'


Private Type TRACKMOUSEEVENTTYPE
  cbSize      As Long
  dwFlags     As Long
  hwndTrack   As Long
  dwHoverTime As Long
End Type


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function CallWindowProc _
                Lib "user32" Alias "CallWindowProcA" ( _
                  ByVal lpPrevWndFunc As Long, _
                  ByVal hWnd As Long, _
                  ByVal Msg As Long, _
                  ByVal wParam As Long, _
                  ByVal lParam As Long _
                ) As Long

Private Declare Function SetWindowLong _
                Lib "user32" Alias "SetWindowLongA" ( _
                  ByVal hWnd As Long, _
                  ByVal nIndex As Long, _
                  ByVal dwNewLong As Long _
                ) As Long

Private Declare Function TrackMouseEvent _
                Lib "user32" ( _
                  lpEventTrack As TRACKMOUSEEVENTTYPE _
                ) As Long



'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


Private m_ButtonCollection As Collection


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Function NewWndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim CButton1 As CustomButton

  On Error Resume Next

  If Not m_ButtonCollection Is Nothing Then

    Set CButton1 = m_ButtonCollection.Item("hWnd: " & hWnd)

    '[  test for mouse leave event  ]'
    If CButton1.Enabled Then
      Select Case uMsg

        Case WM_MOUSELEAVE:
          CButton1.State = cbtnStateNormal
          CButton1.RaiseMouseLeaveEvent

      End Select
    End If

    '[  call original window procedure  ]'
    NewWndProc = CallWindowProc(CButton1.OldWndProc, hWnd, uMsg, wParam, lParam)

  End If

End Function


Public Sub StartSubclassingButton(ByVal CButton1 As CustomButton)

  If CButton1.OldWndProc = 0 Then

    '[  create button collection  ] '
    If m_ButtonCollection Is Nothing Then
      Set m_ButtonCollection = New Collection
    End If

    '[  subclass button  ]'
    m_ButtonCollection.Add CButton1, "hWnd: " & CButton1.hWnd
    CButton1.OldWndProc = SetWindowLong(CButton1.hWnd, GWL_WNDPROC, _
                                        AddressOf NewWndProc)

  End If

End Sub


Public Sub StopSubclassingButton(ByVal CButton1 As CustomButton)

  '[  reconnect button to its original window procedure  ]'
  If CButton1.OldWndProc <> 0 Then
    SetWindowLong CButton1.hWnd, GWL_WNDPROC, CButton1.OldWndProc
    CButton1.OldWndProc = 0
    m_ButtonCollection.Remove "hWnd: " & CButton1.hWnd
  End If

End Sub


Public Sub TrackMouseLeaveEvent(ByVal CButton1 As CustomButton)

  Dim TMEvent As TRACKMOUSEEVENTTYPE

  With TMEvent
    .cbSize = Len(TMEvent)
    .hwndTrack = CButton1.hWnd
    .dwFlags = TME_LEAVE
  End With

  TrackMouseEvent TMEvent

End Sub
