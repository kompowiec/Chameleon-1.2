VERSION 5.00
Begin VB.UserControl CustomButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ClipBehavior    =   0  'None
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
End
Attribute VB_Name = "CustomButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Custom Button User Control                                                  '
'  [CustomButton]                                                              '
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
'  Requires:  modCustomButton.bas                                              '
'                                                                              '
'------------------------------------------------------------------------------'


Option Explicit


'------------------------------------------------------------------------------'
'  Windows API Structure Data Types                                            '
'------------------------------------------------------------------------------'


'[  pixel coordinates structure  ]'
Private Type POINTAPI
  X As Long
  Y As Long
End Type

'[  rectangular coordinates structure  ]'
Private Type RECT
  Left As Long
  top As Long
  Right As Long
  Bottom As Long
End Type

'[  font layout information  ]'
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type


'------------------------------------------------------------------------------'
'  Public Enumerated Data Types                                                '
'------------------------------------------------------------------------------'


'[  mask style enumeration  ]'
Public Enum cbtnMaskStyleConstants
  cbtnMaskStyleNone = 0
  cbtnMaskStyleAuto = 1
  cbtnMaskStyleCustom = 2
End Enum


'[  Button State Enumeration  ]'
Public Enum cbtnStateConstants
  cbtnStateNormal = 0
  cbtnStateHover = 1
  cbtnStatePressed = 2
End Enum


'------------------------------------------------------------------------------'
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constants for "DrawEdge" function "edge" parameter  ]'
Private Const BDR_INNER       As Long = &HC
Private Const BDR_OUTER       As Long = &H3
Private Const BDR_RAISED      As Long = &H5
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKEN      As Long = &HA
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_SUNKENOUTER As Long = &H2

'[  constants for "DrawEdge" function "edge" parameter  ]'
Private Const EDGE_BUMP   As Long = (BDR_RAISEDOUTER + BDR_SUNKENINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER + BDR_RAISEDINNER)
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER + BDR_RAISEDINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER + BDR_SUNKENINNER)

'[  constants for "DrawEdge" function "grfFlags" parameter  ]'
Private Const BF_LEFT   As Long = &H1
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RIGHT  As Long = &H4
Private Const BF_TOP    As Long = &H2
Private Const BF_RECT   As Long = (BF_LEFT + BF_TOP + BF_RIGHT + BF_BOTTOM)

'[  constants for "DrawStatePic" function "flags" parameter  ]'
'[  constants for "DrawStateTxt" function "flags" parameter  ]'
Private Const DSS_NORMAL   As Long = &H0
Private Const DSS_DISABLED As Long = &H20

'[  constants for "DrawStatePic" function "flags" parameter  ]'
'[  constants for "DrawStateTxt" function "flags" parameter  ]'
Private Const DST_PREFIXTEXT As Long = &H2
Private Const DST_ICON       As Long = &H3
Private Const DST_BITMAP     As Long = &H4

'[  constants for "SetBkMode" function "nBkMode" parameter  ]'
Private Const BACKMODE_OPAQUE      As Long = 0
Private Const BACKMODE_TRANSPARENT As Long = 1


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function BitBlt _
        Lib "gdi32" ( _
          ByVal hDestDC As Long, _
          ByVal X As Long, _
          ByVal Y As Long, _
          ByVal nWidth As Long, _
          ByVal nHeight As Long, _
          ByVal hSrcDC As Long, _
          ByVal xSrc As Long, _
          ByVal ySrc As Long, _
          ByVal dwRop As Long _
        ) As Long

Private Declare Function CreateCompatibleBitmap _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal nWidth As Long, _
          ByVal nHeight As Long _
        ) As Long

Private Declare Function CreateCompatibleDC _
        Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateFont _
        Lib "gdi32" Alias "CreateFontA" ( _
          ByVal nHeight As Long, _
          ByVal nWidth As Long, _
          ByVal nEscapement As Long, _
          ByVal nOrientation As Long, _
          ByVal fnWeight As Long, _
          ByVal fdwItalic As Long, _
          ByVal fdwUnderline As Long, _
          ByVal fdwStrikeOut As Long, _
          ByVal fdwCharSet As Long, _
          ByVal fdwOutputPrecision As Long, _
          ByVal fdwClipPrecision As Long, _
          ByVal fdwQuality As Long, _
          ByVal fdwPitchAndFamily As Long, _
          ByVal lpszFace As String _
        ) As Long

Private Declare Function CreateSolidBrush _
        Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject _
        Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DrawEdge _
        Lib "user32" ( _
          ByVal hDC As Long, _
          ByRef qrc As RECT, _
          ByVal edge As Long, _
          ByVal grfFlags As Long _
        ) As Long

Private Declare Function DrawFocusRect _
        Lib "user32" ( _
          ByVal hDC As Long, _
          lpRect As RECT _
        ) As Long

Private Declare Function DrawStatePic _
        Lib "user32" Alias "DrawStateA" ( _
          ByVal hDC As Long, _
          ByVal hBrush As Long, _
          ByVal lpDrawStateProc As Long, _
          ByVal lParam As Long, _
          ByVal wParam As Long, _
          ByVal X As Long, _
          ByVal Y As Long, _
          ByVal cx As Long, _
          ByVal cy As Long, _
          ByVal Flags As Long _
        ) As Long

Private Declare Function DrawStateTxt _
        Lib "user32" Alias "DrawStateA" ( _
          ByVal hDC As Long, _
          ByVal hBrush As Long, _
          ByVal lpDrawStateProc As Long, _
          ByVal lString As String, _
          ByVal wParam As Long, _
          ByVal X As Long, _
          ByVal Y As Long, _
          ByVal cx As Long, _
          ByVal cy As Long, _
          ByVal Flags As Long _
        ) As Long

Private Declare Function FillRect _
        Lib "user32" ( _
          ByVal hDC As Long, _
          lpRect As RECT, _
          ByVal hBrush As Long _
        ) As Long

Private Declare Function GetPixel _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal X As Long, _
          ByVal Y As Long _
        ) As Long

Private Declare Function GetTextMetrics _
        Lib "gdi32" Alias "GetTextMetricsA" ( _
          ByVal hDC As Long, _
          lpMetrics As TEXTMETRIC _
        ) As Long

Private Declare Sub OleTranslateColor _
        Lib "oleaut32.dll" ( _
          ByVal clr As Long, _
          ByVal hpal As Long, _
          ByRef lpcolorref As Long)

Private Declare Function SelectObject _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal hObject As Long _
        ) As Long

Private Declare Function SetBkColor _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal crColor As Long _
        ) As Long

Private Declare Function SetBkMode _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal nBkMode As Long _
        ) As Long

Private Declare Function SetPixel _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal X As Long, _
          ByVal Y As Long, _
          ByVal crColor As Long _
        ) As Long

Private Declare Function SetTextColor _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal crColor As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  Public Events                                                               '
'------------------------------------------------------------------------------'


Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, _
                       Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseHover()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, _
                       Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, _
                     Y As Single)
Attribute MouseUp.VB_UserMemId = -607


'------------------------------------------------------------------------------'
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


Private m_Alignment As AlignmentConstants
Private m_Caption As String
Private m_GotFocus As Boolean
Private m_MaskStyle As cbtnMaskStyleConstants
Private m_OldWndProc As Long
Private m_Padding As Long
Private m_Picture As StdPicture
Private m_PictureOffset As Long
Private m_State As cbtnStateConstants


'------------------------------------------------------------------------------'
'  Public Properties                                                           '
'------------------------------------------------------------------------------'


Public Property Get AccessKeys() As String
  AccessKeys = UserControl.AccessKeys
End Property


Public Property Let AccessKeys(New_AccessKeys As String)
  If UserControl.AccessKeys <> New_AccessKeys Then
    UserControl.AccessKeys = New_AccessKeys
    PropertyChanged "AccessKeys"
    PaintControl
  End If
End Property


Public Property Get Alignment() As AlignmentConstants
  Alignment = m_Alignment
End Property


Public Property Let Alignment(New_Alignment As AlignmentConstants)
  If m_Alignment <> New_Alignment Then
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    PaintControl
  End If
End Property


Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property


Public Property Let BackColor(New_BackColor As OLE_COLOR)
  If UserControl.BackColor <> New_BackColor Then
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    PaintControl
  End If
End Property


Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
  Caption = m_Caption
End Property


Public Property Let Caption(New_Caption As String)
  If m_Caption <> New_Caption Then
    m_Caption = New_Caption
    PropertyChanged "Caption"
    PaintControl
  End If
End Property


Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property


Public Property Let Enabled(New_Enabled As Boolean)
  If UserControl.Enabled <> New_Enabled Then
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    PaintControl
  End If
End Property


Public Property Get Font() As StdFont
  Set Font = UserControl.Font
End Property


Public Property Set Font(New_Font As StdFont)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  PaintControl
End Property


Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property


Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
  If UserControl.ForeColor <> New_ForeColor Then
    UserControl.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    PaintControl
  End If
End Property


Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property


Public Property Get MaskColor() As OLE_COLOR
  MaskColor = UserControl.MaskColor
End Property


Public Property Let MaskColor(New_MaskColor As OLE_COLOR)
  If UserControl.MaskColor <> New_MaskColor Then
    UserControl.MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    PaintControl
  End If
End Property


Public Property Get MaskStyle() As cbtnMaskStyleConstants
  MaskStyle = m_MaskStyle
End Property


Public Property Let MaskStyle(New_MaskStyle As cbtnMaskStyleConstants)
  If m_MaskStyle <> New_MaskStyle Then
    m_MaskStyle = New_MaskStyle
    PropertyChanged "MaskStyle"
    PaintControl
  End If
End Property


Public Property Get Picture() As StdPicture
  Set Picture = m_Picture
End Property


Public Property Set Picture(New_Picture As StdPicture)
  Set m_Picture = New_Picture
  PropertyChanged "Picture"
  PaintControl
End Property


Public Property Get Padding() As Long
  Padding = m_Padding
End Property


Public Property Let Padding(New_Padding As Long)
  If m_Padding <> New_Padding Then
    m_Padding = New_Padding
    PropertyChanged "Padding"
    PaintControl
  End If
End Property


Public Property Get PictureOffset() As Long
  PictureOffset = m_PictureOffset
End Property


Public Property Let PictureOffset(New_PictureOffset As Long)
  If m_PictureOffset <> New_PictureOffset Then
    m_PictureOffset = New_PictureOffset
    PropertyChanged "PictureOffset"
    PaintControl
  End If
End Property


'------------------------------------------------------------------------------'
'  Friend Properties                                                           '
'------------------------------------------------------------------------------'


Friend Property Get OldWndProc() As Long
  OldWndProc = m_OldWndProc
End Property


Friend Property Let OldWndProc(ByVal New_OldWndProc As Long)
  m_OldWndProc = New_OldWndProc
End Property


Friend Property Get State() As cbtnStateConstants
  State = m_State
End Property


Friend Property Let State(ByVal New_State As cbtnStateConstants)
  If m_State <> New_State Then
    m_State = New_State
    PaintControl
  End If
End Property


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub Press()

  On Error Resume Next

  If Extender.Visible Then
    If UserControl.Enabled Then
      UserControl.SetFocus
      DoEvents
      Call UserControl_Click
    End If
  End If

End Sub


Public Sub Refresh()
  PaintControl
End Sub


'------------------------------------------------------------------------------'
'  Friend Procedures                                                           '
'------------------------------------------------------------------------------'


Friend Sub RaiseMouseLeaveEvent()
  RaiseEvent MouseLeave
End Sub


'------------------------------------------------------------------------------'
'  Private Procedures                                                          '
'------------------------------------------------------------------------------'


Private Sub DrawBorder(ByVal hDC As Long, ByRef tRect As RECT)

  Dim tRect2 As RECT
  tRect2.Left = tRect.Left + 1
  tRect2.top = tRect.top + 1
  tRect2.Right = tRect.Right - 1
  tRect2.Bottom = tRect.Bottom - 1

  Select Case m_State
    Case cbtnStateNormal:
      DrawEdge hDC, tRect, EDGE_ETCHED, BF_RECT
    Case cbtnStateHover:
      DrawEdge hDC, tRect2, BDR_RAISEDINNER, BF_RECT
    Case cbtnStatePressed:
      DrawEdge hDC, tRect2, BDR_SUNKENOUTER, BF_RECT
  End Select

End Sub


Private Sub DrawCaption(ByVal hDC As Long)

  If Len(Trim$(m_Caption)) = 0 Then Exit Sub

  Dim CapStr   As String
  Dim CapWidth As Long
  Dim X        As Long
  Dim Y        As Long
  Dim Txt      As TEXTMETRIC
  Dim hNewFont As Long
  Dim hOldFont As Long

  '[  get caption width without mnemonic character  ]'
  CapStr = Replace(m_Caption, "&&", Chr$(254))
  CapStr = Replace(CapStr, "&", vbNullString)
  CapStr = Replace(CapStr, Chr$(0), "&&")
  CapWidth = TextWidth(CapStr)

  '[  compute vertical position  ]'
  Y = (UserControl.ScaleHeight - TextHeight(m_Caption)) \ 2

  '[  compute horizontal position  ]'
  If m_Alignment = vbLeftJustify Then
    If m_Picture Is Nothing Then
      X = m_Padding
    Else
      X = m_Padding + m_PictureOffset + _
          ScaleX(m_Picture.Width, vbHimetric, vbPixels)
    End If
  ElseIf m_Alignment = vbRightJustify Then
    X = UserControl.ScaleWidth - CapWidth - m_Padding
  ElseIf m_Alignment = vbCenter Then
    If m_Picture Is Nothing Then
      X = (UserControl.ScaleWidth - CapWidth) \ 2
    Else
      X = (UserControl.ScaleWidth - CapWidth + m_PictureOffset + _
           ScaleX(m_Picture.Width, vbHimetric, vbPixels)) \ 2
    End If
  End If

  '[  adjust text position according to state  ]'
  If m_State = cbtnStatePressed Then
    X = X + 1
    Y = Y + 1
  End If

  '[  set font  ]'
  GetTextMetrics UserControl.hDC, Txt
  hNewFont = CreateFont(Txt.tmHeight, 0, 0, 0, Txt.tmWeight, Txt.tmItalic, _
                        Txt.tmUnderlined, Txt.tmStruckOut, 0, 0, 16, 0, 0, _
                        UserControl.Font.Name)
  hOldFont = SelectObject(ByVal hDC, hNewFont)

  '[  set text color and background  ]'
  SetTextColor hDC, TranslateColor(UserControl.ForeColor)
  SetBkMode hDC, BACKMODE_TRANSPARENT

  '[  paint Caption  ]'
  If UserControl.Enabled Then
    DrawStateTxt hDC, 0, 0, m_Caption, Len(m_Caption), X, Y, 0, 0, _
                  DST_PREFIXTEXT Or DSS_NORMAL
  Else
    DrawStateTxt hDC, 0, 0, m_Caption, Len(m_Caption), X, Y, 0, 0, _
                  DST_PREFIXTEXT Or DSS_DISABLED
  End If

  '[  restore original font  ]'
  SelectObject hDC, hOldFont
  DeleteObject hNewFont

End Sub


Private Sub DrawFocusRectangle(ByVal hDC As Long, ByRef tRect As RECT)
  Dim tRect2 As RECT
  tRect2.Left = tRect.Left + 3
  tRect2.top = tRect.top + 3
  tRect2.Right = tRect.Right - 4
  tRect2.Bottom = tRect.Bottom - 4
  SetTextColor hDC, vbBlack
  DrawFocusRect hDC, tRect2
End Sub


Private Sub DrawPicture(ByVal hDC As Long)

  If m_Picture Is Nothing Then Exit Sub

  Dim CapStr     As String
  Dim CapWidth   As Long
  Dim offset     As Long
  Dim X          As Long
  Dim Y          As Long
  Dim PicWidth   As Long
  Dim PicHeight  As Long
  Dim hDCTmp     As Long
  Dim MaskClr    As Long
  Dim BackClr    As Long
  Dim Ctr1       As Long
  Dim Ctr2       As Long

  '[  get caption width and offset  ]'
  If Len(Trim$(m_Caption)) = 0 Then
    CapWidth = 0
    offset = 0
  Else
    '[  remove mnemonic characters from total width  ]'
    CapStr = Replace(m_Caption, "&&", Chr$(254))
    CapStr = Replace(CapStr, "&", vbNullString)
    CapStr = Replace(CapStr, Chr$(0), "&&")
    CapWidth = TextWidth(CapStr)
    offset = m_PictureOffset
  End If

  '[  get picture dimensions  ]'
  PicWidth = UserControl.ScaleX(m_Picture.Width, vbHimetric, vbPixels)
  PicHeight = UserControl.ScaleY(m_Picture.Height, vbHimetric, vbPixels)

  '[  compute vertical position  ]'
  Y = (UserControl.ScaleHeight - PicHeight) \ 2

  '[  compute horizontal position  ]'
  If m_Alignment = vbLeftJustify Then
    X = m_Padding
  ElseIf m_Alignment = vbRightJustify Then
    X = UserControl.ScaleWidth - m_Padding - CapWidth - PicWidth - offset
  ElseIf m_Alignment = vbCenter Then
    X = (UserControl.ScaleWidth - CapWidth - PicWidth - offset) \ 2
  End If

  '[  adjust picture position according to state  ]'
  If m_State = cbtnStatePressed Then
    X = X + 1
    Y = Y + 1
  End If

  '[  if icon  ]'
  If m_Picture.Type = vbPicTypeIcon Then

    If UserControl.Enabled Then
      DrawStatePic hDC, 0, 0, m_Picture.Handle, 0, X, Y, 0, 0, _
                   DST_ICON Or DSS_NORMAL
    Else
      DrawStatePic hDC, 0, 0, m_Picture.Handle, 0, X, Y, 0, 0, _
                   DST_ICON Or DSS_DISABLED
    End If

  '[  else if bitmap  ]'
  Else

    '[  create temporary dc for picture  ]'
    hDCTmp = CreateCompatibleDC(ByVal hDC)
    DeleteObject SelectObject(ByVal hDCTmp, m_Picture.Handle)

    '[  if mask enabled  ]'
    If m_MaskStyle <> cbtnMaskStyleNone Then

      '[  get mask color  ]'
      If m_MaskStyle = cbtnMaskStyleAuto Then
        MaskClr = GetPixel(ByVal hDCTmp, 0, 0)
      Else
        MaskClr = TranslateColor(UserControl.MaskColor)
      End If

      BackClr = TranslateColor(UserControl.BackColor)

      '[  clear transparent pixels of picture  ]'
      For Ctr2 = 0 To PicHeight
        For Ctr1 = 0 To PicWidth
          If GetPixel(ByVal hDCTmp, Ctr1, Ctr2) = MaskClr Then
            SetPixel hDCTmp, Ctr1, Ctr2, BackClr
          End If
        Next Ctr1
      Next Ctr2

    End If

    '[  paint picture  ]'
    BitBlt hDC, X, Y, PicWidth, PicHeight, hDCTmp, 0, 0, vbSrcCopy

    '[  cleanup temporary dc  ]'
    DeleteDC hDCTmp

  End If

End Sub


Private Sub PaintControl()

  Dim tRect  As RECT
  Dim hDC    As Long
  Dim hBmp   As Long
  Dim hBrush As Long

  '[  compute control dimensions  ]'
  tRect.Left = 0
  tRect.top = 0
  tRect.Right = UserControl.ScaleWidth
  tRect.Bottom = UserControl.ScaleHeight

  '[  create back buffer  ]'
  hDC = CreateCompatibleDC(UserControl.hDC)
  hBmp = CreateCompatibleBitmap(UserControl.hDC, tRect.Right, tRect.Bottom)
  DeleteObject SelectObject(ByVal hDC, hBmp)

  '[  fill back buffer with background color  ]'
  hBrush = CreateSolidBrush(TranslateColor(UserControl.BackColor))
  FillRect hDC, tRect, hBrush
  DeleteObject hBrush

  '[  paint control on back buffer  ]'
  DrawPicture hDC
  DrawCaption hDC
  DrawBorder hDC, tRect

  If m_GotFocus Then
    DrawFocusRectangle hDC, tRect
  End If

  '[  copy back buffer contents to Control  ]'
  BitBlt UserControl.hDC, tRect.Left, tRect.top, tRect.Right, tRect.Bottom, _
         hDC, 0, 0, vbSrcCopy

  '[  update display  ]'
  DoEvents
  UserControl.Refresh

  '[  cleanup temporary resources  ]'
  DeleteDC hDC
  DeleteObject hBmp

End Sub


Private Function TranslateColor(OLEColor As OLE_COLOR) As Long
  OleTranslateColor OLEColor, UserControl.Palette, TranslateColor
End Function


'------------------------------------------------------------------------------'
'  Control Event Handlers                                                      '
'------------------------------------------------------------------------------'


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  Call UserControl_Click
End Sub


Private Sub UserControl_AmbientChanged(PropertyName As String)
  PaintControl
End Sub


Private Sub UserControl_Click()
  RaiseEvent Click
End Sub


Private Sub UserControl_GotFocus()
  m_GotFocus = True
  PaintControl
End Sub


Private Sub UserControl_Hide()
  modCustomButton.StopSubclassingButton Me
End Sub

Private Sub UserControl_Initialize()
  m_GotFocus = False
  m_State = cbtnStateNormal
  m_OldWndProc = 0
End Sub


Private Sub UserControl_InitProperties()

  m_Alignment = vbCenter
  m_Caption = Replace(Extender.Name, UserControl.Name, "Button")
  m_MaskStyle = cbtnMaskStyleAuto
  m_Padding = 10
  Set m_Picture = Nothing
  m_PictureOffset = 10

  UserControl.AccessKeys = vbNullString
  UserControl.BackColor = vbButtonFace
  UserControl.Enabled = True
  Set UserControl.Font = Ambient.Font
  UserControl.ForeColor = vbButtonText
  UserControl.MaskColor = vbWhite

End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  If (KeyCode = vbKeyReturn) Then
    '[  simulate click  ]'
    RaiseEvent Click
  ElseIf (KeyCode = vbKeySpace) And (m_State <> cbtnStatePressed) Then
    '[  simulate mouse down  ]'
    m_State = cbtnStatePressed
    PaintControl
  ElseIf (KeyCode = vbKeyDown) Or (KeyCode = vbKeyRight) Then
    '[  select next control  ]'
    SendKeys "{TAB}", True
  ElseIf (KeyCode = vbKeyUp) Or (KeyCode = vbKeyLeft) Then
    '[  select previous control  ]'
    SendKeys "+{TAB}", True
  End If

  RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

  If (KeyCode = vbKeySpace) And (m_State = cbtnStatePressed) Then
    '[  simulate click  ]'
    m_State = cbtnStateNormal
    PaintControl
    RaiseEvent Click
  End If

  RaiseEvent KeyUp(KeyCode, Shift)

End Sub


Private Sub UserControl_LostFocus()
  m_GotFocus = False
  PaintControl
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, _
                                  X As Single, Y As Single)

  If Button = vbLeftButton Then
    m_State = cbtnStatePressed
    PaintControl
  End If

  RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, _
                                  X As Single, Y As Single)

  If Ambient.UserMode Then

    RaiseEvent MouseMove(Button, Shift, X, Y)

    If (m_State <> cbtnStateHover) And (Button <> vbLeftButton) Then
      modCustomButton.TrackMouseLeaveEvent Me
      m_State = cbtnStateHover
      PaintControl
      RaiseEvent MouseHover
    End If

  End If

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
  m_State = cbtnStateNormal
  PaintControl
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  With PropBag

    m_Alignment = .ReadProperty("Alignment", vbCenter)
    m_Caption = .ReadProperty("Caption", vbNullString)
    m_MaskStyle = .ReadProperty("MaskStyle", cbtnMaskStyleAuto)
    m_Padding = .ReadProperty("Padding", 10)
    Set m_Picture = .ReadProperty("Picture", Nothing)
    m_PictureOffset = .ReadProperty("PictureOffset", 6)

    UserControl.AccessKeys = .ReadProperty("AccessKeys", vbNullString)
    UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
    UserControl.Enabled = .ReadProperty("Enabled", True)
    Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = .ReadProperty("ForeColor", vbButtonText)
    UserControl.MaskColor = .ReadProperty("MaskColor", vbWhite)

  End With

End Sub


Private Sub UserControl_Resize()
  PaintControl
End Sub


Private Sub UserControl_Show()
  m_State = cbtnStateNormal
  PaintControl
  modCustomButton.StartSubclassingButton Me
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "AccessKeys", UserControl.AccessKeys, vbNullString
    .WriteProperty "Alignment", m_Alignment, vbCenter
    .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
    .WriteProperty "Caption", m_Caption, vbNullString
    .WriteProperty "Enabled", UserControl.Enabled, True
    .WriteProperty "Font", UserControl.Font, Ambient.Font
    .WriteProperty "ForeColor", UserControl.ForeColor, vbButtonText
    .WriteProperty "MaskColor", UserControl.MaskColor, vbWhite
    .WriteProperty "MaskStyle", m_MaskStyle, cbtnMaskStyleAuto
    .WriteProperty "Padding", m_Padding, 10
    .WriteProperty "Picture", m_Picture, Nothing
    .WriteProperty "PictureOffset", m_PictureOffset, 6
  End With
End Sub
