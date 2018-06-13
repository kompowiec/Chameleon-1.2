VERSION 5.00
Begin VB.Form frmView 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   -30
   ClientTop       =   -345
   ClientWidth     =   9600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStegoImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   7200
      Index           =   0
      Left            =   4815
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4785
      Begin VB.PictureBox picStegoImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   150
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   3
         ToolTipText     =   "Stego Image"
         Top             =   150
         Width           =   1500
      End
   End
   Begin VB.PictureBox picCoverImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   7200
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4785
      Begin VB.PictureBox picCoverImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   150
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         ToolTipText     =   "Cover Image"
         Top             =   150
         Width           =   1500
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "Preview Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Image Preview Form                                                          '
'  [frmView]                                                                   '
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
'  Private Variables                                                           '
'------------------------------------------------------------------------------'


'[  freeimage wrapper object  ]'
Private Imager As FreeImageWrapper

'[  image dib handles  ]'
Private m_CoverImageDIB As Long
Private m_StegoImageDIB As Long
Private m_DIBsTemporary As Boolean

'[  number of images to display  ]'
Private m_ImageCount As Long

'[  mouse down coordinate  ]'
Private m_MouseX As Long
Private m_MouseY As Long

'[  size difference between image and container  ]'
Private m_DifX As Long
Private m_DifY As Long


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Function DisplayByDIB( _
                  Optional ByVal CoverImage As Long = 0, _
                  Optional ByVal StegoImage As Long = 0 _
                ) As Boolean

  DisplayByDIB = False

  m_ImageCount = 0
  m_CoverImageDIB = CoverImage
  m_StegoImageDIB = StegoImage
  m_DIBsTemporary = False

  If PaintImage(picCoverImage(1)) Then m_ImageCount = m_ImageCount + 1
  If PaintImage(picStegoImage(1)) Then m_ImageCount = m_ImageCount + 1

  If m_ImageCount > 0 Then
    Call Form_Resize
    Me.Show
    DisplayByDIB = True
  End If

End Function


Public Function DisplayByFilename( _
                  Optional ByVal CoverImage As String = vbNullString, _
                  Optional ByVal StegoImage As String = vbNullString _
                ) As Boolean

  DisplayByFilename = False

  m_ImageCount = 0
  m_CoverImageDIB = Imager.LoadDIB(CoverImage)
  m_StegoImageDIB = Imager.LoadDIB(StegoImage)
  m_DIBsTemporary = True

  If PaintImage(picCoverImage(1)) Then m_ImageCount = m_ImageCount + 1
  If PaintImage(picStegoImage(1)) Then m_ImageCount = m_ImageCount + 1

  If m_ImageCount > 0 Then
    Call Form_Resize
    Me.Show
    DisplayByFilename = True
  End If

End Function


'------------------------------------------------------------------------------'
'  Private Procedures                                                          '
'------------------------------------------------------------------------------'


Private Function PaintImage(ByRef Picture1 As PictureBox) As Boolean

  Dim hDIB As Long

  Select Case Picture1.Name
    Case picCoverImage(1).Name:  hDIB = m_CoverImageDIB
    Case picStegoImage(1).Name:  hDIB = m_StegoImageDIB
  End Select

  If hDIB <> 0 Then

    Picture1.Width = Imager.GetWidth(hDIB)
    Picture1.Height = Imager.GetHeight(hDIB)

    If Imager.PaintDIB(hDIB, Picture1.hDC) Then
      PaintImage = True
      Select Case Picture1.Name
        Case picCoverImage(1).Name:  picCoverImage(0).Visible = True
        Case picStegoImage(1).Name:  picStegoImage(0).Visible = True
      End Select
    Else
      PaintImage = False
    End If

  End If

End Function


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub Form_Deactivate()

  Me.Hide
  picCoverImage(0).Visible = False
  picStegoImage(0).Visible = False

  If m_DIBsTemporary Then
    On Error Resume Next
    Imager.UnloadDIB m_CoverImageDIB
    Imager.UnloadDIB m_StegoImageDIB
  End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode

    Case vbKeyRight:
      If picCoverImage(0).Visible Then
        Call picCoverImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picCoverImage_MouseMove(1, vbLeftButton, 0, -10, 0)
      Else
        Call picStegoImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picStegoImage_MouseMove(1, vbLeftButton, 0, -10, 0)
      End If

    Case vbKeyLeft:
      If picCoverImage(0).Visible Then
        Call picCoverImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picCoverImage_MouseMove(1, vbLeftButton, 0, 10, 0)
      Else
        Call picStegoImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picStegoImage_MouseMove(1, vbLeftButton, 0, 10, 0)
      End If

    Case vbKeyDown:
      If picCoverImage(0).Visible Then
        Call picCoverImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picCoverImage_MouseMove(1, vbLeftButton, 0, 0, -10)
      Else
        Call picStegoImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picStegoImage_MouseMove(1, vbLeftButton, 0, 0, -10)
      End If

    Case vbKeyUp:
      If picCoverImage(0).Visible Then
        Call picCoverImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picCoverImage_MouseMove(1, vbLeftButton, 0, 0, 10)
      Else
        Call picStegoImage_MouseDown(1, vbLeftButton, 0, 0, 0)
        Call picStegoImage_MouseMove(1, vbLeftButton, 0, 0, 10)
      End If

    Case vbKeyEscape:
      Call mnuClose_Click

    Case 93:
      Call Form_MouseDown(vbRightButton, 0, 0, 0)

  End Select

End Sub


Private Sub Form_Load()
  Set Imager = New FreeImageWrapper
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuView
  End If
End Sub


Private Sub Form_Resize()

  Dim W As Long
  Dim H As Long
  Dim L As Long
  Dim T As Long

  Select Case m_ImageCount

    Case 1:
      W = Me.ScaleWidth
      H = Me.ScaleHeight
      If m_CoverImageDIB <> 0 Then
        m_DifX = W - picCoverImage(1).Width
        m_DifY = H - picCoverImage(1).Height
        picCoverImage(0).Move 0, 0
        picCoverImage(0).Width = W
        picCoverImage(0).Height = H
      Else
        m_DifX = W - picStegoImage(1).Width
        m_DifY = H - picStegoImage(1).Height
        picStegoImage(0).Move 0, 0
        picStegoImage(0).Width = W
        picStegoImage(0).Height = H
      End If

    Case 2:
      W = Me.ScaleWidth \ 2
      H = Me.ScaleHeight
      m_DifX = W - picCoverImage(1).Width
      m_DifY = H - picCoverImage(1).Height
      picCoverImage(0).Move 0, 0
      picCoverImage(0).Width = W - 1
      picCoverImage(0).Height = H
      picStegoImage(0).Move W + 1, 0
      picStegoImage(0).Width = W - 1
      picStegoImage(0).Height = H

  End Select

  L = m_DifX \ 2
  T = m_DifY \ 2

  If m_CoverImageDIB <> 0 Then
    With picCoverImage(1)
      If L < 0 Then
        .Left = 0
        .MousePointer = vbSizeAll
      Else
        .Left = L
        .MousePointer = vbDefault
      End If
      If T < 0 Then
        .top = 0
        .MousePointer = vbSizeAll
      Else
        .top = T
      End If
    End With
  End If

  If m_StegoImageDIB <> 0 Then
    With picStegoImage(1)
      If L < 0 Then
        .Left = 0
        .MousePointer = vbSizeAll
      Else
        .Left = L
        .MousePointer = vbDefault
      End If
      If T < 0 Then
        .top = 0
        .MousePointer = vbSizeAll
      Else
        .top = T
      End If
    End With
  End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set Imager = Nothing
End Sub


Private Sub mnuClose_Click()
  Me.Hide
End Sub


Private Sub picCoverImage_MouseDown(Index As Integer, Button As Integer, _
                                    Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    m_MouseX = X
    m_MouseY = Y
  ElseIf Button = vbRightButton Then
    PopupMenu mnuView
  End If
End Sub


Private Sub picStegoImage_MouseDown(Index As Integer, Button As Integer, _
                                    Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    m_MouseX = X
    m_MouseY = Y
  ElseIf Button = vbRightButton Then
    PopupMenu mnuView
  End If
End Sub


Private Sub picCoverImage_MouseMove(Index As Integer, Button As Integer, _
                                    Shift As Integer, X As Single, Y As Single)

  If (Index = 1) Then
    If (Button = vbLeftButton) Then

      Dim X2 As Long
      Dim Y2 As Long
      X2 = picCoverImage(1).Left
      Y2 = picCoverImage(1).top

      If (m_DifX < 0) Then
        X2 = X2 + (X - m_MouseX)
        If X2 > 0 Then
          X2 = 0
        ElseIf X2 < m_DifX Then
          X2 = m_DifX
        End If
      End If

      If (m_DifY < 0) Then
        Y2 = Y2 + (Y - m_MouseY)
        If Y2 > 0 Then
          Y2 = 0
        ElseIf Y2 < m_DifY Then
          Y2 = m_DifY
        End If
      End If

      If (X2 <> picCoverImage(1).Left) Or (Y2 <> picCoverImage(1).top) Then
        picCoverImage(1).Move X2, Y2
        If picStegoImage(0).Visible Then picStegoImage(1).Move X2, Y2
        Me.Refresh
      End If

    End If
  End If

End Sub


Private Sub picStegoImage_MouseMove(Index As Integer, Button As Integer, _
                                    Shift As Integer, X As Single, Y As Single)

  If (Index = 1) Then
    If (Button = vbLeftButton) Then

      Dim X2 As Long
      Dim Y2 As Long
      X2 = picStegoImage(1).Left
      Y2 = picStegoImage(1).top

      If (m_DifX < 0) Then
        X2 = X2 + (X - m_MouseX)
        If X2 > 0 Then
          X2 = 0
        ElseIf X2 < m_DifX Then
          X2 = m_DifX
        End If
      End If

      If (m_DifY < 0) Then
        Y2 = Y2 + (Y - m_MouseY)
        If Y2 > 0 Then
          Y2 = 0
        ElseIf Y2 < m_DifY Then
          Y2 = m_DifY
        End If
      End If

      If (X2 <> picStegoImage(1).Left) Or (Y2 <> picStegoImage(1).top) Then
        picStegoImage(1).Move X2, Y2
        If picCoverImage(0).Visible Then picCoverImage(1).Move X2, Y2
        Me.Refresh
      End If

    End If
  End If

End Sub
