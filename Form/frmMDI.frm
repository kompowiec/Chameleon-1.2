VERSION 5.00
Begin VB.MDIForm frmMDI 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "Chameleon"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6630
   Icon            =   "frmMDI.frx":0000
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Multiple Document Interface (MDI) Form                                      '
'  [frmMDI]                                                                    '
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
'  Windows User Interface Constants                                            '
'------------------------------------------------------------------------------'


Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&

Private Const SC_SIZE = &HF000&
Private Const SC_MOVE = &HF010&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_MAXIMIZE = &HF030&

Private Const WS_CAPTION = &HC00000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000

Private Const GWL_STYLE = -16


'------------------------------------------------------------------------------'
'  Windows User Interface Function Declarations                                '
'------------------------------------------------------------------------------'


Private Declare Function GetSystemMenu _
        Lib "user32" ( _
          ByVal hWnd As Long, _
          ByVal bRevert As Long _
        ) As Long

Private Declare Function RemoveMenu _
        Lib "user32" ( _
          ByVal hMenu As Long, _
          ByVal nPosition As Long, _
          ByVal wFlags As Long _
        ) As Long

Private Declare Function DrawMenuBar _
        Lib "user32" ( _
          ByVal hWnd As Long _
        ) As Long

Private Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
          ByVal hWnd As Long, _
          ByVal nIndex As Long, _
          ByVal dwNewLong As Long _
        ) As Long

Private Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
          ByVal hWnd As Long, _
          ByVal nIndex As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub FormatWindow()

  Dim Res As Long
  Dim Ctr As Long

  '[  remove maximize button from mdi form and disable resizing  ]'
  Res = GetWindowLong(Me.hWnd, GWL_STYLE)
  SetWindowLong Me.hWnd, GWL_STYLE, Res And Not (WS_MAXIMIZEBOX + WS_THICKFRAME)

  '[  remove maximize menu item from system menu  ]'
  Res = GetSystemMenu(Me.hWnd, 0)
  RemoveMenu Res, SC_MAXIMIZE, MF_BYCOMMAND
  DrawMenuBar Me.hWnd

End Sub


Public Sub ResetChildFormControls()
  On Error Resume Next
  Dim Frm As Form
  For Each Frm In VB.Forms
    If Frm.MDIChild Then Frm.ResetControls
  Next Frm
End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub MDIForm_Load()

  '[  check for other instances of chameleon in memory  ]'
  If App.PrevInstance Then
    MsgBox "Another instance of Chameleon is already open." & vbCrLf & _
           "Only one instance of Chameleon may be opened.", vbCritical
    End
  End If

  FormatWindow
  Load frmView

End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If (UnloadMode = vbAppTaskManager) Or (UnloadMode = vbFormControlMenu) Then
    If Not frmMDI.ActiveForm Is frmMainMenu Then

      If MsgBox("Do you wish to cancel the current task and exit Chameleon?", _
                vbQuestion + vbYesNo) = vbNo Then
        Cancel = True
        Exit Sub
      End If

    End If
  End If

  Cancel = False

  '[  hide mdi form  ]'
  frmMDI.WindowState = vbMinimized
  frmMDI.Hide

  CloseHelpFiles

  Unload frmView
  Unload frmMDI

End Sub
