VERSION 5.00
Begin VB.Form frmPassword1 
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6660
   ClipControls    =   0   'False
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
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6660
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPage 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   6660
      Begin VB.Frame fraPageTitle 
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   0.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   450
         TabIndex        =   11
         Top             =   210
         Width           =   5985
         Begin VB.Image imgPageIcon 
            Height          =   480
            Index           =   1
            Left            =   -270
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblPageTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   12
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.Frame fraHelp 
         Caption         =   " Help "
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         TabIndex        =   9
         Top             =   2505
         Width           =   6225
         Begin Chameleon.CustomButton btnHelp 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   4305
            TabIndex        =   4
            ToolTipText     =   " (shortcut: F1) "
            Top             =   180
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   688
            Caption         =   "More Help"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmPassword1.frx":0000
            PictureOffset   =   8
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "Specify a password to protect the data to be hidden."
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   270
            Width           =   3900
         End
      End
      Begin VB.TextBox txtPassword 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1950
         Width           =   6225
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1230
         Width           =   6225
      End
      Begin VB.Image imgPageIcon 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "frmPassword1.frx":059A
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password &Confirmation:"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   960
         Width           =   750
      End
   End
   Begin VB.Frame fraNavBar 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      TabIndex        =   13
      Top             =   3255
      Width           =   6660
      Begin Chameleon.CustomButton btnCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   7
         ToolTipText     =   " (shortcut: Esc) "
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPassword1.frx":0E64
      End
      Begin Chameleon.CustomButton btnNext 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1455
         TabIndex        =   6
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         AccessKeys      =   "n"
         Caption         =   "&Next"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPassword1.frx":1186
      End
      Begin Chameleon.CustomButton btnBack 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   210
         TabIndex        =   5
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         AccessKeys      =   "b"
         Caption         =   "&Back"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPassword1.frx":14A8
         PictureOffset   =   4
      End
   End
End
Attribute VB_Name = "frmPassword1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Password Specification Form                                                 '
'  [frmPassword1]                                                              '
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
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub ResetControls()
  txtPassword(0).Text = vbNullString
  txtPassword(1).Text = vbNullString
End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub btnBack_Click()
  frmCoverImage.Show
  frmCoverImage.SetFocus
End Sub


Private Sub btnCancel_Click()
  frmMainMenu.Show
  frmMainMenu.SetFocus
End Sub


Private Sub btnHelp_Click()
  DisplayHelpFile "hide_wizard.html#Page3"
End Sub


Private Sub btnNext_Click()

  On Error Resume Next
  Me.ValidateControls

  If ActiveControl Is btnNext Then
    frmEncode.Show
    frmEncode.SetFocus
  End If

End Sub


Private Sub Form_Activate()
  On Error Resume Next
  txtPassword(0).SetFocus
End Sub


Private Sub Form_Deactivate()
  If Screen.ActiveForm.MDIChild Then Me.Hide
End Sub


Private Sub Form_Load()
  Set imgPageIcon(1).Picture = imgPageIcon(0).Picture
  SetMargins txtPassword(0), 1
  SetMargins txtPassword(1), 1
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then btnHelp.Press
End Sub


Private Sub txtPassword_GotFocus(Index As Integer)

  With txtPassword(Index)
    If .Tag <> "MouseDown" Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

End Sub


Private Sub txtPassword_KeyDown(Index As Integer, KeyCode As Integer, _
                                Shift As Integer)

  If KeyCode = vbKeyReturn Then
    Select Case Index
      Case 0:  txtPassword(1).SetFocus
      Case 1:  btnNext.SetFocus
    End Select
  End If

End Sub


Private Sub txtPassword_MouseDown(Index As Integer, Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, Y As Single)
  txtPassword(Index).Tag = "MouseDown"
End Sub


Private Sub txtPassword_MouseUp(Index As Integer, Button As Integer, _
                                Shift As Integer, _
                                X As Single, Y As Single)
  txtPassword(Index).Tag = ""
End Sub


Private Sub txtPassword_Validate(Index As Integer, Cancel As Boolean)

  If Len(txtPassword(0).Text) = 0 Then
    MsgBox "No password has been specified." & vbCrLf & _
           "Please specify a password.", vbExclamation
    Cancel = True
  ElseIf Len(txtPassword(1).Text) = 0 Then
    MsgBox "The password has not been confirmed." & vbCrLf & _
           "Please confirm the password.", vbExclamation
    Cancel = True
  ElseIf txtPassword(0).Text <> txtPassword(1).Text Then
    MsgBox "The two passwords specified are different." & vbCrLf & _
           "Please confirm the password carefully.", vbExclamation
    Cancel = True
  End If

End Sub
