VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPassword2 
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
      TabIndex        =   6
      Top             =   -60
      Width           =   6660
      Begin VB.Frame fraStatus 
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
         Height          =   315
         Left            =   210
         TabIndex        =   12
         Top             =   1950
         Visible         =   0   'False
         Width           =   6225
         Begin MSComctlLib.ProgressBar prgStatus 
            Height          =   150
            Left            =   210
            TabIndex        =   13
            Top             =   75
            Width           =   5790
            _ExtentX        =   10213
            _ExtentY        =   265
            _Version        =   393216
            Appearance      =   0
         End
         Begin MSComctlLib.StatusBar staStatus 
            Height          =   225
            Left            =   195
            TabIndex        =   14
            Top             =   30
            Width           =   5820
            _ExtentX        =   10266
            _ExtentY        =   397
            Style           =   1
            ShowTips        =   0   'False
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
                  AutoSize        =   1
                  Object.Width           =   10213
               EndProperty
            EndProperty
         End
      End
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
         TabIndex        =   8
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
            TabIndex        =   9
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
         TabIndex        =   7
         Top             =   2505
         Width           =   6225
         Begin Chameleon.CustomButton btnHelp 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   4305
            TabIndex        =   2
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
            Picture         =   "frmPassword2.frx":0000
            PictureOffset   =   8
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "Press ""Next"" to verify the password."
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   270
            Width           =   3900
         End
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
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1230
         Width           =   6225
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   6390
         TabIndex        =   16
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   1680
         Width           =   45
      End
      Begin VB.Image imgPageIcon 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "frmPassword2.frx":059A
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         Height          =   195
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
      TabIndex        =   10
      Top             =   3255
      Width           =   6660
      Begin Chameleon.CustomButton btnCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   5
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
         Picture         =   "frmPassword2.frx":0E64
      End
      Begin Chameleon.CustomButton btnNext 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1455
         TabIndex        =   4
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
         Picture         =   "frmPassword2.frx":1186
      End
      Begin Chameleon.CustomButton btnBack 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   210
         TabIndex        =   3
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
         Picture         =   "frmPassword2.frx":14A8
         PictureOffset   =   4
      End
   End
End
Attribute VB_Name = "frmPassword2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Password Verification Form                                                  '
'  [frmPassword2]                                                              '
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


'[  user-selected input file  ]'
Private m_StegoImagePath As String

'[  cancel process flag  ]'
Private m_Cancel As Boolean


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub ResetControls()
  txtPassword.Text = vbNullString
End Sub


'------------------------------------------------------------------------------'
'  Private Procedures                                                          '
'------------------------------------------------------------------------------'


Private Sub PerformVerification()

  Dim PasswordHashMD5 As String
  Dim PasswordHashSHA As String

  With frmMainMenu.StegoDecoder

    fraPage.Enabled = False

    If .LoadStegoImage(m_StegoImagePath) Then

      PasswordHashMD5 = GetHashValue(txtPassword.Text, MD5)
      PasswordHashSHA = GetHashValue(txtPassword.Text, SHA)

      lblStatus(0).Caption = "Verifying password..."
      fraStatus.Visible = True

      .DecodeMetadata PasswordHashMD5

      lblStatus(0).Caption = vbNullString
      lblStatus(1).Caption = vbNullString

      If .StoredPassword = PasswordHashSHA Then
        If m_Cancel Then Exit Sub
        frmDecode.Show
        frmDecode.SetFocus
      Else
        If m_Cancel Then Exit Sub
        lblStatus(0).Caption = "Invalid Password."
        MsgBox "The specified password is invalid." & vbCrLf & _
               "Please specify the correct password.", vbExclamation
        fraPage.Enabled = True
        txtPassword.SetFocus
      End If

    Else

      MsgBox "The selected image does not contain valid hidden data." & _
             vbCrLf & "Please select a valid stego image.", vbExclamation

      frmStegoImage.Show
      frmStegoImage.SetFocus

    End If

  End With

End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub btnBack_Click()
  frmStegoImage.Show
  frmStegoImage.SetFocus
End Sub


Private Sub btnCancel_Click()

  If Not btnNext.Visible Then
    If MsgBox("Do you wish to cancel the current task?", _
              vbQuestion + vbYesNo) = vbYes Then
      m_Cancel = True
      frmMainMenu.StegoDecoder.Abort
      frmMainMenu.Show
      frmMainMenu.SetFocus
    End If
  Else
    frmMainMenu.Show
    frmMainMenu.SetFocus
  End If

End Sub


Private Sub btnHelp_Click()
  DisplayHelpFile "extract_wizard.html#Page2"
End Sub


Private Sub btnNext_Click()

  On Error Resume Next
  Me.ValidateControls

  If ActiveControl Is btnNext Then PerformVerification

End Sub


Private Sub Form_Activate()

  m_Cancel = False
  m_StegoImagePath = frmStegoImage.txtPath.Text

  On Error Resume Next
  txtPassword.SetFocus

End Sub


Private Sub Form_Deactivate()

  If Screen.ActiveForm.MDIChild Then

    Me.Hide

    '[  reset controls  ]'
    fraPage.Enabled = True
    fraStatus.Visible = False
    lblStatus(0).Caption = vbNullString
    lblStatus(1).Caption = vbNullString
    prgStatus.Min = 0
    prgStatus.Max = 100
    prgStatus.Value = 0

  End If

End Sub


Private Sub Form_Load()
  Set imgPageIcon(1).Picture = imgPageIcon(0).Picture
  SetMargins txtPassword, 1
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then btnHelp.Press
End Sub


Private Sub txtPassword_GotFocus()

  With txtPassword
    If .Tag <> "MouseDown" Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

End Sub


Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then btnNext.SetFocus
End Sub


Private Sub txtPassword_MouseDown(Button As Integer, Shift As Integer, _
                                  X As Single, Y As Single)
  txtPassword.Tag = "MouseDown"
End Sub


Private Sub txtPassword_MouseUp(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
  txtPassword.Tag = ""
End Sub


Private Sub txtPassword_Validate(Cancel As Boolean)

  If Len(txtPassword.Text) = 0 Then
    MsgBox "No password has been specified." & vbCrLf & _
           "Please specify a password.", vbExclamation
    Cancel = True
  End If

End Sub
