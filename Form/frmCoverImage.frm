VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCoverImage 
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
      Begin MSComDlg.CommonDialog dlgBrowse 
         Left            =   5700
         Top             =   450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Select Cover Image"
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
         Height          =   1440
         Left            =   4335
         TabIndex        =   17
         Top             =   1710
         Width           =   2100
         Begin Chameleon.CustomButton btnHelp 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   180
            TabIndex        =   3
            ToolTipText     =   " (shortcut: F1) "
            Top             =   975
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
            Picture         =   "frmCoverImage.frx":0000
            PictureOffset   =   8
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the image in which to hide the data file."
            ForeColor       =   &H8000000D&
            Height          =   600
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   1725
            WordWrap        =   -1  'True
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
         TabIndex        =   13
         Top             =   210
         Width           =   5985
         Begin VB.Image imgPageIcon 
            Height          =   480
            Index           =   1
            Left            =   -300
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblPageTitle 
            AutoSize        =   -1  'True
            Caption         =   "Select Cover Image"
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
            TabIndex        =   14
            Top             =   180
            Width           =   1650
         End
      End
      Begin VB.Frame fraProperties 
         Caption         =   " Cover Image Properties "
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
         Height          =   1440
         Left            =   210
         TabIndex        =   9
         Top             =   1710
         Width           =   4080
         Begin VB.Frame fraView 
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
            Height          =   1440
            Left            =   3210
            TabIndex        =   12
            Top             =   0
            Width           =   870
            Begin Chameleon.CustomButton btnView 
               CausesValidation=   0   'False
               Height          =   690
               Left            =   90
               TabIndex        =   2
               ToolTipText     =   " Preview  (shortcut: Ctrl+P) "
               Top             =   450
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   1217
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCoverImage.frx":059A
            End
         End
         Begin VB.Label lblFileDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Date:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   690
            UseMnemonic     =   0   'False
            Width           =   690
         End
         Begin VB.Label lblFileDate 
            BackStyle       =   0  'Transparent
            Height          =   600
            Index           =   1
            Left            =   1005
            TabIndex        =   15
            Top             =   690
            Width           =   2010
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblFileSize 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   1005
            TabIndex        =   11
            Top             =   390
            Width           =   2010
         End
         Begin VB.Label lblFileSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Size:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   390
            Width           =   630
         End
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   210
         MaxLength       =   255
         TabIndex        =   0
         Top             =   1230
         Width           =   5820
      End
      Begin Chameleon.CustomButton btnBrowse 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6075
         TabIndex        =   1
         ToolTipText     =   " Open  (shortcut: Ctrl+O) "
         Top             =   1230
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Padding         =   0
         Picture         =   "frmCoverImage.frx":0E74
      End
      Begin VB.Image imgPageIcon 
         Height          =   480
         Index           =   0
         Left            =   150
         Picture         =   "frmCoverImage.frx":140E
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&File Path:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   960
         Width           =   675
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
      TabIndex        =   19
      Top             =   3255
      Width           =   6660
      Begin Chameleon.CustomButton btnCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   6
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
         Picture         =   "frmCoverImage.frx":1CD8
      End
      Begin Chameleon.CustomButton btnNext 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1455
         TabIndex        =   5
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
         Picture         =   "frmCoverImage.frx":1FFA
      End
      Begin Chameleon.CustomButton btnBack 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   210
         TabIndex        =   4
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
         Picture         =   "frmCoverImage.frx":231C
         PictureOffset   =   4
      End
   End
End
Attribute VB_Name = "frmCoverImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Cover Image Selection Form                                                  '
'  [frmCoverImage]                                                             '
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
  lblFileSize(1).Caption = vbNullString
  lblFileDate(1).Caption = vbNullString
  txtPath.Text = vbNullString
End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub btnBack_Click()
  frmDataFile.Show
  frmDataFile.SetFocus
End Sub


Private Sub btnBrowse_Click()
  
  If Len(txtPath.Text) > 0 Then
    dlgBrowse.InitDir = txtPath.Text
    If FileExists(txtPath.Text) Then dlgBrowse.FileName = txtPath.Text
  End If

  On Error Resume Next
  dlgBrowse.ShowOpen

  If Err.Number = 0 Then
    txtPath.Text = dlgBrowse.FileName
    Call txtPath_LostFocus
  End If

  txtPath.SetFocus

End Sub


Private Sub btnCancel_Click()
  frmMainMenu.Show
  frmMainMenu.SetFocus
End Sub


Private Sub btnHelp_Click()
  DisplayHelpFile "hide_wizard.html#Page2"
End Sub


Private Sub btnNext_Click()

  On Error Resume Next
  Me.ValidateControls

  If ActiveControl Is btnNext Then
    frmPassword1.Show
    frmPassword1.SetFocus
  End If

End Sub


Private Sub btnView_Click()

  On Error Resume Next
  Me.ValidateControls

  If ActiveControl Is btnView Then
    If Not frmView.DisplayByFilename(, txtPath.Text) Then
      MsgBox "The specified file cannot be opened as an image." & vbCrLf & _
             "Please select a valid image file.", vbExclamation
    End If
  End If

End Sub


Private Sub Form_Activate()
  On Error Resume Next
  txtPath.SetFocus
End Sub


Private Sub Form_Deactivate()
  If Screen.ActiveForm.MDIChild Then Me.Hide
End Sub


Private Sub Form_Load()

  Set imgPageIcon(1).Picture = imgPageIcon(0).Picture
  SetMargins txtPath, 1

  dlgBrowse.Filter = "(All supported image formats)|" & _
                       "*.bmp;*.png;*.tif;*.tiff;*.tga;*.ppm;" & _
                       "*.jpg;*.jpeg;*.pcx;*.psd;*.ras|" & _
                     "Adobe Photoshop (*.psd)|*.psd|" & _
                     "JPEG File Interchange Format (*.jpg)" & _
                       "|*.jpg;*.jpeg|" & _
                     "PC Paintbrush (*.pcx)|*.pcx|" & _
                     "Portable Network Graphics (*.png)|*.png|" & _
                     "Portable Pixelmap (*.ppm)|*.ppm|" & _
                     "Tagged Image File Format (*.tif)|*.tif;*.tiff|" & _
                     "TARGA Bitmap (*.tga)|*.tga|" & _
                     "Sun Rasterfile (*.ras)|*.ras|" & _
                     "Windows Bitmap (*.bmp)|*.bmp|"
  dlgBrowse.FilterIndex = 1
  dlgBrowse.Flags = cdlOFNHideReadOnly + cdlOFNLongNames + _
                    cdlOFNPathMustExist + cdlOFNFileMustExist
  dlgBrowse.InitDir = GetPrimaryDrive

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If Shift = vbCtrlMask Then

    Select Case KeyCode
      Case vbKeyO:  btnBrowse.Press
      Case vbKeyP:  btnView.Press
    End Select

  ElseIf KeyCode = vbKeyF1 Then

    btnHelp.Press

  End If

End Sub


Private Sub txtPath_Change()

  If Screen.ActiveControl Is txtPath Then
    lblFileSize(1).Caption = vbNullString
    lblFileDate(1).Caption = vbNullString
  End If

End Sub


Private Sub txtPath_GotFocus()

  With txtPath
    If .Tag <> "MouseDown" Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

End Sub


Private Sub txtPath_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then btnNext.SetFocus
End Sub


Private Sub txtPath_LostFocus()
  txtPath.Text = GetAbsolutePath(Trim$(txtPath.Text))
  lblFileSize(1).Caption = GetFileSize(txtPath.Text)
  lblFileDate(1).Caption = GetFileDateTime(txtPath.Text)
End Sub


Private Sub txtPath_MouseDown(Button As Integer, Shift As Integer, _
                              X As Single, Y As Single)
  txtPath.Tag = "MouseDown"
End Sub


Private Sub txtPath_MouseUp(Button As Integer, Shift As Integer, _
                            X As Single, Y As Single)
  txtPath.Tag = ""
End Sub


Private Sub txtPath_Validate(Cancel As Boolean)

  If Len(txtPath.Text) = 0 Then
    MsgBox "No file has been specified." & vbCrLf & _
           "Please specify an existing file.", vbExclamation
    Cancel = True
  ElseIf Not FileExists(txtPath.Text) Then
    MsgBox "The specified file cannot be found." & vbCrLf & _
           "Please specify an existing file.", vbExclamation
    Cancel = True
  End If

End Sub
