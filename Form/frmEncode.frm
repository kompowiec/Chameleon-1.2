VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEncode 
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
      TabIndex        =   10
      Top             =   -60
      Width           =   6660
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
         TabIndex        =   28
         Top             =   1710
         Width           =   2100
         Begin Chameleon.CustomButton btnHelp 
            CausesValidation=   0   'False
            Height          =   390
            Left            =   180
            TabIndex        =   2
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
            Picture         =   "frmEncode.frx":0000
            PictureOffset   =   8
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "Press ""Cancel"" to abort the encoding process."
            ForeColor       =   &H8000000D&
            Height          =   600
            Index           =   1
            Left            =   180
            TabIndex        =   30
            Top             =   270
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblHelp 
            BackStyle       =   0  'Transparent
            Caption         =   "Press ""Next"" to begin encoding the data inside a stego image."
            ForeColor       =   &H8000000D&
            Height          =   600
            Index           =   0
            Left            =   180
            TabIndex        =   29
            Top             =   270
            UseMnemonic     =   0   'False
            Width           =   1725
            WordWrap        =   -1  'True
         End
      End
      Begin MSComDlg.CommonDialog dlgBrowse 
         Left            =   5700
         Top             =   450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Save Stego Image"
      End
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
         Top             =   1230
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
            Left            =   -240
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lblPageTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Encode Stego Image"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   300
            TabIndex        =   9
            Top             =   180
            UseMnemonic     =   0   'False
            Width           =   1740
         End
      End
      Begin VB.Frame fraProperties 
         Caption         =   " Input Files "
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
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   1710
         Width           =   4080
         Begin VB.Label lblCoverImage 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   1290
            TabIndex        =   18
            Top             =   690
            UseMnemonic     =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblCoverImage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cover Image:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   690
            Width           =   990
         End
         Begin VB.Label lblDataFile 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   1290
            TabIndex        =   16
            Top             =   390
            UseMnemonic     =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblDataFile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data File:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   15
            Top             =   390
            Width           =   690
         End
      End
      Begin VB.Frame fraProperties 
         Caption         =   " Encoding Results "
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
         Index           =   1
         Left            =   210
         TabIndex        =   19
         Top             =   1710
         Visible         =   0   'False
         Width           =   6225
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
            Left            =   5355
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   870
            Begin Chameleon.CustomButton btnView 
               Height          =   690
               Left            =   90
               TabIndex        =   5
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
               Picture         =   "frmEncode.frx":059A
            End
         End
         Begin VB.Frame fraSave 
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
            Left            =   4515
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   870
            Begin Chameleon.CustomButton btnSave 
               Height          =   690
               Left            =   90
               TabIndex        =   4
               ToolTipText     =   " Save  (shortcut: Ctrl+S) "
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
               Picture         =   "frmEncode.frx":0E74
            End
         End
         Begin VB.Label lblCapacity 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   23
            Top             =   690
            Width           =   2340
         End
         Begin VB.Label lblCapacity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Image Hiding Capacity:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   22
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label lblDataSize 
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   1980
            TabIndex        =   21
            Top             =   390
            Width           =   2340
         End
         Begin VB.Label lblDataSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compressed Data Size:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   390
            Width           =   1665
         End
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   26
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   6390
         TabIndex        =   25
         Top             =   960
         Width           =   45
      End
      Begin VB.Image imgPageIcon 
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "frmEncode.frx":17C6
         Top             =   270
         Width           =   480
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
      TabIndex        =   7
      Top             =   3255
      Width           =   6660
      Begin Chameleon.CustomButton btnCancel 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   1
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
         Picture         =   "frmEncode.frx":2090
      End
      Begin Chameleon.CustomButton btnNext 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1455
         TabIndex        =   0
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
         Picture         =   "frmEncode.frx":23B2
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
         Picture         =   "frmEncode.frx":26D4
         PictureOffset   =   4
      End
      Begin Chameleon.CustomButton btnClose 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   6
         Top             =   75
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         AccessKeys      =   "c"
         Caption         =   "&Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmEncode.frx":29F6
      End
   End
End
Attribute VB_Name = "frmEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Encoding Process Form                                                       '
'  [frmEncode]                                                                 '
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


'[  user-selected input files  ]'
Private m_CoverImagePath As String
Private m_DataFilePath   As String

'[  temporary file  ]'
Private m_TemporaryFilePath As String

'[  user-selected output file  ]'
Private m_StegoImagePath As String

'[  user password  ]'
Private m_Password        As String
Private m_PasswordHashMD5 As String
Private m_PasswordHashSHA As String

'[  cancel process flag  ]'
Private m_Cancel As Boolean


'------------------------------------------------------------------------------'
'  Private Procedures                                                          '
'------------------------------------------------------------------------------'


Private Function PerformCompression() As Boolean

  lblStatus(0).Caption = "Compressing data file..."
  lblStatus(1).Caption = vbNullString
  prgStatus.Value = 0

  On Error GoTo ZlibError
  With frmMainMenu.ZlibCompressor
    .Level = Standard
    .InputFile = m_DataFilePath
    .OutputFile = m_TemporaryFilePath
    .Compress
  End With

ZlibError:
  If UCase(frmMainMenu.ZlibCompressor.Status) = "SUCCESS" Then
    PerformCompression = True
  Else
    PerformCompression = False
  End If

End Function


Private Function PerformEmbedding()

  lblStatus(0).Caption = "Hiding data file..."
  lblStatus(1).Caption = vbNullString
  prgStatus.Value = 0

  PerformEmbedding = False

  With frmMainMenu.StegoEncoder

    If .LoadDataFile(m_TemporaryFilePath, _
                     GetPathFileName(m_DataFilePath), _
                     FileDateTime(m_DataFilePath), _
                     GetHashValueOfFile(m_TemporaryFilePath, MD5)) Then

      If .LoadCoverImage(m_CoverImagePath) Then
        PerformEmbedding = .Encode(m_PasswordHashMD5, m_PasswordHashSHA)
        lblDataSize(1).Caption = FormatSize(.DataFileSize)
        lblCapacity(1).Caption = FormatSize(.ImageCapacity)
      Else
        m_Cancel = True
        MsgBox "The selected cover image cannot be loaded." & vbCrLf & _
               "Please select a valid image file.", vbInformation
        frmCoverImage.Show
        frmCoverImage.SetFocus
      End If

    Else
      m_Cancel = True
      MsgBox "The selected data file cannot be loaded." & vbCrLf & _
             "Please select a valid data file.", vbInformation
      frmDataFile.Show
      frmDataFile.SetFocus
    End If

  End With

End Function


Private Sub PerformEncryption()

  lblStatus(0).Caption = "Encrypting data file..."
  lblStatus(1).Caption = vbNullString
  prgStatus.Value = 0

  '--------------------------------------------------------------------------'
  '                                                                          '
  '  Special Note:                                                           '
  '    For some reason, the last hash function executed by EzCryptoApi       '
  '    affects the succeeding encryption/decryption operation. Because of    '
  '    this, it must be made sure that last hash function called before      '
  '    a decryption operation is the same as the last hash function called   '
  '    before its corresponding encryption operation.                        '
  '                                                                          '
  '--------------------------------------------------------------------------'
  GetHashValue m_Password, SHA

  On Error Resume Next

  With frmMainMenu.EzCrypto
    .EncryptionAlgorithm = RC4
    .Speed = [1KB]
    .Password = m_Password
    .EncryptFile m_TemporaryFilePath
  End With

End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub btnBack_Click()
  frmPassword1.Show
  frmPassword1.SetFocus
End Sub


Private Sub btnCancel_Click()

  If Not btnNext.Visible Then
    If MsgBox("Do you wish to cancel the current task?", _
              vbQuestion + vbYesNo) = vbYes Then
      m_Cancel = True
      frmMainMenu.ZlibCompressor.Abort
      frmMainMenu.StegoEncoder.Abort
      frmMainMenu.Show
      frmMainMenu.SetFocus
    End If
  Else
    frmMainMenu.Show
    frmMainMenu.SetFocus
  End If

End Sub


Private Sub btnClose_Click()

  If fraSave.Visible Then
    If Len(m_StegoImagePath) = 0 Then
      If MsgBox("Would you like to save the created stego image first?", _
                vbQuestion + vbYesNo) = vbYes Then
        btnSave.Press
      End If
    End If
  End If

  frmMainMenu.Show
  frmMainMenu.SetFocus

End Sub


Private Sub btnHelp_Click()
  DisplayHelpFile "hide_wizard.html#Page4"
End Sub


Private Sub btnNext_Click()

  Dim Completed As Boolean

  m_Cancel = False

  fraStatus.Visible = True
  btnNext.Visible = False
  btnBack.Visible = False
  lblHelp(0).Visible = False
  lblHelp(1).Visible = True

  '[  prepare timer  ]'
  Dim Tmr As RealTimer
  Set Tmr = New RealTimer

  PerformCompression
  If m_Cancel Then Exit Sub

  PerformEncryption
  If m_Cancel Then Exit Sub

  Completed = PerformEmbedding
  If m_Cancel Then Exit Sub

  '[  display elapsed time  ]'
  Tmr.Mark
  lblStatus(1).Caption = "Processing Time: " & Tmr.ElapsedTimeInMinutes

  fraProperties(1).Visible = True
  fraProperties(0).Visible = False
  fraHelp.Visible = False
  btnCancel.Visible = False
  btnClose.Visible = True
  btnClose.Refresh

  If Completed Then

    '[  indicate completion ]'
    lblStatus(0).Caption = "Encoding Complete."
    fraSave.Visible = True
    fraView.Visible = True
    btnSave.Press
    btnClose.SetFocus

  Else

    '[  indicate failure  ]'
    lblStatus(0).Caption = "Encoding Failed."

    Dim Res As VbMsgBoxResult
    Res = MsgBox("There is not enough space on the selected image." & vbCrLf & _
                 "Would you like to select another cover image?", _
                 vbQuestion + vbYesNo)

    If Res = vbYes Then
      frmCoverImage.Show
      frmCoverImage.SetFocus
    Else
      MsgBox "Encoding Canceled.", vbExclamation
      btnClose.SetFocus
    End If

  End If

End Sub


Private Sub btnSave_Click()

  On Error Resume Next
  dlgBrowse.ShowSave

  If Err.Number = 0 Then

    '[  set stego image file format  ]'
    Dim ImageFormat As FREE_IMAGE_FORMAT
    Select Case dlgBrowse.FilterIndex
      Case 1:  ImageFormat = FIF_BMP
      Case 2:  ImageFormat = FIF_PNG
      Case 3:  ImageFormat = FIF_PPMRAW
      Case 4:  ImageFormat = FIF_TIFF
      Case 5:  ImageFormat = FIF_TARGA
    End Select

    '[  save stego image  ]'
    m_StegoImagePath = dlgBrowse.FileName
    frmMainMenu.StegoEncoder.SaveStegoImage m_StegoImagePath, ImageFormat
    btnView.SetFocus

  End If

End Sub


Private Sub btnView_Click()
  With frmMainMenu.StegoEncoder
    If Not frmView.DisplayByDIB(.CoverImageDIB, .StegoImageDIB) Then
      MsgBox "The images cannot be displayed." & _
             "The encoding process may have failed.", vbExclamation
    End If
  End With
End Sub


Private Sub Form_Activate()

  m_Cancel = False
  m_DataFilePath = frmDataFile.txtPath.Text
  m_CoverImagePath = frmCoverImage.txtPath.Text
  m_TemporaryFilePath = GenerateTempFile
  m_StegoImagePath = vbNullString

  lblDataFile(1).Caption = m_DataFilePath
  lblCoverImage(1).Caption = m_CoverImagePath

  CompactCaptionWithEllipses lblDataFile(1)
  CompactCaptionWithEllipses lblCoverImage(1)

  m_Password = frmPassword1.txtPassword(0).Text
  m_PasswordHashMD5 = GetHashValue(m_Password, MD5)
  m_PasswordHashSHA = GetHashValue(m_Password, SHA)

  If Not ActiveControl Is btnNext Then btnNext.SetFocus

End Sub


Private Sub Form_Deactivate()

  If Screen.ActiveForm.MDIChild Then

    Me.Hide

    '[  reset controls  ]'
    fraStatus.Visible = False
    fraProperties(0).Visible = True
    fraProperties(1).Visible = False
    fraHelp.Visible = True
    fraSave.Visible = False
    fraView.Visible = False

    lblStatus(0).Caption = vbNullString
    lblStatus(1).Caption = vbNullString
    lblDataFile(1).Caption = vbNullString
    lblCoverImage(1).Caption = vbNullString
    lblDataSize(1).Caption = vbNullString
    lblCapacity(1).Caption = vbNullString
    lblHelp(0).Visible = True
    lblHelp(1).Visible = False

    prgStatus.Min = 0
    prgStatus.Max = 100
    prgStatus.Value = 0

    btnNext.Visible = True
    btnBack.Visible = True
    btnCancel.Visible = True
    btnClose.Visible = False

    '[  cleanup temporary files  ]'
    WipeFile m_TemporaryFilePath, False

  End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If Shift = vbCtrlMask Then

    Select Case KeyCode
      Case vbKeyS:  btnSave.Press
      Case vbKeyP:  btnView.Press
    End Select

  ElseIf KeyCode = vbKeyF1 Then

    btnHelp.Press

  End If

End Sub


Private Sub Form_Load()

  Set imgPageIcon(1).Picture = imgPageIcon(0).Picture

  dlgBrowse.Filter = "Windows Bitmap (*.bmp)|*.bmp|" & _
                     "Portable Network Graphics (*.png)|*.png|" & _
                     "Portable Pixelmap (*.ppm)|*.ppm|" & _
                     "Tagged Image File Format (*.tif)|*.tif;*.tiff|" & _
                     "TARGA Bitmap (*.tga)|*.tga|"
  dlgBrowse.FilterIndex = 1
  dlgBrowse.Flags = cdlOFNHideReadOnly + cdlOFNLongNames + _
                    cdlOFNPathMustExist + cdlOFNOverwritePrompt
  dlgBrowse.InitDir = GetPrimaryDrive

End Sub


Private Sub Form_Unload(Cancel As Integer)
  WipeFile m_TemporaryFilePath, False
End Sub
