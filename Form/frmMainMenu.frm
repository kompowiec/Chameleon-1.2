VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E88121A0-9FA9-11CF-9D9F-00AA003A3AA3}#1.0#0"; "ZLIBTOOL.OCX"
Object = "{DCDA41A2-02F6-4FB7-82BE-A39D10159364}#1.0#0"; "EZCRYPTOAPI.OCX"
Begin VB.Form frmMainMenu 
   AutoRedraw      =   -1  'True
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
   Icon            =   "frmMainMenu.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6660
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraWipe 
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
      Height          =   960
      Left            =   3345
      TabIndex        =   13
      Top             =   1275
      Width           =   3315
      Begin Chameleon.CustomButton btnWipe 
         Height          =   690
         Left            =   210
         TabIndex        =   2
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1217
         AccessKeys      =   "w"
         Alignment       =   0
         Caption         =   "&Wipe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Padding         =   12
         Picture         =   "frmMainMenu.frx":0CCA
      End
      Begin VB.Label lblWipe 
         BackStyle       =   0  'Transparent
         Caption         =   "Destroy a file and render it unrecoverable."
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   1860
         TabIndex        =   14
         Top             =   180
         Width           =   1230
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraHelp 
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
      Height          =   960
      Left            =   3345
      TabIndex        =   11
      Top             =   2265
      Width           =   3315
      Begin Chameleon.CustomButton btnHelp 
         Height          =   690
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   " (shortcut: F1) "
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1217
         Alignment       =   0
         Caption         =   "Help"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Padding         =   12
         Picture         =   "frmMainMenu.frx":15A4
      End
      Begin VB.Label lblHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "Learn about Chameleon and steganography."
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   1860
         TabIndex        =   12
         Top             =   180
         Width           =   1230
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   960
      Left            =   0
      TabIndex        =   8
      Top             =   2265
      Width           =   3315
      Begin Chameleon.CustomButton btnExtract 
         Height          =   690
         Left            =   210
         TabIndex        =   1
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1217
         AccessKeys      =   "e"
         Alignment       =   0
         Caption         =   "&Extract"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Padding         =   12
         Picture         =   "frmMainMenu.frx":1E7E
      End
      Begin VB.Label lblExtract 
         BackStyle       =   0  'Transparent
         Caption         =   "Decode a data file hidden inside an image."
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   1860
         TabIndex        =   10
         Top             =   180
         Width           =   1230
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraHide 
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
      Height          =   960
      Left            =   0
      TabIndex        =   7
      Top             =   1275
      Width           =   3315
      Begin Chameleon.CustomButton btnHide 
         Height          =   690
         Left            =   210
         TabIndex        =   0
         Top             =   135
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1217
         AccessKeys      =   "h"
         Alignment       =   0
         Caption         =   "&Hide"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Padding         =   12
         Picture         =   "frmMainMenu.frx":2758
      End
      Begin VB.Label lblHide 
         BackStyle       =   0  'Transparent
         Caption         =   "Encode a data file inside an image."
         ForeColor       =   &H8000000D&
         Height          =   600
         Left            =   1860
         TabIndex        =   9
         Top             =   180
         Width           =   1230
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      Height          =   1260
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6660
      Begin MSComDlg.CommonDialog dlgBrowse 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Wipe File"
      End
      Begin VB.Image imgAuthor 
         Height          =   1200
         Left            =   3300
         Picture         =   "frmMainMenu.frx":3032
         Top             =   0
         Visible         =   0   'False
         Width           =   6600
      End
      Begin VB.Image imgTitle 
         Height          =   1200
         Left            =   0
         Picture         =   "frmMainMenu.frx":3CF9
         Top             =   0
         Visible         =   0   'False
         Width           =   6600
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
      TabIndex        =   6
      Top             =   3255
      Width           =   6660
      Begin Chameleon.CustomButton btnExit 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   5235
         TabIndex        =   4
         Top             =   75
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         AccessKeys      =   "x"
         Caption         =   "E&xit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMainMenu.frx":5393
      End
   End
   Begin ZLIBTOOLLib.ZlibTool ZlibCompressor 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   0
   End
   Begin CryptoApi.EzCryptoApi EzCrypto 
      Left            =   0
      Top             =   0
      _ExtentX        =   1640
      _ExtentY        =   1905
      Password        =   ""
      EncryptionAlgorithm=   1
   End
   Begin ZLIBTOOLLib.ZlibTool ZlibDecompressor 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1050
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Main Menu Form                                                              '
'  [frmMainMenu]                                                               '
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


Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long


'------------------------------------------------------------------------------'
'  Public Variables                                                            '
'------------------------------------------------------------------------------'


'[  stegosystem objects  ]'
Public WithEvents StegoEncoder As StegosystemEncoder
Attribute StegoEncoder.VB_VarHelpID = -1
Public WithEvents StegoDecoder As StegosystemDecoder
Attribute StegoDecoder.VB_VarHelpID = -1


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub btnExit_Click()
  Unload frmMDI
End Sub


Private Sub btnExtract_Click()
  CurrentTask = TASK_DECODE
  frmStegoImage.Show
  frmStegoImage.SetFocus
End Sub


Private Sub btnHelp_Click()
  DisplayHelpFile "index.html"
End Sub


Private Sub btnHide_Click()
  CurrentTask = TASK_ENCODE
  frmDataFile.Show
  frmDataFile.SetFocus
End Sub


Private Sub btnWipe_Click()

  Dim Msg As String

  On Error Resume Next
  dlgBrowse.ShowOpen

  Msg = "A file can no longer be recovered after being wiped." & vbCrLf & _
        "Are you sure you wish to permanently remove the following file?" & _
        vbCrLf & vbCrLf & _
        dlgBrowse.FileName

  If Err.Number = 0 Then
    If MsgBox(Msg, vbQuestion + vbYesNo) = vbYes Then
      If WipeFile(dlgBrowse.FileName, True) Then
        MsgBox "The selected file has been successfully wiped."
      Else
        MsgBox "The selected file cannot be wiped."
      End If
    End If
  End If

End Sub


Private Sub EzCrypto_DecryptionFileStatus(ByVal lBytesProcessed As Long, _
                                          ByVal lTotalBytes As Long)
  With frmDecode
    .prgStatus.Value = (lBytesProcessed * 100) \ lTotalBytes
    .lblStatus(1).Caption = .prgStatus.Value & "% complete"
  End With
End Sub


Private Sub EzCrypto_EncryptionFileStatus(ByVal lBytesProcessed As Long, _
                                          ByVal lTotalBytes As Long)
  With frmEncode
    .prgStatus.Value = (lBytesProcessed * 100) \ lTotalBytes
    .lblStatus(1).Caption = .prgStatus.Value & "% complete"
  End With
End Sub


Private Sub Form_Activate()
  CurrentTask = TASK_NONE
  frmMDI.ResetChildFormControls
End Sub


Private Sub Form_Deactivate()
  If Screen.ActiveForm.MDIChild Then Me.Hide
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then btnHelp.Press
End Sub


Private Sub Form_Load()

  Set picHeader.Picture = imgTitle.Picture

  dlgBrowse.InitDir = GetPrimaryDrive
  dlgBrowse.Filter = "(All files)|*.*|"
  dlgBrowse.Flags = cdlOFNHideReadOnly + cdlOFNLongNames + _
                    cdlOFNPathMustExist + cdlOFNFileMustExist

  Set StegoEncoder = New StegosystemEncoder
  Set StegoDecoder = New StegosystemDecoder

  Me.Show

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set StegoEncoder = Nothing
  Set StegoDecoder = Nothing
End Sub


Private Sub picHeader_LostFocus()
  If picHeader.Picture <> imgTitle.Picture Then
    Set picHeader.Picture = imgTitle.Picture
  End If
End Sub


Private Sub picHeader_MouseDown(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
  Set picHeader.Picture = imgAuthor.Picture
End Sub


Private Sub picHeader_MouseUp(Button As Integer, Shift As Integer, _
                              X As Single, Y As Single)
  If picHeader.Picture <> imgTitle.Picture Then
    Set picHeader.Picture = imgTitle.Picture
  End If
End Sub


Private Sub StegoDecoder_DataFileProgress(ByVal Processed As Long, _
                                          ByVal Total As Long)
  With frmDecode
    .prgStatus.Value = (Processed * 100) \ Total
    .lblStatus(1).Caption = .prgStatus.Value & "% complete"
  End With
End Sub


Private Sub StegoDecoder_MetadataProgress(ByVal Processed As Long, _
                                          ByVal Total As Long)
  With frmPassword2
    .prgStatus.Value = (Processed * 100) \ Total
    .lblStatus(1).Caption = .prgStatus.Value & "% complete"
  End With
End Sub


Private Sub StegoEncoder_Progress(ByVal Processed As Long, ByVal Total As Long)
  With frmEncode
    .prgStatus.Value = (Processed * 100) \ Total
    .lblStatus(1).Caption = .prgStatus.Value & "% complete"
  End With
End Sub


Private Sub ZlibCompressor_Progress(ByVal percent_complete As Integer)
  With frmEncode
    .prgStatus.Value = percent_complete
    .lblStatus(1).Caption = percent_complete & "% complete"
  End With
End Sub


Private Sub ZlibDecompressor_Progress(ByVal percent_complete As Integer)
  With frmDecode
    .prgStatus.Value = percent_complete
    .lblStatus(1).Caption = percent_complete & "% complete"
  End With
End Sub
