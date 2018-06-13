Attribute VB_Name = "modApplication"
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  Miscellaneous Application Procedures Module                                 '
'  [modApplication]                                                            '
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
'  Public Enumerated Data Types                                                '
'------------------------------------------------------------------------------'


Public Enum appTaskConstants
  TASK_NONE = 0
  TASK_ENCODE = 1
  TASK_DECODE = 2
End Enum


'------------------------------------------------------------------------------'
'  Windows API Structure Data Types                                            '
'------------------------------------------------------------------------------'


Public Type RECT
  Left   As Long
  top    As Long
  Right  As Long
  Bottom As Long
End Type


'------------------------------------------------------------------------------'
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constants for setting textbox margins with "SendMessageLong" function  ]'
Private Const EC_LEFTMARGIN  As Long = &H1
Private Const EC_RIGHTMARGIN As Long = &H2
Private Const EM_SETMARGINS  As Long = &HD3

'[  constants for "DrawText" function "uFormat" parameter  ]'
Private Const DT_CALCRECT      As Long = &H400
Private Const DT_END_ELLIPSIS  As Long = &H8000
Private Const DT_MODIFYSTRING  As Long = &H10000
Private Const DT_NOPREFIX      As Long = &H800
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_WORD_ELLIPSIS As Long = &H40000

'[  constants for "ShellExecute" function return values  ]'
Private Const SE_ERR_ACCESSDENIED    As Long = 5
Private Const SE_ERR_ASSOCINCOMPLETE As Long = 27
Private Const SE_ERR_DDEBUSY         As Long = 30
Private Const SE_ERR_DDEFAIL         As Long = 29
Private Const SE_ERR_DDETIMEOUT      As Long = 28
Private Const SE_ERR_DLLNOTFOUND     As Long = 32
Private Const SE_ERR_FNF             As Long = 2
Private Const SE_ERR_NOASSOC         As Long = 31
Private Const SE_ERR_OOM             As Long = 8
Private Const SE_ERR_PNF             As Long = 3
Private Const SE_ERR_SHARE           As Long = 26
Private Const ERROR_BAD_FORMAT       As Long = 11&

'[  constants for "ShellExecute" function "nShowCmd" parameter  ]'
Private Const SW_SHOW            As Long = 5
Private Const SW_SHOWDEFAULT     As Long = 10
Private Const SW_SHOWMAXIMIZED   As Long = 3
Private Const SW_SHOWMINIMIZED   As Long = 2
Private Const SW_SHOWMINNOACTIVE As Long = 7
Private Const SW_SHOWNA          As Long = 8
Private Const SW_SHOWNOACTIVATE  As Long = 4
Private Const SW_SHOWNORMAL      As Long = 1

'------------------------------------------------------------------------------'
'  HTML Help Constants                                                         '
'------------------------------------------------------------------------------'


Public Const HH_DISPLAY_TOPIC       As Long = &H0
Public Const HH_SET_WIN_TYPE        As Long = &H4
Public Const HH_GET_WIN_TYPE        As Long = &H5
Public Const HH_GET_WIN_HANDLE      As Long = &H6
Public Const HH_DISPLAY_TEXT_POPUP  As Long = &HE
Public Const HH_HELP_CONTEXT        As Long = &HF
Public Const HH_TP_HELP_CONTEXTMENU As Long = &H10
Public Const HH_TP_HELP_WM_HELP     As Long = &H11
Public Const HH_CLOSE_ALL           As Long = &H12


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function DestroyWindow _
        Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function DrawText _
        Lib "user32" Alias "DrawTextA" ( _
          ByVal hDC As Long, _
          ByVal lpString As String, _
          ByVal nCount As Long, _
          lpRect As RECT, _
          ByVal uFormat As Long _
        ) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function SendMessageLong _
        Lib "user32" Alias "SendMessageA" ( _
          ByVal hWnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          ByVal lParam As Long _
        ) As Long

Private Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
          ByVal hWnd As Long, _
          ByVal lpOperation As String, _
          ByVal lpFile As String, _
          ByVal lpParameters As String, _
          ByVal lpDirectory As String, _
          ByVal nShowCmd As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  HTML Help Function Declarations                                             '
'------------------------------------------------------------------------------'


Private Declare Function HtmlHelp _
        Lib "hhctrl.ocx" Alias "HtmlHelpA" ( _
          ByVal hwndCaller As Long, _
          ByVal pszFile As String, _
          ByVal uCommand As Long, _
          ByVal dwData As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  Public Variables                                                            '
'------------------------------------------------------------------------------'


Public CurrentTask As appTaskConstants


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'


Public Sub CloseHelpFiles()
  Call HtmlHelp(0, "", HH_CLOSE_ALL, 0)
End Sub


Public Sub CompactCaptionWithEllipses(Label1 As Control)

  Dim Cap As String
  Dim RC As RECT

  With Label1

    '[  compute boundary  ]'
    RC.Right = frmView.ScaleX(.Width, vbTwips, vbPixels)
    RC.Bottom = frmView.ScaleY(.Height, vbTwips, vbPixels)

    Cap = .Caption
    Set frmMainMenu.Font = Label1.Font

    '[  use ellipses to compact the caption  ]'
    Call DrawText(frmMainMenu.hDC, Cap, -1, RC, _
                  DT_CALCRECT + DT_MODIFYSTRING + DT_NOPREFIX + _
                  DT_PATH_ELLIPSIS)

    If .Caption = Cap Then
      .ToolTipText = vbNullString
    Else
      '[  set the tooltip to the original caption and display new caption  ]'
      .ToolTipText = .Caption
      .Caption = Cap
    End If

  End With

End Sub


Public Sub DisplayHelpFile(ByVal TopicFile As String)
  Call HtmlHelp(GetDesktopWindow, App.Path & "\Chameleon.chm::/" & TopicFile & _
                ">default", HH_DISPLAY_TOPIC, 0)
End Sub


Public Function FormatDate(ByVal DateTime As Date) As String
  FormatDate = Format$(DateTime, "mmmm d, yyyy" & vbCrLf & "dddd" & vbCrLf & _
                                 "h:Nn AMPM")
End Function


Public Function FormatSize(ByVal Bytes As Long) As String
  FormatSize = Format$(Bytes, "#,##0") & IIf(Bytes = 1, " byte", " bytes")
End Function


Public Function GetHashValue(ByVal Data As String, _
                             ByVal Algorithm As EC_HASH_ALG_ID) As String
  With frmMainMenu.EzCrypto
    .HashAlgorithm = Algorithm
    .CreateHash
    .HashDigestData Data
    GetHashValue = .GetDigestedData(EC_HF_ASCII)
    .DestroyHash
  End With
End Function


Public Function GetHashValueOfFile(ByVal FileName As String, _
                                   ByVal Algorithm As EC_HASH_ALG_ID) As String
  With frmMainMenu.EzCrypto
    .HashAlgorithm = Algorithm
    .CreateHash
    .HashDigestFile FileName
    GetHashValueOfFile = .GetDigestedData(EC_HF_ASCII)
    .DestroyHash
  End With
End Function


Public Function OpenFile(ByVal FileName As String) As Boolean

  Dim RV As Long

  On Error Resume Next
  OpenFile = (ShellExecute(0, "", FileName, "", "", SW_SHOW) > 32)

  If Not OpenFile Then

    RV = ShellExecute(0, "open", FileName, "", "", SW_SHOW)

    If RV > 32 Then

      OpenFile = True

    Else

      OpenFile = False

      Select Case RV
        Case SE_ERR_NOASSOC, SE_ERR_ASSOCINCOMPLETE:
          MsgBox "The file " & FileName & " cannot be opened." & vbCrLf & _
                 "The file may not be associated with any application.", _
                 vbExclamation
        Case Else:
          MsgBox "The file " & FileName & " cannot be opened." & vbCrLf & _
                 "An error occured while accessing the file.", vbExclamation
      End Select

    End If

  End If

End Function


Public Sub SetMargins(ByRef Text1 As TextBox, ByVal Margin As Long)
  SendMessageLong Text1.hWnd, EM_SETMARGINS, EC_LEFTMARGIN, Margin
  SendMessageLong Text1.hWnd, EM_SETMARGINS, EC_RIGHTMARGIN, Margin * &H10000
End Sub
