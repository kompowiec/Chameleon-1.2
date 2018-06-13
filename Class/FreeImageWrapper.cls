VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeImageWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------'
'                                                                              '
'  Chameleon Image Steganography v1.2                                          '
'                                                                              '
'  FreeImage Library Wrapper Class                                             '
'  [FreeImageWrapper]                                                          '
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
'  Requires:  FreeImage.dll                                                    '
'                                                                              '
'  The FreeImage Library (FreeImage.dll) is written by Floris van den Berg.    '
'                                                                              '
'------------------------------------------------------------------------------'


Option Explicit


'------------------------------------------------------------------------------'
'  Windows API Structure Data Types                                            '
'------------------------------------------------------------------------------'


'[  rgb color components ]'
Private Type RGBQUAD
  rgbBlue     As Byte
  rgbGreen    As Byte
  rgbRed      As Byte
  rgbReserved As Byte
End Type

'[  dib dimensions and color format  ]'
Private Type BITMAPINFOHEADER
  biSize          As Long
  biWidth         As Long
  biHeight        As Long
  biPlanes        As Integer
  biBitCount      As Integer
  biCompression   As Long
  biSizeImage     As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed       As Long
  biClrImportant  As Long
End Type

'[  dib information  ]'
Private Type BITMAPINFO
  bmiHeader   As BITMAPINFOHEADER
  bmiColors() As RGBQUAD
End Type

'[  bitmap structure  ]'
Private Type BITMAP
  bmType       As Long
  bmWidth      As Long
  bmHeight     As Long
  bmWidthBytes As Long
  bmPlanes     As Integer
  bmBitsPixel  As Integer
  bmBits       As Long
End Type


'------------------------------------------------------------------------------'
'  Windows API Constants                                                       '
'------------------------------------------------------------------------------'


'[  constant for "CreateDIBitmap" function "BitFlags" parameter  ]'
Private Const CBM_INIT As Long = &H4

'[  color representation  ]'
Private Const DIB_RGB_COLORS As Long = 0  '[  colors as rgb components  ]'
Private Const DIB_PAL_COLORS As Long = 1  '[  colors as palette indexes  ]'


'------------------------------------------------------------------------------'
'  FreeImage Constants                                                         '
'------------------------------------------------------------------------------'


'[  default load/save flags  ]'
Private Const BMP_DEFAULT   As Long = 0
Private Const CUT_DEFAULT   As Long = 0
Private Const ICO_DEFAULT   As Long = 0
Private Const IFF_DEFAULT   As Long = 0
Private Const JPEG_DEFAULT  As Long = 0
Private Const KOALA_DEFAULT As Long = 0
Private Const LBM_DEFAULT   As Long = 0
Private Const MNG_DEFAULT   As Long = 0
Private Const PCD_DEFAULT   As Long = 0
Private Const PCX_DEFAULT   As Long = 0
Private Const PNG_DEFAULT   As Long = 0
Private Const PNM_DEFAULT   As Long = 0
Private Const PSD_DEFAULT   As Long = 0
Private Const RAS_DEFAULT   As Long = 0
Private Const TARGA_DEFAULT As Long = 0
Private Const TIFF_DEFAULT  As Long = 0
Private Const WBMP_DEFAULT  As Long = 0

'[  special load flags  ]'
Private Const ICO_FIRST         As Long = 0
Private Const ICO_SECOND        As Long = 0
Private Const ICO_THIRD         As Long = 0
Private Const JPEG_FAST         As Long = 1
Private Const JPEG_ACCURATE     As Long = 2
Private Const PCD_BASE          As Long = 1  '[  PhotoCD, size 768 x 512  ]'
Private Const PCD_BASEDIV4      As Long = 2  '[  PhotoCD, size 384 x 256  ]'
Private Const PCD_BASEDIV16     As Long = 3  '[  PhotoCD, size 192 x 128  ]'
Private Const PNG_IGNOREGAMMA   As Long = 1
Private Const TARGA_LOAD_RGB888 As Long = 1

'[  special save flags  ]'
Private Const JPEG_QUALITYSUPERB  As Long = &H80
Private Const JPEG_QUALITYGOOD    As Long = &H100
Private Const JPEG_QUALITYNORMAL  As Long = &H200
Private Const JPEG_QUALITYAVERAGE As Long = &H400
Private Const JPEG_QUALITYBAD     As Long = &H800
Private Const PNM_SAVE_RAW        As Long = 0
Private Const PNM_SAVE_ascii      As Long = 1


'------------------------------------------------------------------------------'
'  FreeImage Enumerated Data Types                                             '
'------------------------------------------------------------------------------'


'[  image file format  ]'
Public Enum FREE_IMAGE_FORMAT
  FIF_UNKNOWN = -1  '[  unidentified bitmap type             ]'
  FIF_BMP = 0       '[  windows or OS/2 bitmap file (*.bmp)  ]'
  FIF_ICO = 1       '[  windows icon (*.ico)                 ]'
  FIF_JPEG = 2      '[  independent jpeg froup (*.jpg)       ]'
  FIF_JNG = 3       '[  jpeg network graphics (*.jng)        ]'
  FIF_KOALA = 4     '[  commodore 64 koala format (*.koa)    ]'
  FIF_IFF = 5       '[  amiga iff (*.iff, *.lbm)             ]'
  FIF_MNG = 6       '[  multiple network graphics (*.mng)    ]'
  FIF_PBM = 7       '[  portable bitmap (ascii) (*.pbm)      ]'
  FIF_PBMRAW = 8    '[  portable bitmap (binary) (*.pbm)     ]'
  FIF_PCD = 9       '[  kodak photocd (*.pcd)                ]'
  FIF_PCX = 10      '[  pcx bitmap format (*.pcx)            ]'
  FIF_PGM = 11      '[  portable graymap (ascii) (*.pgm)     ]'
  FIF_PGMRAW = 12   '[  portable graymap (binary) (*.pgm)    ]'
  FIF_PNG = 13      '[  portable network Graphics (*.PNG)    ]'
  FIF_PPM = 14      '[  portable pixelmap (ascii) (*.PPM)    ]'
  FIF_PPMRAW = 15   '[  portable pixelmap (binary) (*.PPM)   ]'
  FIF_RAS = 16      '[  sun rasterfile (*.ras)               ]'
  FIF_TARGA = 17    '[  targa files (*.tga)                  ]'
  FIF_TIFF = 18     '[  tagged image file format (*.tiff)    ]'
  FIF_WBMP = 19     '[  wireless bitmap (*.wbmp)             ]'
  FIF_PSD = 20      '[  photoshop (*.psd)                    ]'
  FIF_CUT = 21      '[  dr. halo (*.cut)                     ]'
End Enum


'------------------------------------------------------------------------------'
'  Windows API Function Declarations                                           '
'------------------------------------------------------------------------------'


Private Declare Function CreateDIBitmap _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByRef lpBITMAPINFOHEADER As Any, _
          ByVal BitFlags As Long, _
          ByRef BmpBits As Any, _
          ByRef lpBitmapInfo As Any, _
          ByVal fuUsage As Long _
        ) As Long

Private Declare Function DeleteObject _
        Lib "gdi32" (ByVal hObject As Long) As Boolean

Private Declare Function GetBitmapBits _
        Lib "gdi32" ( _
          ByVal hBitmap As Long, _
          ByVal dwCount As Long, _
          BmpBits As Any _
        ) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetDIBits _
        Lib "gdi32" ( _
          ByVal hDC As Long, _
          ByVal hBitmap As Long, _
          ByVal StartingScanLine As Long, _
          ByVal ScanLineCount As Long, _
          ByRef lpBits As Any, _
          ByRef lpBitmapInfo As Any, _
          ByVal ColorUsage As Long _
        ) As Long

Private Declare Function GetObject _
        Lib "gdi32" Alias "GetObjectA" ( _
          ByVal hObject As Long, _
          ByVal nCount As Long, _
          ByRef lpObject As Any _
        ) As Long

Private Declare Function ReleaseDC _
        Lib "user32" ( _
          ByVal hWnd As Long, _
          ByVal hDC As Long _
        ) As Long

Private Declare Function SetBitmapBits _
        Lib "gdi32" ( _
          ByVal hBitmap As Long, _
          ByVal dwCount As Long, _
          ByRef BmpBits As Any _
        ) As Long

Private Declare Function SetDIBitsToDevice _
        Lib "gdi32" ( _
          ByVal Dest_hDC As Long, _
          ByVal Dest_X As Long, _
          ByVal Dest_Y As Long, _
          ByVal Src_Width As Long, _
          ByVal Src_Height As Long, _
          ByVal Src_X As Long, _
          ByVal Src_Y As Long, _
          ByVal StartingScanLine As Long, _
          ByVal ScanLineCount As Long, _
          ByRef BmpBits As Any, ByRef BmpInfo As Any, _
          ByVal ColorUse As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  FreeImage Library General Function Declarations                             '
'------------------------------------------------------------------------------'


Private Declare Sub FreeImage_DeInitialise _
        Lib "FreeImage" Alias "_FreeImage_DeInitialise@0" ()

Private Declare Sub FreeImage_Initialise _
        Lib "FreeImage" Alias "_FreeImage_Initialise@4" ( _
          Optional ByVal load_local_plugins_only As Boolean = False)


'------------------------------------------------------------------------------'
'  FreeImage Library Bitmap Management Function Declarations                   '
'------------------------------------------------------------------------------'


Private Declare Function FreeImage_Allocate _
        Lib "FreeImage" Alias "_FreeImage_Allocate@24" ( _
          ByVal Width As Long, _
          ByVal Height As Long, _
          ByVal BitsPerPixel As Long, _
          Optional ByVal Red_Mask As Long = 0, _
          Optional ByVal Green_Mask As Long = 0, _
          Optional ByVal Blue_Mask As Long = 0 _
        ) As Long

Private Declare Function FreeImage_Load _
        Lib "FreeImage" Alias "_FreeImage_Load@12" ( _
          ByVal FIF As FREE_IMAGE_FORMAT, _
          ByVal FileName As String, _
          Optional ByVal Flags As Long = 0 _
        ) As Long

Private Declare Function FreeImage_Save _
        Lib "FreeImage" Alias "_FreeImage_Save@16" ( _
          ByVal FIF As FREE_IMAGE_FORMAT, _
          ByVal hDIB As Long, _
          ByVal FileName As String, _
          Optional ByVal Flags As Long = 0 _
        ) As Boolean

Private Declare Sub FreeImage_Unload _
        Lib "FreeImage" Alias "_FreeImage_Unload@4" (ByVal hDIB As Long)


'------------------------------------------------------------------------------'
'  FreeImage Library Bitmap Information Function Declarations                  '
'------------------------------------------------------------------------------'


Private Declare Function FreeImage_GetBits _
        Lib "FreeImage" Alias "_FreeImage_GetBits@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetBPP _
        Lib "FreeImage" Alias "_FreeImage_GetBPP@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetHeight _
        Lib "FreeImage" Alias "_FreeImage_GetHeight@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetInfo _
        Lib "FreeImage" Alias "_FreeImage_GetInfo@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetInfoHeader _
        Lib "FreeImage" Alias "_FreeImage_GetInfoHeader@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetPitch _
        Lib "FreeImage" Alias "_FreeImage_GetPitch@4" ( _
          ByVal hDIB As Long _
        ) As Long

Private Declare Function FreeImage_GetWidth _
        Lib "FreeImage" Alias "_FreeImage_GetWidth@4" ( _
          ByVal hDIB As Long _
        ) As Long


'------------------------------------------------------------------------------'
'  FreeImage Library Filetype Function Declarations                            '
'------------------------------------------------------------------------------'


Private Declare Function FreeImage_GetFileType _
        Lib "FreeImage" Alias "_FreeImage_GetFileType@8" ( _
          ByVal FileName As String, _
          Optional ByVal size As Long = 16 _
        ) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_GetFileTypeFromExt _
        Lib "FreeImage" Alias "_FreeImage_GetFileTypeFromExt@4" ( _
          ByVal FileName As String _
        ) As FREE_IMAGE_FORMAT

Private Declare Function FreeImage_GetFileTypeFromFormat _
        Lib "FreeImage" Alias "_FreeImage_GetFileTypeFromFormat@4" ( _
          ByVal FIF As FREE_IMAGE_FORMAT _
        ) As String


'------------------------------------------------------------------------------'
'  Public Procedures                                                           '
'------------------------------------------------------------------------------'

Public Function CreateBitmapFromDIB(ByVal hDIB As Long) As Long

  Dim BmpHead As Long
  Dim BmpInfo As Long
  Dim BmpBits As Long
  Dim hDeskDC As Long
  Dim hBmp    As Long

  '[  get desktop dc  ]'
  hDeskDC = GetDC(GetDesktopWindow)

  '[  get dib information  ]'
  BmpHead = FreeImage_GetInfoHeader(hDIB)
  BmpInfo = FreeImage_GetInfo(hDIB)
  BmpBits = FreeImage_GetBits(hDIB)

  '[  create screen-compatible bitmap based on dib  ]'
  hBmp = CreateDIBitmap(hDeskDC, ByVal BmpHead, CBM_INIT, ByVal BmpBits, _
                        ByVal BmpInfo, DIB_RGB_COLORS)

  '[  return bitmap handle  ]'
  CreateBitmapFromDIB = hBmp

  '[  cleanup desktop dc  ]'
  ReleaseDC GetDesktopWindow, hDeskDC

End Function


Public Function CreateDIBFromBitmap(ByVal hBmp As Long) As Long

  Dim BMP      As BITMAP
  Dim W        As Long
  Dim H        As Long
  Dim BPP      As Long
  Dim hDeskDC  As Long
  Dim hTmpDIB  As Long
  Dim BmpInfo  As Long
  Dim BmpBits  As Long
  Dim Res      As Long

  CreateDIBFromBitmap = 0

  '[  get bitmap information  ]'
  If GetObject(hBmp, Len(BMP), BMP) <> 0 Then

    '[  get desktop dc  ]'
    W = BMP.bmWidth
    H = BMP.bmHeight
    BPP = 24

    '[  get desktop dc  ]'
    hDeskDC = GetDC(GetDesktopWindow)

    '[  allocate new dib  ]'
    hTmpDIB = FreeImage_Allocate(W, H, BPP)

    '[  get dib information  ]'
    BmpInfo = FreeImage_GetInfo(hTmpDIB)
    BmpBits = FreeImage_GetBits(hTmpDIB)

    '[  set dib bits  ]'
    Res = GetDIBits(hDeskDC, hBmp, 0, H, ByVal BmpBits, ByVal BmpInfo, _
                   DIB_RGB_COLORS)

    If False Then
      '[  cleanup dib  ]'
      FreeImage_Unload hTmpDIB
    Else
      '[  return dib handle  ]'
      CreateDIBFromBitmap = hTmpDIB
    End If

    '[  cleanup desktop dc  ]'
    ReleaseDC GetDesktopWindow, hDeskDC

  End If

End Function


Public Function DeleteBitmap(ByVal hBmp As Long) As Boolean
  DeleteBitmap = DeleteObject(hBmp)
End Function


Public Sub GetBitmapPixels(ByVal hBmp As Long, PixelArray() As Byte)

  Dim BMP As BITMAP

  '[  get bitmap structure  ]'
  GetObject hBmp, Len(BMP), BMP

  '[  resize bitmap array  ]'
  ReDim PixelArray(0 To (BMP.bmWidthBytes \ BMP.bmWidth) - 1, _
                   0 To BMP.bmWidth - 1, _
                   0 To BMP.bmHeight - 1) As Byte

  '[  load bitmap bits to array  ]'
  GetBitmapBits hBmp, BMP.bmWidthBytes * BMP.bmHeight, _
                PixelArray(0, 0, 0)

End Sub


Public Function GetHeight(ByVal hDIB As Long) As Long
  GetHeight = FreeImage_GetHeight(hDIB)
End Function


Public Function GetWidth(ByVal hDIB As Long) As Long
  GetWidth = FreeImage_GetWidth(hDIB)
End Function


Public Function GetFileType(ByVal FileName As String) As FREE_IMAGE_FORMAT

  GetFileType = FreeImage_GetFileType(FileName)

  If GetFileType = FIF_UNKNOWN Then
    GetFileType = FreeImage_GetFileTypeFromExt(FileName)
  End If

End Function


Public Function GetFileTypeString(ByVal FIF As FREE_IMAGE_FORMAT) As String
  GetFileTypeString = FreeImage_GetFileTypeFromFormat$(FIF)
End Function


Public Function LoadDIB(ByVal FileName As String, _
                        Optional ByRef FIF As Long) As Long

  Dim Flg As Long

  '[  get format  ]'
  FIF = GetFileType(FileName)

  '[  set load flags based on format  ]'
  Select Case FIF
    Case FIF_JPEG:  Flg = JPEG_ACCURATE
    Case FIF_PCD:   Flg = PCD_BASE
    Case FIF_PNG:   Flg = PNG_IGNOREGAMMA
    Case Else:      Flg = 0
  End Select

  '[  if format recognized, then load image  ]'
  If FIF = FIF_UNKNOWN Then
    LoadDIB = 0
  Else
    LoadDIB = FreeImage_Load(FIF, FileName, Flg)
  End If

End Function


Public Function PaintDIB(ByVal hDIB As Long, ByVal Dest_hDC As Long, _
                         Optional ByVal Dest_X As Long = 0, _
                         Optional ByVal Dest_Y As Long = 0, _
                         Optional ByVal Src_Width As Long = -1, _
                         Optional ByVal Src_Height As Long = -1) As Boolean

  Dim W       As Long
  Dim H       As Long
  Dim BmpBits As Long
  Dim BmpInfo As Long
  Dim Res     As Long

  '[  get dimensions of dib  ]'
  W = IIf(Src_Width > 0, Src_Width, FreeImage_GetWidth(hDIB))
  H = IIf(Src_Height > 0, Src_Height, FreeImage_GetHeight(hDIB))

  '[  get memory address of dib bits  ]'
  BmpBits = FreeImage_GetBits(hDIB)

  '[  get memory address of dib bitmap information  ]'
  BmpInfo = FreeImage_GetInfo(hDIB)

  '[  paint as bitmap  ]'
  If (W > 0) And (H > 0) And (BmpBits) And (BmpInfo) Then
    Res = SetDIBitsToDevice(Dest_hDC, Dest_X, Dest_Y, W, H, 0, 0, 0, H, _
                           ByVal BmpBits, ByVal BmpInfo, DIB_RGB_COLORS)
  Else
    Res = 0
  End If

  PaintDIB = (Res <> 0)

End Function


Public Function SaveDIB(ByVal hDIB As Long, ByVal FileName As String, _
                        ByVal FIF As FREE_IMAGE_FORMAT) As Long

  Dim Flg As Long

  '[  set save flags  ]'
  Select Case FIF
    Case FIF_JPEG:  Flg = JPEG_QUALITYSUPERB
    Case FIF_PBM:   Flg = PNM_SAVE_RAW
    Case FIF_PGM:   Flg = PNM_SAVE_RAW
    Case FIF_PPM:   Flg = PNM_SAVE_RAW
    Case Else:      Flg = 0
  End Select

  SaveDIB = FreeImage_Save(FIF, hDIB, FileName, Flg)

End Function


Public Sub SetBitmapPixels(ByVal hBmp As Long, PixelArray() As Byte)

  Dim BMP As BITMAP

  '[  get bitmap structure  ]'
  GetObject hBmp, Len(BMP), BMP

  '[  resize bitmap array  ]'
  ReDim Preserve PixelArray(0 To (BMP.bmWidthBytes \ BMP.bmWidth) - 1, _
                            0 To BMP.bmWidth - 1, _
                            0 To BMP.bmHeight - 1) As Byte

  '[  set bitmap bits from array  ]'
  SetBitmapBits hBmp, BMP.bmWidthBytes * BMP.bmHeight, _
                PixelArray(0, 0, 0)

End Sub


Public Sub UnloadDIB(hDIB As Long)
  FreeImage_Unload hDIB
End Sub


'------------------------------------------------------------------------------'
'  Event Handlers                                                              '
'------------------------------------------------------------------------------'


Private Sub Class_Initialize()
  ChDir App.Path
  FreeImage_Initialise
End Sub


Private Sub Class_Terminate()
  FreeImage_DeInitialise
End Sub
