Attribute VB_Name = "modSDKDeclares"
'***************************************************************************
' Module:     modSDKDeclares
'
' SDK/API specific declares
'
Option Explicit

Public Const DIB_RGB_COLORS = 0     ' color table in RGBs
Public Const DIB_PAL_COLORS = 1     ' color table in palette
Public Const BFT_BITMAP = &H4D42    ' bfType
Public Const PALVERSION = &H300     ' version of LOGPALETTE
Public Const BI_RGB = 0&            ' bitmap compression
Public Const BI_RLE4 = 2&           ' bitmap compression
Public Const BI_RLE8 = 1&           ' bitmap compression
Public Const BI_BITFIELDS = 3&      ' bitmap compression
Public Const GMEM_MOVEABLE = &H2    ' GlobalAlloc flag
Public Const GMEM_ZEROINIT = &H40   ' GlobalAlloc flag
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GMEM_DDESHARE = &H2000 ' GlobalAlloc flag
Public Const RASTERCAPS = 38        ' Bitblt capabilities
Public Const RC_PALETTE = &H100     ' supports a palette

Public Type BITMAPFILEHEADER
    bfType          As Integer
    bfSize          As Long
    bfReserved1     As Integer
    bfReserved2     As Integer
    bfOffBits       As Long
End Type ' BITMAPFILEHEADER

Public Type BITMAPCOREHEADER
    bcSize          As Long
    bcWidth         As Integer
    bcHeight        As Integer
    bcPlanes        As Integer
    bcBitCount      As Integer
End Type ' BITMAPCOREHEADER

Public Type BITMAPINFOHEADER
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
End Type ' BITMAPINFOHEADER

Public Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type ' RGBQUAD

Public Type RGBTRIPLE
    rgbtBlue        As Byte
    rgbtGreen       As Byte
    rgbtRed         As Byte
End Type ' RGBTRIPLE

Public Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors()     As RGBQUAD
End Type ' BITMAPINFO

Public Type PALETTEENTRY
    peRed           As Byte
    peGreen         As Byte
    peBlue          As Byte
    peFlags         As Byte
End Type ' PALETTEENTRY

Public Type LOGPALETTE
    PALVERSION      As Integer
    palNumEntries   As Integer
    palPalEntry(0 To 0) As PALETTEENTRY
End Type ' LOGPALETTE

' You will notice that some of the declares below have been modified.  Usually this is because
' the routine has optional parameters (usually a structure) but the API declare doesn't allow
' this (because structures can't be optional).  When I have changed a declare I've aliased it
' to "routinename"My so you can see the differences

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
'Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function CreateDIBSectionMy Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
'Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBitmapMy Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As Any, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As Any, ByVal wUsage As Long) As Long

'Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsMy Lib "gdi32" Alias "SetDIBits" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
'Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDIBitsMy Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function GdiFlush Lib "gdi32" () As Long

'Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function CreatePaletteMy Lib "gdi32" Alias "CreatePalette" (ByVal lpLogPalette As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal lnBytes As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function CreateFileMappingMy Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, lpName As Any) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



